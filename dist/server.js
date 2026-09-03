import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import { createHmac, timingSafeEqual } from 'node:crypto';
import {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    TurnContext
} from 'botbuilder';
import { z } from 'zod';
import { PersistentStore } from './store.js';

const store = new PersistentStore();
const app = express();

const origins = (process.env.ALLOWED_ORIGINS ?? '')
    .split(',')
    .map((x) => x.trim())
    .filter(Boolean);

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors({ origin: origins.length ? origins : false }));
app.use(express.json({ limit: '32kb' }));

const authentication = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: process.env.MICROSOFT_APP_ID,
    MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
    MicrosoftAppType:
        process.env.MICROSOFT_APP_TYPE ?? 'SingleTenant',
    MicrosoftAppTenantId:
        process.env.MICROSOFT_APP_TENANT_ID,
});

const adapter = new CloudAdapter(authentication);

adapter.onTurnError = async (context, error) => {
    console.error(error);
    await context.sendActivity(
        'The live-chat bridge encountered an error.'
    );
};

function expandIdentifiers(raw) {
    const values = raw.filter(
        (value) =>
            typeof value === 'string' &&
            value.length > 0
    );

    const keys = new Set();

    for (const value of values) {
        keys.add(value);

        let decoded = value;

        try {
            decoded = decodeURIComponent(value);
        } catch {
            // Keep original.
        }

        keys.add(decoded);

        for (const candidate of [value, decoded]) {
            const match = candidate.match(
                /(?:messageid|messageId|replyToId)=([^;&]+)/i
            );

            if (match?.[1]) {
                keys.add(match[1]);

                try {
                    keys.add(decodeURIComponent(match[1]));
                } catch {
                    // Keep captured value.
                }
            }
        }
    }

    return [...keys];
}

function replyKeys(activity) {
    const data = activity.channelData ?? {};

    return expandIdentifiers([
        activity.replyToId,
        activity.conversation?.id,
        data.replyToId,
        data.messageId,
        data.teamsMessageId,
    ]);
}

function sessionForReply(activity) {
    const keys = replyKeys(activity);

    for (const key of keys) {
        const direct = store.rootToSession.get(key);

        if (direct) {
            return direct;
        }
    }

    for (const [rootId, sessionId] of store.rootToSession) {
        if (keys.some((key) => key.includes(rootId))) {
            return sessionId;
        }
    }

    const unthreaded = [...store.sessions.values()].filter(
        (session) => !session.rootActivityId
    );

    if (unthreaded.length === 1) {
        const rootId = keys.find((key) => /^\d+$/.test(key));

        if (rootId) {
            store.setRoot(unthreaded[0], rootId);
            return unthreaded[0].id;
        }
    }

    return undefined;
}

function authorised(req, res, next) {
    const supplied = req.get('x-chat-key');

    if (
        !process.env.CHAT_API_KEY ||
        supplied !== process.env.CHAT_API_KEY
    ) {
        return res.status(401).json({
            error: 'Unauthorised',
        });
    }

    next();
}

function makeSessionToken(
    sessionId,
    expires = Date.now() + 30 * 24 * 60 * 60 * 1000
) {
    const payload = `${sessionId}.${expires}`;

    const signature = createHmac(
        'sha256',
        process.env.CHAT_API_KEY ?? ''
    )
        .update(payload)
        .digest('base64url');

    return `${expires}.${signature}`;
}

function authorisedSession(req, res, next) {
    const sessionId = String(req.params.id);

    const supplied =
        req.get('x-chat-token') ??
        (typeof req.query.token === 'string'
            ? req.query.token
            : '');

    const [expiryText, signature = ''] =
        supplied.split('.');

    const expiry = Number(expiryText);

    const expected = makeSessionToken(
        sessionId,
        expiry
    ).split('.')[1];

    const valid =
        Number.isFinite(expiry) &&
        expiry >= Date.now() &&
        signature.length === expected.length &&
        timingSafeEqual(
            Buffer.from(signature),
            Buffer.from(expected)
        );

    if (!valid) {
        return res.status(401).json({
            error: 'Chat session expired',
        });
    }

    next();
}

app.get('/health', (_req, res) => {
    res.json({
        ok: true,
        teamsChannelConfigured: Boolean(store.channel),
    });
});

app.get('/availability', (_req, res) => {
    res.json({
        online: store.online,
    });
});

app.post('/api/messages', (req, res) => {
    adapter.process(req, res, async (context) => {
        if (context.activity.type !== 'message') {
            return;
        }

        if (
            context.activity.from?.id ===
            context.activity.recipient?.id
        ) {
            return;
        }

        const text = (context.activity.text ?? '')
            .replace(/<at>.*?<\/at>/g, '')
            .trim();

        const command = text.toLowerCase();

        const mentioned = (
            context.activity.entities ?? []
        ).some((entity) => entity.type === 'mention');

        if (mentioned && command === 'setup') {
            const channel =
                TurnContext.getConversationReference(
                    context.activity
                );

            if (channel.conversation?.id) {
                channel.conversation.id =
                    channel.conversation.id.replace(
                        /;messageid=[^;]+/i,
                        ''
                    );
            }

            channel.activityId = undefined;
            store.channel = channel;

            await context.sendActivity(
                'Website live chat is connected to this channel.'
            );

            return;
        }

        if (
            mentioned &&
            (command === 'online' || command === 'offline')
        ) {
            store.setOnline(command === 'online');

            await context.sendActivity(
                command === 'online'
                    ? 'Website live chat is now online.'
                    : 'Website live chat is now offline. Visitors can still leave a message.'
            );

            return;
        }

        if (mentioned && command === 'status') {
            await context.sendActivity(
                `Website live chat is currently **${
                    store.online ? 'online' : 'offline'
                }**.`
            );

            return;
        }

        const sessionId =
            sessionForReply(context.activity);

        console.info(
            'Teams reply lookup',
            JSON.stringify({
                activityId: context.activity.id,
                conversationId:
                    context.activity.conversation?.id,
                keys: replyKeys(context.activity),
                matched: Boolean(sessionId),
            })
        );

        if (!sessionId) {
            return;
        }

        if (text) {
            store.add(sessionId, 'agent', text);
        }
    });
});

const startSchema = z.object({
    name: z.string().trim().min(1).max(80),
    email: z.string().email().max(254).optional(),
    page: z.string().max(500).default(''),
});

app.post('/chat/start', authorised, (req, res) => {
    const parsed = startSchema.safeParse(req.body);

    if (!parsed.success) {
        return res.status(400).json({
            error: 'Invalid chat details',
        });
    }

    const session = store.create(
        parsed.data.name,
        parsed.data.page,
        parsed.data.email
    );

    res.status(201).json({
        id: session.id,
        token: makeSessionToken(session.id),
        online: store.online,
        messages: session.messages,
    });
});

const messageSchema = z.object({
    text: z.string().trim().min(1).max(2000),
});

app.post(
    '/chat/:id/messages',
    authorised,
    async (req, res) => {
        const session = store.sessions.get(
            String(req.params.id)
        );

        const parsed = messageSchema.safeParse(req.body);

        if (!session) {
            return res.status(404).json({
                error: 'Chat not found',
            });
        }

        if (!parsed.success) {
            return res.status(400).json({
                error: 'Invalid message',
            });
        }

        if (!store.channel) {
            return res.status(503).json({
                error: 'Teams channel has not been configured.',
            });
        }

        const visitorMessage = store.add(
            session.id,
            'visitor',
            parsed.data.text
        );

        const target = structuredClone(store.channel);

        if (
            session.rootActivityId &&
            target.conversation?.id &&
            !/;messageid=/i.test(target.conversation.id)
        ) {
            target.conversation.id =
                `${target.conversation.id};messageid=` +
                session.rootActivityId;
        }

        await adapter.continueConversationAsync(
            process.env.MICROSOFT_APP_ID ?? '',
            target,
            async (context) => {
                const details = session.visitorEmail
                    ? `\n\n**Email:** ${session.visitorEmail}`
                    : '';

                const activity = session.rootActivityId
                    ? {
                          type: 'message',
                          text:
                              `**${session.visitorName}:** ` +
                              parsed.data.text,
                          replyToId:
                              session.rootActivityId,
                      }
                    : {
                          type: 'message',
                          text:
                              `**New ${
                                  store.online
                                      ? 'website chat'
                                      : 'offline enquiry'
                              } · ${session.id.slice(
                                  0,
                                  8
                              )}**\n\n` +
                              `**Visitor:** ${
                                  session.visitorName
                              }${details}\n\n` +
                              `**Page:** ${
                                  session.page
                              }\n\n` +
                              parsed.data.text,
                      };

                const sent =
                    await context.sendActivity(activity);

                if (
                    !session.rootActivityId &&
                    sent?.id
                ) {
                    const keys =
                        expandIdentifiers([sent.id]);

                    const rootId =
                        keys.find((key) =>
                            /^\d+$/.test(key)
                        ) ?? sent.id;

                    store.setRoot(
                        session,
                        rootId,
                        keys
                    );

                    console.info(
                        'Teams root sent',
                        JSON.stringify({
                            sentId: sent.id,
                            keys,
                        })
                    );
                }
            }
        );

        res.status(201).json(visitorMessage);
    }
);

app.get(
    '/chat/:id/history',
    authorisedSession,
    (req, res) => {
        const session = store.sessions.get(
            String(req.params.id)
        );

        if (!session) {
            return res.status(404).json({
                error: 'Chat not found',
            });
        }

        res.json({
            messages: session.messages,
            online: store.online,
        });
    }
);

app.get(
    '/chat/:id/events',
    authorisedSession,
    (req, res) => {
        const sessionId = String(req.params.id);
        const session = store.sessions.get(sessionId);

        if (!session) {
            return res.status(404).end();
        }

        res.setHeader(
            'Content-Type',
            'text/event-stream'
        );

        res.setHeader(
            'Cache-Control',
            'no-cache, no-transform'
        );

        res.setHeader('Connection', 'keep-alive');
        res.flushHeaders();

        res.write(': connected\n\n');

        session.messages.forEach((message) => {
            res.write(
                `data: ${JSON.stringify(message)}\n\n`
            );
        });

        const unsubscribe = store.subscribe(
            sessionId,
            (message) => {
                res.write(
                    `data: ${JSON.stringify(message)}\n\n`
                );
            }
        );

        const heartbeat = setInterval(() => {
            res.write(': heartbeat\n\n');
        }, 25000);

        req.on('close', () => {
            clearInterval(heartbeat);
            unsubscribe();
        });
    }
);

app.listen(
    Number(process.env.PORT ?? 3978),
    () => {
        console.log(
            `Teams live-chat bridge listening on ${
                process.env.PORT ?? 3978
            }`
        );
    }
);
