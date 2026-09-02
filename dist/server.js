import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import { CloudAdapter, ConfigurationBotFrameworkAuthentication, TurnContext } from 'botbuilder';
import { z } from 'zod';
import { MemoryStore } from './store.js';
const store = new MemoryStore();
const app = express();
const origins = (process.env.ALLOWED_ORIGINS ?? '').split(',').map((x) => x.trim()).filter(Boolean);
app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors({ origin: origins.length ? origins : false }));
app.use(express.json({ limit: '32kb' }));
const authentication = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: process.env.MICROSOFT_APP_ID,
    MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
    MicrosoftAppType: process.env.MICROSOFT_APP_TYPE ?? 'SingleTenant',
    MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});
const adapter = new CloudAdapter(authentication);
adapter.onTurnError = async (context, error) => { console.error(error); await context.sendActivity('The live-chat bridge encountered an error.'); };
function authorised(req, res, next) {
    const supplied = req.get('x-chat-key') ?? (typeof req.query.key === 'string' ? req.query.key : undefined);
    if (!process.env.CHAT_API_KEY || supplied !== process.env.CHAT_API_KEY)
        return res.status(401).json({ error: 'Unauthorised' });
    next();
}
app.get('/health', (_req, res) => res.json({ ok: true, teamsChannelConfigured: Boolean(store.channel) }));
app.post('/api/messages', (req, res) => {
    adapter.process(req, res, async (context) => {
        if (context.activity.type !== 'message')
            return;
        const text = (context.activity.text ?? '').replace(/<at>.*?<\/at>/g, '').trim();
        if (text.toLowerCase() === 'setup') {
            store.channel = TurnContext.getConversationReference(context.activity);
            await context.sendActivity('Website live chat is connected to this channel.');
            return;
        }
        const rootId = context.activity.replyToId;
        const sessionId = rootId ? store.rootToSession.get(rootId) : undefined;
        if (!sessionId)
            return;
        store.add(sessionId, 'agent', text);
    });
});
const startSchema = z.object({ name: z.string().trim().min(1).max(80), page: z.string().max(500).default('') });
app.post('/chat/start', authorised, (req, res) => {
    const parsed = startSchema.safeParse(req.body);
    if (!parsed.success)
        return res.status(400).json({ error: 'Invalid chat details' });
    res.status(201).json(store.create(parsed.data.name, parsed.data.page));
});
const messageSchema = z.object({ text: z.string().trim().min(1).max(2000) });
app.post('/chat/:id/messages', authorised, async (req, res) => {
    const sessionId = String(req.params.id);
    const session = store.sessions.get(sessionId);
    const parsed = messageSchema.safeParse(req.body);
    if (!session)
        return res.status(404).json({ error: 'Chat not found' });
    if (!parsed.success)
        return res.status(400).json({ error: 'Invalid message' });
    if (!store.channel)
        return res.status(503).json({ error: 'Teams channel has not been configured. Mention the bot and send setup in the chosen channel.' });
    const visitorMessage = store.add(session.id, 'visitor', parsed.data.text);
    await adapter.continueConversationAsync(process.env.MICROSOFT_APP_ID ?? '', store.channel, async (context) => {
        const activity = session.rootActivityId
            ? { type: 'message', text: `**${session.visitorName}:** ${parsed.data.text}`, replyToId: session.rootActivityId }
            : { type: 'message', text: `**New website chat · ${session.id.slice(0, 8)}**\n\n**Visitor:** ${session.visitorName}\n\n**Page:** ${session.page}\n\n${parsed.data.text}` };
        const sent = await context.sendActivity(activity);
        if (!session.rootActivityId && sent?.id) {
            session.rootActivityId = sent.id;
            store.rootToSession.set(sent.id, session.id);
        }
    });
    res.status(201).json(visitorMessage);
});
app.get('/chat/:id/events', authorised, (req, res) => {
    const sessionId = String(req.params.id);
    if (!store.sessions.has(sessionId))
        return res.status(404).end();
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();
    res.write(': connected\n\n');
    const unsubscribe = store.subscribe(sessionId, (message) => res.write(`data: ${JSON.stringify(message)}\n\n`));
    const heartbeat = setInterval(() => res.write(': heartbeat\n\n'), 25000);
    req.on('close', () => { clearInterval(heartbeat); unsubscribe(); });
});
app.listen(Number(process.env.PORT ?? 3978), () => console.log(`Teams live-chat bridge listening on ${process.env.PORT ?? 3978}`));
