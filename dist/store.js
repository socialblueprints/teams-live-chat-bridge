import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

export class PersistentStore {
    file = process.env.STATE_FILE ?? './data/state.json';
    _channel;
    online = true;
    sessions = new Map();
    rootToSession = new Map();
    listeners = new Map();

    constructor() {
        this.load();
    }

    get channel() {
        return this._channel;
    }

    set channel(value) {
        this._channel = value;
        this.persist();
    }

    load() {
        try {
            if (!existsSync(this.file)) {
                return;
            }

            const state = JSON.parse(
                readFileSync(this.file, 'utf8')
            );

            this._channel = state.channel;
            this.online = state.online ?? true;

            this.sessions = new Map(
                (state.sessions ?? []).map((session) => [
                    session.id,
                    session,
                ])
            );

            this.rootToSession = new Map(
                state.rootToSession ?? []
            );
        } catch (error) {
            console.error(
                'Could not load persistent chat state',
                error
            );
        }
    }

    persist() {
        const directory = dirname(this.file);
        mkdirSync(directory, { recursive: true });

        const temporary = `${this.file}.tmp`;

        const state = {
            channel: this._channel,
            online: this.online,
            sessions: [...this.sessions.values()],
            rootToSession: [...this.rootToSession.entries()],
        };

        writeFileSync(
            temporary,
            JSON.stringify(state),
            { mode: 0o600 }
        );

        renameSync(temporary, this.file);
    }

    setOnline(online) {
        this.online = online;
        this.persist();
    }

    create(visitorName, page, visitorEmail) {
        const now = new Date().toISOString();

        const session = {
            id: crypto.randomUUID(),
            visitorName,
            visitorEmail,
            page,
            createdAt: now,
            updatedAt: now,
            messages: [],
        };

        this.sessions.set(session.id, session);
        this.persist();

        return session;
    }

    setRoot(session, rootId, keys = [rootId]) {
        session.rootActivityId = rootId;

        keys.forEach((key) => {
            this.rootToSession.set(key, session.id);
        });

        session.updatedAt = new Date().toISOString();
        this.persist();
    }

    add(sessionId, sender, text) {
        const session = this.sessions.get(sessionId);

        if (!session) {
            throw new Error('Chat session not found');
        }

        const message = {
            id: crypto.randomUUID(),
            sender,
            text,
            at: new Date().toISOString(),
        };

        session.messages.push(message);
        session.updatedAt = message.at;
        this.persist();

        this.listeners
            .get(sessionId)
            ?.forEach((listener) => listener(message));

        return message;
    }

    subscribe(sessionId, listener) {
        const listeners =
            this.listeners.get(sessionId) ?? new Set();

        listeners.add(listener);
        this.listeners.set(sessionId, listeners);

        return () => listeners.delete(listener);
    }
}
