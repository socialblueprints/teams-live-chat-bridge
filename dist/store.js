export class MemoryStore {
    channel;
    sessions = new Map();
    rootToSession = new Map();
    listeners = new Map();
    create(visitorName, page) {
        const id = crypto.randomUUID();
        const session = { id, visitorName, page, messages: [] };
        this.sessions.set(id, session);
        return session;
    }
    add(sessionId, sender, text) {
        const session = this.sessions.get(sessionId);
        if (!session)
            throw new Error('Chat session not found');
        const message = { id: crypto.randomUUID(), sender, text, at: new Date().toISOString() };
        session.messages.push(message);
        this.listeners.get(sessionId)?.forEach((listener) => listener(message));
        return message;
    }
    subscribe(sessionId, listener) {
        const set = this.listeners.get(sessionId) ?? new Set();
        set.add(listener);
        this.listeners.set(sessionId, set);
        return () => set.delete(listener);
    }
}
