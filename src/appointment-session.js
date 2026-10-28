import { assertAppointmentContext } from "./outlook.js";

/** In-memory scope. Invalidating also blocks A -> B -> A during an operation. */
export class AppointmentSession {
  constructor() { this.revision = 0; this.context = null; this.invalid = true; this.drafts = new Map(); }
  invalidate() { this.revision++; this.invalid = true; }
  use(context) { this.revision++; this.context = context; this.invalid = false; }
  capture() {
    const context = this.context;
    const revision = this.revision;
    return { context, ensure: () => {
      if (this.invalid || this.revision !== revision || this.context !== context) {
        const error = new Error("Le rendez-vous Outlook a changé. Cliquez sur Actualiser pour le recharger.");
        error.code = "OUTLOOK_CONTEXT_CHANGED";
        throw error;
      }
      assertAppointmentContext(context);
    } };
  }
  rememberDraft(subject, text) {
    if (!this.context?.itemId) return;
    const key = this.key(this.context);
    if (text.trim()) this.drafts.set(key, { subject, text });
    else this.drafts.delete(key);
  }
  draft() { return this.drafts.get(this.key(this.context)); }
  key(context) { return JSON.stringify([context?.mailboxEmail?.toLowerCase(), context?.restItemId || context?.itemId]); }
}
