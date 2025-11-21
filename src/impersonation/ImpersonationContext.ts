import { AsyncLocalStorage } from 'async_hooks';

interface RequestContext {
  impersonatedUser?: string;
  meta?: Record<string, any>;
}

class ImpersonationContext {
  private static storage = new AsyncLocalStorage<RequestContext>();

  static setImpersonatedUser(email: string): void {
    if (!email) return;
    const current = this.storage.getStore() || {};
    this.storage.enterWith({ ...current, impersonatedUser: email });
  }

  static getImpersonatedUser(): string | undefined {
    return this.storage.getStore()?.impersonatedUser;
  }

  static setMeta(meta: Record<string, any>): void {
    if (!meta) return;
    const current = this.storage.getStore() || {};
    this.storage.enterWith({ ...current, meta });
  }

  static getMeta(): Record<string, any> | undefined {
    return this.storage.getStore()?.meta;
  }

  static withUser<T>(email: string, fn: () => Promise<T>): Promise<T> {
    return this.storage.run({ impersonatedUser: email }, fn);
  }

  static withContext<T>(context: RequestContext, fn: () => Promise<T>): Promise<T> {
    return this.storage.run(context, fn);
  }
}

export default ImpersonationContext;
