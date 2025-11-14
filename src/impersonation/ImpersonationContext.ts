import { AsyncLocalStorage } from 'async_hooks';

class ImpersonationContext {
  private static storage = new AsyncLocalStorage<string>();

  static setImpersonatedUser(email: string): void {
    if (!email) return;
    this.storage.enterWith(email);
  }

  static getImpersonatedUser(): string | undefined {
    return this.storage.getStore();
  }

  static withUser<T>(email: string, fn: () => Promise<T>): Promise<T> {
    return this.storage.run(email, fn);
  }
}

export default ImpersonationContext;


