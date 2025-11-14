import MailboxDiscoveryCache from './MailboxDiscoveryCache.js';

export class AccessValidator {
  private cache: MailboxDiscoveryCache;

  constructor(cache: MailboxDiscoveryCache) {
    this.cache = cache;
  }

  async isAllowedMailbox(impersonatedUser: string, targetMailbox?: string): Promise<boolean> {
    const allowed = await this.cache.getMailboxes(impersonatedUser);
    const target = (targetMailbox || impersonatedUser).toLowerCase();
    return allowed.some((m) => m.email.toLowerCase() === target || m.id.toLowerCase() === target);
  }
}

export default AccessValidator;


