import { MockRegistry } from './registry.js';
import logger from '../logger.js';
import { MockResponse } from './MockResponse.js';

// Default mocks are intentionally minimal/empty.
// Use MS365_MCP_DRYRUN_FILE=mocks.json to load rich example data.
export function loadDefaultMocks(reg: MockRegistry, _seed?: number): void {
  logger.info('[dryrun] registering default mocks for two users (a.viets and b.viets)');

  const ok = (body: unknown) => new MockResponse(body, { status: 200 });
  const withList = (items: unknown[]) => ok({ value: items, '@odata.count': items.length });

  const USER_A = 'a.viets@casadomus.de';
  const USER_B = 'b.viets@casadomus.de';

  const messagesByUser: Record<string, Array<{ id: string; subject: string }>> = {
    [USER_A.toLowerCase()]: Array.from({ length: 8 }).map((_, i) => ({
      id: `msg_a_${String(i + 1).padStart(3, '0')}`,
      subject: `A: Dryrun message #${i + 1}`,
    })),
    [USER_B.toLowerCase()]: Array.from({ length: 8 }).map((_, i) => ({
      id: `msg_b_${String(i + 1).padStart(3, '0')}`,
      subject: `B: Dryrun message #${i + 1}`,
    })),
  };

  const calendarsCommon = [
    { id: 'cal_default', name: 'Calendar' },
    { id: 'cal_team', name: 'Team Calendar' },
  ];

  const eventsByUser: Record<string, Array<any>> = {
    [USER_A.toLowerCase()]: [
      {
        id: 'evt_a_001',
        subject: 'A: Dryrun Event 1',
        start: { dateTime: '2025-01-16T09:00:00Z', timeZone: 'UTC' },
        end: { dateTime: '2025-01-16T10:00:00Z', timeZone: 'UTC' },
        location: { displayName: 'Room 100' },
      },
      {
        id: 'evt_a_002',
        subject: 'A: Dryrun Event 2',
        start: { dateTime: '2025-01-17T09:00:00Z', timeZone: 'UTC' },
        end: { dateTime: '2025-01-17T10:00:00Z', timeZone: 'UTC' },
        location: { displayName: 'Room 101' },
      },
    ],
    [USER_B.toLowerCase()]: [
      {
        id: 'evt_b_001',
        subject: 'B: Dryrun Event 1',
        start: { dateTime: '2025-02-16T09:00:00Z', timeZone: 'UTC' },
        end: { dateTime: '2025-02-16T10:00:00Z', timeZone: 'UTC' },
        location: { displayName: 'Room 200' },
      },
      {
        id: 'evt_b_002',
        subject: 'B: Dryrun Event 2',
        start: { dateTime: '2025-02-17T09:00:00Z', timeZone: 'UTC' },
        end: { dateTime: '2025-02-17T10:00:00Z', timeZone: 'UTC' },
        location: { displayName: 'Room 201' },
      },
    ],
  };

  const getImpersonated = (): string | undefined => {
    const v = (process.env.MS365_MCP_IMPERSONATE_USER || '').trim().toLowerCase();
    return v || undefined;
  };

  // Mail messages
  reg.registerMock('GET', '/users/:userId/messages', (ctx) => {
    const email = decodeURIComponent(ctx.params.userId || '').toLowerCase();
    const src = messagesByUser[email] || [];
    const top = Number((ctx.query && ctx.query.get('$top')) || '0');
    const slice = top > 0 ? src.slice(0, top) : src;
    return withList(slice);
  });
  reg.registerMock('GET', '/me/messages', (ctx) => {
    const imp = getImpersonated();
    if (imp) {
      const src = messagesByUser[imp.toLowerCase()] || [];
      const top = Number((ctx.query && ctx.query.get('$top')) || '0');
      const slice = top > 0 ? src.slice(0, top) : src;
      return withList(slice);
    }
    // No impersonation: return all users' messages combined
    const all = Object.values(messagesByUser).flat();
    const top = Number((ctx.query && ctx.query.get('$top')) || '0');
    const slice = top > 0 ? all.slice(0, top) : all;
    return withList(slice);
  });

  // Mail folders (simple static)
  reg.registerMock('GET', '/users/:userId/mailFolders', () =>
    withList([
      { id: 'inbox', displayName: 'Inbox', totalItemCount: 8 },
      { id: 'sentitems', displayName: 'Sent Items', totalItemCount: 2 },
      { id: 'archive', displayName: 'Archive', totalItemCount: 0 },
    ])
  );

  // Calendars and events
  reg.registerMock('GET', '/me/calendars', () => withList(calendarsCommon));
  reg.registerMock('GET', '/me/events', () => {
    const imp = getImpersonated();
    if (imp) {
      return withList(eventsByUser[imp.toLowerCase()] || []);
    }
    return withList([...eventsByUser[USER_A.toLowerCase()], ...eventsByUser[USER_B.toLowerCase()]]);
  });
  reg.registerMock('GET', '/me/calendars/:calendarId/events', () => {
    const imp = getImpersonated();
    if (imp) {
      return withList(eventsByUser[imp.toLowerCase()] || []);
    }
    return withList([...eventsByUser[USER_A.toLowerCase()], ...eventsByUser[USER_B.toLowerCase()]]);
  });
  reg.registerMock('GET', '/users/:userId/events', (ctx) => {
    const email = decodeURIComponent(ctx.params.userId || '').toLowerCase();
    return withList(eventsByUser[email] || []);
  });
  reg.registerMock('GET', '/users/:userId/calendars/:calendarId/events', (ctx) => {
    const email = decodeURIComponent(ctx.params.userId || '').toLowerCase();
    return withList(eventsByUser[email] || []);
  });

  // Users listing (two users)
  reg.registerMock('GET', '/users', () =>
    withList([
      { id: USER_A, displayName: 'A Viets', userPrincipalName: USER_A, mail: USER_A },
      { id: USER_B, displayName: 'B Viets', userPrincipalName: USER_B, mail: USER_B },
    ])
  );

  // Write operations: accept or no-content for safety
  reg.registerMock(
    'POST',
    '/users/:userId/sendMail',
    () => new MockResponse('', { status: 202, statusText: 'Accepted' })
  );
  reg.registerMock(
    'DELETE',
    '/users/:userId/messages/:messageId',
    () => new MockResponse('', { status: 204, statusText: 'No Content' })
  );
}


