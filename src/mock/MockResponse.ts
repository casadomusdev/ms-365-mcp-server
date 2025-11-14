export class MockResponse {
  private bodyText: string;
  private headersObj: Record<string, string>;
  status: number;
  statusText: string;

  constructor(
    body: unknown,
    init?: {
      status?: number;
      statusText?: string;
      headers?: Record<string, string>;
    }
  ) {
    this.status = init?.status ?? 200;
    this.statusText = init?.statusText ?? 'OK';
    this.bodyText =
      typeof body === 'string' ? body : body == null ? '' : JSON.stringify(body);
    this.headersObj = { ...(init?.headers ?? {}) };
    if (!this.headersObj['Content-Type']) {
      this.headersObj['Content-Type'] = 'application/json; charset=utf-8';
    }
  }

  get ok(): boolean {
    return this.status >= 200 && this.status < 300;
  }

  headers = {
    get: (name: string): string | null => {
      const key = Object.keys(this.headersObj).find(
        (k) => k.toLowerCase() === String(name).toLowerCase()
      );
      return key ? this.headersObj[key] : null;
    },
  };

  async text(): Promise<string> {
    return this.bodyText;
  }
}

export type MockResponseInit = ConstructorParameters<typeof MockResponse>[1];


