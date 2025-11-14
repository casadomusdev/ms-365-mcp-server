export type MatchResult = {
  params: Record<string, string>;
};

export function stripQuery(path: string): string {
  const idx = path.indexOf('?');
  return idx >= 0 ? path.slice(0, idx) : path;
}

export function matchPattern(pattern: string, path: string): MatchResult | null {
  const cleanPath = stripQuery(path);
  const patternParts = pattern.split('/').filter(Boolean);
  const pathParts = cleanPath.split('/').filter(Boolean);
  if (patternParts.length !== pathParts.length) return null;

  const params: Record<string, string> = {};

  for (let i = 0; i < patternParts.length; i++) {
    const pp = patternParts[i];
    const pv = pathParts[i];
    if (pp.startsWith(':')) {
      params[pp.slice(1)] = decodeURIComponent(pv);
      continue;
    }
    if (pp !== pv) return null;
  }
  return { params };
}


