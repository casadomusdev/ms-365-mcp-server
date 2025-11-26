#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="${SCRIPT_DIR}"

# Load local environment overrides if present so feature toggles are respected.
if [[ -f "${PROJECT_ROOT}/.env" ]]; then
  # shellcheck disable=SC1090
  set -a
  source "${PROJECT_ROOT}/.env"
  set +a
fi

ENDPOINTS_FILE="${PROJECT_ROOT}/dist/endpoints.json"
if [[ ! -f "${ENDPOINTS_FILE}" ]]; then
  ENDPOINTS_FILE="${PROJECT_ROOT}/src/endpoints.json"
fi

if [[ ! -f "${ENDPOINTS_FILE}" ]]; then
  echo "Unable to locate endpoints.json in dist/ or src/. Please build the project first." >&2
  exit 1
fi

ENDPOINTS_FILE_PATH="${ENDPOINTS_FILE}" node --input-type=module <<'NODE'
import fs from 'node:fs';

const parseBool = (value, defaultValue = false) => {
  if (value == null || value === '') {
    return defaultValue;
  }
  const normalized = String(value).trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes' || normalized === 'on';
};

const endpointsPath = process.env.ENDPOINTS_FILE_PATH;
if (!endpointsPath || !fs.existsSync(endpointsPath)) {
  console.error('ENDPOINTS_FILE_PATH is not set or does not exist.');
  process.exit(1);
}

const endpoints = JSON.parse(fs.readFileSync(endpointsPath, 'utf8'));

const toggles = {
  mail: parseBool(process.env.MS365_MCP_ENABLE_MAIL, false),
  calendar: parseBool(process.env.MS365_MCP_ENABLE_CALENDAR, false),
  files: parseBool(process.env.MS365_MCP_ENABLE_FILES, false),
  teams: parseBool(process.env.MS365_MCP_ENABLE_TEAMS, false),
  excel: parseBool(process.env.MS365_MCP_ENABLE_EXCEL_POWERPOINT, false),
  onenote: parseBool(process.env.MS365_MCP_ENABLE_ONENOTE, false),
  tasks: parseBool(process.env.MS365_MCP_ENABLE_TASKS, false),
  contacts: parseBool(process.env.MS365_MCP_ENABLE_CONTACTS, false),
  user: parseBool(process.env.MS365_MCP_ENABLE_USER, false),
  search: parseBool(process.env.MS365_MCP_ENABLE_SEARCH, false),
};

const orgMode = parseBool(process.env.MS365_MCP_ORG_MODE, false);

const WRAPPER_ALIASES = new Set([
  'list-calendar-events',
  'get-calendar-event',
  'create-calendar-event',
  'update-calendar-event',
  'delete-calendar-event',
  'list-mail-messages',
  'get-mail-message',
  'send-mail',
]);

const INTERNAL_ONLY_ALIASES = new Set([
  'list-specific-calendar-events',
  'get-specific-calendar-event',
  'create-specific-calendar-event',
  'update-specific-calendar-event',
  'delete-specific-calendar-event',
  'list-mail-folder-messages',
  'list-shared-mailbox-messages',
  'list-shared-mailbox-folder-messages',
  'get-shared-mailbox-message',
  'send-shared-mailbox-mail',
]);

const getCategory = (pathPattern) => {
  const p = String(pathPattern || '').toLowerCase();

  if (
    p.includes('/me/messages') ||
    p.includes('/users/{user-id}/messages') ||
    p.includes('/mailfolders') ||
    p.includes('/sendmail') ||
    p.includes('/attachments')
  ) {
    return 'mail';
  }

  if (
    p.includes('/events') ||
    p.includes('/calendarview') ||
    p.includes('/calendars') ||
    p.includes('/findmeetingtimes')
  ) {
    return 'calendar';
  }

  if (p.includes('/drives') || p.startsWith('/sites')) {
    return 'files';
  }

  if (p.includes('/workbook/')) {
    return 'excel';
  }

  if (p.includes('/onenote/')) {
    return 'onenote';
  }

  if (p.includes('/todo/') || p.includes('/planner/')) {
    return 'tasks';
  }

  if (p.includes('/chats') || p.includes('/joinedteams') || p.startsWith('/teams')) {
    return 'teams';
  }

  if (p.includes('/contacts')) {
    return 'contacts';
  }

  if (p === '/me' || p === '/users') {
    return 'user';
  }

  if (p === '/search/query') {
    return 'search';
  }

  return undefined;
};

const categoryLabels = {
  mail: 'Mail',
  calendar: 'Calendar',
  files: 'Files',
  teams: 'Teams',
  excel: 'Excel/PowerPoint',
  onenote: 'OneNote',
  tasks: 'Tasks/Planner',
  contacts: 'Contacts',
  user: 'User',
  search: 'Search',
};

const colors = {
  mail: '\u001b[32m', // green
  calendar: '\u001b[36m', // cyan
  files: '\u001b[94m', // bright blue
  teams: '\u001b[31m', // red
  excel: '\u001b[33m', // yellow
  onenote: '\u001b[95m', // bright magenta
  tasks: '\u001b[93m', // bright yellow-orange
  contacts: '\u001b[92m', // bright green
  user: '\u001b[37m', // white
  search: '\u001b[90m', // gray
};
const resetColor = '\u001b[0m';

const isCategoryEnabled = (pathPattern) => {
  const category = getCategory(pathPattern);
  if (!category) return false;
  return toggles[category] ?? false;
};

const enabledTools = [];
const disabledTools = [];

for (const endpoint of endpoints) {
  const { pathPattern, toolName, scopes, workScopes } = endpoint;

  if (!toolName || typeof toolName !== 'string') {
    continue;
  }

  if (INTERNAL_ONLY_ALIASES.has(toolName)) {
    continue;
  }

  const category = getCategory(pathPattern) || 'unknown';
  const isOrgModeRequired = !orgMode && (!scopes || scopes.length === 0) && workScopes && workScopes.length > 0;
  const isCategoryEnabled = toggles[category] ?? false;

  if (isOrgModeRequired || !isCategoryEnabled) {
    let reason = '';
    if (isOrgModeRequired) {
      reason = 'org-mode-required';
    } else if (!isCategoryEnabled) {
      reason = 'category-disabled';
    }
    disabledTools.push({ name: toolName, category, reason });
    continue;
  }

  enabledTools.push({ name: toolName, category });
}

const uniqueEnabledMap = new Map();
for (const tool of enabledTools) {
  if (!uniqueEnabledMap.has(tool.name)) {
    uniqueEnabledMap.set(tool.name, tool);
  }
}

const uniqueEnabled = Array.from(uniqueEnabledMap.values()).sort((a, b) =>
  a.name.localeCompare(b.name)
);

const uniqueDisabledMap = new Map();
for (const tool of disabledTools) {
  if (!uniqueDisabledMap.has(tool.name)) {
    uniqueDisabledMap.set(tool.name, tool);
  }
}

const uniqueDisabled = Array.from(uniqueDisabledMap.values()).sort((a, b) =>
  a.name.localeCompare(b.name)
);

if (uniqueEnabled.length === 0 && uniqueDisabled.length === 0) {
  console.log('No tools found.');
  process.exit(0);
}

console.log('Tool origins:');
console.log(
  Object.entries(colors)
    .map(([key, color]) => `${color}${categoryLabels[key] || key}${resetColor}`)
    .join('  ')
);
console.log();

// Group enabled tools
if (uniqueEnabled.length > 0) {
  const enabledGrouped = uniqueEnabled.reduce((acc, tool) => {
    const key = tool.category || 'unknown';
    if (!acc[key]) acc[key] = [];
    acc[key].push(tool.name);
    return acc;
  }, {});

  const sortedEnabledCategories = Object.keys(enabledGrouped).sort((a, b) => {
    const labelA = categoryLabels[a] || a;
    const labelB = categoryLabels[b] || b;
    return labelA.localeCompare(labelB);
  });

  for (const category of sortedEnabledCategories) {
    const color = colors[category] || '';
    const label = categoryLabels[category] || category || 'Unknown';
    console.log(`${color}${label}${resetColor}:`);
    const tools = enabledGrouped[category].sort((a, b) => a.localeCompare(b));
    for (const name of tools) {
      const wrapperTag = WRAPPER_ALIASES.has(name) ? ' (wrapper)' : '';
      console.log(`  ${color}${name}${wrapperTag}${resetColor}`);
    }
    console.log();
  }
} else {
  console.log('No enabled tools with current feature toggles.');
  console.log();
}

// Group disabled tools
if (uniqueDisabled.length > 0) {
  const disabledGrouped = uniqueDisabled.reduce((acc, tool) => {
    const key = tool.category || 'unknown';
    if (!acc[key]) acc[key] = [];
    acc[key].push(tool.name);
    return acc;
  }, {});

  const sortedDisabledCategories = Object.keys(disabledGrouped).sort((a, b) => {
    const labelA = categoryLabels[a] || a;
    const labelB = categoryLabels[b] || b;
    return labelA.localeCompare(labelB);
  });

  const dimColor = '\u001b[2m'; // dim
  const grayColor = '\u001b[90m'; // gray

  console.log(`${dimColor}Disabled tools:${resetColor}`);
  console.log();

  for (const category of sortedDisabledCategories) {
    const color = colors[category] || '';
    const label = categoryLabels[category] || category || 'Unknown';
    console.log(`${dimColor}${grayColor}${label}${resetColor}${dimColor} (disabled):${resetColor}`);
    const tools = disabledGrouped[category].sort((a, b) => a.localeCompare(b));
    for (const name of tools) {
      console.log(`  ${dimColor}${grayColor}${name}${resetColor}`);
    }
    console.log();
  }
}
NODE

AUTH_TOOLS=(
  "login"
  "logout"
  "verify-login"
  "list-accounts"
  "select-account"
  "remove-account"
  "list-impersonated-mailboxes"
)

if [[ ${#AUTH_TOOLS[@]} -gt 0 ]]; then
  echo
  echo "Auth tools (only when auth tools are registered):"
  for tool in "${AUTH_TOOLS[@]}"; do
    printf '\033[33m%s\033[0m\n' "$tool"
  done
fi

