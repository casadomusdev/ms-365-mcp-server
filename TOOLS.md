## MCP Tooling Notes

### Consolidation strategy

- We generate the full Microsoft Graph surface (see `src/generated/client.ts`), but only expose a curated subset to MCP clients.
- `src/tool-blacklist.ts` enumerates the raw aliases that overlap (specific calendar/mailbox variants, shared mailbox helpers, etc.). `registerGraphTools` loads every generated handler, stores it in-memory, then skips calling `server.tool()` for anything in the blacklist.
- Public wrappers (calendar + mail) are registered manually inside `registerGraphTools`. They introduce optional routing params:
  - Calendar wrappers (`list/get/create/update/delete-calendar-event`) accept `calendarId?` and call the specific alias when provided.
  - Mail wrappers (`list-mail-messages`, `get-mail-message`, `send-mail`) accept `folderId?`, `sharedMailboxId?`, `sharedMailboxEmail?` and dispatch to the correct internal handler (own mailbox, folder, shared mailbox, shared mailbox folder).
- Wrappers reuse the stored handler metadata so behavior stays identical; only the exposed surface shrinks. See `approach.md` for the full rationale and mapping table.

### Tool descriptions & docs

- Wrapper descriptions call out the optional routing params in `src/tool-descriptions.ts`.
- `README.md` contains a dedicated “Outlook Tool Routing” section describing the new selector semantics, plus updated shared mailbox instructions.

### Listing available tools

- Run `./util-list-tools.sh` after setting the relevant `MS365_MCP_ENABLE_*` toggles (and `MS365_MCP_ORG_MODE` if needed). The script loads `dist/endpoints.json` (falling back to `src/endpoints.json`) and prints colored groups of enabled/disabled tools.
- Wrapper aliases are annotated with ` (wrapper)` in the listing output; internal-only aliases are filtered out so the table mirrors what MCP clients actually see.
  ```bash
  MS365_MCP_ENABLE_MAIL=true MS365_MCP_ENABLE_CALENDAR=true ./util-list-tools.sh
  ```
