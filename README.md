# Microsoft Outlook MCP

An [MCP](https://modelcontextprotocol.io) server that connects Claude (or any
MCP client) to Microsoft Outlook — **Mail, Calendar, and Contacts** — through
the [Microsoft Graph API](https://learn.microsoft.com/graph/) using OAuth 2.0
**device code flow**.

- Cloud-based via Microsoft Graph — works anywhere, for Microsoft 365 and
  personal Microsoft accounts.
- No client secret required. Device code flow uses a **public client** app.
- Tokens are cached on disk and refreshed silently, so you sign in once.

## Tools

| Area | Tools |
| --- | --- |
| Account | `whoami`, `auth_status` |
| Mail | `list_messages`, `get_message`, `send_mail`, `reply_to_message`, `update_message`, `move_message`, `delete_message` |
| Calendar | `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `respond_to_event` |
| Contacts | `list_contacts`, `get_contact`, `create_contact`, `update_contact`, `delete_contact` |
| Folders | `list_mail_folders`, `create_mail_folder`, `rename_mail_folder`, `delete_mail_folder` |
| Inbox rules | `list_message_rules`, `get_message_rule`, `create_message_rule`, `update_message_rule`, `delete_message_rule` |

## 1. Register an app in Azure AD (Entra ID)

1. Go to the [Azure Portal](https://portal.azure.com) →
   **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it e.g. `Outlook MCP`.
3. **Supported account types**: choose based on the accounts you'll use:
   - *Accounts in any organizational directory and personal Microsoft accounts*
     → tenant `common`
   - *This organizational directory only* → your specific tenant GUID
4. Leave **Redirect URI** blank (device code flow doesn't need one). Register.
5. On the app's **Overview** page, copy the **Application (client) ID** and
   **Directory (tenant) ID**.
6. **Authentication** → **Advanced settings** → set
   **Allow public client flows** to **Yes**. (Required for device code flow.)
7. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions**, and add:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `MailboxSettings.ReadWrite` (required for inbox rules)
   - `offline_access` (added automatically for refresh tokens)

   Click **Grant admin consent** if you're in an org tenant that requires it.

## 2. Configure

```powershell
cd C:\Users\zlalime\Documents\Github\Microsoft-Outlook-MCP
copy .env.example .env
```

Edit `.env` and set at least `OUTLOOK_CLIENT_ID` and `OUTLOOK_TENANT_ID`.

## 3. Build & sign in

```powershell
npm install
npm run build
npm run login
```

`npm run login` prints a URL and a code. Open
<https://microsoft.com/devicelogin>, enter the code, and approve. The session
is cached in `.token-cache.json` (git-ignored). Useful variants:

```powershell
npm run login -- --status   # check whether a session is cached
npm run login -- --logout   # clear the cached session
```

## 4. Connect a client

### Claude Code

```powershell
claude mcp add outlook --scope user -- node "C:\Users\zlalime\Documents\Github\Microsoft-Outlook-MCP\dist\index.js"
```

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:\\Users\\zlalime\\Documents\\Github\\Microsoft-Outlook-MCP\\dist\\index.js"]
    }
  }
}
```

The server reads `.env` from its own folder, so no `env` block is needed. If
you prefer, you can instead pass the values inline:

```json
"env": {
  "OUTLOOK_CLIENT_ID": "…",
  "OUTLOOK_TENANT_ID": "common"
}
```

Restart the client. Ask Claude to run `auth_status` or `whoami` to confirm.

## How auth works

- The MCP server runs **non-interactively** over stdio, so it never prompts.
  At request time it acquires tokens **silently** from the on-disk cache
  (`acquireTokenSilent`), refreshing with the stored refresh token as needed.
- Interactive **device-code sign-in** happens only in the separate
  `npm run login` step. Run it again any time a tool reports the session
  expired.
- The token cache is stored in **OS-native encrypted storage** via
  [`@azure/msal-node-extensions`](https://www.npmjs.com/package/@azure/msal-node-extensions):
  DPAPI on Windows, Keychain on macOS, libsecret on Linux. The backend is
  verified at startup; if the native layer can't be loaded, the server logs a
  warning and falls back to a restricted-permission plaintext file so sign-in
  still works.
- All diagnostics go to **stderr**; **stdout** carries only the MCP JSON-RPC
  stream.

## Environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `OUTLOOK_CLIENT_ID` | *(required)* | Azure AD application (client) ID |
| `OUTLOOK_TENANT_ID` | *(required)* | **Tenant GUID** (e.g. `12345678-1234-1234-1234-123456789abc`). This hardened fork **requires** a real tenant GUID and rejects `common`, `organizations`, or `consumers`. |
| `OUTLOOK_SCOPES` | `User.Read Mail.ReadWrite Mail.Send Calendars.ReadWrite Contacts.ReadWrite MailboxSettings.ReadWrite` | Space-separated delegated Graph scopes. **Scopes are allowlisted in code**. Any scope not on the allowlist will cause startup to fail. Do not request `*.Shared`, `Directory.*`, `Sites.*`, or `Files.*` scopes. |
| `OUTLOOK_TOKEN_CACHE_PATH` | `.token-cache.json` | Where the token cache is stored. **OS-native encrypted storage is required** (Windows DPAPI / macOS Keychain / Linux libsecret). This fork refuses to fall back to plaintext. |
| `OUTLOOK_EXPECTED_USERNAME` | *(recommended)* `rlourens@opsmateai.com` | The UPN of the expected signed-in user. When set, silent token acquisition will fail if a different account is cached. **Strongly recommended for work-tenant forks to prevent accidental use of the wrong cached account.** |
| `OUTLOOK_ALLOW_WRITES` | `false` | **Write gate**. When `false` (default), dangerous write operations are disabled or restricted. See "Write gates" below. Set to `true` only when you need to send mail, delete resources, or create events with attendees. |

## Security hardening (this fork)

This is a **hardened fork** of the original `taddiemason/Microsoft-Outlook-MCP`
designed to be safer when run as a local Graph proxy for a single work user:
**rlourens@opsmateai.com**.

Key changes:

1. **Graph pagination origin check**: `@odata.nextLink` URLs are validated to
   ensure they point to `https://graph.microsoft.com` before the Bearer token
   is sent. Prevents token leakage to malicious redirect hosts.

2. **Path sanitization**: Message IDs, folder IDs, contact IDs, event IDs, and
   folder display names are validated and encoded. Rejects `/` and `..` to
   prevent path traversal attacks that could escape `/me/` scope.

3. **Scope allowlist**: Graph scopes are hard-coded in `src/config.ts`. The
   allowed set is:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Contacts.ReadWrite`
   - `MailboxSettings.ReadWrite`
   - `offline_access` (added automatically by MSAL)

   Any scope not on this list (especially `*.Shared`, `Directory.*`, `Sites.*`,
   or `Files.*`) causes startup to fail. Do not request extra scopes in
   `OUTLOOK_SCOPES`.

4. **Encrypted token cache enforced**: OS-native encrypted storage (Windows
   DPAPI / macOS Keychain / Linux libsecret) is **required**. If the native
   backend is unavailable, startup fails. No plaintext fallback. See
   troubleshooting for how to enable encrypted storage on your OS.

5. **Tenant GUID required**: `OUTLOOK_TENANT_ID` must be set to your
   organization's tenant GUID. This fork rejects `common`, `organizations`, and
   `consumers` to prevent sign-in with the wrong account type.

6. **Account binding**: Set `OUTLOOK_EXPECTED_USERNAME` (e.g.
   `rlourens@opsmateai.com`) to bind silent token acquisition to a specific
   UPN. The server refuses to use cached accounts that don't match. Prevents
   accidentally picking the wrong cached account (e.g. a personal account when
   you meant to use a work account).

7. **Write gates**: Dangerous write operations are gated behind
   `OUTLOOK_ALLOW_WRITES=true` (default: `false`). When writes are disabled:
   - Inbox rule write operations (`create_message_rule`, `update_message_rule`,
     `delete_message_rule`) are **completely removed** to prevent durable
     mail-forwarding rules with `forwardTo`.
   - `send_mail` and `reply_to_message` are not registered (sending disabled).
   - `delete_message` with `permanent=true` returns an error (soft delete to
     trash is still allowed).
   - `create_event` with attendees returns an error (creating events without
     attendees is still allowed).
   - `delete_event`, `delete_contact`, and `delete_mail_folder` are not
     registered.

   Set `OUTLOOK_ALLOW_WRITES=true` only when you need these operations. Read,
   list, get, mark read/unread, move messages, and update flags remain
   available.

## Write gates

By default, `OUTLOOK_ALLOW_WRITES=false`. In this mode:

| Tool | Availability |
| --- | --- |
| **Mail** | |
| `list_messages`, `get_message` | ✅ Available |
| `update_message` (read/flag) | ✅ Available |
| `move_message` | ✅ Available |
| `delete_message` (soft) | ✅ Available (moves to trash) |
| `delete_message` (permanent) | ❌ Requires `OUTLOOK_ALLOW_WRITES=true` |
| `send_mail` | ❌ Not registered |
| `reply_to_message` | ❌ Not registered |
| **Mail (drafts)** | |
| `create_draft` | ✅ Available (creates draft without sending) |
| `create_reply_draft` | ✅ Available (creates reply draft without sending) |
| `create_reply_all_draft` | ✅ Available (creates reply-all draft without sending) |
| `create_forward_draft` | ✅ Available (creates forward draft without sending) |
| `update_draft` | ✅ Available (updates draft message) |
| **Calendar** | |
| `list_events`, `get_event` | ✅ Available |
| `create_event` (no attendees) | ✅ Available |
| `create_event` (with attendees) | ❌ Requires `OUTLOOK_ALLOW_WRITES=true` |
| `update_event` | ✅ Available |
| `respond_to_event` | ✅ Available |
| `delete_event` | ❌ Not registered |
| **Contacts** | |
| `list_contacts`, `get_contact` | ✅ Available |
| `create_contact`, `update_contact` | ✅ Available |
| `delete_contact` | ❌ Not registered |
| **Folders** | |
| `list_mail_folders` | ✅ Available |
| `create_mail_folder`, `rename_mail_folder` | ✅ Available |
| `delete_mail_folder` | ❌ Not registered |
| **Inbox rules** | |
| `list_message_rules`, `get_message_rule` | ✅ Available |
| `create_message_rule`, `update_message_rule`, `delete_message_rule` | ❌ Completely disabled (not registered even with `OUTLOOK_ALLOW_WRITES=true` — too dangerous for a local proxy) |

When `OUTLOOK_ALLOW_WRITES=true`, the tools marked with ❌ become available
(except inbox rule writes, which remain disabled). Use this setting only when
you need to send mail, delete resources, or create meetings with attendees.

## Security notes

- The token cache holds live refresh/access tokens. **OS-native encrypted
  storage is required** in this fork (Windows DPAPI / macOS Keychain / Linux
  libsecret). If the native backend is unavailable, the server refuses to
  start. See troubleshooting below for setup instructions.
- **Client ID and tenant ID are not secrets** — this is a **public client**
  app, so no client secret is stored anywhere.
- Scopes are delegated: the server can only do what your signed-in account can.
- Run `npm run login -- --logout` to clear the token cache and revoke local
  access.

## Troubleshooting

### Encrypted storage isn't being used (native module)

The encrypted cache is provided by `@azure/msal-node-extensions`, which relies
on a **native (N-API) addon**. `npm install` normally downloads a prebuilt
binary for your platform. If none is available for your OS/CPU/Node version,
the build falls back to compiling from source, which needs a toolchain.

You'll know the native layer failed to load if the startup log shows:

```
[auth] secure token storage unavailable (...); falling back to a plaintext file...
```

instead of:

```
[auth] token cache: OS-native encrypted storage (msal-node-extensions)
```

The server still works in this state — it just stores tokens in a
restricted-permission plaintext file. To get encryption working:

- **Windows** — install the
  [Visual Studio Build Tools](https://visualstudio.microsoft.com/downloads/)
  ("Desktop development with C++"), then `npm rebuild`.
- **macOS** — install the Xcode command line tools: `xcode-select --install`,
  then `npm rebuild`.
- **Linux** — install `build-essential`, `python3`, and libsecret headers
  (Debian/Ubuntu: `sudo apt-get install build-essential python3 libsecret-1-dev`);
  ensure a Secret Service provider (e.g. GNOME Keyring) is running, then
  `npm rebuild`.

After rebuilding, run `npm run login -- --status` and confirm the log reports
OS-native encrypted storage.

### `AADSTS7000218` or "public client flow" errors during login

Your app registration is missing the public-client setting. In Azure →
**Authentication** → **Advanced settings**, set **Allow public client flows**
to **Yes**.

### Tools return HTTP 401 / 403 after adding scopes

The cached token predates the new permission. Re-run `npm run login` (and grant
admin consent in Azure if your tenant requires it) so a fresh token carries the
added scope.

### `whoami` / tools say "Not signed in"

Run `npm run login` in a terminal from the project folder. The MCP server never
prompts for sign-in itself — it only reads the cached session.

## Project layout

```
src/
  index.ts          MCP server entry: registers tools, connects stdio transport
  login.ts          Standalone device-code sign-in (npm run login)
  config.ts         .env loader + config
  auth.ts           MSAL public client, device code + silent refresh, encrypted cache
  graph.ts          Minimal fetch-based Graph client (auth, paging, errors)
  tools/
    mail.ts         Mail tools
    calendar.ts     Calendar tools
    contacts.ts     Contacts tools
    rules.ts        Mail folder management + inbox rule tools
    util.ts         ok()/fail() result helpers
```

## License

MIT
