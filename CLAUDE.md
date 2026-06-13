# CLAUDE.md - AI Maestro Gateways

## Project Overview

Multi-platform messaging gateway monorepo connecting AI agents to Discord, Slack, Email, and WhatsApp via the AMP (Agent Messaging Protocol). Part of the [AI Maestro](https://github.com/23blocks-OS/ai-maestro) ecosystem by 23blocks-OS.

## Tech Stack

- **Language:** TypeScript 5.7 (strict mode)
- **Runtime:** Node.js 22+ with tsx (no build step needed for dev)
- **Module system:** ESM (`"type": "module"`) — import paths use `.js` extensions
- **HTTP:** Express 4
- **Test framework:** `node:test` (Node.js built-in)
- **Lint:** ESLint 9 with typescript-eslint
- **Process manager:** PM2 (ecosystem.config.cjs in each gateway)
- **Containerization:** Docker multi-stage builds, orchestrated via docker-compose.yml

## Repository Structure

```
aimaestro-gateways/
├── discord-gateway/    # discord.js v14, port 3023
├── slack-gateway/      # Bolt SDK + Socket Mode, port 3022
├── email-gateway/      # Mandrill webhooks + SMTP, port 3020 (has Vue.js UI in ui/)
├── whatsapp-gateway/   # Baileys (WhatsApp Web), port 3021 (Beta)
├── docker-compose.yml
├── eslint.config.js    # Shared lint config
└── package.json        # Root scripts for lint/typecheck/test
```

Each gateway follows the same internal layout:
```
src/
├── server.ts           # Express server + platform client bootstrap + main()
├── inbound.ts          # Platform event -> content security -> AMP route
├── outbound.ts         # Poll AMP inbox -> format -> send to platform
├── config.ts           # Env loading + AMP auto-registration
├── types.ts            # GatewayConfig, AMPEnvelope, AMPPayload, ResolvedUser, etc.
├── content-security.ts # Trust resolution (user directory + legacy env var fallback) + 34-pattern injection scanner
├── amp-bootstrap.ts    # Ed25519 keygen + /api/v1/register
├── agent-resolver.ts   # @AIM:agent-name parsing + cached resolution
├── user-resolver.ts    # User directory resolution via Maestro /api/users/resolve (Discord only, Phase 2)
└── api/                # /health, /api/config, /api/stats, /api/activity, /api/gateway/dm
```

Email gateway additionally has `router.ts` (address->agent mapping) and `ui/` (Vue.js management dashboard).
WhatsApp gateway has `session.ts` (Baileys socket/QR), `normalize.ts` (E.164), `router.ts`, and `scripts/login.ts`.

## Commands

```bash
# Root-level
npm run lint                # ESLint all gateways
npm run typecheck           # tsc --noEmit (currently only discord + slack!)
npm test                    # Runs discord-gateway tests only

# Per-gateway
cd discord-gateway
npm start                   # tsx src/server.ts
npm run dev                 # tsx watch src/server.ts
npm run typecheck           # tsc --noEmit
npm test                    # node --import tsx --test src/__tests__/*.test.ts
```

## Code Conventions

- 2-space indentation
- camelCase for variables/functions, PascalCase for types/interfaces
- Console logging with `[CONTEXT]` prefixes (e.g., `[DISCORD]`, `[SECURITY]`, `[AMP]`)
- Timing-safe comparison (`crypto.timingSafeEqual`) for all token/auth checks
- Graceful shutdown via SIGTERM/SIGINT handlers with 10s timeout
- Management API endpoints require Bearer token auth via ADMIN_TOKEN env var
- Activity logging uses an in-memory ring buffer (500 events max)

## Architecture: Message Flow

**Inbound:** Platform event -> user directory resolve (cache + auto-create unknown senders) -> trust resolution (user directory role/trustLevel, legacy env var fallback) -> injection scan -> `<external-content>` wrapping if untrusted -> AMP route request to `/api/v1/route` with enriched context (sender, thread, topicHints)

**Outbound (replies):** Filesystem poll of `~/.agent-messaging/agents/{uuid}/messages/inbox/` -> format for platform -> deliver response

**Outbound (DMs):** Maestro `/api/users/:id/notify` -> resolves preferred platform -> calls gateway `POST /api/gateway/dm` -> Discord DM delivered

Agent targeting uses `@AIM:agent-name` syntax in messages, resolved via cached lookups against AMP hosts.

## Content Security Model

- Operator trust now resolved via Maestro user directory (`role === 'operator'` or `trustLevel === 'full'`), with fallback to `OPERATOR_*_IDS` env vars (legacy)
- Operator messages bypass scanning entirely
- External messages scanned against 34 regex patterns (injection, extraction, exfiltration, etc.)
- Flagged content wrapped in `<external-content source="..." sender="..." trust="none">` tags
- Scanner short-circuits after 5 flags or 10K chars (DoS protection)
- Unicode normalization (NFKD + zero-width char stripping) applied before scanning

## Key Environment Variables (shared across gateways)

```
PORT                    # Gateway HTTP port
AIMAESTRO_URL           # AI Maestro API base URL (default http://127.0.0.1:23000)
DEFAULT_AGENT           # Default target agent name
ADMIN_TOKEN             # Bearer token for management API + DM endpoint auth
AMP_API_KEY             # AMP authentication (auto-generated if missing via bootstrap)
AMP_AGENT_ADDRESS       # AMP address (auto-assigned during bootstrap)
CACHE_USER_TTL_MS       # User directory cache TTL (default 300000 = 5min)
DEBUG                   # Enable debug logging
```

## Observations and Known Issues

### Bugs Found and Fixed
1. **Email webhook auth bypass** — `server.ts:350-363` had missing else branch allowing unverified webhooks when webhook key was empty/falsy. Fixed with explicit reject-all else clause.
2. **Email outbound sender spoofing** — `outbound.ts` passed agent-supplied `from` to Mandrill with no validation. Fixed with `allowedFromDomains` config + domain validation + rewrite to `defaultFrom`.
3. **AMP bootstrap inbox path mismatch** — `amp-bootstrap.ts` saves inbox as `agents/{agentName}/` but Maestro delivers to `agents/{agentId}/` (UUID). Workaround: manually set `AMP_INBOX_DIR` in .env. Code fix still needed in bootstrap.
4. **`.env.example` uses wrong var name** — says `DEFAULT_AGENT` but config reads `AMP_DEFAULT_AGENT`.
5. **README claims slash commands** for Discord gateway but none are implemented.
6. **Discord gateway auth bypass when ADMIN_TOKEN unset** — `server.ts:authMiddleware` had `if (!adminToken) return next()`, leaving `/api/*` (incl. `/api/gateway/dm`) unauthenticated whenever the token was missing. Same shape as bug #1. Fixed by removing the bypass branch and making `config.ts` throw at startup if `ADMIN_TOKEN` is empty (fail-closed). Slack/email/whatsapp gateways likely have the same copy-pasted middleware — audit pending.
7. **Discord gateway loopback-only bind** — `server.ts` hardcoded `httpApp.listen(port, '127.0.0.1', …)`, blocking cross-host calls (e.g. another Maestro/agent on the tailscale mesh). Fixed by adding a `HOST` env var that accepts comma-separated addresses (e.g. `HOST=127.0.0.1,<tailscale-ip>`); `server.ts` now opens one listener per host. Default remains `127.0.0.1`. Maestro currently hardcodes `localhost` for gateway URL (`services/users-service.ts:272`), so dual-bind is needed when the gateway must serve both Maestro-on-same-host and remote agents.

### Code Duplication
`content-security.ts` is copy-pasted across all 4 gateways with only the trust field name differing (operatorDiscordIds, operatorSlackIds, operatorEmails, operatorPhones). Similarly, `amp-bootstrap.ts`, `types.ts`, and `cache.ts` are near-identical. These should be extracted into a shared `packages/common/` or similar.

### Incomplete Root Scripts
- `npm run typecheck` only runs discord + slack — email and whatsapp are defined but not wired in
- `npm test` only runs discord-gateway tests — no aggregate test runner

### Missing Test Coverage
Only discord-gateway has tests (`src/__tests__/content-security.test.ts`). Slack, email, and whatsapp have zero tests despite having identical security-critical code.

### Email Inbound Not Separated
Email gateway's inbound webhook handling lives in `server.ts` (~18KB) instead of a separate `inbound.ts` like the other gateways. This makes it harder to follow the consistent pattern.

### WhatsApp Beta Gaps
Marked as Beta. Has an `ARCHITECTURE.md` with design rationale but no tests, and the `api/` directory only has activity router (missing config and stats endpoints that other gateways have).

### Potential Security Review Items
- The 34 injection patterns in content-security.ts should be reviewed for bypass techniques (unicode tricks beyond NFKD, homoglyph substitution, token-boundary exploits)
- Tag escape logic for `</external-content>` injection should be verified
- Mandrill webhook HMAC verification in email gateway needs test coverage

## User Directory Integration (Phase 2+3 — completed 2026-04-08)

Discord gateway now integrates with Maestro's centralized user directory:

- **user-resolver.ts**: HTTP-backed resolver with local TTL cache. Calls `GET /api/users/resolve?platform=discord&platformUserId=...`. Auto-creates unknown senders via `POST /api/users/auto-create`. Note: Maestro API wraps responses in `{ "user": {...} }` — must unwrap.
- **Trust migration**: `content-security.ts:resolveTrust()` accepts optional `ResolvedUser` param. Prefers user directory trust (`role/trustLevel`), falls back to `OPERATOR_DISCORD_IDS` env var.
- **AMP envelope enrichment**: Inbound messages include `context.sender` (userId, displayName, trustLevel, role), `context.thread` (threadId, inReplyTo, isNewConversation), and `context.topicHints` (max 3 keywords). These feed Maestro's memory retrieval middleware.
- **Outbound DMs**: `POST /api/gateway/dm` accepts `{ platformUserId, message }`, validates mutual guild membership, sends Discord DM. Auth via ADMIN_TOKEN Bearer token.
- **isNewConversation heuristic**: DMs with no thread activity in 30min = new. Guild @mentions = always new.

Maestro-side counterparts:
- `POST /api/users/:id/notify` routes outbound to user's preferred platform via gateway DM endpoint
- Memory retrieval middleware consumes enriched context fields for trigger heuristic + entity extraction
