# Teams Gateway — Adding a Bot

How to stand up a new Microsoft Teams bot that routes DMs to a mesh agent
through this gateway. The gateway is **multi-bot**: one process (port 3024)
hosts any number of bots, each at its own path `/api/<slug>/messages`. In v1,
**one Azure bot = one mesh agent** (the bot's `defaultAgent`); to reach several
agents (e.g. an ops agent, an admin agent) you stand up one Azure bot per agent.

## Mental model — what's "per bot" vs "shared"

| Thing | Scope |
|-------|-------|
| Azure Bot app registration (appId + secret) | **per bot** |
| `TEAMS_BOTS[]` entry in `.env` | **per bot** |
| NPM path allow-rule `/api/<slug>/messages` | **per bot** (one regex, add each slug) |
| Per-bot AMP identity `teams-<slug>-bot@…` + inbox dir | **per bot** (minted on first live register) |
| Teams app manifest / sideload package | **per bot** |
| The gateway process, port 3024, `/health` | **shared** (one for all bots) |
| Your Maestro **user** record (keyed by your Teams tenant+aadObjectId) | **shared** — you are ONE user across every bot |

Corollary: operator trust is set once on your user record and applies to every
bot you DM.

---

## Prerequisites

- Access to the Azure tenant that owns the bots (single-tenant model).
- The gateway host on the tailnet; public ingress via NPM already configured
  (`https://<your-host>/` → `http://<gateway-tailscale-ip>:3024`). See the live
  bring-up runbook for the NPM + tailnet setup.
- `AMP_TENANT` set in `.env` (e.g. `example`) — required on a tenant-scoped
  Maestro, else registration 400s.

---

## Step 1 — Create the Azure Bot

1. **Azure Portal** (`portal.azure.com`) → create an **Azure Bot** resource.
   - **Type of App:** Single Tenant.
   - **Creation type:** create a new Microsoft App ID.
2. Note the **App (client) ID** — this is `appId`.
3. **Microsoft Entra ID → App registrations →** your bot app **→ Certificates &
   secrets → New client secret**. Copy the **Value** (NOT the Secret ID — the
   Value is shown once only). This is `appPassword`.
4. Note the **Tenant ID** (Entra → Overview) — this is `appTenantId`.
5. On the Azure Bot resource → **Channels** → add the **Microsoft Teams** channel.
6. Leave the **Messaging endpoint** for Step 5 (set it AFTER the gateway is live,
   or Azure will POST to a route that 404s until the bot is registered).

## Step 2 — Add the bot to the gateway `.env`

Add an entry to the `TEAMS_BOTS` JSON array in `teams-gateway/.env`:

```bash
TEAMS_BOTS='[{"slug":"leoai","appId":"…","appPassword":"…","appTenantId":"…","defaultAgent":"ops-agent@example.aimaestro.local"},
             {"slug":"myagent","appId":"…","appPassword":"…","appTenantId":"…","defaultAgent":"dev-gateway-agent@example.aimaestro.local"}]'
```

- **`slug`** — lowercase alnum + dashes, must be unique, and NOT reserved
  (`admin`, `health`, `api`). It becomes the public path `/api/<slug>/messages`
  and the AMP identity `teams-<slug>-bot`.
- **`defaultAgent`** — the full mesh address of the agent this bot routes DMs to.
- **`agentName`** (optional) — defaults to `teams-<slug>-bot`.
- ⚠️ **Single-quote the whole value.** `start.sh` does `source .env`, and bash
  strips the inner double-quotes from an unquoted JSON value → the gateway dies
  at boot with `TEAMS_BOTS is not valid JSON`.

## Step 3 — Open the bot's path in NPM

The public location is locked to specific paths. Edit the proxy host's
**Advanced** tab (Nginx custom config) and extend the allow-list regex to include
the new slug:

```nginx
location ~ ^/(?!api/leoai/messages$|api/myagent/messages$|health$) {
    return 404;
}
```

Add `|api/<slug>/messages$` for each bot. (Do NOT use a bare `location / { … }`
in the Advanced tab — NPM already generates one and a duplicate breaks the reload.)

## Step 4 — Restart the gateway (mints the identity)

```bash
pm2 restart teams-gateway --update-env
```

On boot in live mode (`TEAMS_DRY_RUN=0`) the gateway registers each bot's AMP
identity and creates its inbox:

```
[BOOTSTRAP] <slug>: registered teams-<slug>-bot@<org>.aimaestro.local (<agentId>)
[OUTBOUND]  <slug>: /home/<user>/.agent-messaging/agents/<agentId>/messages/inbox
```

`TEAMS_DRY_RUN=1` logs the plan with no network register (useful for a dry boot).
First live registration mints a new directory identity — heads-up the Maestro
owner (Maestro core) so it isn't a surprise.

## Step 5 — Set the Azure messaging endpoint

Azure Portal → the bot's **Azure Bot → Configuration → Messaging endpoint**:

```
https://<your-host>/api/<slug>/messages
```

## Step 6 — Build & sideload the Teams app manifest

Each bot needs its own Teams app package (zip of `manifest.json` + `color.png` +
`outline.png`, files at the zip root). Use `teams-app/manifest.json` as the
template; change `id`, `bots[].botId` (both = the new bot's `appId`), `name`,
`description`, and `validDomains`. Then in Teams: **Apps → Manage your apps →
Upload a custom app** → select the zip.

> If "Upload a custom app" is greyed out: **Teams admin center → Teams apps →
> Setup policies → Global → Upload custom apps = On.**

## Step 7 — Target agent awake

The bot routes DMs to its `defaultAgent`. That agent must be awake on the mesh to
process the message and reply, or nothing comes back.

---

## Verify

```bash
curl -s http://127.0.0.1:3024/health | python3 -m json.tool   # bots[] lists your slug, ampAddress populated
pm2 logs teams-gateway --lines 50                              # watch the round-trip
```

DM the bot in Teams → expect `[TEAMS] (<slug>) routed activity → <defaultAgent>`
inbound, then `[-> Teams] (<slug>) reply from <agent>` outbound.

## Gotchas (learned the hard way)

- **Single-quote `TEAMS_BOTS`** (Step 2) — the #1 boot failure.
- **`AMP_TENANT` is required** on a tenant-scoped Maestro — without it the bot
  tries to register under `default` and Maestro 400s.
- **Client secret Value, not Secret ID** (Step 1.3) — the Value is shown once.
- **Operator trust** is keyed on `(tenantId, aadObjectId)` via
  `OPERATOR_AAD_OBJECT_IDS=<tenantId>:<aadObjectId>` — both halves, tenant-scoped.
- **Reserved slugs:** `admin`, `health`, `api` are rejected (path collisions).
