#!/usr/bin/env bash
#
# Public-repo leak scanner — fails if private mesh / agent / client / operator
# tokens are committed. swickson/ai-maestro is a PUBLIC fork (can't go private),
# so committed content must stay generic. See CLAUDE.md "Public-repo hygiene".
#
# DESIGN (so the guard itself never re-leaks): the denylist of ACTUAL private
# tokens (host names, agent names, client/project names, operator handle + email
# domain, the real Tailscale mesh IPs) is supplied at runtime via $LEAK_DENYLIST
# (newline-separated, regex-allowed), sourced from a CI secret and NEVER
# committed. Committing the token list here would itself re-leak it. We do NOT
# use broad structural patterns (e.g. "any 100.64/10 IP"): the repo legitimately
# documents EXAMPLE Tailscale IPs, so only the SPECIFIC real mesh IPs (in the
# secret) are leaks.
#
# TIERS: the denylist may carry a SOFT marker line (a comment containing the word
# SOFT, e.g. "# --- SOFT (warn-only) ---"). Tokens BEFORE the marker are HARD
# (real mesh structure: IPs / hostnames / secrets) → FAIL the build. Tokens AFTER
# it are SOFT (agent + operator NAMES) → WARN only, non-blocking: a name has no
# attack value (the operator's name is already on every commit; agent handles are
# invented), so we surface it as a reminder without forcing lossy comment-scrubs
# on every PR. A denylist with NO marker is treated as ALL-HARD (backward compat).
#
# Local pre-PR:  LEAK_DENYLIST="$(cat ~/.aimaestro/leak-denylist.txt)" scripts/leak-scan.sh
# CI:            LEAK_DENYLIST from a repo secret (.github/workflows/leak-scan.yml)
set -uo pipefail
cd "$(git rev-parse --show-toplevel)" || exit 2

# KEEP — public terms that are NOT leaks. We STRIP these from each candidate
# line before re-testing the denylist, so a real token sharing a line with a
# KEEP term is still caught (dropping whole KEEP lines would be a false-negative
# hole — credit: review of this scanner):
#   swickson/<repo>       GitHub org in URLs/remotes/submodule
#   23blocks<...>         product / upstream brand
#   '<Name>IA' / '<Name>AI'  agent-name-generation alias-pool entries (generic
#                            options in FEMALE_ALIASES/MALE_ALIASES — not refs)
strip_keep() {
  sed -E "s@swickson/[A-Za-z0-9._-]+@@g; s/23blocks[A-Za-z0-9-]*//g; s/'[A-Za-z]+(IA|AI)'//g"
}

# Exclude vendored/build/memory + this scanner's own files (they reference the
# mechanism, not real data).
files() { git ls-files | grep -viE 'node_modules|\.next/|package-lock|yarn\.lock|/memory/|scripts/leak-scan\.sh|\.github/workflows/leak-scan\.yml'; }

if [ -z "${LEAK_DENYLIST:-}" ]; then
  echo "⚠️  LEAK_DENYLIST not set — cannot enforce the private-token denylist."
  echo "    Set the LEAK_DENYLIST CI secret (or export it locally from ~/.aimaestro/leak-denylist.txt)."
  echo "    Skipping (non-blocking) so unconfigured forks/CI don't fail spuriously."
  exit 0
fi

# Split the denylist into HARD / SOFT tiers on the SOFT marker line (a comment
# line containing the word SOFT). Tokens before the marker → HARD; after → SOFT.
# No marker → everything HARD (backward compatible). Comment (#) and blank lines
# are dropped from the patterns either way.
soft_marker_re='^[[:space:]]*#.*SOFT'
clean_patterns() { grep -vE '^[[:space:]]*(#|$)' | paste -sd '|' -; }
hard_pat=$(printf '%s\n' "$LEAK_DENYLIST" | awk -v m="$soft_marker_re" '$0 ~ m {f=1; next} !f' | clean_patterns)
soft_pat=$(printf '%s\n' "$LEAK_DENYLIST" | awk -v m="$soft_marker_re" '$0 ~ m {f=1; next}  f' | clean_patterns)

# Emit "file:line:content" hits where a denylist token SURVIVES the KEEP-strip.
scan() {
  local pat="$1"
  [ -z "$pat" ] && return 0
  files | tr '\n' '\0' | xargs -0 grep -InHE -- "($pat)" 2>/dev/null | while IFS= read -r hit; do
    if printf '%s\n' "$hit" | strip_keep | grep -qE -- "($pat)"; then printf '%s\n' "$hit"; fi
  done
}

hard_hits=$(scan "$hard_pat")
soft_hits=$(scan "$soft_pat")

# SOFT — warn only, never blocks.
if [ -n "$soft_hits" ]; then
  printf '\n⚠️  leak-scan WARN — agent/operator NAME tokens in committed content (cosmetic, non-blocking):\n%s\n' "$soft_hits"
  echo "These are warn-tier (names, not mesh structure). Genericize if convenient — not required to merge."
fi

# HARD — fail.
if [ -n "$hard_hits" ]; then
  printf '\n❌ leak-scan FOUND HARD private tokens (mesh IPs / hostnames / secrets) in committed content:\n%s\n\n' "$hard_hits"
  echo "Scrub to generic placeholders before merge (CLAUDE.md → Public-repo hygiene)."
  exit 1
fi

if [ -n "$soft_hits" ]; then
  echo "✅ leak-scan: no HARD tokens (soft name-warnings above are non-blocking)."
else
  echo "✅ leak-scan: clean."
fi
exit 0
