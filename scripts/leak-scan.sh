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

pat=$(printf '%s\n' "$LEAK_DENYLIST" | grep -vE '^[[:space:]]*$' | paste -sd '|' -)
if [ -z "$pat" ]; then echo "✅ leak-scan: empty denylist, nothing to check."; exit 0; fi

# For each candidate hit, strip KEEP terms then re-test the denylist on what
# remains — a hit survives only if a real token is present after KEEP removal.
out=$(files | tr '\n' '\0' | xargs -0 grep -InHE -- "($pat)" 2>/dev/null | while IFS= read -r hit; do
  if printf '%s\n' "$hit" | strip_keep | grep -qE -- "($pat)"; then printf '%s\n' "$hit"; fi
done)
if [ -n "$out" ]; then
  printf '\n❌ leak-scan FOUND private tokens in committed content:\n%s\n\n' "$out"
  echo "Scrub to generic placeholders before merge (CLAUDE.md → Public-repo hygiene)."
  exit 1
fi
echo "✅ leak-scan: clean."
