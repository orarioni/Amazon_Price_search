# Codex project instructions (Amazon_Price_search)

## Runtime / environment
- Target runtime is Windows PowerShell 5.1 unless explicitly stated otherwise.
- Do not assume pwsh (PowerShell 7) is installed.
- Prefer solutions compatible with PS5.1 (no ResponseHeadersVariable on Invoke-RestMethod in PS5.1).

## Repo conventions
- Main update entrypoint: run_update.bat -> scripts/10_update_excel.ps1
- Keep existing log format stable (logs/run.log).
- Do not change Excel column meanings without updating README and scripts consistently.

## Excel COM performance rules
- Avoid redundant COM writes.
- Never write a cell and immediately overwrite it (e.g., set then blank).
- Prefer writing each cell at most once per row.
- Keep column I as "unused/blank" unless the requirements change.

## SP-API behavior
- Assume rate limit headers may be missing; code must be safe without them.
- Handle 429 with backoff; avoid bursty request patterns.
- Avoid selecting multipack/bundle ASINs when resolving JAN if possible; add debug hooks to diagnose.

## Pull request expectations
- Make minimal diffs.
- Add a short comment in code when a guard is preventing a known crash.
- If behavior changes, update README in the same PR.