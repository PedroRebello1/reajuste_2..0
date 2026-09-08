# Repo conventions — read before touching git

## This folder is always `main`
This working directory must always stay checked out on the `main` branch.
Never `git checkout` another branch here, and never leave the folder on a
detached HEAD or another branch, even temporarily. If work on another branch
is needed, use a separate `git worktree` (e.g. under `$env:TEMP`) or a
cherry-pick workflow that never touches the files in this folder, then remove
the worktree when done.

## `main` vs `PcDaniel`
`PcDaniel` is a sibling branch that runs the exact same code and logic as
`main`. The **only** intentional difference between the branches is screen
coordinates, because:
- `main` (this folder) targets a **2-screen** computer.
- `PcDaniel` targets a **1-screen** computer.

The coordinate-only differences live in:
- `autofill_cassi.py` — the `COORD` dict near the top, and the `x_fim`
  variable inside `copiar_texto_mouse`.
- `searcher.py` — the `COORD` dict near the top.

Everything else (functions, control flow, strings, README content, etc.) is
expected to be identical in spirit between the two branches, aside from
PcDaniel's own machine-specific README notes.

### Rule: every push to `main` must also reach `PcDaniel`
Whenever you commit and push logic/code changes to `main`, replicate the same
non-coordinate changes onto `PcDaniel` and push that too — **without**
overwriting PcDaniel's coordinate values or its machine-specific README
notes. Do this without checking out `PcDaniel` in this folder:

1. `git fetch origin`
2. `git worktree add <temp-dir> PcDaniel` (use a temp dir, e.g. under `$env:TEMP`)
3. `git -C <temp-dir> cherry-pick <the main commit(s)>`
   (cherry-picking is safe here because the coordinate lines are never part
   of the logic diff, so there should be no conflicts touching `COORD`/`x_fim`)
4. Verify the `COORD` dicts and `x_fim` in `<temp-dir>` still match PcDaniel's
   original values (diff against the previous PcDaniel commit if unsure).
5. `git -C <temp-dir> push origin PcDaniel`
6. `git worktree remove <temp-dir> --force`

If a change actually needs different behavior per machine (not just a
coordinate), stop and ask the user before applying it to both branches.

## Commit/push identity
All commits and pushes in this project must be authored under one of:
1. `pedrorebellozf@gmail.com` — **preferred, always use this by default.**
2. `e80004428@cassi.com.br` — fallback only if there's a specific reason to
   use the corporate identity.

This repo's local git config (`user.email`) is already set to the preferred
address. On a fresh clone/machine, set it explicitly rather than relying on
git's auto-detected corporate identity:

```
git config user.email "pedrorebellozf@gmail.com"
```
