# App Modernization Workflow

A repeatable pipeline for turning a legacy Access app into a modern web app, built around two repos and one contract.

## The core idea: two repos, one contract

```
┌─────────────────────────────┐        app-spec/ bundle        ┌─────────────────────────────┐
│  REPO A — the Factory        │  ───────────────────────────▶ │  REPO B — the Forge          │
│  (this repo)                 │   (the versioned contract)     │  (Claude-Code-native)        │
│                              │                                │                              │
│  Access .accdb               │   app_spec.json                │  CLAUDE.md  (conventions)    │
│   → extract_access_db.py     │   app_spec.md                  │  .claude/skills/  (skillset) │
│   → inspect_artifacts.py     │   vba/<module>.vba             │  .claude/agents/  (subagents)│
│   → split_vba.py  (new)      │   fixtures/<table>.csv         │  → generates the app         │
└─────────────────────────────┘                                └─────────────────────────────┘
```

- **Repo A is a *factory*:** deterministic scripts that emit a self-contained `app-spec/` bundle. No app code lives here.
- **Repo B is a *forge*:** a reusable template repo whose "skillset" (Skills + subagents + `CLAUDE.md`) knows how to turn *any* `app-spec/` bundle into an app.
- **The `app-spec/` bundle is the API between them.** If it's correct and complete, Repo B is mechanical. This is why the casing bug above matters — fix the contract before building anything that consumes it.

The payoff: the first migration is work; every migration after reuses Repo B's skillset and just swaps in a new bundle.

---

## Phase 1 — Factory: emit the spec bundle (Repo A)

**Goal:** produce one versioned, self-contained `app-spec/` folder per app.

| Step | Command | Produces |
|------|---------|----------|
| Extract | `python migration/extract_access_db.py` | `schema.json`, `forms_vba.json`, table CSVs + SQL |
| Inspect | `python generators/inspect_artifacts.py --db-name <App>` | `app_spec.json`, `app_spec.md` |
| Split VBA *(new)* | `python generators/split_vba.py --db-name <App>` | `app-spec/vba/<module>.vba` (one readable file per module) |
| Bundle *(new)* | `python generators/make_bundle.py --db-name <App>` | assembles `app-spec/` + a `spec_version` stamp |

Why the two new scripts:

- **`split_vba.py`** — `forms_vba.json` stores VBA as escaped one-liners (`\r\n`, binary macros). Nobody — human or Claude — should port from that. Materialize each module as a real `.vba` file so it's diffable, greppable, and citable (`@app-spec/vba/modOrders.vba`).
- **`make_bundle.py`** — gathers `app_spec.*`, `vba/`, and the CSVs (as **golden fixtures**, see the parity loop) into one folder with a version stamp, so Repo B always knows which spec it built from.

### The spec contract — what Repo B is allowed to rely on

| Field in `app_spec.json` | Consumer in Repo B | Guarantee |
|--------------------------|--------------------|-----------|
| `schema.entities[].fields[].csharp_name` / `csharp_type` | model + EF config generation | **Must** be correct casing (see bug fix) |
| `schema.entities[].scope` (`v1_core` / `v1_support` / `later`) | sprint scoping | drives what gets built first |
| `schema.entities[].foreign_keys` | EF relationships, FK-safe seed order | complete & acyclic |
| `forms[].suggested_route` / `mapped_entity` | React route + page scaffolding | nullable; null = no route |
| `vba_modules[].classification` (`domain`/`ui_glue`/`infrastructure`) | what to port vs. discard | only `domain` gets ported |
| `vba/<module>.vba` | the actual port source | full source, human-readable |
| `fixtures/<table>.csv` | parity tests | exact source rows |

**Exit gate:** `app_spec.json` validates, every `v1_core` entity has correct `csharp_name`s, every `domain` module has a `.vba` file, every table has a fixture CSV.

---

## Phase 2 — Architecture Decision Record (bootstrap Repo B)

Don't re-pick a stack each time — the spec already implies one (`inspect_artifacts` emits C# types and React routes). Record the decision once as an ADR and move on. Default:

```
Frontend : React 18 + TypeScript + Vite + TanStack Query
Backend  : C# ASP.NET Core 8 + EF Core 8 + FluentValidation
Database : SQL Server / Fabric SQL (seeded from fixtures)
Tests    : xUnit (backend), Vitest + Playwright (frontend)
```

Then derive scope mechanically from the spec — no guessing:

- **Build set = `scope == "v1_core"`** plus the `v1_support` lookups they FK into.
- **Routes = `forms[].suggested_route` where non-null.**
- **Services to port = `vba_modules[].classification == "domain"`.**

Capture this in Repo B as `docs/ADR-001-stack.md` and `docs/scope.md`. That's the whole planning phase — minutes, not a document.

---

## Phase 3 — Forge: the Claude Code skillset (Repo B)

This is the "skillset for creating apps." Three layers, all checked into Repo B so they're reusable and team-shareable.

### Layer 1 — `CLAUDE.md` (persistent context, auto-loaded)

Auto-loads every session. Encodes the conventions so you never re-explain them:

```markdown
# Modernization Forge

The spec for the app we're building lives in `app-spec/` (treat it as read-only source of truth).
- Entity definitions: @app-spec/app_spec.json
- Human overview + routes: @app-spec/app_spec.md
- Legacy logic to port: app-spec/vba/*.vba
- Golden data for parity tests: app-spec/fixtures/*.csv

## Conventions
- Backend: ASP.NET Core 8, EF Core 8, FluentValidation. One service per entity.
- Frontend: React + TS, TanStack Query for all server state, one page per route.
- Property names & types come from `csharp_name`/`csharp_type` in the spec — never invent them.
- THE PARITY RULE: a feature is not "done" until its parity test passes against app-spec/fixtures.
```

### Layer 2 — Skills (reusable, parameterized commands)

Live in `.claude/skills/<name>/SKILL.md`. Invoke as `/scaffold-entity Companies`. Example:

```markdown
---
name: scaffold-entity
description: Generate the C# model, EF config, service, controller, and tests for one entity in the spec.
argument-hint: "<EntityName>"
arguments: entity
allowed-tools: Read, Write, Edit, Bash(dotnet *)
---
Read `schema.entities.$entity` from @app-spec/app_spec.json.

Following @CLAUDE.md conventions, generate:
1. backend/Models/$entity.cs        — property per field (use csharp_name / csharp_type; [Key] on PKs)
2. backend/Data/$entityConfig.cs    — table=table_name, max_length, FK rels from foreign_keys
3. backend/Services/${entity}Service.cs — CRUD + validation ported from @app-spec/vba/modValidation.vba
4. backend/Controllers/${entity}Controller.cs — REST endpoints per app_spec.md routes
5. backend/Tests/${entity}ServiceTests.cs

Run `dotnet build`; do not finish if it fails.
```

Recommended skill set:

| Skill | Does |
|-------|------|
| `/scaffold-entity <Entity>` | model + EF config + service + controller + tests |
| `/build-page <Form>` | React list or detail page from a form's `suggested_route` |
| `/port-vba <module>` | hands off to the `vba-porter` subagent (below) |
| `/verify-parity <Entity>` | seed from fixtures, assert row counts + spot values |

### Layer 3 — Subagents (isolated specialists)

Live in `.claude/agents/<name>.md`, run in their own context with restricted tools. Two earn their keep:

```markdown
---
name: vba-porter
description: Ports one legacy VBA module to an idiomatic C# service + xUnit tests.
tools: Read, Write, Edit, Glob, Grep, Bash
model: sonnet
---
1. Read app-spec/vba/<module>.vba in full.
2. Keep pure business logic (calculations, validation, workflow/state). Drop UI plumbing (DoCmd, form/control manipulation).
3. Write the service in backend/Services/ using the existing entity models.
4. Write xUnit tests encoding the original rules as cases.
5. Mark every assumption with `// PORT NOTE:`.
Report: functions ported, functions intentionally dropped (+ why), assumptions made.
```

```markdown
---
name: parity-checker
description: Read-only. Verifies generated app data matches app-spec/fixtures golden CSVs.
tools: Read, Glob, Grep, Bash(dotnet test*)
model: haiku
---
For the given entity: seed a test DB from app-spec/fixtures/<table>.csv, then assert
row count matches and 5 sampled rows match field-for-field. Report PASS/FAIL with diffs.
Never modify application code — only report.
```

**Why this split:** Skills = repeatable procedures you trigger; subagents = messy, exploratory work you want quarantined in its own context (porting tangled VBA, validating data) so it doesn't pollute the main session. `CLAUDE.md` = the things both should always know.

---

## The verification loop (this is what makes it trustworthy)

The CSV exports aren't just data — they're **golden fixtures**. Wire them into a loop so "done" is provable, not asserted:

```
scaffold-entity  →  port-vba (if domain logic)  →  build-page  →  verify-parity
                                                                        │
                                              FAIL ◀────────────────────┤
                                               │                        ▼
                                          fix & repeat              PASS → next entity
```

`verify-parity` seeds a throwaway DB from `fixtures/<table>.csv`, hits the new API, and asserts row counts + sampled values against the source. No entity ships until it's green. This catches the casing bug, dropped columns, bad FK wiring, and VBA logic that was ported wrong — automatically.

---

## Sequenced execution (gated)

Run per entity, `v1_core` first, in FK-dependency order (the spec's "Table Insertion Order"):

1. `/scaffold-entity <Entity>` → `dotnet build` green
2. `/port-vba <module>` for any `domain` module that entity needs → tests green
3. `/build-page <Form>` for that entity's routes → app renders
4. `/verify-parity <Entity>` → **gate**: must pass before the next entity
5. Commit as one coherent unit referencing the entity + `spec_version`

Finish all `v1_core`, then `v1_support`, then reassess `later`.

---

## Handing the bundle to Repo B

`app-spec/` is the unit of handoff. Pick one (in order of preference):

- **Release artifact** — `make_bundle.py` zips `app-spec/` with its `spec_version`; Repo B pulls a pinned version. Cleanest; Repo B always knows its source of truth.
- **Git submodule** — Repo B includes Repo A's `app-spec/` as a submodule. Good if specs change often.
- **Copy script** — simplest; a `sync-spec.ps1` that copies the bundle. Fine to start.

Whichever you choose, **stamp `spec_version`** so a generated app can always name the spec it came from.

---

## Checklists

**Phase 1 — Factory exit gate**
- [ ] `_to_pascal` casing bug fixed; `csharp_name`s verified on `v1_core`
- [ ] `app_spec.json` validates; every form mapped (or explicitly null)
- [ ] every `domain` module has a readable `.vba` file
- [ ] every table has a fixture CSV; FK graph is acyclic
- [ ] bundle stamped with `spec_version`

**Phase 3 — per-entity "done"**
- [ ] model/types match spec `csharp_name`/`csharp_type` exactly
- [ ] `domain` logic ported with tests; `// PORT NOTE`s reviewed
- [ ] routes render; CRUD works end to end
- [ ] **`verify-parity` passes** against fixtures
- [ ] committed referencing entity + `spec_version`

---

## What changed from the first draft (and why)

- **Reframed as Factory + Forge + a spec contract** — names the handoff explicitly and makes Repo B reusable across apps instead of a one-off.
- **Flagged the `_to_pascal` casing bug** — the spec is the contract; a corrupt contract poisons every downstream generation.
- **Replaced copy-paste prompts with real Claude Code primitives** — `CLAUDE.md` + Skills + subagents. This *is* the requested "skillset," and it's version-controlled and team-shareable.
- **Added `split_vba.py` + fixtures** — porting from escaped JSON one-liners is untenable; golden CSVs turn "data parity" from a wish into a test.
- **Collapsed the architecture menu into one ADR** — the spec already implies the stack; choosing it every time is wasted motion.
- **Made every entity gated on a passing parity test** — "done" becomes provable.
```
