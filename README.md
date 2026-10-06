# dsh-solidworks

Drive SolidWorks from DeepSeek Harness. Five model-facing tools over the
SolidWorks COM API, designed so that **the open part of the problem stays
open** and the **verified part stays verified**.

## Requirements

| | Requirement | Why |
|---|---|---|
| OS | **Windows only** | The plugin reaches SolidWorks through COM (`cscript` / `WScript` late binding), which exists on no other platform. |
| CAD | A **locally installed SolidWorks** (developed against 2026 SP2.1) | Every tool drives the running desktop application; there is no headless or file-format fallback. |
| Runtime | `cscript.exe` (ships with Windows) and the SolidWorks COM registration for the current user | A SolidWorks install performed for another user, or a portable/registry-free install, is not reachable. |
| Node | Node 18+ to run the plugin and the tests | ESM, `import.meta.dirname`, top-level `await`. |
| Locale/paths | An **ASCII** scratch directory and part template | COM mangles non-ASCII paths, so the plugin refuses them up front. `%TEMP%\dsh-solidworks` is the default and is normally ASCII. |

There is **no Linux/macOS path and no CI path**: verification is render-based
because `GetMassProperties`, `GetBodyBox` and `GetCurves` return `Empty`
through late binding on this host, so a green run requires a real SolidWorks
session with a GUI. Every tool call fails on a machine without SolidWorks, by
design rather than by accident.

| Tool | Role |
|---|---|
| `solidworks_run` | Escape hatch. Runs any VBScript against SolidWorks; handles the ANSI/CRLF encoding trap, the ASCII template, the ASCII scratch dir and a hard timeout. This is why the plugin never narrows what can be modelled. |
| `solidworks_verify` | Validation harness. Reads back what a script actually built (semantic feature kinds, body count, renders) and compares it with an expected shape. **A recipe is only trustworthy once this reports `ok:true`.** |
| `solidworks_capabilities` | Capability probe. Reports which API routes work on *this* machine, cached 6 h, so a strategy can be routed around a broken call instead of discovering it by failing. |
| `solidworks_recipe` | Runs a stored recipe by name, with parameter overrides, and verifies the result against the recipe's own recorded expectation. |
| `solidworks_recipes` | The data surface: list / save / remove recipes. Promotion happens here, as a data write — never a source edit. |

## Architecture: open composition, closed primitives

```
        open composition                     closed primitives
   ┌──────────────────────────┐        ┌──────────────────────────────┐
   │ solidworks_run           │        │ solidworks_verify            │
   │ any VBScript you can     │  ───▶  │ feature kinds + body count   │
   │ write — no ceiling       │        │ + rendered views             │
   └──────────────────────────┘        └──────────────┬───────────────┘
                ▲                                     │ accepted
                │                                     ▼
   ┌────────────┴─────────────┐        ┌──────────────────────────────┐
   │ solidworks_capabilities  │        │ solidworks_recipes.save      │
   │ routes around broken API │        │ freezes it as JSON DATA      │
   └──────────────────────────┘        └──────────────┬───────────────┘
                                                      │ next load
                                                      ▼
                                       ┌──────────────────────────────┐
                                       │ solidworks_recipe <name>     │
                                       │ runs + self-verifies         │
                                       └──────────────────────────────┘
```

The flexibility cost of a plugin is paid in two places, and both are handled
here rather than accepted:

1. **Capability discovery.** `solidworks_capabilities` probes the live API
   surface instead of hard-coding "what works", so the closed set adapts to the
   machine (see `capabilities.sweptCut` below).
2. **Recipe promotion.** A shape the primitives cannot express is written with
   `solidworks_run`, validated with `solidworks_verify`, and only then frozen
   into a reusable recipe — the catalogue grows from real demand instead of
   having to be complete up front.

### Recipes are data, not code

This is the part that makes iteration safe. A recipe is a JSON file:

```json
{
  "name": "flat_disc",
  "description": "Flat disc: one extruded circle on the front plane.",
  "parameters": { "R": { "value": 0.04, "description": "disc radius" },
                  "T": { "value": 0.012, "description": "thickness" } },
  "verify": { "bodyCount": 1, "featureTypes": ["Extrude"] },
  "script": "…Const R = {{R}}\nConst T = {{T}}…"
}
```

Adding a recipe therefore costs one `solidworks_recipes` call, and it can never
take the plugin down:

| | source edit | recipe data |
|---|---|---|
| bad content effect | `apply()` throws → **all five tools disappear** | that one recipe is skipped with a warning |
| activation risk | plugin entry aborts at composition | none; `apply()` still returns |
| iteration latency | needs an app reload (`hmr.root` is opt-in and off) | next load; a `save` is usable immediately |

Substitution is validated before anything runs: the script must be ASCII, every
`{{NAME}}` must have a declared finite-number parameter, and every declared
parameter must actually appear. `%NAME%` is reserved for the Windows environment
variables the scripts already use (`%SWPARTTPL%`, `%SWOUTDIR%`), so it is
rejected as a placeholder with a message naming the correct form.

Where recipes live: user recipes in the plugin's `recipeDir` (default
`<scratch>/recipes`, set it to a stable path to keep them across restarts), plus
the shipped defaults in `lib/recipes.js`. A user recipe with a shipped name
overrides it.

```
solidworks_recipes { action: "list" }
solidworks_recipes { action: "save", name, description, script, parameters, verify }
solidworks_recipe  { name: "threaded_shaft", params: { REVS: 12 } }
```

A recipe run auto-verifies against its own recorded expectation, so a recipe
that stops producing the right shape reports a failure instead of a result.

### Adding a primitive

`solidworks_build` (declarative feature emitters) remains the next step and is
deliberately absent: build it once enough recipes exist to see which emitters
earn their place. The recipe store is the natural feeder for it.

## Runtime facts encoded here

Measured on SolidWorks 2026 SP2.1, zh-CN Windows, late binding through
`cscript`:

- **WSH reads `.vbs` as ANSI.** A UTF-8 source containing non-ASCII text fails
  to parse. Scripts are staged through an ANSI code page with forced CRLF, and
  `solidworks_run` rejects a non-ASCII script with that reason instead of
  producing a confusing parse error.
- **COM mangles non-ASCII paths.** Templates, part files and output
  directories must be ASCII, so work happens in an ASCII scratch directory.
- **Sketch geometry must be `Call`ed.** A bare `sm.CreateLine(...)` statement
  is a silent no-op.
- **`InsertSketch2` / `InsertHelix` / `InsertFeatureChamfer` live on
  `IModelDoc2`**, not on `FeatureManager` or `SketchManager`.
- **`GetMassProperties`, `GetBodyBox` and `GetCurves` return `Empty`** through
  late binding, which is why verification is render-based.
- **The swept-cut route does not complete**: `InsertCutSwept*` reports an
  invalid parameter count and the `CreateDefinition(18)` pipeline fails inside
  `AccessSelections`. Threads are therefore built as a **core body at the minor
  diameter plus a merged swept boss** — how a lathe cuts a thread.

## Install

The `desktop` profile is owned by the Electron application, so it is installed
through the application, not by editing files:

1. Open the Plugins page in the Web sidebar.
2. Install by **absolute path** to this directory
   (`...\.dsh\plugins\dsh-solidworks`). Local paths are read from their own
   `package.json` for the compatibility check, and a path install needs no
   registry.
3. Select the `dsh-solidworks` bundle. With HMR enabled the composition updates
   live; otherwise reload the app.

The same operation is available to an agent through the `plugin_manager` tool
(`install_bundle` with the absolute path, then enable the bundle), which
requires `danger-full-access` or an approval for that call.

Configuration (`cordis.patch.yml`):

| Field | Default | Meaning |
|---|---|---|
| `partTemplate` | auto | `.prtdot` path; empty scans `C:\ProgramData\SOLIDWORKS\<year>\templates` for the newest **ASCII-named** template |
| `scratchDir` | `%TEMP%\dsh-solidworks` | ASCII scratch for staged scripts, renders, reports; must stay ASCII |
| `defaultTimeoutMs` | `900000` | Hard cap per `solidworks_run` |
| `cacheDir` | scratch dir | where `capabilities.json` is cached |
| `recipeDir` | `<scratch>/recipes` | where user recipes are stored; point it at a stable ASCII path to keep them across restarts |

## Usage

```
solidworks_capabilities {}                       # what works on this machine
solidworks_run { script: "<ascii VBScript>" }    # build, export, set anything
solidworks_verify { expectBodyCount: 1,
                    featureTypes: ["Extrude","Helix","SweepBoss"] }
```

`featureTypes` matches the **semantic kind**, never the localized feature name:
`Extrude`, `Revolve`, `Cut`, `Helix`, `SweepBoss`, `SweepCut`, `Chamfer`,
`Fillet`, `Hole`, `Pattern`, `Shell`, `Draft`. Reference planes, sketches and
feature folders are filtered out.

A verified recipe, the thread pattern:

```vbscript
' core body at the minor diameter, extruded circle (axis unambiguous) ...
Const CORE_R = 0.024 : Const MAJ_R = 0.025 : Const TIP_R = 0.026
Const PIT = 0.005 : Const REVS = 40 : Const W_CREST = 0.00125 : Const W_ROOT = 0.00225
' 1. extrude a circle of CORE_R along +Z
' 2. helix from a MAJ_R base circle on the same plane, pitch PIT, REVS turns
' 3. trapezoid on that same plane, spanning CORE_R..TIP_R across the helix start
'    (a multi-contour sketch would need preselection; four CreateLine calls are fine)
' 4. sweep it as a BOSS with Merge = True, so thread and core become one body
Set fTh = fm.InsertProtrusionSwept4(True, False, 0, False, False, 0, 0, False, _
                                    0.01, 0.01, 0, 1, True, False, False, 0.0, _
                                    False, False, 0.01, 1)
```

`solidworks_verify` accepting `bodyCount = 1` with
`["Extrude","Helix","SweepBoss"]` is what turns that script into a recipe.

## Tests

```
node test/schema-guard.mjs           # schemas inside the Harness's enforced subset
node test/drive-tools.mjs            # run + verify + probe against live SolidWorks
node test/promote-thread-recipe.mjs  # the recipe gate
node test/recipes.mjs                # the recipe data layer (save / run / isolation)
```

`schema-guard.mjs` exists because a schema violation does not fail one tool — it
aborts the **whole plugin entry** at composition time with
`JsonSchemaError: unsupported JSON schema`, which the GUI surfaces only as
"Could not enable". Two constraints are easy to get wrong:

- `parameters` and `output.schema` are plain JSON Schema, and **`output.schema`
  is required**: the registry asserts it on register, so omitting it throws
  `schema must be a schema object`. An annotation-only `{}` declares
  unconstrained JSON and is the legal form for a tool returning free-form values.
- `items`, `properties`, `required`, `additionalProperties`, `enum` and `const`
  are constraint keywords and must be **bare schema objects**: annotation
  keywords such as `description` belong on the annotated node, never inside
  `items`.

Drives the three registered tools against the live SolidWorks session without
the host: capability probe, a disc build through `solidworks_run`, a passing
verify, a deliberately failing verify (to prove mismatches surface), and a
non-ASCII script (to prove the encoding guard fires).

## Companion skill

The domain knowledge — the full API surface, verified signatures, recipes and
the failure catalogue — lives in the `solidworks-modeling` skill. The plugin
owns mechanics (encoding, process, probe, verification); the skill owns
knowledge. Load the skill before modelling.
