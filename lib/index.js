/**
 * @deepseek-ai/dsh-solidworks — Drive SolidWorks from DeepSeek Harness.
 *
 * Architecture ("open composition, closed primitives"):
 *
 *   solidworks_run     the escape hatch. Runs any VBScript against the
 *                      SolidWorks COM API. Handles the ANSI/CRLF encoding
 *                      trap, the ASCII template, the ASCII scratch dir, and
 *                      a hard timeout. This is why the plugin never narrows
 *                      what can be modelled.
 *
 *   solidworks_verify  the validation harness. Reads back what a script
 *                      actually built (feature tree, body count, renders)
 *                      and compares it against an expected shape. This is
 *                      what makes a new primitive trustworthy: new geometry
 *                      only becomes a "closed primitive" after verify accepts
 *                      it here.
 *
 *   solidworks_capabilities
 *                      the capability probe. Reports which API routes work
 *                      on this specific machine, so a strategy can be routed
 *                      around a broken call instead of discovering it by
 *                      failing. Cached on disk with a TTL.
 *
 *   solidworks_build   the closed primitive catalogue: declarative features
 *                      (extrude, cut, revolve, chamfer, helix, swept boss)
 *                      that emit verified API calls. New shapes are added by
 *                      dropping an emitter in features/ — not by editing this
 *                      file — and by promoting a script that verify accepted.
 *
 * Runtime facts this plugin encodes (measured on SolidWorks 2026 SP2.1,
 * zh-CN Windows, late binding through cscript):
 *
 *   - WSH reads .vbs as ANSI: a UTF-8 source with non-ASCII text fails to
 *     parse. Every script is re-encoded to the ANSI code page, CRLF forced.
 *   - COM mangles non-ASCII paths: templates, part files and output dirs must
 *     be ASCII, so work happens in an ASCII scratch dir and is copied out.
 *   - Sketch geometry must be invoked with `Call`: a bare
 *     `sm.CreateLine(...)` statement is silently a no-op.
 *   - InsertSketch2 / InsertHelix / InsertFeatureChamfer live on ModelDoc2.
 *   - Feature.GetCurves, Body2.GetBodyBox and ModelDoc2.GetMassProperties
 *     return Empty through late binding, so geometry is verified by
 *     rendering and reading the image, not by querying the model.
 *   - The swept-cut route (InsertCutSwept* and the CreateDefinition(18)
 *     pipeline) does not complete on this host, so threads are built as a
 *     core body plus a merged swept BOSS.
 *
 * @module @deepseek-ai/dsh-solidworks
 */

import { spawnSync } from 'node:child_process'
import {
  existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync,
} from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { randomUUID } from 'node:crypto'

import {
  describeRecipes, loadShipped, loadUserRecipes, materialize, mergeRecipes,
  removeRecipeFile, saveRecipeFile, validateRecipe,
} from './recipes.js'

export const name = 'solidworks'

/** Iteration marker, surfaced in capability reports so a live reload is observable. */
const PLUGIN_REVISION = 'recipes-8'
export const inject = ['tools']

/** Directory holding this module, so bundled scripts resolve wherever installed. */
const here = dirname(fileURLToPath(import.meta.url))
const assetDir = resolve(here, '..', 'assets')

const DEFAULT_TIMEOUT_MS = 900_000
const CAPABILITY_TTL_MS = 6 * 60 * 60 * 1000

/**
 * Stable machine-readable failure kinds. A caller (or the model) can branch on
 * these instead of pattern-matching prose: `no-template` and `cscript-missing`
 * are environment faults worth changing configuration for, `timeout` suggests a
 * modal dialog, `script-error` means the script itself failed, and the rest are
 * caller mistakes. `error` stays human-readable beside the code; `code` is the
 * contract, so these strings must not drift — `test/error-codes.mjs` pins them.
 */
const ERROR_CODES = {
  NO_SOLIDWORKS: 'no-solidworks',
  NO_TEMPLATE: 'no-template',
  NON_ASCII_SCRIPT: 'non-ascii-script',
  CSPRIPT_MISSING: 'cscript-missing',
  TIMEOUT: 'timeout',
  SCRIPT_ERROR: 'script-error',
  PROBE_FAILED: 'probe-failed',
  INSPECT_FAILED: 'inspect-failed',
  VERIFICATION_FAILED: 'verification-failed',
  RECIPE_UNKNOWN: 'recipe-unknown',
  RECIPE_INVALID: 'recipe-invalid',
  INVALID_INPUT: 'invalid-input',
  UNKNOWN: 'unknown',
}

/**
 * Where SolidWorks records its installed templates. `resolveTemplate` scans the
 * newest subdirectory here, and the capability cache is scoped by the same
 * value so a different install cannot read another install's probe result.
 */
const SW_VERSION_DIR = 'C:\\ProgramData\\SOLIDWORKS'

/**
 * The SolidWorks install a capability report belongs to, for cache scoping.
 *
 * A probe result describes the API surface of ONE installed release, so it must
 * never be served to a different one: an upgrade or a Service Pack can flip
 * exactly the routes the probe measures (`sweptCut`, the fillet option bits,
 * `InsertFeatureShell`), and the 6-hour TTL alone would hand the new install a
 * report measured on the old one. The key is derived from the *installed*
 * version rather than from the probe itself, because reading the probe's own
 * `RevisionNumber` would mean launching SolidWorks — which is the cost the
 * cache exists to avoid.
 *
 * Year granularity is deliberate: the path carries the year and nothing finer.
 * A Service Pack within the same year is therefore still covered only by the
 * TTL, so use `solidworks_capabilities { force: true }` after installing one.
 *
 * @returns a short key such as `2026`, or `unknown` when no install is locatable.
 */
function swVersionKey(templatePath) {
  // The path is the cheapest authority. The guard is the `SOLIDWORKS` label
  // followed by a year, not a bare four-digit number, so a year inside a
  // template file name (…\gb_part 2020.prtdot) cannot be mistaken for the
  // install year. The label is separated from the year by a space or a
  // separator: real installs write "C:\ProgramData\SOLIDWORKS\SOLIDWORKS
  // 2026\templates\…", and a template path may also end at the year itself.
  const marked = /SOLIDWORKS[ _-]{0,2}(\d{4})(?=[\\/]|$)/i.exec(String(templatePath ?? ''))
  if (marked) return marked[1]
  // The same roots `resolveTemplate` scans, so the scope and the template can
  // never disagree about which install they describe.
  for (const root of swInstallRoots()) {
    try {
      const newest = readdirSync(root, { withFileTypes: true })
        .filter((entry) => entry.isDirectory() && /^\d{4}$/.test(entry.name))
        .map((entry) => entry.name)
        .sort()
        .reverse()[0]
      if (newest) return newest
    } catch {
      // Unreadable root: try the next candidate rather than giving up.
    }
  }
  return 'unknown'
}

// ---------------------------------------------------------------- helpers

function asciiOnly(text) {
  return typeof text === 'string' && /^[\x20-\x7E]*$/.test(text)
}

function ensureDir(path) {
  mkdirSync(path, { recursive: true })
  return path
}

/**
 * Snapshot directory. Kept ASCII-only because every path that crosses COM
 * must be; the caller's requested output is copied out afterwards.
 */
function scratchDir(config, sessionKey) {
  const base = config.scratchDir && asciiOnly(config.scratchDir)
    ? config.scratchDir
    : join(process.env.TEMP ?? process.env.TMP ?? '.', 'dsh-solidworks')
  const dir = join(base, sessionKey)
  // Prune only when this invocation opens a NEW directory, so the sweep runs
  // once per invocation instead of on every path resolution inside one.
  const fresh = !existsSync(dir)
  ensureDir(join(base, sessionKey))
  if (fresh) pruneScratch(config, sessionKey)
  return dir
}

/** How many per-invocation scratch directories to keep before pruning. */
const SCRATCH_RETENTION = 40
/** Never pruned: they are stable data, not per-invocation scratch. */
const SCRATCH_KEEP = new Set(['recipes'])

/**
 * Bound the scratch directory.
 *
 * Isolation is per invocation, so directories accumulate; this keeps the newest
 * `SCRATCH_RETENTION` and removes the rest, oldest first. `recipes` and the
 * `capabilities-*.json` cache files live in the same base and are never touched
 * (recipes by name, the cache because it is a file rather than a directory).
 * Failures are swallowed: pruning is housekeeping and must not fail a tool run.
 */
function pruneScratch(config, protect) {
  const base = config.scratchDir && asciiOnly(config.scratchDir)
    ? config.scratchDir
    : join(process.env.TEMP ?? process.env.TMP ?? '.', 'dsh-solidworks')
  try {
    const candidates = readdirSync(base, { withFileTypes: true })
      .filter((entry) => entry.isDirectory() && !SCRATCH_KEEP.has(entry.name) && entry.name !== protect)
      .map((entry) => ({ name: entry.name, mtime: statSync(join(base, entry.name)).mtimeMs }))
      .sort((a, b) => b.mtime - a.mtime)
    for (const stale of candidates.slice(SCRATCH_RETENTION)) {
      rmSync(join(base, stale.name), { recursive: true, force: true })
    }
  } catch {
    // no scratch base yet, or a directory vanished mid-scan
  }
}

/**
 * Roots that may hold SolidWorks installs, newest-last-undetermined. The env
 * override exists for a deployment that moved the install root and for the test
 * that pins a fake install; `resolveTemplate` and `swVersionKey` read the same
 * helper so a capability report can never be scoped to a different install than
 * the template it was produced with.
 */
function swInstallRoots() {
  return [process.env.SW_VERSION_DIR, SW_VERSION_DIR].filter((p) => typeof p === 'string' && p.length > 0)
}

/** Resolve the part template: configured path first, then ASCII scan. */
function resolveTemplate(config) {
  if (config.partTemplate && existsSync(config.partTemplate)) return config.partTemplate
  const roots = swInstallRoots().filter(existsSync)
  for (const root of roots) {
    let years = []
    try {
      years = readdirSafe(root)
        .filter((entry) => entry.isDirectory())
        .map((entry) => entry.name)
        .sort()
        .reverse()
    } catch {
      years = []
    }
    for (const year of years) {
      const tplDir = join(root, year, 'templates')
      if (!existsSync(tplDir)) continue
      const tpls = readdirSafe(tplDir).filter((e) => e.isFile() && e.name.endsWith('.prtdot'))
      if (tpls.length === 0) continue
      const ascii = tpls.find((e) => asciiOnly(e.name))
      return join(tplDir, (ascii ?? tpls[0]).name)
    }
  }
  return ''
}

function readdirSafe(dir) {
  try {
    return readdirSync(dir, { withFileTypes: true })
  } catch {
    return []
  }
}

/**
 * Re-encode a UTF-8 VBScript source into the ANSI code page and force CRLF,
 * which is the only form Windows Script Host parses on a zh-CN host.
 */
function stageAnsi(source, destPath) {
  const text = String(source).replace(/\r\n/g, '\n').replace(/\n/g, '\r\n')
  // latin1 round-trip would corrupt CJK; write bytes for the ANSI code page
  // via Buffer with 'binary' semantics is not enough, so use the host's
  // built-in conversion through writeFileSync with 'latin1' only for ASCII.
  // Instead: PowerShell is avoided, so encode manually for cp936-safe text by
  // keeping non-ASCII out of the scripts we stage. Scripts authored inline by
  // the model are ASCII by contract (the tool description says so).
  const nonAscii = [...text].filter((ch) => ch.charCodeAt(0) > 0x7e)
  if (nonAscii.length > 0) {
    return { ok: false, reason: 'script contains non-ASCII characters; WSH reads .vbs as ANSI, so keep script text ASCII (comments included)' }
  }
  writeFileSync(destPath, Buffer.from(text, 'latin1'))
  return { ok: true }
}

function decodeJson(path) {
  try {
    const raw = readFileSync(path, 'utf8')
    return { ok: true, value: JSON.parse(raw) }
  } catch (error) {
    return { ok: false, reason: error instanceof Error ? error.message : String(error) }
  }
}

function tail(text, limit = 4000) {
  if (typeof text !== 'string') return ''
  return text.length <= limit ? text : `— ${text.length - limit} chars truncated)\n${text.slice(-limit)}`
}

/**
 * Classify a failure string into one of `ERROR_CODES`.
 *
 * The signals are the ones the runtime actually produces: `spawnSync` reports a
 * missing executable as `ENOENT`, the probe reports a COM failure by name, and
 * everything else that reached a live interpreter is a script-level failure.
 *
 * @param text - the failure message, or '' when there is no message (a plain
 *               non-zero exit still means the script failed).
 * @returns one of the `ERROR_CODES` values, and `script-error` for any
 *          failure that cannot be attributed more precisely.
 */
function classify(text) {
  const value = typeof text === 'string' ? text : ''
  const lower = value.toLowerCase()
  if (lower.includes('enoent') || lower.includes('cannot find')) return ERROR_CODES.CSPRIPT_MISSING
  if (lower.includes('cannot create sldworks')) return ERROR_CODES.NO_SOLIDWORKS
  if (lower.includes('non-ascii') || lower.includes('not ascii')) return ERROR_CODES.NON_ASCII_SCRIPT
  if (lower.includes('template')) return ERROR_CODES.NO_TEMPLATE
  return ERROR_CODES.SCRIPT_ERROR
}

/** `{ error, code }` — always both, so the code is part of the contract. */
function errorResult(error, code) {
  return { error: String(error ?? ''), code: code ?? classify(String(error ?? '')) }
}

/** Prepend the code to an error message, so a human log line carries it too. */
function withCode(code, message) {
  return `${code}: ${message}`
}

/** The one-line failure rendering shared by every tool. */
function errorText(label, value) {
  return `${label}: ${value.code ? `[${value.code}] ` : ''}${value.error}`
}

// ------------------------------------------------------------------- tools

/** Resolve a bundled asset (probe.vbs / inspect.vbs) to an absolute path. */
function assetPath(file) {
  const path = join(assetDir, file)
  if (!existsSync(path)) throw new Error(`missing bundled asset: ${path}`)
  return path
}
/**
 * The scratch key for one tool invocation. `callId` is the host's per-call
 * identity; without it (a direct `apply()` call from a test or a script) a
 * random key is still better than a fixed one, because the whole point is that
 * two invocations never share a workspace.
 */
function invocationKey(exec) {
  const callId = exec?.callId
  return typeof callId === 'string' && callId.length > 0 ? callId : `local-${randomUUID()}`
}

function buildTools(config, log, instanceKey) {
  /**
   * Paths + env shared by every script run.
   *
   * The scratch directory is keyed per INVOCATION, not per session: SolidWorks
   * has one active document, so two runs cannot usefully share a workspace even
   * within one session, and a fixed `inspect.json` in a shared directory let one
   * verification read another invocation's report. `sessionKey` is therefore
   * the per-call id from `invocationKey`.
   */
  function environment(sessionKey) {
    const dir = scratchDir(config, sessionKey)
    const template = resolveTemplate(config)
    return {
      dir,
      template,
      env: {
        ...process.env,
        SWPARTTPL: template,
        SWOUTDIR: `${dir}\\`,
      },
    }
  }

  /**
   * Scratch environment for one tool invocation. The host hands the tool body an
   * `exec` carrying the session agent and a `callId`; a caller without either —
   * a direct `apply()` from a test — still gets per-invocation isolation rather
   * than a directory shared with every other call.
   */
  const environmentFor = (exec) => environment(invocationKey(exec))

  /** Run an already-ANSI-encoded script through cscript with a hard timeout. */
  function runCscript(vbsPath, args, env, timeoutMs) {
    const result = spawnSync('cscript.exe', ['//nologo', vbsPath, ...args], {
      env,
      encoding: 'utf8',
      timeout: timeoutMs,
      windowsHide: true,
      maxBuffer: 32 * 1024 * 1024,
    })
    if (result.error) {
      const code = result.error.code
      if (code === 'ETIMEDOUT') {
        return { timedOut: true, stdout: result.stdout ?? '', stderr: result.stderr ?? '', status: null }
      }
      // ENOENT is the usual "no interpreter on PATH"; EPERM is what a sandboxed
      // or locked-down host reports instead. Both mean the same thing to a
      // caller — no script can run on this machine — so they classify
      // identically rather than sending the caller after the script.
      const failed = `${code ?? ''} ${result.error.message}`.trim()
      const missing = code === 'ENOENT' || code === 'EPERM' || /cannot find/i.test(result.error.message ?? '')
      return { failed, missing, stdout: result.stdout ?? '', stderr: result.stderr ?? '', status: null }
    }
    return { stdout: result.stdout ?? '', stderr: result.stderr ?? '', status: result.status }
  }

  /**
   * Stage and run a script. Shared by solidworks_run / _verify / _capabilities.
   *
   * The staged file name carries a monotonic counter: Date.now() alone collides
   * when two tool calls land in the same millisecond, and the loser would run
   * the other call's script.
   */
  let stagedSeq = 0
  function executeScript(source, scriptArgs, exec, timeoutMs) {
    const { dir, template, env } = environmentFor(exec)
    const staged = join(dir, `staged_${Date.now()}_${stagedSeq++}.vbs`)
    const stagedResult = stageAnsi(source, staged)
    if (!stagedResult.ok) return { ...errorResult(stagedResult.reason, ERROR_CODES.NON_ASCII_SCRIPT), dir, template }
    const run = runCscript(staged, scriptArgs, env, timeoutMs)
    if (run.missing) return { ...errorResult(withCode(ERROR_CODES.CSPRIPT_MISSING, `cannot launch cscript: ${run.failed}`), ERROR_CODES.CSPRIPT_MISSING), dir, template }
    if (run.failed) return { ...errorResult(withCode(ERROR_CODES.SCRIPT_ERROR, `cannot launch cscript: ${run.failed}`), ERROR_CODES.SCRIPT_ERROR), dir, template, log: run.stdout }
    return { ...run, dir, template, staged }
  }

  /**
   * Inspect the active document and compare it against expectations.
   *
   * Shared by `solidworks_verify` and by every promoted recipe, so a recipe run
   * is held to exactly the same gate a hand-driven build is.
   */
  function verifyModel(expect = {}, exec = {}) {
    const { dir } = environmentFor(exec)
    const reportPath = join(dir, 'inspect.json')
    try {
      rmSync(reportPath, { force: true })
    } catch {
      // a stale report cannot be trusted, so failing later is correct
    }
    const script = readFileSync(assetPath('inspect.vbs'), 'utf8')
    const run = executeScript(
      script,
      [reportPath, dir, expect.views ?? '7,5', expect.includeReference === true ? '1' : '0'],
      exec,
      300_000,
    )
    if (run.error) return { ...errorResult(run.error, run.code), log: run.stdout ?? '' }
    const parsed = decodeJson(reportPath)
    if (!parsed.ok) {
      return {
        ...errorResult(
          withCode(ERROR_CODES.INSPECT_FAILED, `inspect produced no readable JSON (${parsed.reason})`),
          ERROR_CODES.INSPECT_FAILED,
        ),
        log: tail(run.stdout),
        exitCode: run.status,
      }
    }
    const report = parsed.value
    if (report.error) {
      return {
        ...errorResult(report.error, /no part template/i.test(String(report.error)) ? ERROR_CODES.NO_TEMPLATE : undefined),
        log: tail(run.stdout),
      }
    }

    const mismatches = []
    const actualKinds = (report.features ?? []).map((f) => f.kind).filter(Boolean)
    if (expect.expectBodyCount !== undefined && report.bodyCount !== expect.expectBodyCount) {
      mismatches.push(`solid bodies: found ${report.bodyCount}, expected ${expect.expectBodyCount}`)
    }
    for (const wanted of expect.featureTypes ?? []) {
      if (!actualKinds.includes(wanted)) {
        mismatches.push(`missing feature kind ${wanted} (found: ${actualKinds.join(', ') || 'none'})`)
      }
    }
    if (report.bodyCount === 0) mismatches.push('the document has no solid body at all')

    return {
      ok: mismatches.length === 0,
      title: report.title ?? '',
      bodyCount: report.bodyCount,
      expectedBodyCount: expect.expectBodyCount,
      featureKinds: actualKinds,
      features: report.features ?? [],
      boundingBox: report.boundingBox ?? null,
      renders: report.renders ?? [],
      mismatches,
      log: tail(run.stdout),
    }
  }

  return {
    environment,
    environmentFor,
    executeScript,
    runCscript,
    verifyModel,
  }
}

// ------------------------------------------------------------------ apply

/**
 * Register the SolidWorks tools on `ctx.tools`.
 *
 * @param ctx - registrant context carrying the tool registry and logger.
 * @param config - deployment configuration (paths, timeouts, defaults).
 */
export function apply(ctx, config = {}) {
  const cfg = {
    // The env override is what test/verify-active.mjs and test/rebuild-via-plugin.mjs
    // document (SW_TEST_TEMPLATE), and it is also how a caller pins the install a
    // capability report belongs to. Resolved once, at registration, so a run
    // cannot switch installs midway.
    partTemplate: config.partTemplate || process.env.SW_TEST_TEMPLATE || '',
    scratchDir: config.scratchDir ?? '',
    defaultTimeoutMs: config.defaultTimeoutMs ?? DEFAULT_TIMEOUT_MS,
    cacheDir: config.cacheDir ?? '',
    recipeDir: config.recipeDir ?? '',
  }
  const log = (message) => {
    try {
      ctx.logger?.info?.(`[solidworks] ${message}`)
    } catch {
      // logging must never break a tool call
    }
  }
  /**
   * Identity of THIS plugin instance, used for the two things that must stay
   * stable across a session rather than per call: the capability cache (a 6-hour
   * TTL that would never be hit in a fresh directory) and the user recipe store.
   * Two sessions therefore share one recipe catalogue and one probe cache, which
   * is what makes `save` usable from the next turn — the isolation is applied
   * per invocation instead, where the shared-SolidWorks hazard actually lives.
   */
  const instanceKey = randomUUID()
  const { environment, environmentFor, executeScript, verifyModel } = buildTools(cfg, log, instanceKey)

  /**
   * Where user recipes live. ASCII only, because a recipe path is echoed into
   * logs and the scratch directory; recipes are plain JSON, never executed.
   */
  const recipeDir = cfg.recipeDir && asciiOnly(cfg.recipeDir)
    ? cfg.recipeDir
    : join(environment(instanceKey).dir, 'recipes')

  /**
   * Validate a tool definition with the Harness's own `defineTool`, which
   * asserts the enforced JSON-Schema subset and the output contract.
   *
   * Authoring the definitions by hand means an unsupported keyword (for
   * example a `description` inside array `items`, which is a constraint
   * keyword and must stay a bare schema object) would otherwise only surface
   * as `JsonSchemaError` when the composition mounts — a whole-profile
   * activation failure. Resolving `defineTool` from the host here turns that
   * into a precise, per-tool load-time error, and falls back to the raw
   * definition when the peer is not resolvable.
   */
  let hostDefineTool
  const validated = (definition) => {
    if (hostDefineTool === undefined) {
      try {
        hostDefineTool = createRequire(import.meta.url)('@deepseek-ai/dsh-tools').defineTool
      } catch (error) {
        log(`defineTool unavailable, registering unvalidated: ${error instanceof Error ? error.message : String(error)}`)
        hostDefineTool = null
      }
    }
    if (typeof hostDefineTool !== 'function') return definition
    try {
      return hostDefineTool(definition)
    } catch (error) {
      throw new Error(`tool ${definition.name} has an invalid definition: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  const register = (definition) => {
    const tool = validated(definition)
    try {
      ctx.tools.register(tool)
      log(`registered tool ${tool.name ?? definition.name}`)
    } catch (error) {
      log(`tool registration failed for ${definition.name}: ${error instanceof Error ? error.message : String(error)}`)
      throw error
    }
  }

  // ------------------------------------------------------------ assets
  const asset = assetPath


  // ------------------------------------------------- capability probing
  /**
   * The install whose API surface a capability report describes. Every reader
   * (cache path, tool return, renderer) goes through here, so they cannot drift
   * apart and disagree about which release a cached report belongs to.
   */
  const swVersionScope = () => swVersionKey(cfg.partTemplate || resolveTemplate(cfg))

  /**
   * Scope the cached probe to the installed release: one install's API surface
   * is not another's, and the TTL alone would leak a report across an upgrade.
   */
  const cachePath = () => {
    const base = cfg.cacheDir && asciiOnly(cfg.cacheDir)
      ? cfg.cacheDir
      : environment(instanceKey).dir
    return join(ensureDir(base), `capabilities-${swVersionScope()}.json`)
  }

  function probeCapabilities({ force = false, timeoutMs = 120_000, exec } = {}) {
    const path = cachePath()
    if (!force && existsSync(path)) {
      const age = Date.now() - statSync(path).mtimeMs
      if (age < CAPABILITY_TTL_MS) {
        const cached = decodeJson(path)
        if (cached.ok) return { ...cached.value, fromCache: true, path }
      }
    }
    const script = readFileSync(assetPath('probe.vbs'), 'utf8')
    const run = executeScript(script, [path], exec, timeoutMs)
    if (run.error) return { ...errorResult(run.error, run.code), log: run.stdout ?? '' }
    const parsed = decodeJson(path)
    if (!parsed.ok) {
      return {
        ...errorResult(
          withCode(ERROR_CODES.PROBE_FAILED, `probe produced no readable JSON (${parsed.reason})`),
          ERROR_CODES.PROBE_FAILED,
        ),
        log: tail(run.stdout),
      }
    }
    return { ...parsed.value, fromCache: false, path, log: tail(run.stdout) }
  }

  /** Routing guidance derived from the capability report. */
  function routingAdvice(caps) {
    const advice = []
    const swept = caps?.capabilities?.sweptCut
    if (swept === true) {
      advice.push('swept cut: CreateDefinition(18) returned an object, but the full pipeline (AccessSelections + CreateFeature) has failed to complete on this host — prefer the core-body + swept-boss thread strategy unless you verify otherwise.')
    } else if (swept === false) {
      advice.push('swept cut: unavailable (CreateDefinition(18) returned nothing). Do NOT use InsertCutSwept* — build threads as a core body plus a merged swept BOSS (InsertProtrusionSwept4).')
    }
    if (caps?.api?.doc_get_mass_properties?.ok === false) {
      advice.push('GetMassProperties returns Empty: verify geometry with renders and feature-tree checks, and report dimensions from the script constants — do not try to read mass properties. For a real measurement use ext.CreateMassProperty2 and read .Volume (SI, m^3): do NOT call AddBodies (fails here) and Density is read-only.')
    }
    if (caps?.api?.doc_get_bodies2?.ok === false) {
      advice.push('GetBodies2 returns Empty on a fresh part (documents with a solid body do report bodies); treat a body count read as best-effort.')
    }

    // The three silent-failure routes. Each returns Nothing with Err=0, or
    // aborts the script outright, so a build that does not know about them
    // reports success while producing the wrong part — which is why they are
    // probed and surfaced here rather than rediscovered each session.
    const fillet = caps?.capabilities?.filletOptionsBit2Required
    if (fillet === false) {
      advice.push('fillet: FeatureFillet3 returned Nothing for EVERY Options value tested — do not rely on preselected edges + FeatureFillet3 here; fall back to a chamfer, or try FeatureFillet2 / the 7-arg FeatureFillet in a throwaway document.')
    } else if (fillet === true) {
      advice.push(`fillet: works, but ONLY with the uniform-radius option bit — use Options ${caps?.capabilities?.filletWorkingOption ?? 3} (swFeatureFilletUniformRadius=2, optionally | swFeatureFilletPropagate=1). Options 0, 1 and 4 return Nothing with Err=0.`)
    }
    if (caps?.capabilities?.negZFaceCut === false) {
      advice.push('cuts: FeatureCut4 returns Nothing when the sketch sits on a face whose outward normal is -Z (the back face of a +Z extrusion). Put every cut sketch on the Front plane or on a +Z-facing face.')
    }
    const broken = []
    if (caps?.capabilities?.insertFeatureShellCreatesFeature === false) broken.push('IModelDoc2.InsertFeatureShell')
    if (caps?.capabilities?.insertAxis2CreatesFeature === false) broken.push('IFeatureManager.InsertAxis2')
    if (broken.length > 0) {
      advice.push(`${broken.join(' and ')} did NOT create a feature here (they return Nothing/Empty, and InsertAxis2 can return a sketch-like object that fails on use). Build the shape another way: several extrudes instead of a shell, one multi-contour sketch/cut instead of a circular pattern.`)
    }
    // Applies to every build, not just the broken routes above: this is the
    // failure mode that loses features silently.
    advice.push('put `On Error Resume Next` at the TOP OF EVERY Sub/Function that calls the API: a file-level handler does NOT apply inside a procedure, so an error raised there aborts that procedure while the statements after the call in the caller keep running — a build silently ends up with everything after the failing feature missing.')
    return advice.filter((line) => typeof line === 'string' && line.length > 0)
  }

  // -------------------------------------------------------------- tools
  register(validated({
    name: 'solidworks_capabilities',
    description: [
      'Report what the SolidWorks COM API actually supports on this machine, cached for 6 hours.',
      'Use it before choosing a modelling strategy for anything unusual (threads, sweeps, lofts, patterns): the plugin routes around API routes that are broken here instead of letting a build fail silently.',
      'It may open one throwaway scratch part and closes only that part; your documents are left alone.',
    ].join(' '),
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        force: { type: 'boolean', description: 'Ignore the cache and probe again.' },
      },
    },
    isConcurrencySafe: () => true,
    output: {
      // An annotation-only schema declares unconstrained JSON, which is what
      // this tool returns (the probe report, or `{ error, log }`). The host
      // registry asserts output.schema on register, so omitting it fails the
      // whole plugin activation with "schema must be a schema object".
      schema: {},
      render: (_args, value) => [{
        type: 'text',
        text: value?.error
          ? [
            errorText('SolidWorks capability probe failed', value),
            `  cache scope: ${value?.swVersionScope ?? 'unknown'}`,
            value?.path ? `  cache file: ${value.path}` : '',
          ].filter(Boolean).join('\n')
          : [
            `SolidWorks ${value?.revision ?? '?'} capabilities${value?.fromCache ? ' (cached)' : ''} [plugin ${value?.pluginRevision ?? 'unknown'}]`,
            `  cache scope: ${value?.swVersionScope ?? 'unknown'}${value?.fromCache ? ' (served from cache)' : ''}`,
            `  swept cut route usable: ${value?.capabilities?.sweptCut === true ? 'definition object created (pipeline still unproven)' : 'NO — use core body + swept boss'}`,
            ...(routingAdvice(value).map((line) => `  ${line}`)),
            value?.path ? `  report: ${value.path}` : '',
          ].filter(Boolean).join('\n'),
      }],
    },
    execute: async ({ force }, exec) => {
      const caps = probeCapabilities({ force: force === true, exec })
      // The scope and the cache location are reported in both outcomes: which
      // install a report belongs to, and which file a stale one would be read
      // from, are exactly what a confusing probe result needs to be diagnosed.
      const scope = swVersionScope()
      const path = cachePath()
      if (caps.error) return { ...errorResult(caps.error, caps.code), swVersionScope: scope, path, log: caps.log }
      return { ...caps, pluginRevision: PLUGIN_REVISION, swVersionScope: scope, path }
    },
    presentCall: (args) => ({
      card: 'generic',
      title: args?.force ? 'Probe SolidWorks capabilities (forced)' : 'Probe SolidWorks capabilities',
      kind: 'other',
      rawInput: args ?? {},
    }),
  }))

  register(validated({
    name: 'solidworks_run',
    description: [
      'Run a VBScript against SolidWorks over COM — the general escape hatch for any geometry, setting or export the other tools cannot express.',
      'The script must be ASCII (WSH reads .vbs as ANSI, so non-ASCII text fails to parse); write units in metres (the SolidWorks API is metre-based).',
      'The tool stages the script into an ASCII scratch directory, sets %SWPARTTPL% (part template) and %SWOUTDIR% (ASCII output dir), runs it with cscript and returns the log.',
      'SolidWorks traps to respect: sketch geometry must be invoked with `Call` (a bare sm.CreateLine statement is silently ignored); sw.InsertSketch2 / sw.InsertHelix live on ModelDoc2, not SketchManager; swept cuts and GetMassProperties do not work here.',
      'Save renders with sw.SaveBMP("<ascii path>",1400,900) — reading the image is the only reliable way to judge geometry; the tool returns a summary of the .bmp files produced.',
    ].join(' '),
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        script: { type: 'string', description: 'The full VBScript source (ASCII).' },
        args: {
          type: 'array',
          description: 'Arguments passed to the script (WScript.Arguments).',
          // `items` is a constraint keyword: it must be a bare schema object,
          // because the Harness validator rejects annotation keywords there.
          items: { type: 'string' },
        },
        timeoutMs: { type: 'integer', description: 'Hard timeout in ms (default 900000).' },
      },
      required: ['script'],
    },
    output: {
      schema: {},
      render: (_args, value) => [{
        type: 'text',
        text: value?.error
          ? `${errorText('solidworks_run failed', value)}\n${tail(value.log ?? '', 1500)}`
          : [
            `solidworks_run exit=${value?.exitCode ?? '?'}${value?.timedOut ? ' (TIMED OUT — SolidWorks may be showing a modal dialog)' : ''}`,
            `  template : ${value?.template || '<none>'}`,
            `  out dir  : ${value?.dir}`,
            value?.artifacts?.length ? `  artifacts: ${value.artifacts.join(', ')}` : '  artifacts: none (no .bmp/.SLDPRT produced)',
            value?.stderr ? `  stderr   : ${tail(value.stderr, 500)}` : '',
            '',
            tail(value?.log ?? '', 6000),
          ].filter(Boolean).join('\n'),
      }],
    },
    execute: async ({ script, args, timeoutMs }, exec) => {
      const { dir, template } = environmentFor(exec)
      const before = new Set(listArtifacts(dir))
      const run = executeScript(
        script,
        (args ?? []).map(String),
        exec,
        timeoutMs ?? cfg.defaultTimeoutMs,
      )
      if (run.error) return { ...errorResult(run.error, run.code), log: run.stdout ?? '', dir, template }
      const artifacts = listArtifacts(dir).filter((file) => !before.has(file))
      // Every field here must be lossless JSON: the host rejects the whole
      // result otherwise ("value is not lossless JSON"), so the log, the
      // artifacts and the exit code are all discarded and the caller sees a
      // tool error even though the script ran. `exec.signal` is an
      // AbortSignal object (and `undefined` when absent), which is exactly
      // what used to happen; the cancellation state is now reported as a
      // JSON-safe boolean instead of the signal itself.
      return {
        exitCode: run.status,
        timedOut: run.timedOut === true,
        aborted: exec?.signal?.aborted === true,
        // A run that reached the interpreter still reports a machine-readable
        // outcome: a timeout and a non-zero exit are the two recoverable
        // failures. Conditionally spread, because an explicit `code: undefined`
        // is NOT lossless JSON and the host then discards the whole result (the
        // same trap the abortSignal note below records).
        ...(run.timedOut === true
          ? { code: ERROR_CODES.TIMEOUT }
          : run.status === 0 ? {} : { code: ERROR_CODES.SCRIPT_ERROR }),
        stdout: run.stdout,
        stderr: run.stderr,
        log: run.stdout,
        dir,
        template,
        artifacts,
      }
    },
    presentCall: (args) => ({
      card: 'generic',
      title: 'Run VBScript in SolidWorks',
      kind: 'other',
      rawInput: { script: `${(args?.script ?? '').slice(0, 400)}…`, args: args?.args },
    }),
  }))

  register(validated({
    name: 'solidworks_verify',
    description: [
      'Read back what the ACTIVE SolidWorks document actually contains and compare it with an expected result: semantic feature kinds, solid-body count, and rendered views.',
      'Use it after every build. Feature NAMES are language-dependent (this UI is zh-CN), so assert on semantic kinds instead — each feature is reported as Extrude, Revolve, Cut, Helix, SweepBoss, SweepCut, Chamfer, Fillet, Hole, Pattern, Shell or Draft, while reference planes, sketches and folders are filtered out.',
      'Example: a threaded shaft built as core + helix + swept boss is expectBodyCount=1 with featureTypes=["Extrude","Helix","SweepBoss"].',
      'Returns structured differences plus the paths of rendered .bmp files — look at those images, because GetMassProperties and GetBodyBox are unavailable on this host.',
      'A build is only trustworthy once this tool reports ok:true; that is the gate a new modelling recipe must pass before it is worth reusing.',
    ].join(' '),
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        expectBodyCount: { type: 'integer', description: 'Expected solid-body count (1 for a single part).' },
        featureTypes: {
          type: 'array',
          items: { type: 'string' },
          description: 'Expected feature kinds; each must be present. One of: Extrude, Revolve, Cut, Helix, SweepBoss, SweepCut, Chamfer, Fillet, Hole, Pattern, Shell, Draft.',
        },
        views: { type: 'string', description: 'Comma-separated view ids to render (default "7,5"). 7=isometric, 2=front, 5=front(XZ), 6=top.' },
        includeReference: { type: 'boolean', description: 'Also list reference planes/axes, for debugging only (default false).' },
      },
    },
    // Declaring this makes the host schedule the call as `parallel`; every tool
    // that drives the SolidWorks UI deliberately omits it, so the host treats
    // those as `exclusive` and serializes them for us. Reading the active
    // document and probing capabilities do not mutate the model, so they are
    // safe to run alongside other read-only work.
    isConcurrencySafe: () => true,
    output: {
      schema: {},
      render: (_args, value) => [{
        type: 'text',
        text: value?.error
          ? errorText('solidworks_verify failed', value)
          : [
            `solidworks_verify: ${value.ok ? 'OK' : 'MISMATCH'} — ${value.title || '<untitled>'}`,
            `  solid bodies : ${value.bodyCount}${value.expectedBodyCount === undefined ? '' : ` (expected ${value.expectedBodyCount})`}`,
            `  feature kinds: ${(value.featureKinds ?? []).join(', ') || 'none'}`,
            value.boundingBox ? `  bounding box : ${value.boundingBox.size.join(' x ')} mm` : '  bounding box : unavailable (host limitation)',
            value.renders?.length ? `  renders      : ${value.renders.map((r) => r.bmp).join(', ')}` : '  renders      : none',
            ...(value.mismatches?.length ? ['', ...value.mismatches.map((m) => `  ! ${m}`)] : []),
          ].filter(Boolean).join('\n'),
      }],
    },
    execute: async ({ expectBodyCount, featureTypes, views, includeReference }, exec) =>
      verifyModel({ expectBodyCount, featureTypes, views, includeReference }, exec),
    presentCall: (args) => ({
      card: 'generic',
      title: 'Verify SolidWorks model',
      kind: 'other',
      rawInput: args ?? {},
    }),
  }))

  // ------------------------------------------------------- recipe data
  /**
   * The plugin's data surface. Recipes are loaded from disk at composition
   * time: shipped defaults come from `lib/recipes.js`, user recipes from
   * `recipeDir`. A recipe that fails validation is skipped and logged — it
   * cannot take the plugin down, which is the whole reason this is data rather
   * than code.
   */
  const recipes = mergeRecipes(loadShipped(log), loadUserRecipes(recipeDir, log), log)
  const recipeByName = new Map(recipes.map((recipe) => [recipe.name, recipe]))
  log(`recipes loaded: ${recipes.length === 0 ? 'none' : recipes.map((r) => `${r.name}${r.shipped ? '(shipped)' : ''}`).join(', ')} from ${recipeDir}`)

  register(validated({
    name: 'solidworks_recipe',
    description: [
      'Run a stored build recipe: a validated script plus the shape it must produce.',
      'A recipe is data, not code — it carries parameter defaults, so the same recipe produces a different part by passing `params`.',
      'The run always verifies itself against the recipe\'s recorded expectations (body count and semantic feature kinds) and reports whether it passed; a failing verification is an error, not a warning.',
      'Available recipes:',
      describeRecipes(recipes),
      'Use solidworks_recipes to list them, save a new one from a script that solidworks_verify has accepted, or remove one.',
    ].join(' '),
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        name: { type: 'string', description: 'Recipe name, as listed in this description.' },
        params: {
          type: 'object',
          description: 'Parameter overrides, e.g. {"CORE_R":0.03}. Must match the recipe\'s declared parameters; values are finite numbers.',
        },
        views: { type: 'string', description: 'Comma-separated view ids to render during verification (default "7,5").' },
        timeoutMs: { type: 'integer', description: 'Hard timeout in ms for the build script (default 900000).' },
      },
      required: ['name'],
    },
    output: {
      schema: {},
      render: (_args, value) => [{
        type: 'text',
        text: value?.error
          ? errorText('solidworks_recipe failed', value)
          : [
            `solidworks_recipe ${value.ok ? 'OK' : 'VERIFICATION FAILED'} — ${value.name}${value.shipped ? ' (shipped)' : ''}`,
            `  parameters   : ${Object.entries(value.values ?? {}).map(([k, v]) => `${k}=${v}`).join(' ')}`,
            `  solid bodies : ${value.verification?.bodyCount ?? '?'}${value.verification?.expectedBodyCount === undefined ? '' : ` (expected ${value.verification.expectedBodyCount})`}`,
            `  feature kinds: ${(value.verification?.featureKinds ?? []).join(', ') || 'none'}`,
            value.verification?.renders?.length ? `  renders      : ${value.verification.renders.map((r) => r.bmp).join(', ')}` : '',
            ...(value.verification?.mismatches?.length ? ['', ...value.verification.mismatches.map((m) => `  ! ${m}`)] : []),
            '',
            tail(value.log ?? '', 4000),
          ].filter(Boolean).join('\n'),
      }],
    },
    execute: async ({ name, params, views, timeoutMs }, exec) => {
      const recipe = recipeByName.get(String(name))
      if (recipe === undefined) {
        return errorResult(`unknown recipe "${name}" (available: ${[...recipeByName.keys()].join(', ') || 'none'})`, ERROR_CODES.RECIPE_UNKNOWN)
      }
      const built = materialize(recipe, params ?? {})
      if (!built.ok) return errorResult(built.reason, ERROR_CODES.RECIPE_INVALID)
      const run = executeScript(built.script, [], exec, timeoutMs ?? cfg.defaultTimeoutMs)
      if (run.error) return { ...errorResult(run.error, run.code), log: run.stdout ?? '', name: recipe.name }
      const verification = verifyModel({
        expectBodyCount: recipe.verify.bodyCount,
        featureTypes: recipe.verify.featureTypes,
        views,
      }, exec)
      const logText = `${run.stdout ?? ''}\n--- verification ---\n${verification.log ?? verification.error ?? ''}`
      if (verification.error) {
        return {
          ...errorResult(
            withCode(verification.code ?? ERROR_CODES.INSPECT_FAILED, `recipe ran but verification could not read the model: ${verification.error}`),
            verification.code ?? ERROR_CODES.INSPECT_FAILED,
          ),
          log: tail(logText),
          name: recipe.name,
        }
      }
      return {
        ok: verification.ok,
        // Only present on failure. An explicit `error: undefined` is NOT
        // lossless JSON for the host, which then discards the whole result -
        // the same trap that `abortSignal: undefined` set for solidworks_run.
        ...(verification.ok
          ? {}
          : {
            ...errorResult(
              `recipe "${recipe.name}" built a model that does not match its recorded expectations`,
              ERROR_CODES.VERIFICATION_FAILED,
            ),
          }),
        name: recipe.name,
        shipped: recipe.shipped === true,
        values: built.values,
        exitCode: run.status,
        timedOut: run.timedOut === true,
        ...(run.timedOut === true
          ? { code: ERROR_CODES.TIMEOUT }
          : run.status === 0 ? {} : { code: ERROR_CODES.SCRIPT_ERROR }),
        artifacts: listArtifacts(run.dir ?? environmentFor(exec).dir),
        verification,
        log: tail(logText),
      }
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Run recipe ${args?.name ?? '?'}`,
      kind: 'other',
      rawInput: args ?? {},
    }),
  }))

  register(validated({
    name: 'solidworks_recipes',
    description: [
      'Manage stored build recipes: list them, save a new one, remove one.',
      'A recipe is a validated script plus its parameter defaults and the model shape it must produce; solidworks_recipe then runs it with no source edit involved.',
      'The save gate is intentional: the script must be ASCII and every %NAME% placeholder must have a declared numeric parameter, so a recipe cannot be stored half-formed.',
      'Promotion procedure: write and iterate the script with solidworks_run, get solidworks_verify to report ok:true for the shape you want, then save it here with that expectation (verify.bodyCount / verify.featureTypes). It becomes solidworks_recipe name=<name>.',
    ].join(' '),
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        action: { type: 'string', enum: ['list', 'save', 'remove'], description: 'list (default), save, or remove.' },
        name: { type: 'string', description: 'Recipe name for save/remove: lowercase letters, digits, _ or - (max 31 chars).' },
        description: { type: 'string', description: 'One-line description shown in listings (save).' },
        script: {
          type: 'string',
          description: 'The full VBScript, ASCII only, with numeric literals replaced by %NAME% placeholders (save).',
        },
        parameters: {
          type: 'object',
          description: 'Parameter defaults for save: {"CORE_R":{"value":0.024,"description":"core radius"}}. Every placeholder needs one.',
        },
        verify: {
          type: 'object',
          description: 'Recorded expectation checked by every future run: {"bodyCount":1,"featureTypes":["Extrude","Helix","SweepBoss"]}.',
        },
      },
      required: ['action'],
    },
    isConcurrencySafe: () => true,
    output: {
      schema: {},
      render: (_args, value) => [{
        type: 'text',
        text: value?.error
          ? errorText('solidworks_recipes failed', value)
          : [
            `solidworks_recipes ${value.action}: ${value.summary}`,
            value.path ? `  saved to: ${value.path}` : '',
            value.recipeDir ? `  directory: ${value.recipeDir}` : '',
            value.catalogue ? `\n${value.catalogue}` : '',
          ].filter(Boolean).join('\n'),
      }],
    },
    execute: async ({ action, name, description, script, parameters, verify }) => {
      const act = action ?? 'list'
      if (act === 'list') {
        return {
          action: act,
          summary: `${recipes.length} recipe(s)`,
          recipeDir,
          catalogue: describeRecipes(recipes),
        }
      }
      if (act === 'remove') {
        if (typeof name !== 'string' || name.length === 0) return errorResult('remove needs "name"', ERROR_CODES.INVALID_INPUT)
        if (recipeByName.get(name)?.shipped === true && !existsSync(join(recipeDir, `${name}.json`))) {
          return errorResult(`"${name}" is a shipped recipe; save an overridden version under the same name first, or remove a user recipe`, ERROR_CODES.INVALID_INPUT)
        }
        const removed = removeRecipeFile(recipeDir, name)
        recipeByName.delete(name)
        const index = recipes.findIndex((recipe) => recipe.name === name)
        if (index >= 0) recipes.splice(index, 1)
        return {
          action: act,
          summary: removed ? `removed "${name}" (takes effect from the next load for shipped defaults)` : `no user recipe named "${name}"`,
          recipeDir,
          catalogue: describeRecipes(recipes),
        }
      }
      if (act !== 'save') return errorResult(`unknown action "${act}"`, ERROR_CODES.INVALID_INPUT)

      const candidate = {
        name,
        description: description ?? '',
        script,
        parameters: parameters ?? {},
        verify: verify ?? {},
      }
      const result = validateRecipe(candidate, 'save')
      if (!result.ok) return errorResult(result.reason, ERROR_CODES.RECIPE_INVALID)
      if (recipeByName.has(result.recipe.name) && recipeByName.get(result.recipe.name).shipped) {
        log(`recipe "${result.recipe.name}" overrides the shipped version`)
      }
      const path = saveRecipeFile(recipeDir, result.recipe)
      recipeByName.set(result.recipe.name, result.recipe)
      const existing = recipes.findIndex((recipe) => recipe.name === result.recipe.name)
      if (existing >= 0) recipes[existing] = result.recipe
      else recipes.push(result.recipe)
      recipes.sort((a, b) => a.name.localeCompare(b.name))
      return {
        action: act,
        summary: `saved "${result.recipe.name}" with ${Object.keys(result.recipe.parameters).length} parameter(s)`,
        path,
        recipeDir,
        catalogue: describeRecipes(recipes),
      }
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Recipes: ${args?.action ?? 'list'}`,
      kind: 'other',
      rawInput: args ?? {},
    }),
  }))

  // --------------------------------------------------- prompt section
  try {
    ctx.systemPrompt?.section?.({
      name: 'tool:solidworks',
      text: [
        'SolidWorks is driven through the solidworks_* tools (VBScript over COM).',
        'Non-negotiables: scripts must be ASCII; units are metres; sketch geometry must be invoked with `Call`;',
        'verify every build with solidworks_verify and look at the rendered image, because model queries such as',
        'GetMassProperties and GetBodyBox return Empty on this host.',
        'Every failure carries a stable machine-readable `code` beside its message — branch on the code, not the prose:',
        'no-solidworks (COM could not start SolidWorks), no-template (part template missing), non-ascii-script,',
        'cscript-missing, timeout (often a modal dialog in the SolidWorks UI), script-error, probe-failed,',
        'inspect-failed, verification-failed, recipe-unknown, recipe-invalid, invalid-input.',
      ].join(' '),
    })
  } catch (error) {
    log(`system prompt section skipped: ${error instanceof Error ? error.message : String(error)}`)
  }
}

/** New artifacts (.bmp / .SLDPRT / .json reports) in the scratch dir. */
function listArtifacts(dir) {
  try {
    return readdirSync(dir).filter((entry) => /\.(bmp|png|sldprt|json|stl|step|stp|igs)$/i.test(entry))
  } catch {
    return []
  }
}
