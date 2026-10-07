/**
 * capability-cache.mjs — the probe cache must be scoped to the installed
 * SolidWorks release.
 *
 * Why this is a real failure mode and not a nicety: a capability report
 * describes the API surface of ONE install. An upgrade or Service Pack can
 * flip exactly the routes the probe measures (`sweptCut`, the fillet option
 * bits, `InsertFeatureShell`, `InsertAxis2`), and the 6-hour TTL alone would
 * serve the new install a report measured on the old one — silently routing
 * modelling away from a route that works, or into one that does not.
 *
 * Each case registers a fresh plugin instance, because the install a report
 * belongs to is resolved once at registration and must not drift mid-run.
 *
 *   node test/capability-cache.mjs
 */
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const plugin = await import('../lib/index.js')

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

const base = join(tmpdir(), 'dsh-solidworks', 'cachetest')
rmSync(base, { recursive: true, force: true })

/** A year marker that exists on disk, so the configured-template path is taken. */
const fakeTemplate = (year) => {
  const dir = join(base, 'tpl', `SOLIDWORKS ${year}`, 'templates')
  mkdirSync(dir, { recursive: true })
  return join(dir, 'gb_part.prtdot')
}
const t2026 = fakeTemplate(2026)
const t2025 = fakeTemplate(2025)

/** A fake install root holding two years, each with a usable template. */
const fakeProgramData = join(base, 'programdata')
mkdirSync(join(fakeProgramData, '2027', 'templates'), { recursive: true })
mkdirSync(join(fakeProgramData, '2026', 'templates'), { recursive: true })
// The template file matters: `resolveTemplate` only accepts a version directory
// that actually holds a `.prtdot`, so empty directories fall through to the
// real install and would make this case assert the wrong thing.
for (const year of ['2027', '2026']) {
  writeFileSync(join(fakeProgramData, year, 'templates', 'gb_part.prtdot'), '')
}
/** A directory with no year subdirectories: the "cannot locate an install" case. */
const emptyRoot = join(base, 'empty-root')
mkdirSync(emptyRoot, { recursive: true })

/** Register a fresh instance with a pinned template / install root. */
const registered = ({ template, versionDir }) => {
  const tools = new Map()
  if (template) process.env.SW_TEST_TEMPLATE = template
  else delete process.env.SW_TEST_TEMPLATE
  if (versionDir) process.env.SW_VERSION_DIR = versionDir
  else delete process.env.SW_VERSION_DIR
  plugin.apply(
    {
      tools: { register: (tool) => tools.set(tool.name, tool) },
      logger: { info: () => {} },
      systemPrompt: { section: () => {} },
    },
    { scratchDir: join(base, 'scratch') },
  )
  return tools
}

/**
 * Call the tool once. `force` makes the probe genuinely run, so a host without
 * SolidWorks yields `{ error }` — every cache-scope assertion below holds in
 * both outcomes, which is what keeps this test runnable off a CAD machine.
 */
const call = async (options) => {
  const tool = registered(options).get('solidworks_capabilities')
  const value = await tool.execute({ force: options.force === true }, {})
  const rendered = tool.output.render({}, value).map((c) => c.text).join('\n')
  delete process.env.SW_TEST_TEMPLATE
  delete process.env.SW_VERSION_DIR
  return { value, rendered }
}

// 1. the configured template's year keys the cache
const a = await call({ template: t2026 })
check('configured template 2026 keys the cache as 2026', a.value.swVersionScope === '2026', `scope=${a.value.swVersionScope}`)
check('report discloses the cache scope', /cache scope: 2026/.test(a.rendered), a.rendered.split('\n')[1] ?? '')

// 2. a different install is a different scope — the whole point
const b = await call({ template: t2025 })
check('downgrade to 2025 keys the cache as 2025', b.value.swVersionScope === '2025', `scope=${b.value.swVersionScope}`)
check('the two installs do not share a cache file', a.value.path !== b.value.path, `${a.value.path} vs ${b.value.path}`)

// 3. an upgrade invalidates by itself, with no TTL wait and no `force`
const c = await call({ versionDir: fakeProgramData })
check('newest installed release wins when no template is configured', c.value.swVersionScope === '2027', `scope=${c.value.swVersionScope}`)

// 4. a year inside a template FILE name is not the install year
const decoy = join(base, 'tpl', 'gb_part 2020.prtdot')
const d = await call({ template: decoy, versionDir: fakeProgramData })
check('a year in the file name cannot be mistaken for the install year', d.value.swVersionScope === '2027', `scope=${d.value.swVersionScope}`)

// 5. with no locatable install the key degrades to "unknown", never to another
//    install's. Only assertable on a host with no real SolidWorks install: a
//    probe machine correctly falls back to the install it actually has.
const dsh = await call({ template: decoy, versionDir: emptyRoot })
if (existsSync('C:\\ProgramData\\SOLIDWORKS')) {
  check(
    'no fake install does not leak into the scope when a real one exists',
    dsh.value.swVersionScope !== '2020' && dsh.value.swVersionScope !== '2027',
    `scope=${dsh.value.swVersionScope} (real install present, so it resolves there by design)`,
  )
} else {
  check('unlocatable install degrades to "unknown"', dsh.value.swVersionScope === 'unknown', `scope=${dsh.value.swVersionScope}`)
}

// 6. the cache file name actually carries the scope
check('cache file name is scoped, not a bare capabilities.json', /capabilities-2026\.json$/.test(String(a.value.path)), String(a.value.path))

// 7. the cache lives at INSTANCE level while the scratch workspace is per
//    INVOCATION. Both halves matter: a fresh directory per call would mean the
//    6-hour TTL is never reached, and a shared directory per instance would put
//    two invocations' inspect.json in one place.
const scoped = registered({ template: t2026 })
const runA = scoped.get('solidworks_run')
const runB = scoped.get('solidworks_run')
const caps = scoped.get('solidworks_capabilities')
const w1 = (await runA.execute({ script: 'WScript.Echo "a"', timeoutMs: 15000 }, { callId: 'w-one' })).dir
const w2 = (await runB.execute({ script: 'WScript.Echo "b"', timeoutMs: 15000 }, { callId: 'w-two' })).dir
const c1 = await caps.execute({}, { callId: 'c-one' })
const c2 = await caps.execute({}, { callId: 'c-two' })
check('two invocations get two different workspaces', w1 !== w2, `${w1} vs ${w2}`)
check('the cache path is stable across those same invocations', c1.path === c2.path, `${c1.path} vs ${c1.path === c2.path ? '' : c2.path}`)
check('the cache is not inside an invocation workspace', !String(c1.path).includes('w-one') && /capabilities-2026\.json$/.test(String(c1.path)), String(c1.path))

// 8. a real probe (when this host has SolidWorks) writes the scoped file
const f = await call({ template: t2026, force: true })
if (f.value.error) {
  console.log(`SKIP  the probe itself did not run here (${f.value.error}) — the scope logic is asserted above`)
} else {
  check('forced probe wrote the scoped cache file', existsSync(String(f.value.path)), String(f.value.path))
}

const failures = results.filter((r) => !r.ok)
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} (${results.length} checks)`)
process.exit(failures.length === 0 ? 0 : 1)
