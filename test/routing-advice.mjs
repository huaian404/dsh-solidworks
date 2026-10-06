/**
 * routing-advice.mjs — exercise routingAdvice through the real tool renderer.
 *
 * routingAdvice is what turns a capability report into "do not use this route"
 * guidance. The three silent-failure routes added in this revision are only
 * useful if a report containing them actually produces advice, so each is fed
 * through solidworks_capabilities' own render function here, with a fabricated
 * report, and the rendered text is asserted.
 *
 *   node test/routing-advice.mjs
 */
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const plugin = await import('../lib/index.js')

const registered = new Map()
plugin.apply(
  {
    tools: { register: (tool) => registered.set(tool.name, tool) },
    logger: { info: () => {} },
    systemPrompt: { section: () => {} },
  },
  { scratchDir: join(tmpdir(), 'dsh-solidworks', 'routing') },
)

const tool = registered.get('solidworks_capabilities')
if (tool === undefined) {
  console.error('solidworks_capabilities did not register')
  process.exit(1)
}

const render = (report) => tool.output.render({}, report).map((c) => c.text).join('\n')

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

const base = {
  revision: '34.2.1',
  pluginRevision: 'recipes-4',
  capabilities: {
    sweptCut: false,
    filletOptionsBit2Required: true,
    filletWorkingOption: 3,
    negZFaceCut: false,
    insertFeatureShellCreatesFeature: false,
    insertAxis2CreatesFeature: false,
  },
  api: {
    doc_get_mass_properties: { ok: false, detail: 'Empty' },
    doc_get_bodies2: { ok: false, detail: 'Empty' },
  },
}

const text = render(base)
check('fillet advice names the required Options', /Options 3/.test(text) && /uniform-radius/i.test(text))
check('fillet advice warns about the failing Options', /Options 0, 1 and 4/.test(text))
check('negZ advice is present', /outward normal is -Z/.test(text))
check('shell advice says no feature was created', /InsertFeatureShell/.test(text) && /did NOT create a feature/.test(text))
check('axis advice names InsertAxis2', /InsertAxis2/.test(text))
check('advice states the procedure-scope error rule', /On Error Resume Next/.test(text) && /does NOT apply inside a procedure/.test(text))
check('swept-cut advice is preserved', /core body plus a merged swept BOSS/.test(text))
check('mass-property advice mentions CreateMassProperty2', /CreateMassProperty2/.test(text))
check('no empty advice line leaked in', !text.includes('\n  \n'))

// A report from a host where everything works must not invent problems.
const healthy = render({
  revision: '34.2.1',
  pluginRevision: 'recipes-4',
  capabilities: {
    sweptCut: true,
    filletOptionsBit2Required: true,
    filletWorkingOption: 2,
    negZFaceCut: true,
    insertFeatureShellCreatesFeature: true,
    insertAxis2CreatesFeature: true,
  },
  api: { doc_get_bodies2: { ok: true } },
})
check('healthy host: no shell/axis warning', !/did NOT create a feature/.test(healthy))
check('healthy host: no negZ warning', !/outward normal is -Z/.test(healthy))
// the procedure-scope rule is a VBScript fact, not a host capability: always on
check('procedure-scope rule is always present', /does NOT apply inside a procedure/.test(healthy))

const failures = results.filter((r) => !r.ok)
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} (${results.length} checks)`)
process.exit(failures.length === 0 ? 0 : 1)
