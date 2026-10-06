// Data-layer test: recipes are data, the code stays put.
//
// Verifies the four properties that make recipe storage worth having:
//   1. a shipped recipe runs and passes its own recorded verification;
//   2. a new recipe is saved and immediately runnable via solidworks_recipe;
//   3. a malformed save is REJECTED with a precise reason;
//   4. a malformed recipe FILE on disk is skipped at load and does not stop the
//      plugin from loading — the property that a source edit does not have.
import { existsSync, mkdirSync, writeFileSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { apply } from '../lib/index.js'

const base = join(tmpdir(), 'dsh-solidworks', 'recipetest')
const recipeDir = join(base, 'recipes')
rmSync(base, { recursive: true, force: true })
mkdirSync(recipeDir, { recursive: true })

const logs = []
const tools = new Map()
apply({
  tools: { register: (definition) => tools.set(definition.name, definition) },
  systemPrompt: { section: () => {} },
  logger: { info: (m) => { logs.push(m); console.log('[log]', m) } },
}, { scratchDir: base, recipeDir })

const call = async (name, args) => {
  const tool = tools.get(name)
  if (!tool) throw new Error(`no tool ${name}`)
  const value = await tool.execute(args ?? {}, {})
  return { value, rendered: tool.output.render(args ?? {}, value).map((c) => c.text).join('\n') }
}

const fail = (message) => { console.error(`\nFAIL: ${message}`); process.exit(1) }
const ok = (message) => console.log(`  ok  ${message}`)

console.log(`registered tools: ${[...tools.keys()].join(', ')}`)
if (tools.size !== 5) fail(`expected 5 tools, got ${tools.size}`)

// 1) list ---------------------------------------------------------------
console.log('\n===== list =====')
const listed = await call('solidworks_recipes', { action: 'list' })
console.log(listed.rendered)
if (!listed.value.catalogue?.includes('threaded_shaft')) fail('shipped recipe missing from the catalogue')

// 2) run the shipped recipe --------------------------------------------
console.log('\n===== run shipped threaded_shaft =====')
const ran = await call('solidworks_recipe', { name: 'threaded_shaft', params: { REVS: 12 }, views: '7' })
console.log(ran.rendered)
if (ran.value.error) fail(`recipe run errored: ${ran.value.error}`)
if (!ran.value.ok) fail(`recipe verification failed: ${(ran.value.verification?.mismatches ?? []).join('; ')}`)
if (ran.value.values.REVS !== 12) fail('parameter override was not applied')
ok('shipped recipe ran, self-verified, and honoured the param override')

// 3) reject a malformed save -------------------------------------------
console.log('\n===== save with non-ASCII script (must be rejected) =====')
const rejected = await call('solidworks_recipes', {
  action: 'save',
  name: 'bad_recipe',
  script: 'WScript.Echo "中文注释"',
  parameters: {},
  verify: {},
})
console.log(rejected.rendered)
if (!rejected.value.error?.includes('non-ASCII')) fail('non-ASCII save was not rejected with a non-ASCII reason')
ok('non-ASCII save rejected')

console.log('\n===== save with an undeclared placeholder (must be rejected) =====')
const rejected2 = await call('solidworks_recipes', {
  action: 'save',
  name: 'bad_recipe2',
  script: 'Dim x\nx = %NOPE%\nWScript.Echo x',
  parameters: {},
  verify: {},
})
console.log(rejected2.rendered)
if (!rejected2.value.error?.includes('NOPE')) fail('undeclared placeholder was not rejected')
ok('undeclared placeholder rejected')

// 4) save a real recipe and run it -------------------------------------
const discScript = [
  'Option Explicit',
  'Const R = {{R}}',
  'Const T = {{T}}',
  'Dim swApp, sw, ext, sm, fm, planes(2), f, n, cc, fe',
  'Set swApp = CreateObject("SldWorks.Application")',
  'swApp.Visible = True',
  'Do While swApp.GetDocumentCount > 0',
  '    Set sw = swApp.ActiveDoc',
  '    If IsEmpty(sw) Or TypeName(sw) = "Nothing" Then Exit Do',
  '    swApp.CloseDoc sw.GetTitle',
  'Loop',
  'swApp.NewDocument CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SWPARTTPL%"), 0, 0, 0',
  'WScript.Sleep 2500',
  'Set sw = swApp.ActiveDoc',
  'Set ext = sw.Extension',
  'Set sm = sw.SketchManager',
  'Set fm = sw.FeatureManager',
  'n = 0',
  'Set f = sw.FirstFeature',
  'Do While Not f Is Nothing',
  '    If f.GetTypeName2 = "RefPlane" Then',
  '        If n < 3 Then planes(n) = f.Name : n = n + 1',
  '    End If',
  '    Set f = f.GetNextFeature',
  'Loop',
  'ext.SelectByID2 planes(0), "PLANE", 0, 0, 0, False, 0, Nothing, 0',
  'sw.InsertSketch2 True',
  'Set cc = sm.CreateCircleByRadius(0.0, 0.0, 0.0, R)',
  'sw.InsertSketch2 True',
  'Set fe = fm.FeatureExtrusion2(True, False, False, 0, 0, T, 0.01, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)',
  'If IsEmpty(fe) Or TypeName(fe) = "Nothing" Then WScript.Echo "FATAL extrude" : WScript.Quit 0',
  'fe.Name = "Disc"',
  'WScript.Echo "disc built: R=" & R & " T=" & T',
  'sw.ClearSelection2 True',
  'sw.ForceRebuild3 False',
].join('\n')

console.log('\n===== save a new recipe =====')
const saved = await call('solidworks_recipes', {
  action: 'save',
  name: 'flat_disc',
  description: 'Flat disc: one extruded circle on the front plane.',
  script: discScript,
  parameters: { R: { value: 0.04, description: 'disc radius' }, T: { value: 0.012, description: 'thickness' } },
  verify: { bodyCount: 1, featureTypes: ['Extrude'] },
})
console.log(saved.rendered)
if (saved.value.error) fail(`save failed: ${saved.value.error}`)
if (!existsSync(join(recipeDir, 'flat_disc.json'))) fail('recipe file was not written')
ok('recipe saved as JSON on disk')

console.log('\n===== run the saved recipe with params =====')
const ranDisc = await call('solidworks_recipe', { name: 'flat_disc', params: { R: 0.05, T: 0.02 }, views: '7' })
console.log(ranDisc.rendered)
if (ranDisc.value.error) fail(`saved recipe errored: ${ranDisc.value.error}`)
if (!ranDisc.value.ok) fail(`saved recipe verification failed: ${(ranDisc.value.verification?.mismatches ?? []).join('; ')}`)
ok('saved recipe ran and self-verified')

// 5) a corrupt recipe file must be skipped, not fatal -------------------
console.log('\n===== corrupt recipe file on disk (skipped, plugin still loads) =====')
writeFileSync(join(recipeDir, 'corrupt.json'), '{ this is not json', 'utf8')
writeFileSync(join(recipeDir, 'badparams.json'), JSON.stringify({
  name: 'badparams', script: 'Dim a\na = %MISSING%', parameters: {}, verify: {},
}), 'utf8')

const logs2 = []
const tools2 = new Map()
apply({
  tools: { register: (definition) => tools2.set(definition.name, definition) },
  systemPrompt: { section: () => {} },
  logger: { info: (m) => { logs2.push(m); console.log('[log]', m) } },
}, { scratchDir: base, recipeDir })
if (tools2.size !== 5) fail(`plugin did not load with corrupt recipe files present (tools=${tools2.size})`)
ok(`plugin loaded all 5 tools despite 2 bad recipe files`)
const skipped = logs2.filter((m) => m.includes('skipping'))
if (skipped.length < 2) fail(`expected 2 skip warnings, got ${skipped.length}`)
ok(`${skipped.length} bad recipe(s) skipped with warnings`)

// 6) remove -------------------------------------------------------------
console.log('\n===== remove =====')
const removed = await call('solidworks_recipes', { action: 'remove', name: 'flat_disc' })
console.log(removed.rendered)
if (removed.value.error) fail(`remove failed: ${removed.value.error}`)
if (existsSync(join(recipeDir, 'flat_disc.json'))) fail('recipe file still present after remove')
ok('recipe removed')

console.log('\nRECIPE DATA LAYER VERIFIED')
