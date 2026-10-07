/**
 * assembly-gate.mjs — the part/assembly verification judgement, pure logic.
 *
 * What this locks down (the bug, measured): `solidworks_verify` compared every
 * document against a solid-body count, but on an assembly `GetBodies2` does not
 * bind, so inspect.vbs reported `bodyCount: -1`, the "has no solid body at all"
 * rule fired, and every correctly built assembly came back MISMATCH. An assembly
 * needs its own expectations (components, mates, suppressed features) and its
 * body count must not be judged at all.
 *
 * It drives the real `verifyModel` through the exported `buildTools` factory,
 * injecting a RECORDED inspect report instead of running inspect.vbs, so the
 * judgement is tested with no SolidWorks and no cscript — which is exactly why
 * it can run in CI and inside a sandboxed agent. The live side (inspect.vbs over
 * COM, and the build itself) is covered by test/assembly-verify.mjs.
 *
 *   node test/assembly-gate.mjs
 */
const plugin = await import('../lib/index.js')

const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ` — ${detail}` : ''}`)
}

const tools = plugin.buildTools({ scratchDir: '', partTemplate: '', assemblyTemplate: '' }, () => {}, 'gate')

/** The shape inspect.vbs emits for a part with one body. */
const partReport = (bodyCount) => ({
  version: 1,
  documentType: 'part',
  documentTypeCode: 1,
  title: 'some_part',
  bodyCount,
  boundingBox: null,
  assembly: null,
  features: [
    { name: 'Core', type: 'Extrusion', kind: 'Extrude', suppressed: false },
    { name: 'Bore', type: 'ICE', kind: 'Cut', suppressed: false },
  ],
  renders: [{ view: 7, bmp: 'C:\\tmp\\view7.bmp' }],
})

/** The shape inspect.vbs emits for an assembly of N instances. */
const assemblyReport = ({ components, references, mates = null, available = false, suppressed = 0 }) => ({
  version: 1,
  documentType: 'assembly',
  documentTypeCode: 2,
  title: 'some_assembly',
  bodyCount: -1,
  boundingBox: null,
  assembly: {
    componentCount: components,
    referenceFeatureCount: references,
    mateCount: mates,
    mateCountAvailable: available,
    suppressedFeatureCount: suppressed,
    componentDetailsAvailable: false,
  },
  features: [
    { name: 'asm_part-1', type: 'Reference', kind: '', suppressed: false },
  ],
  renders: [{ view: 7, bmp: 'C:\\tmp\\view7.bmp' }],
})

const judge = (report, expect) => tools.verifyModel(expect, { inspect: report })
const lossless = (value) => {
  try {
    const text = JSON.stringify(value)
    return typeof text === 'string'
  } catch {
    return false
  }
}

// ---------------------------------------------------------------- part path
const partOk = judge(partReport(1), { expectBodyCount: 1, featureTypes: ['Extrude', 'Cut'] })
check('a part with the expected body count passes', partOk.ok === true, (partOk.mismatches ?? []).join('; '))
check('the part body count is judged', partOk.bodyCount === 1)
const partBad = judge(partReport(2), { expectBodyCount: 1 })
check('a wrong part body count fails', partBad.ok === false, (partBad.mismatches ?? []).join('; '))
const partEmpty = judge(partReport(0), {})
check('an empty part fails even with no expectation', partEmpty.ok === false, (partEmpty.mismatches ?? []).join('; '))

// ----------------------------------------------------------- assembly path
const asm3 = assemblyReport({ components: 3, references: 3 })
const asmOk = judge(asm3, { expectComponentCount: 3 })
check('a 3-component assembly passes expectComponentCount=3', asmOk.ok === true, (asmOk.mismatches ?? []).join('; '))
check('the document type is reported', asmOk.documentType === 'assembly')
check('the body count is NOT judged on an assembly', asmOk.bodyCount === -1 && asmOk.ok === true)
const asmWrong = judge(asm3, { expectComponentCount: 4 })
check('a wrong component count fails', asmWrong.ok === false, (asmWrong.mismatches ?? []).join('; '))
check('the mismatch names both counts', /found 3, expected 4/.test((asmWrong.mismatches ?? []).join(' ')))

// expectBodyCount is ignored on an assembly rather than failing it.
const asmBodyIgnored = judge(asm3, { expectBodyCount: 1, expectComponentCount: 3 })
check('expectBodyCount is ignored for an assembly', asmBodyIgnored.ok === true, (asmBodyIgnored.mismatches ?? []).join('; '))

// Mates: assertable only when the inspector actually saw mate features, which
// on this host it never does — mate state is unreadable. The judge must then
// refuse to confirm the check rather than pass it.
const asmMatesSeen = judge(assemblyReport({ components: 2, references: 2, mates: 1, available: true }), { expectMateCount: 1 })
check('a visible mate satisfies expectMateCount=1', asmMatesSeen.ok === true, (asmMatesSeen.mismatches ?? []).join('; '))
const asmNoMates = judge(assemblyReport({ components: 2, references: 2, mates: 0, available: true }), { expectMateCount: 1 })
check('a visible absence of mates fails expectMateCount=1', asmNoMates.ok === false, (asmNoMates.mismatches ?? []).join('; '))
const asmMatesUnknown = judge(assemblyReport({ components: 2, references: 2 }), { expectMateCount: 1 })
check('unreadable mate state reports the check as unavailable, not as a pass', asmMatesUnknown.ok === false && /cannot be inspected/.test((asmMatesUnknown.mismatches ?? []).join(' ')), (asmMatesUnknown.mismatches ?? []).join('; '))

// A suppressed feature is the only missing-reference signal available here.
const asmSuppressed = judge(assemblyReport({ components: 3, references: 3, suppressed: 1 }), { expectSuppressedCount: 1 })
check('a suppressed feature is counted', asmSuppressed.ok === true, (asmSuppressed.mismatches ?? []).join('; '))
const asmClean = judge(assemblyReport({ components: 3, references: 3 }), { expectSuppressedCount: 0 })
check('a clean assembly reports 0 suppressed', asmClean.ok === true, (asmClean.mismatches ?? []).join('; '))

// ------------------------------------------------------------------ misuse
const wrongDoc = judge(partReport(1), { expectComponentCount: 3 })
check('expectComponentCount on a part fails with an explanation', wrongDoc.ok === false && /not an assembly/.test((wrongDoc.mismatches ?? []).join(' ')), (wrongDoc.mismatches ?? []).join('; '))
const noAssemblyBlock = judge({ ...assemblyReport({ components: 2, references: 2 }), assembly: null }, { expectComponentCount: 2 })
check('a missing assembly block fails rather than passing silently', noAssemblyBlock.ok === false, (noAssemblyBlock.mismatches ?? []).join('; '))

// ------------------------------------------------------- payload contract
// A report RECORDED from inspect.vbs on 2026-02-XX (SW 2026 SP2.1, a 3-instance
// assembly). Pinned verbatim so a JSON-shape change in the inspector cannot
// silently stop matching the judgement — the injected fixtures above are
// hand-written, and this one is the real thing.
const recordedAssemblyReport = {
  version: 1,
  documentType: 'assembly',
  documentTypeCode: 2,
  title: 'asm_probe',
  bodyCount: -1,
  boundingBox: null,
  assembly: {
    componentCount: 3,
    referenceFeatureCount: 3,
    mateCount: null,
    mateCountAvailable: false,
    suppressedFeatureCount: 0,
    componentDetailsAvailable: false,
  },
  features: [],
  renders: [{ view: 7, bmp: 'C:\\tmp\\view7.bmp' }],
}
const recorded = judge(recordedAssemblyReport, { expectComponentCount: 3, expectSuppressedCount: 0 })
check('a report recorded from inspect.vbs passes the gate', recorded.ok === true, (recorded.mismatches ?? []).join('; '))
check('the recorded report is recognised as an assembly', recorded.documentType === 'assembly')
check('the recorded mate count stays null', recorded.assembly?.mateCount === null && recorded.assembly?.mateCountAvailable === false)

for (const [label, value] of [['part', partOk], ['assembly', asmOk], ['mismatch', asmWrong], ['recorded', recorded]]) {
  check(`${label} result is lossless JSON`, lossless(value))
  check(`${label} result has no undefined-valued expectation fields`, Object.entries(value).every(([, v]) => v !== undefined), Object.keys(value).join(', '))
}

const failures = results.filter((r) => !r.ok)
console.log(`\n${failures.length === 0 ? 'ALL PASS' : `${failures.length} FAILED`} — ${results.length} checks`)
process.exit(failures.length === 0 ? 0 : 1)
