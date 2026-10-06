/**
 * emit-recipe-json.mjs — turn the two validated candidates into the exact JSON
 * a recipe is stored as, so they can be promoted without a running GUI:
 * copy them into the plugin's recipeDir (default `<scratch>/recipes`) or paste
 * the same object into `solidworks_recipes { action: "save", ... }`.
 *
 *   node test/emit-recipe-json.mjs
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { validateRecipe } from '../lib/recipes.js'

const repoRoot = resolve(import.meta.dirname, '..', '..', '..', '..')
const outDir = join(repoRoot, 'parts', 'recipes')
mkdirSync(outDir, { recursive: true })

const candidates = [
  {
    name: 'plate_linear_holes',
    description: 'Mounting plate with a linear pattern of through holes (extrude + blind cut from the +Z face + FeatureLinearPattern4).',
    file: 'parts/recipes/plate_linear_holes.vbs.txt',
    parameters: {
      PLATE_X: { value: 0.065, description: 'plate length along X (m)' },
      PLATE_Y: { value: 0.05, description: 'plate width along Y (m)' },
      THK: { value: 0.012, description: 'plate thickness (m)' },
      HOLE_D: { value: 0.009, description: 'hole diameter (m)' },
      HX: { value: 0.05, description: 'seed hole X (m)' },
      HY: { value: 0.012, description: 'seed hole Y (m)' },
      PITCH: { value: 0.012, description: 'pattern pitch along X (m)' },
      COUNT: { value: 3, description: 'number of holes' },
    },
    verify: { bodyCount: 1, featureTypes: ['Extrude', 'Cut', 'Pattern'] },
  },
  {
    name: 'disc_multicontour_cut',
    description: 'Disc with a central bore and N radial slots removed by one multi-contour sketch and one cut (no axis/pattern API needed).',
    file: 'parts/recipes/disc_multicontour_cut.vbs.txt',
    parameters: {
      R_OUT: { value: 0.06, description: 'disc radius (m)' },
      THK: { value: 0.012, description: 'disc thickness (m)' },
      R_BORE: { value: 0.014, description: 'bore radius (m)' },
      SLOT_R0: { value: 0.038, description: 'slot inner radius (m)' },
      SLOT_R1: { value: 0.046, description: 'slot outer radius (m)' },
      SLOT_HW: { value: 0.004, description: 'slot half width (m)' },
      SLOTS: { value: 8, description: 'number of radial slots' },
    },
    verify: { bodyCount: 1, featureTypes: ['Extrude', 'Cut'] },
  },
]

for (const candidate of candidates) {
  const script = readFileSync(join(repoRoot, candidate.file), 'utf8')
  const validated = validateRecipe(
    { name: candidate.name, description: candidate.description, script, parameters: candidate.parameters, verify: candidate.verify },
    `candidate ${candidate.name}`,
  )
  if (!validated.ok) {
    console.error(`INVALID ${candidate.name}: ${validated.reason}`)
    process.exitCode = 1
    continue
  }
  // the exact payload saveRecipeFile writes
  const payload = {
    name: validated.recipe.name,
    description: validated.recipe.description,
    parameters: validated.recipe.parameters,
    verify: validated.recipe.verify,
    script: validated.recipe.script,
  }
  const path = join(outDir, `${candidate.name}.json`)
  writeFileSync(path, `${JSON.stringify(payload, null, 2)}\n`, 'utf8')
  console.log(`wrote ${path} (${Object.keys(payload.parameters).length} parameters, verify ${JSON.stringify(payload.verify)})`)
}
