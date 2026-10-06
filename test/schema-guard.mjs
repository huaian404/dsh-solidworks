// Schema guard. The Harness rejects any tool schema outside its enforced JSON
// Schema subset, and a violation aborts the WHOLE plugin entry at composition
// time (`JsonSchemaError: unsupported JSON schema`), not just the one tool.
//
// This mirrors the subset from @deepseek-ai/dsh-tools so an authoring mistake
// fails here instead of on the user's machine. Subset, verbatim:
//   constraint keywords: type, oneOf, properties, required,
//                        additionalProperties, items, enum, const
//   annotation keywords: description, title, default, examples
//   types: object, array, string, number, integer, boolean, null
// `items`, `properties`, `required`, `additionalProperties`, `enum` and `const`
// must be value-correct for their parent `type`, and `items` entries must be
// bare schema objects — annotation keywords live on the annotated node only.
import { apply } from '../lib/index.js'

const CONSTRAINT = new Set(['type', 'oneOf', 'properties', 'required', 'additionalProperties', 'items', 'enum', 'const'])
const ANNOTATION = new Set(['description', 'title', 'default', 'examples'])
const TYPES = ['object', 'array', 'string', 'number', 'integer', 'boolean', 'null']

const isRecord = (v) => typeof v === 'object' && v !== null && !Array.isArray(v)

function check(node, path, out) {
  if (!isRecord(node)) {
    out.push(`${path} must be a schema object`)
    return
  }
  for (const key of Object.keys(node)) {
    if (!CONSTRAINT.has(key) && !ANNOTATION.has(key)) {
      out.push(`${path}.${key} is not a supported keyword`)
    }
  }
  if ('description' in node && typeof node.description !== 'string') out.push(`${path}.description must be a string`)
  const hasType = 'type' in node
  const hasOneOf = 'oneOf' in node
  if (hasType && hasOneOf) {
    out.push(`${path} cannot declare both type and oneOf`)
    return
  }
  if (!hasType && !hasOneOf) return
  if (hasOneOf) {
    if (!Array.isArray(node.oneOf) || node.oneOf.length < 2) out.push(`${path}.oneOf must hold at least two schemas`)
    else node.oneOf.forEach((sub, i) => check(sub, `${path}.oneOf[${i}]`, out))
    return
  }
  if (!TYPES.includes(node.type)) {
    out.push(`${path}.type must be one of ${TYPES.join('/')}`)
    return
  }
  const shape = {
    properties: ['object'], required: ['object'], additionalProperties: ['object'],
    items: ['array'],
    enum: ['string', 'number', 'integer', 'boolean', 'null'],
    const: ['string', 'number', 'integer', 'boolean', 'null'],
  }
  for (const [key, owners] of Object.entries(shape)) {
    if (key in node && !owners.includes(node.type)) out.push(`${path}.${key} requires type ${owners.join('/')}`)
  }
  if ('additionalProperties' in node && typeof node.additionalProperties !== 'boolean') {
    out.push(`${path}.additionalProperties must be a boolean`)
  }
  if (isRecord(node.properties)) {
    for (const [key, sub] of Object.entries(node.properties)) check(sub, `${path}.properties.${key}`, out)
  }
  if ('items' in node) {
    if (Array.isArray(node.items)) out.push(`${path}.items must be a single schema, not a tuple`)
    else check(node.items, `${path}.items`, out)
  }
}

const captured = []
apply({
  tools: { register: (definition) => captured.push(definition) },
  systemPrompt: { section: () => {} },
  logger: { info: () => {} },
}, {})

let failed = 0
for (const tool of captured) {
  const violations = []
  check(tool.parameters, `${tool.name}.parameters`, violations)

  // The host registry asserts output.schema on register, so a missing schema
  // fails the WHOLE plugin activation ("schema must be a schema object").
  // An annotation-only `{}` is the legal way to declare unconstrained JSON.
  if (tool.output === undefined || typeof tool.output !== 'object') {
    violations.push(`${tool.name}.output must be declared`)
  } else {
    if (typeof tool.output.render !== 'function') violations.push(`${tool.name}.output.render must be a function`)
    if (tool.output.schema === undefined) {
      violations.push(`${tool.name}.output.schema is REQUIRED — the registry asserts it (use {} for unconstrained JSON)`)
    } else {
      check(tool.output.schema, `${tool.name}.output.schema`, violations)
    }
  }

  if (violations.length === 0) {
    console.log(`OK    ${tool.name} (${Object.keys(tool.parameters.properties).length} params)`)
  } else {
    failed++
    console.log(`FAIL  ${tool.name}`)
    for (const v of violations) console.log(`        ${v}`)
  }
}
console.log(failed === 0 ? '\nALL TOOL SCHEMAS INSIDE THE SUPPORTED SUBSET' : `\n${failed} tool schema(s) would break composition`)
process.exit(failed === 0 ? 0 : 1)
