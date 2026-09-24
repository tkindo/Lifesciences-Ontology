import ExcelJS from 'exceljs'
import type { BusinessFunction, DataEntity, Process, Sector } from './sectors'

export type SectorIdentity = Pick<Sector, 'id' | 'name' | 'eyebrow' | 'description' | 'accent'>

const SHEETS = {
  overview: 'Overview',
  stages: 'L1 - Stages',
  processes: 'L2 - Processes',
  substeps: 'L3 - Sub-steps',
  entities: 'Entity Catalog',
  entityRelationships: 'Entity Relationships',
  mapping: 'Entity-to-L3 Mapping',
  taxonomy: 'Functions & Capabilities',
  relationships: 'Relationships',
  processSteps: 'L3 - Process Steps',
} as const

function text(value: ExcelJS.CellValue | undefined) {
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') {
    if ('result' in value) return String(value.result ?? '')
    if ('richText' in value) return value.richText.map((part) => part.text).join('')
    if ('text' in value) return String(value.text)
  }
  return String(value).trim()
}

function records(sheet: ExcelJS.Worksheet | undefined) {
  if (!sheet) return []
  const headers = sheet.getRow(1).values as ExcelJS.CellValue[]
  const output: Array<Record<string, string>> = []
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    const record: Record<string, string> = {}
    for (let column = 1; column < headers.length; column += 1) {
      record[text(headers[column])] = text(row.getCell(column).value)
    }
    if (Object.values(record).some(Boolean)) output.push(record)
  })
  return output
}

function slug(value: string) {
  return value.toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
}

function splitList(value: string | undefined) {
  if (!value) return []
  return value.split(/;|\n/).map((item) => item.trim()).filter(Boolean)
}

function capabilityKey(functionName: string, capabilityName: string) {
  return `${functionName}\u0000${capabilityName}`
}

function capabilityId(sectorId: string, functionName: string, capabilityName: string) {
  return `${sectorId}-${slug(functionName)}-${slug(capabilityName)}`
}

function requireSheet(workbook: ExcelJS.Workbook, name: string) {
  const sheet = workbook.getWorksheet(name)
  if (!sheet) throw new Error(`The workbook is missing the required "${name}" sheet.`)
  return sheet
}

export async function parseSectorWorkbook(data: ArrayBuffer | Uint8Array, identity: SectorIdentity): Promise<Sector> {
  const workbook = new ExcelJS.Workbook()
  const workbookData = data instanceof Uint8Array
    ? data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength)
    : data
  await workbook.xlsx.load(workbookData as never)

  const taxonomyRows = records(requireSheet(workbook, SHEETS.taxonomy))
  const l2Rows = records(requireSheet(workbook, SHEETS.processes))
  const l3Rows = records(requireSheet(workbook, SHEETS.substeps))
  const stepRows = records(requireSheet(workbook, SHEETS.processSteps))
  const entityRows = records(requireSheet(workbook, SHEETS.entities))
  const mappingRows = records(requireSheet(workbook, SHEETS.mapping))
  const entityRelationshipRows = records(requireSheet(workbook, SHEETS.entityRelationships))

  const l2ById = new Map(l2Rows.map((row) => [row['Process ID'], row]))
  const stepsById = new Map(stepRows.map((row) => [row['Sub-step ID'], [row['Step 1'], row['Step 2'], row['Step 3'], row['Step 4'], row['Step 5']].filter(Boolean)]))
  const l3ById = new Map(l3Rows.map((row) => [row['Sub-step ID'], row]))
  const mappingByEntity = new Map(mappingRows.map((row) => [row.Entity, row]))
  const dataObjectIdsByL3 = new Map<string, string[]>()
  mappingRows.forEach((row) => {
    if (!row.Entity || !row['Primary L3 ID']) return
    const current = dataObjectIdsByL3.get(row['Primary L3 ID']) ?? []
    current.push(slug(row.Entity))
    dataObjectIdsByL3.set(row['Primary L3 ID'], current)
  })

  const taxonomy = new Map<string, { functionName: string; capabilityName: string }>()
  taxonomyRows.forEach((row) => {
    if (row['Function Name'] && row.Capability) taxonomy.set(capabilityKey(row['Function Name'], row.Capability), { functionName: row['Function Name'], capabilityName: row.Capability })
  })
  l3Rows.forEach((row) => {
    if (row['Primary Function'] && row['Primary Capability']) taxonomy.set(capabilityKey(row['Primary Function'], row['Primary Capability']), { functionName: row['Primary Function'], capabilityName: row['Primary Capability'] })
  })

  const processesByCapability = new Map<string, Process[]>()
  l3Rows.forEach((row) => {
    const functionName = row['Primary Function']
    const capabilityName = row['Primary Capability']
    if (!row['Sub-step ID'] || !functionName || !capabilityName) return
    const parent = l2ById.get(row['Parent Process ID'])
    const key = capabilityKey(functionName, capabilityName)
    const process: Process = {
      id: row['Sub-step ID'],
      name: row.L3,
      description: row.Objective,
      owner: functionName,
      stageId: row['Stage #'] ? `S${row['Stage #']}` : undefined,
      stageName: row['Stage Name'],
      parentProcessId: row['Parent Process ID'],
      parentProcessName: row['Parent Process Name'] || parent?.L2,
      parentProcessDescription: parent?.Objective,
      contributingFunctions: splitList(row['Contributing Functions']),
      contributingCapabilities: splitList(row['Contributing Capabilities']),
      dataObjectIds: dataObjectIdsByL3.get(row['Sub-step ID']) ?? [],
      inputs: splitList(row.Inputs),
      outputs: splitList(row['Outputs / Deliverables']),
      steps: stepsById.get(row['Sub-step ID']) ?? [row.L3],
    }
    const current = processesByCapability.get(key) ?? []
    current.push(process)
    processesByCapability.set(key, current)
  })

  const functionOrder: string[] = []
  taxonomyRows.forEach((row) => {
    if (row['Function Name'] && !functionOrder.includes(row['Function Name'])) functionOrder.push(row['Function Name'])
  })
  const groupedTaxonomy = new Map<string, string[]>()
  taxonomy.forEach(({ functionName, capabilityName }) => {
    if (!functionOrder.includes(functionName)) functionOrder.push(functionName)
    const capabilities = groupedTaxonomy.get(functionName) ?? []
    if (!capabilities.includes(capabilityName)) capabilities.push(capabilityName)
    groupedTaxonomy.set(functionName, capabilities)
  })

  const functions: BusinessFunction[] = functionOrder.map((functionName) => ({
    id: `${identity.id}-${slug(functionName)}`,
    name: functionName,
    description: `Processes and governed data owned by ${functionName}.`,
    owner: functionName,
    capabilities: (groupedTaxonomy.get(functionName) ?? []).map((capabilityName) => {
      const processList = processesByCapability.get(capabilityKey(functionName, capabilityName)) ?? []
      const systems = new Set<string>()
      mappingRows.forEach((row) => {
        if (row['Primary Function'] === functionName && row['Primary Capability'] === capabilityName && row.System) systems.add(row.System)
      })
      return {
        id: capabilityId(identity.id, functionName, capabilityName),
        name: capabilityName,
        description: processList[0]?.parentProcessName ? `Includes ${processList[0].parentProcessName} and related sub-steps.` : `Operational processes for ${capabilityName}.`,
        systems: Array.from(systems),
        processes: processList,
      }
    }).filter((capability) => capability.processes.length > 0),
  })).filter((businessFunction) => businessFunction.capabilities.length > 0)

  const relationshipsByEntity = new Map<string, string[]>()
  const entityRelationships = entityRelationshipRows.map((row) => ({ from: row['From Entity'], relationship: row.Relationship, to: row['To Entity'] }))
  entityRelationships.forEach((relationship) => {
    const outgoing = relationshipsByEntity.get(relationship.from) ?? []
    outgoing.push(`${relationship.relationship.replace(/_/g, ' ').toLowerCase()} ${relationship.to}`)
    relationshipsByEntity.set(relationship.from, outgoing)
    const incoming = relationshipsByEntity.get(relationship.to) ?? []
    incoming.push(`${relationship.from} ${relationship.relationship.replace(/_/g, ' ').toLowerCase()}`)
    relationshipsByEntity.set(relationship.to, incoming)
  })

  const entities: DataEntity[] = entityRows.map((row) => {
    const mapping = mappingByEntity.get(row.Entity)
    const l3 = mapping ? l3ById.get(mapping['Primary L3 ID']) : undefined
    const functionName = mapping?.['Primary Function'] || l3?.['Primary Function'] || functions[0]?.name || ''
    const capabilityName = mapping?.['Primary Capability'] || l3?.['Primary Capability'] || functions[0]?.capabilities[0]?.name || ''
    return {
      id: slug(row.Entity),
      name: row.Entity,
      description: row.Description,
      system: row.System,
      capabilityId: capabilityId(identity.id, functionName, capabilityName),
      attributes: row.Attributes.split(',').map((item) => item.trim()).filter(Boolean),
      relationships: relationshipsByEntity.get(row.Entity) ?? [],
      primaryFunction: functionName,
      primaryCapability: capabilityName,
      l1Id: l3?.['Stage #'] ? `S${l3['Stage #']}` : undefined,
      l1Name: l3?.['Stage Name'],
      l2Id: l3?.['Parent Process ID'],
      l2Name: l3?.['Parent Process Name'] || (l3?.['Parent Process ID'] ? l2ById.get(l3['Parent Process ID'])?.L2 : undefined),
      l3Id: mapping?.['Primary L3 ID'],
      l3Name: l3?.L3,
    }
  })

  const overviewSheet = requireSheet(workbook, SHEETS.overview)
  const overview: string[] = []
  overviewSheet.eachRow((row) => {
    const values = (row.values as ExcelJS.CellValue[]).slice(1).map((value) => text(value)).filter(Boolean)
    if (values.length) overview.push(values.join(' — '))
  })

  const stages = records(requireSheet(workbook, SHEETS.stages)).map((row) => ({ id: row['Stage ID'], order: Number(row.Order), name: row['Stage Name'], description: row['Objective / Description'] }))
  const processRelationships = records(workbook.getWorksheet(SHEETS.relationships)).map((row) => ({ fromId: row['From ID'], relationship: row.Relationship, toId: row['To ID'], kind: row.Kind }))

  const capabilityCount = new Set(
    taxonomyRows
      .map((row) => row.Capability.trim().toLowerCase())
      .filter(Boolean),
  ).size

  return { ...identity, overview, stages, entityRelationships, processRelationships, capabilityCount, functions, entities }
}

function addRows(workbook: ExcelJS.Workbook, name: string, headers: string[], rows: Array<Array<string | number>>) {
  const sheet = workbook.addWorksheet(name)
  sheet.addRow(headers)
  rows.forEach((row) => sheet.addRow(row))
  sheet.views = [{ state: 'frozen', ySplit: 1 }]
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } }
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF202522' } }
  sheet.columns.forEach((column) => { column.width = 24 })
  return sheet
}

export async function writeSectorWorkbook(sector: Sector): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()
  const overview = workbook.addWorksheet(SHEETS.overview)
  overview.addRow([`${sector.name} End-to-End Process Inventory — with Data Ontology Mapping`])
  overview.addRow(['Organization', 'L1 (Stage) -> L2 (Process) -> L3 (Sub-step), with ontology entities mapped to their home sub-step.'])
  overview.addRow(['How to use', 'Start from Entity-to-L3 Mapping for data ownership. Start from L2/L3 for process execution and data touchpoints.'])
  ;(sector.overview ?? []).forEach((line) => overview.addRow(['Source note', line]))
  overview.columns = [{ width: 24 }, { width: 100 }]

  const allProcesses = sector.functions.flatMap((func) => func.capabilities.flatMap((capability) => capability.processes.map((process) => ({ func, capability, process }))))
  const parentProcesses = new Map<string, (typeof allProcesses)[number]>()
  allProcesses.forEach((item) => parentProcesses.set(item.process.parentProcessId ?? item.process.id, item))

  addRows(workbook, SHEETS.stages, ['Stage ID', 'Order', 'Stage Name', 'Objective / Description'], (sector.stages ?? [{ id: 'S1', order: 1, name: sector.name, description: sector.description }]).map((stage) => [stage.id, stage.order, stage.name, stage.description]))
  addRows(workbook, SHEETS.processes, ['Process ID', 'Stage #', 'Stage Name', 'L2', 'Primary Function', 'Primary Capability', 'Contributing Functions', 'Objective', 'Inputs', 'Outputs / Deliverables', 'References', 'Gate?'], Array.from(parentProcesses.values()).map(({ func, capability, process }) => [process.parentProcessId ?? process.id, process.stageId?.replace(/^S/, '') ?? 1, process.stageName ?? sector.name, process.parentProcessName ?? process.name, func.name, capability.name, '', process.description, process.inputs.join('; '), process.outputs.join('; '), '', '']))
  addRows(workbook, SHEETS.substeps, ['Sub-step ID', 'Stage #', 'Stage Name', 'Parent Process ID', 'Parent Process Name', 'L3', 'Primary Function', 'Primary Capability', 'Contributing Functions', 'Objective', 'Inputs', 'Outputs / Deliverables', 'References', 'Gate?', 'Has Entity?'], allProcesses.map(({ func, capability, process }) => [process.id, process.stageId?.replace(/^S/, '') ?? 1, process.stageName ?? sector.name, process.parentProcessId ?? process.id, process.parentProcessName ?? process.name, process.name, func.name, capability.name, '', process.description, process.inputs.join('; '), process.outputs.join('; '), '', '', sector.entities.some((entity) => entity.capabilityId === capability.id) ? 'Yes' : 'No']))
  addRows(workbook, SHEETS.entities, ['Entity', 'System', 'Attributes', 'Description'], sector.entities.map((entity) => [entity.name, entity.system, entity.attributes.join(', '), entity.description]))
  addRows(workbook, SHEETS.entityRelationships, ['From Entity', 'Relationship', 'To Entity', 'From System', 'To System'], (sector.entityRelationships ?? []).map((relationship) => [relationship.from, relationship.relationship, relationship.to, sector.entities.find((entity) => entity.name === relationship.from)?.system ?? '', sector.entities.find((entity) => entity.name === relationship.to)?.system ?? '']))
  addRows(workbook, SHEETS.mapping, ['Entity', 'System', 'Primary L3 ID', 'Primary L3 Name', 'Primary L2', 'Primary Stage', 'Primary Function', 'Primary Capability', 'Contributing L3s', 'Contributing Functions', 'Reclassified vs. earlier build?'], sector.entities.map((entity) => {
    const item = allProcesses.find(({ capability }) => capability.id === entity.capabilityId)
    return [entity.name, entity.system, item?.process.id ?? '', item?.process.name ?? '', item ? `${item.process.parentProcessId ?? item.process.id} (${item.process.parentProcessName ?? item.process.name})` : '', item?.process.stageName ?? '', item?.func.name ?? '', item?.capability.name ?? '', '', '', '']
  }))
  addRows(workbook, SHEETS.taxonomy, ['Function Code', 'Function Name', 'Capability'], sector.functions.flatMap((func) => func.capabilities.map((capability) => [slug(func.name).toUpperCase().replace(/-/g, '_'), func.name, capability.name])))
  addRows(workbook, SHEETS.relationships, ['From ID', 'From Name', 'Relationship', 'To ID', 'To Name', 'Kind'], (sector.processRelationships ?? []).map((relationship) => [relationship.fromId, '', relationship.relationship, relationship.toId, '', relationship.kind]))
  addRows(workbook, 'New Entities Detail', ['Entity', 'System', 'Attributes', 'Description / Gap Closed', 'Home L3 (creates/owns)', 'Home L3 Name'], [])
  addRows(workbook, SHEETS.processSteps, ['Sub-step ID', 'Sub-step Name', 'Step 1', 'Step 2', 'Step 3', 'Step 4', 'Step 5', 'Step Count'], allProcesses.map(({ process }) => [process.id, process.name, ...Array.from({ length: 5 }, (_, index) => process.steps[index] ?? ''), process.steps.length]))

  const result = await workbook.xlsx.writeBuffer()
  return new Uint8Array(result)
}