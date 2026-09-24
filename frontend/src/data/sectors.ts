import catalog from './sectors.json'

export type Process = {
  id: string
  name: string
  description: string
  owner: string
  stageId?: string
  stageName?: string
  parentProcessId?: string
  parentProcessName?: string
  parentProcessDescription?: string
  contributingFunctions?: string[]
  contributingCapabilities?: string[]
  dataObjectIds?: string[]
  inputs: string[]
  outputs: string[]
  steps: string[]
}

export type Capability = {
  id: string
  name: string
  description: string
  systems: string[]
  processes: Process[]
}

export type BusinessFunction = {
  id: string
  name: string
  description: string
  owner: string
  capabilities: Capability[]
}

export type DataEntity = {
  id: string
  name: string
  description: string
  system: string
  capabilityId: string
  attributes: string[]
  relationships: string[]
  primaryFunction?: string
  primaryCapability?: string
  l1Id?: string
  l1Name?: string
  l2Id?: string
  l2Name?: string
  l3Id?: string
  l3Name?: string
}

export type Sector = {
  id: 'pharma' | 'food-beverage' | 'oil-gas'
  name: string
  eyebrow: string
  description: string
  accent: string
  capabilityCount?: number
  overview?: string[]
  stages?: Array<{ id: string; order: number; name: string; description: string }>
  entityRelationships?: Array<{ from: string; relationship: string; to: string }>
  processRelationships?: Array<{ fromId: string; relationship: string; toId: string; kind: string }>
  functions: BusinessFunction[]
  entities: DataEntity[]
}

export const sectorCatalog = catalog as Sector[]

export function uniqueCapabilityCount(sector: Sector) {
  const documentedCount = sector.overview?.find((line) => /Functions & Capabilities/i.test(line))?.match(/(\d+)\s+rows/i)?.[1]
  return documentedCount ? Number(documentedCount) : sector.capabilityCount ?? new Set(sector.functions.flatMap((businessFunction) => businessFunction.capabilities.map((capability) => capability.name.trim().toLowerCase()))).size
}

export function isSectorCatalog(value: unknown): value is Sector[] {
  if (!Array.isArray(value) || value.length === 0) return false

  return value.every((sector) => {
    if (!sector || typeof sector !== 'object') return false
    const candidate = sector as Partial<Sector>
    const capabilityIds = new Set(
      candidate.functions?.flatMap((businessFunction) =>
        businessFunction.capabilities.map((capability) => capability.id),
      ) ?? [],
    )

    return Boolean(
      candidate.id &&
        candidate.name &&
        candidate.accent &&
        candidate.functions?.every(
          (businessFunction) =>
            businessFunction.id &&
            businessFunction.name &&
            businessFunction.capabilities.length > 0 &&
            businessFunction.capabilities.every(
              (capability) =>
                capability.id &&
                capability.name &&
                capability.processes.length > 0 &&
                capability.processes.every(
                  (process) => process.id && process.name && process.steps.length > 0,
                ),
            ),
        ) &&
        candidate.entities?.every(
          (entity) =>
            entity.id &&
            entity.name &&
            entity.attributes.length > 0 &&
            capabilityIds.has(entity.capabilityId),
        ),
    )
  })
}