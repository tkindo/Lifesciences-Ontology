import { describe, expect, it } from 'vitest'
import { sectorCatalog } from './sectors'

describe('sector catalog', () => {
  it('contains complete, internally consistent sector models', () => {
    expect(sectorCatalog.map((sector) => sector.id)).toEqual([
      'pharma',
      'food-beverage',
      'oil-gas',
    ])

    for (const sector of sectorCatalog) {
      expect(sector.functions.length).toBeGreaterThanOrEqual(4)
      expect(sector.entities.length).toBeGreaterThanOrEqual(6)

      const capabilityIds = new Set(
        sector.functions.flatMap((businessFunction) =>
          businessFunction.capabilities.map((capability) => capability.id),
        ),
      )

      for (const businessFunction of sector.functions) {
        expect(businessFunction.capabilities.length).toBeGreaterThan(0)
        for (const capability of businessFunction.capabilities) {
          expect(capability.processes.length).toBeGreaterThan(0)
          expect(capability.processes.every((process) => process.steps.length > 0)).toBe(true)
        }
      }

      for (const entity of sector.entities) {
        expect(capabilityIds.has(entity.capabilityId)).toBe(true)
        expect(entity.attributes.length).toBeGreaterThan(0)
      }
    }
  })
})