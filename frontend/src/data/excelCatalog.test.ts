import { describe, expect, it } from 'vitest'
import { parseSectorWorkbook, writeSectorWorkbook } from './excelCatalog'
import { sectorCatalog } from './sectors'

describe('Excel sector catalog', () => {
  it('round-trips the common workbook template', async () => {
    const source = sectorCatalog[0]
    const workbook = await writeSectorWorkbook(source)
    const parsed = await parseSectorWorkbook(workbook, {
      id: source.id,
      name: source.name,
      eyebrow: source.eyebrow,
      description: source.description,
      accent: source.accent,
    })

    expect(parsed.name).toBe(source.name)
    expect(parsed.functions.map((item) => item.name)).toEqual(
      source.functions.map((item) => item.name),
    )
    expect(parsed.entities).toHaveLength(source.entities.length)
    expect(parsed.functions[0].capabilities[0].processes[0].steps).toEqual(
      source.functions[0].capabilities[0].processes[0].steps,
    )
  })
})