import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { writeSectorWorkbook } from '../src/data/excelCatalog'
import { sectorCatalog } from '../src/data/sectors'

const outputDirectory = resolve('public/data')
await mkdir(outputDirectory, { recursive: true })

for (const sector of sectorCatalog.filter((item) => item.id !== 'pharma')) {
  const workbook = await writeSectorWorkbook(sector)
  await writeFile(resolve(outputDirectory, `${sector.id}.xlsx`), workbook)
  console.log(`Generated ${sector.id}.xlsx`)
}