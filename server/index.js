import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { createApiServer } from './api.js'
import { createExcelStorage, parseProperties } from './excelStorage.js'

const root = process.cwd()
const properties = parseProperties(await readFile(resolve(root, 'application.properties'), 'utf8'))
const workbookPath = resolve(root, properties['excel.file.path'] || './data/habits.xlsx')
const worksheetName = properties['excel.sheet.name'] || 'Habits'
const storage = createExcelStorage(workbookPath, worksheetName)
const server = createApiServer({ storage, workbookPath, worksheetName })

const port = Number(process.env.PORT || 3001)
server.listen(port, '127.0.0.1', () => {
  console.log(`Excel storage API listening on http://127.0.0.1:${port}`)
})