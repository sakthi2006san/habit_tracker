import { createServer } from 'node:http'
import { access, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { createExcelStorage, parseProperties } from './excelStorage.js'

const root = process.cwd()
const properties = parseProperties(await readFile(resolve(root, 'application.properties'), 'utf8'))
const workbookPath = resolve(root, properties['excel.file.path'] || './data/habits.xlsx')
const worksheetName = properties['excel.sheet.name'] || 'Habits'
const storage = createExcelStorage(workbookPath, worksheetName)
let operationQueue = Promise.resolve()

function serialize(operation) {
  const result = operationQueue.then(operation, operation)
  operationQueue = result.catch(() => {})
  return result
}

function send(response, status, body) {
  response.writeHead(status, { 'content-type': 'application/json; charset=utf-8' })
  response.end(JSON.stringify(body))
}

async function readBody(request) {
  let body = ''
  for await (const chunk of request) body += chunk
  if (body.length > 5_000_000) throw new Error('Request body is too large.')
  return JSON.parse(body)
}

async function workbookExists() {
  try {
    await access(workbookPath)
    return true
  } catch (error) {
    if (error.code === 'ENOENT') return false
    throw error
  }
}

const server = createServer(async (request, response) => {
  if (request.method === 'GET' && request.url === '/api/config') {
    send(response, 200, { workbookPath, worksheetName })
    return
  }

  if (request.method === 'GET' && request.url === '/api/habits') {
    try {
      const initialized = await workbookExists()
      const habits = await serialize(() => storage.readHabits())
      send(response, 200, { habits, initialized })
    } catch (error) {
      send(response, 500, { error: error.message })
    }
    return
  }

  if (request.method === 'PUT' && request.url === '/api/habits') {
    try {
      const { habits } = await readBody(request)
      await serialize(() => storage.writeHabits(habits))
      send(response, 200, { saved: habits.length })
    } catch (error) {
      send(response, 400, { error: error.message })
    }
    return
  }

  send(response, 404, { error: 'Not found.' })
})

const port = Number(process.env.PORT || 3001)
server.listen(port, '127.0.0.1', () => {
  console.log(`Excel storage API listening on http://127.0.0.1:${port}`)
})