import { createServer } from 'node:http'
import { access } from 'node:fs/promises'
import { InvalidHabitDataError } from './excelStorage.js'

class RequestBodyError extends Error {}

function send(response, status, body) {
  response.writeHead(status, { 'content-type': 'application/json; charset=utf-8' })
  response.end(JSON.stringify(body))
}

async function readBody(request) {
  let body = ''
  for await (const chunk of request) {
    body += chunk
    if (body.length > 5_000_000) throw new RequestBodyError('Request body is too large.')
  }
  return JSON.parse(body)
}

async function workbookExists(workbookPath) {
  try {
    await access(workbookPath)
    return true
  } catch (error) {
    if (error.code === 'ENOENT') return false
    throw error
  }
}

function completionHistory(habit) {
  return JSON.stringify(Object.entries(habit.completions ?? {}).sort(([left], [right]) => left.localeCompare(right)))
}

function matchesSavedHabits(requested, saved) {
  return requested.length === saved.length && requested.every((habit, index) => (
    habit.id === saved[index].id && completionHistory(habit) === completionHistory(saved[index])
  ))
}

export function createApiServer({ storage, workbookPath, worksheetName }) {
  let operationQueue = Promise.resolve()

  function serialize(operation) {
    const result = operationQueue.then(operation, operation)
    operationQueue = result.catch(() => {})
    return result
  }

  return createServer(async (request, response) => {
    const pathname = new URL(request.url, 'http://localhost').pathname

    if (request.method === 'GET' && pathname === '/api/health') {
      send(response, 200, { ok: true })
      return
    }

    if (request.method === 'GET' && pathname === '/api/config') {
      send(response, 200, { workbookPath, worksheetName })
      return
    }

    if (request.method === 'GET' && pathname === '/api/habits') {
      try {
        const [initialized, habits] = await Promise.all([
          workbookExists(workbookPath),
          serialize(() => storage.readHabits()),
        ])
        send(response, 200, { habits, initialized })
      } catch (error) {
        send(response, 500, { error: error.message })
      }
      return
    }

    if (request.method === 'PUT' && pathname === '/api/habits') {
      try {
        const body = await readBody(request)
        if (!body || !Array.isArray(body.habits)) {
          throw new InvalidHabitDataError('Request body must include a habits array.')
        }

        const savedHabits = await serialize(async () => {
          await storage.writeHabits(body.habits)
          const saved = await storage.readHabits()
          if (!matchesSavedHabits(body.habits, saved)) {
            throw new Error('Excel write verification failed.')
          }
          return saved
        })
        send(response, 200, { saved: savedHabits.length, verified: true })
      } catch (error) {
        const status = error instanceof InvalidHabitDataError || error instanceof RequestBodyError || error instanceof SyntaxError ? 400 : 500
        send(response, status, { error: error.message })
      }
      return
    }

    send(response, 404, { error: 'Not found.' })
  })
}