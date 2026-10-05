import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { createApiServer } from './api.js'
import { createExcelStorage } from './excelStorage.js'

let temporaryDirectory
let server
let baseUrl

afterEach(async () => {
  if (server?.listening) await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()))
  if (temporaryDirectory) await rm(temporaryDirectory, { recursive: true, force: true })
  server = undefined
  temporaryDirectory = undefined
})

async function startApi() {
  temporaryDirectory = await mkdtemp(join(tmpdir(), 'habit-api-'))
  const workbookPath = join(temporaryDirectory, 'habits.xlsx')
  server = createApiServer({
    storage: createExcelStorage(workbookPath, 'Habits'),
    workbookPath,
    worksheetName: 'Habits',
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  baseUrl = `http://127.0.0.1:${server.address().port}`
}

describe('habit API', () => {
  it('verifies a save by reading the completion back from Excel', async () => {
    await startApi()
    const habit = {
      id: 'reading', name: 'Read', icon: 'book', category: 'learning', color: '#4778b9', frequency: 'daily',
      customDays: [], target: 1, createdAt: '2026-09-28T00:00:00.000Z', paused: false, archived: false,
      completions: { '2026-09-28': true },
    }

    const response = await fetch(`${baseUrl}/api/habits`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ habits: [habit] }),
    })

    expect(response.status).toBe(200)
    expect(await response.json()).toEqual({ saved: 1, verified: true })
    const saved = await fetch(`${baseUrl}/api/habits`).then((result) => result.json())
    expect(saved.habits).toEqual([habit])
    expect(await fetch(`${baseUrl}/api/health`).then((result) => result.json())).toEqual({ ok: true })
  })

  it('returns a client error for malformed habit payloads', async () => {
    await startApi()
    const response = await fetch(`${baseUrl}/api/habits`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ habits: 'invalid' }),
    })

    expect(response.status).toBe(400)
    expect(await response.json()).toEqual({ error: 'Request body must include a habits array.' })
  })
})