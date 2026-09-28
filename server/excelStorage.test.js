import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import ExcelJS from 'exceljs'
import { createExcelStorage, parseProperties } from './excelStorage.js'

let temporaryDirectory

afterEach(async () => {
  if (temporaryDirectory) await rm(temporaryDirectory, { recursive: true, force: true })
  temporaryDirectory = undefined
})

describe('Excel storage', () => {
  it('creates the configured sheet and round-trips habit history', async () => {
    temporaryDirectory = await mkdtemp(join(tmpdir(), 'habit-tracker-'))
    const storage = createExcelStorage(join(temporaryDirectory, 'habits.xlsx'), 'Daily habits')
    const habit = {
      id: 'reading', name: 'Read', icon: 'book', category: 'learning', color: '#4778b9', frequency: 'daily',
      customDays: [], target: 1, createdAt: '2026-09-28T00:00:00.000Z', paused: false, archived: false,
      completions: { '2026-09-28': true, '2026-09-27': true },
    }

    expect(await storage.readHabits()).toEqual([])
    await storage.writeHabits([habit])

    expect(await storage.readHabits()).toEqual([habit])

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(join(temporaryDirectory, 'habits.xlsx'))
    const dailyLog = workbook.getWorksheet('Daily Log')
    expect(dailyLog.getRow(1).values.slice(1)).toEqual(['Date', 'Habit ID', 'Habit', 'Completed'])
    expect([dailyLog.getRow(2), dailyLog.getRow(3)].map((row) => row.values.slice(1))).toEqual([
      ['2026-09-27', 'reading', 'Read', true],
      ['2026-09-28', 'reading', 'Read', true],
    ])
  })

  it('parses key-value properties', () => {
    expect(parseProperties('# comment\nexcel.file.path=./data/habits.xlsx\nexcel.sheet.name= Daily Habits ')).toEqual({
      'excel.file.path': './data/habits.xlsx',
      'excel.sheet.name': 'Daily Habits',
    })
  })
})