import { access, mkdir, rename } from 'node:fs/promises'
import { dirname } from 'node:path'
import ExcelJS from 'exceljs'

const columns = [
  { header: 'Habit ID', key: 'id', width: 22 },
  { header: 'Name', key: 'name', width: 28 },
  { header: 'Description', key: 'description', width: 42 },
  { header: 'Icon', key: 'icon', width: 12 },
  { header: 'Category', key: 'category', width: 18 },
  { header: 'Color', key: 'color', width: 14 },
  { header: 'Frequency', key: 'frequency', width: 18 },
  { header: 'Custom Days JSON', key: 'customDays', width: 24 },
  { header: 'Weekly Target', key: 'weeklyTarget', width: 16 },
  { header: 'Target', key: 'target', width: 12 },
  { header: 'Reminder Time', key: 'reminderTime', width: 18 },
  { header: 'Created At', key: 'createdAt', width: 28 },
  { header: 'Paused', key: 'paused', width: 12 },
  { header: 'Archived', key: 'archived', width: 12 },
  { header: 'Completions JSON', key: 'completions', width: 48 },
]
const dailyLogSheetName = 'Daily Log'
const dailyLogColumns = [
  { header: 'Date', key: 'date', width: 16 },
  { header: 'Habit ID', key: 'habitId', width: 22 },
  { header: 'Habit', key: 'habitName', width: 28 },
  { header: 'Completed', key: 'completed', width: 14 },
]

const requiredFields = ['id', 'name', 'icon', 'category', 'color', 'frequency', 'target', 'createdAt']

function parseJson(value, fallback) {
  if (typeof value !== 'string' || !value) return fallback
  try {
    return JSON.parse(value)
  } catch {
    return fallback
  }
}

function rowToHabit(row) {
  const value = Object.fromEntries(columns.map(({ key }, index) => [key, row.getCell(index + 1).value]))
  const habit = {
    id: String(value.id ?? ''),
    name: String(value.name ?? ''),
    icon: String(value.icon ?? ''),
    category: String(value.category ?? ''),
    color: String(value.color ?? ''),
    frequency: String(value.frequency ?? ''),
    target: Number(value.target),
    createdAt: String(value.createdAt ?? ''),
    paused: value.paused === true || String(value.paused).toLowerCase() === 'true',
    archived: value.archived === true || String(value.archived).toLowerCase() === 'true',
    customDays: parseJson(value.customDays, []),
    completions: parseJson(value.completions, {}),
  }

  if (value.description != null && value.description !== '') habit.description = String(value.description)
  if (value.weeklyTarget != null && value.weeklyTarget !== '') habit.weeklyTarget = Number(value.weeklyTarget)
  if (value.reminderTime != null && value.reminderTime !== '') habit.reminderTime = String(value.reminderTime)
  return habit
}

function validateHabits(habits) {
  if (!Array.isArray(habits) || !habits.every((habit) => habit && typeof habit === 'object'
    && requiredFields.every((field) => field in habit)
    && typeof habit.id === 'string' && typeof habit.name === 'string'
    && typeof habit.target === 'number' && typeof habit.createdAt === 'string'
    && habit.completions && typeof habit.completions === 'object')) {
    throw new Error('Habit data must be an array of valid habit records.')
  }
}

export function createExcelStorage(filePath, sheetName) {
  async function openWorksheet() {
    const workbook = new ExcelJS.Workbook()
    let fileExists = true
    try {
      await access(filePath)
    } catch (error) {
      if (error.code === 'ENOENT') fileExists = false
      else throw error
    }
    if (fileExists) await workbook.xlsx.readFile(filePath)

    let worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) worksheet = workbook.addWorksheet(sheetName)

    const firstRow = worksheet.getRow(1)
    const headers = columns.map(({ header }) => header)
    if (worksheet.rowCount === 0 || firstRow.cellCount === 0) {
      worksheet.columns = columns
      firstRow.values = headers
    } else {
      const existingHeaders = headers.map((_, index) => String(firstRow.getCell(index + 1).value ?? ''))
      if (existingHeaders.some((header, index) => header !== headers[index])) {
        throw new Error(`Worksheet "${sheetName}" does not have the expected habit columns.`)
      }
    }

    let dailyLog = workbook.getWorksheet(dailyLogSheetName)
    if (!dailyLog) dailyLog = workbook.addWorksheet(dailyLogSheetName)

    const dailyLogHeaders = dailyLogColumns.map(({ header }) => header)
    const dailyLogFirstRow = dailyLog.getRow(1)
    if (dailyLog.rowCount === 0 || dailyLogFirstRow.cellCount === 0) {
      dailyLog.columns = dailyLogColumns
      dailyLogFirstRow.values = dailyLogHeaders
    } else {
      const existingHeaders = dailyLogHeaders.map((_, index) => String(dailyLogFirstRow.getCell(index + 1).value ?? ''))
      if (existingHeaders.some((header, index) => header !== dailyLogHeaders[index])) {
        throw new Error(`Worksheet "${dailyLogSheetName}" does not have the expected daily log columns.`)
      }
    }

    return { workbook, worksheet, dailyLog }
  }

  return {
    async readHabits() {
      const { worksheet } = await openWorksheet()
      const habits = []
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1 && row.getCell(1).value != null) habits.push(rowToHabit(row))
      })
      return habits
    },
    async writeHabits(habits) {
      validateHabits(habits)
      const { workbook, worksheet, dailyLog } = await openWorksheet()
      if (worksheet.rowCount > 1) worksheet.spliceRows(2, worksheet.rowCount - 1)
      if (dailyLog.rowCount > 1) dailyLog.spliceRows(2, dailyLog.rowCount - 1)

      const dailyCompletions = habits.flatMap((habit) => Object.entries(habit.completions ?? {})
        .filter(([, completed]) => completed === true)
        .map(([date]) => ({ date, habitId: habit.id, habitName: habit.name, completed: true })))
        .sort((left, right) => left.date.localeCompare(right.date)
          || left.habitName.localeCompare(right.habitName)
          || left.habitId.localeCompare(right.habitId))

      for (const habit of habits) {
        worksheet.addRow({
          ...habit,
          customDays: JSON.stringify(habit.customDays ?? []),
          completions: JSON.stringify(habit.completions ?? {}),
        })
      }
      for (const completion of dailyCompletions) dailyLog.addRow(completion)
      await mkdir(dirname(filePath), { recursive: true })
      const temporaryPath = `${filePath}.tmp`
      await workbook.xlsx.writeFile(temporaryPath)
      await rename(temporaryPath, filePath)
    },
  }
}

export function parseProperties(source) {
  return Object.fromEntries(source.split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith('#'))
    .map((line) => {
      const separator = line.indexOf('=')
      if (separator < 0) return [line, '']
      return [line.slice(0, separator).trim(), line.slice(separator + 1).trim()]
    }))
}