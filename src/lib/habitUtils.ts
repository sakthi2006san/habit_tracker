import { addDays, eachDayOfInterval, format, isBefore, startOfDay, subDays } from 'date-fns'
import type { Habit, HabitFrequency, DailyStats } from '../types'

export function isHabit(value: unknown): value is Habit {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Record<string, unknown>
  return typeof candidate.id === 'string' && typeof candidate.name === 'string' && typeof candidate.icon === 'string' && typeof candidate.color === 'string' && typeof candidate.createdAt === 'string' && typeof candidate.paused === 'boolean' && typeof candidate.archived === 'boolean' && Array.isArray(candidate.customDays) && typeof candidate.completions === 'object' && candidate.completions !== null
}

export const dateKey = (date: Date): string => format(date, 'yyyy-MM-dd')

export function isHabitScheduledForDate(habit: Habit, date: Date): boolean {
  if (habit.paused || habit.archived || isBefore(startOfDay(date), startOfDay(new Date(habit.createdAt)))) return false
  const day = date.getDay()
  if (habit.frequency === 'daily') return true
  if (habit.frequency === 'weekdays') return day > 0 && day < 6
  if (habit.frequency === 'weekends') return day === 0 || day === 6
  if (habit.frequency === 'custom') return habit.customDays.includes(day)
  return true
}

const isCompletedOn = (habit: Habit, date: Date): boolean => habit.completions[dateKey(date)] === true

export function calculateCurrentStreak(habit: Habit, referenceDate = new Date()): number {
  if (habit.paused || habit.archived) return 0
  let cursor = startOfDay(referenceDate)
  if (!isHabitScheduledForDate(habit, cursor) || !isCompletedOn(habit, cursor)) {
    cursor = subDays(cursor, 1)
  }
  let streak = 0
  let guard = 0
  while (guard < 500 && isHabitScheduledForDate(habit, cursor) && isCompletedOn(habit, cursor)) {
    streak += 1
    cursor = subDays(cursor, 1)
    guard += 1
  }
  return streak
}

export function calculateLongestStreak(habit: Habit): number {
  const dates = Object.keys(habit.completions).filter((key) => habit.completions[key]).sort()
  let longest = 0
  let current = 0
  let previous: Date | undefined
  dates.forEach((key) => {
    const date = new Date(`${key}T12:00:00`)
    if (!previous) current = 1
    else if (dateKey(addDays(previous, 1)) === key || !isHabitScheduledForDate(habit, previous)) current += 1
    else current = 1
    longest = Math.max(longest, current)
    previous = date
  })
  return longest
}

export function calculateCompletionRate(habit: Habit, startDate: Date, endDate: Date): number {
  const days = eachDayOfInterval({ start: startOfDay(startDate), end: startOfDay(endDate) }).filter((day) => isHabitScheduledForDate(habit, day))
  if (!days.length) return 0
  return Math.round((days.filter((day) => isCompletedOn(habit, day)).length / days.length) * 100)
}

export function buildDailyStats(habits: Habit[], startDate: Date, endDate: Date): DailyStats[] {
  return eachDayOfInterval({ start: startOfDay(startDate), end: startOfDay(endDate) }).map((date) => {
    const active = habits.filter((habit) => isHabitScheduledForDate(habit, date))
    const completed = active.filter((habit) => isCompletedOn(habit, date)).length
    return { date: dateKey(date), completed, total: active.length, percentage: active.length ? Math.round((completed / active.length) * 100) : 0 }
  })
}

export const frequencyLabel = (frequency: HabitFrequency): string => frequency === 'times-per-week' ? 'Weekly target' : frequency[0].toUpperCase() + frequency.slice(1)
