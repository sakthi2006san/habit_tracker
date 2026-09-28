import { describe, expect, it } from 'vitest'
import { calculateCompletionRate, calculateCurrentStreak, isHabit, isHabitScheduledForDate } from './habitUtils'
import type { Habit } from '../types'

const habit = (overrides: Partial<Habit> = {}): Habit => ({
  id: 'test', name: 'Test habit', icon: 'x', category: 'health', color: '#000', frequency: 'daily', customDays: [], target: 1,
  createdAt: '2026-01-01T12:00:00.000Z', paused: false, archived: false, completions: {}, ...overrides,
})

describe('habit scheduling and metrics', () => {
  it('calculates a daily streak from the reference date', () => {
    const completions = { '2026-01-05': true, '2026-01-06': true, '2026-01-07': true }
    expect(calculateCurrentStreak(habit({ completions }), new Date('2026-01-07T12:00:00'))).toBe(3)
  })

  it('supports weekdays and excludes weekends', () => {
    const weekdayHabit = habit({ frequency: 'weekdays' })
    expect(isHabitScheduledForDate(weekdayHabit, new Date('2026-01-09T12:00:00'))).toBe(true)
    expect(isHabitScheduledForDate(weekdayHabit, new Date('2026-01-10T12:00:00'))).toBe(false)
  })

  it('calculates completion rate only across scheduled days', () => {
    const weekdayHabit = habit({ frequency: 'weekdays', completions: { '2026-01-05': true, '2026-01-06': true, '2026-01-07': false } })
    expect(calculateCompletionRate(weekdayHabit, new Date('2026-01-05T12:00:00'), new Date('2026-01-11T12:00:00'))).toBe(40)
  })

  it('rejects malformed imported habit data', () => {
    expect(isHabit({ id: 'missing-fields' })).toBe(false)
    expect(isHabit(habit())).toBe(true)
  })
})