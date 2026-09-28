import { beforeEach, describe, expect, it } from 'vitest'
import { useHabitStore } from './habitStore'
import type { Habit } from '../../types'

const newHabit: Habit = { id: 'store-test', name: 'Store test', icon: 'x', category: 'personal', color: '#000', frequency: 'daily', customDays: [], target: 1, createdAt: new Date().toISOString(), paused: false, archived: false, completions: {} }

describe('habit store actions', () => {
  beforeEach(() => useHabitStore.setState({ habits: [], onboardingCompleted: true }))

  it('creates and completes a habit', () => {
    useHabitStore.getState().addHabit(newHabit)
    useHabitStore.getState().completeHabit(newHabit.id, '2026-01-07')
    expect(useHabitStore.getState().habits[0].completions['2026-01-07']).toBe(true)
  })

  it('deletes a habit and its history', () => {
    useHabitStore.getState().addHabit(newHabit)
    useHabitStore.getState().deleteHabit(newHabit.id)
    expect(useHabitStore.getState().habits).toHaveLength(0)
  })
})