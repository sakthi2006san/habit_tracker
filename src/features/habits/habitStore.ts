import { create } from 'zustand'
import { persist } from 'zustand/middleware'
import { dateKey } from '../../lib/habitUtils'
import { defaultPreferences, demoHabits } from '../../data/demoData'
import type { AppState, Habit, UserPreferences } from '../../types'

interface HabitStore extends AppState {
  excelStorageStatus: 'connecting' | 'connected' | 'unavailable'
  setExcelStorageStatus: (status: HabitStore['excelStorageStatus']) => void
  addHabit: (habit: Habit) => void
  updateHabit: (id: string, updates: Partial<Habit>) => void
  deleteHabit: (id: string) => void
  completeHabit: (id: string, date?: string) => void
  uncompleteHabit: (id: string, date?: string) => void
  pauseHabit: (id: string) => void
  archiveHabit: (id: string) => void
  setPreferences: (preferences: Partial<UserPreferences>) => void
  setOnboardingCompleted: (completed: boolean) => void
  reset: () => void
}

export const useHabitStore = create<HabitStore>()(persist((set) => ({
  habits: demoHabits,
  preferences: defaultPreferences,
  onboardingCompleted: true,
  excelStorageStatus: 'connecting',
  setExcelStorageStatus: (excelStorageStatus) => set({ excelStorageStatus }),
  addHabit: (habit) => set((state) => ({ habits: [...state.habits, habit] })),
  updateHabit: (id, updates) => set((state) => ({ habits: state.habits.map((habit) => habit.id === id ? { ...habit, ...updates } : habit) })),
  deleteHabit: (id) => set((state) => ({ habits: state.habits.filter((habit) => habit.id !== id) })),
  completeHabit: (id, date = dateKey(new Date())) => set((state) => ({ habits: state.habits.map((habit) => habit.id === id ? { ...habit, completions: { ...habit.completions, [date]: true } } : habit) })),
  uncompleteHabit: (id, date = dateKey(new Date())) => set((state) => ({ habits: state.habits.map((habit) => { if (habit.id !== id) return habit; const completions = { ...habit.completions }; delete completions[date]; return { ...habit, completions } }) })),
  pauseHabit: (id) => set((state) => ({ habits: state.habits.map((habit) => habit.id === id ? { ...habit, paused: !habit.paused } : habit) })),
  archiveHabit: (id) => set((state) => ({ habits: state.habits.map((habit) => habit.id === id ? { ...habit, archived: true } : habit) })),
  setPreferences: (preferences) => set((state) => ({ preferences: { ...state.preferences, ...preferences } })),
  setOnboardingCompleted: (onboardingCompleted) => set({ onboardingCompleted }),
  reset: () => set({ habits: demoHabits, preferences: defaultPreferences, onboardingCompleted: true }),
}), {
  name: 'momentum-habits',
  partialize: ({ habits, preferences, onboardingCompleted }) => ({ habits, preferences, onboardingCompleted }),
}))
