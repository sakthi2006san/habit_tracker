export type HabitFrequency = 'daily' | 'weekdays' | 'weekends' | 'custom' | 'times-per-week'
export type HabitCategory = 'health' | 'fitness' | 'mindfulness' | 'learning' | 'productivity' | 'personal' | 'finance' | 'social' | 'other'

export interface Habit {
  id: string
  name: string
  description?: string
  icon: string
  category: HabitCategory
  color: string
  frequency: HabitFrequency
  customDays: number[]
  weeklyTarget?: number
  target: number
  reminderTime?: string
  createdAt: string
  paused: boolean
  archived: boolean
  completions: Record<string, boolean>
}

export interface HabitStats { currentStreak: number; longestStreak: number; completionRate: number; totalCompletions: number }
export interface DailyStats { date: string; completed: number; total: number; percentage: number }
export interface WeeklyStats { week: string; completionRate: number }
export interface UserPreferences { theme: 'light' | 'dark' | 'system'; firstDayOfWeek: 0 | 1; notificationsEnabled: boolean }
export interface AppState { habits: Habit[]; preferences: UserPreferences; onboardingCompleted: boolean }
export interface ChartDataPoint { label: string; value: number }
