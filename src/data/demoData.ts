import { subDays, format } from 'date-fns'
import type { Habit, UserPreferences } from '../types'

const completionHistory = (seed: number, days = 48): Record<string, boolean> => {
  const result: Record<string, boolean> = {}
  for (let index = 0; index < days; index += 1) {
    const date = subDays(new Date(), index)
    result[format(date, 'yyyy-MM-dd')] = (index * 7 + seed * 11) % 10 > (seed % 3 + 1)
  }
  return result
}

const habit = (id: string, name: string, icon: string, category: Habit['category'], color: string, frequency: Habit['frequency'], seed: number, description: string): Habit => ({
  id, name, icon, category, color, frequency, customDays: [], target: 1, createdAt: subDays(new Date(), 50).toISOString(), paused: false, archived: false, description, completions: completionHistory(seed),
})

export const demoHabits: Habit[] = [
  habit('workout', 'Morning workout', '◒', 'fitness', '#e86a4a', 'daily', 1, 'Start the day with movement and energy.'),
  habit('reading', 'Read 10 pages', '▤', 'learning', '#7c70d8', 'daily', 2, 'A small page count that compounds over time.'),
  habit('meditation', 'Meditation', '◌', 'mindfulness', '#3e9d8b', 'daily', 3, 'Create a little room between you and the noise.'),
  habit('water', 'Drink 8 glasses of water', '⌁', 'health', '#4b8dcc', 'daily', 4, 'Keep your baseline energy steady.'),
  habit('deep-work', 'Deep work block', '◈', 'productivity', '#d19a42', 'weekdays', 5, 'Protect one focused block for meaningful work.'),
  habit('typescript', 'Learn TypeScript', '◇', 'learning', '#4778b9', 'weekdays', 6, 'Make a little progress on a durable skill.'),
  habit('sleep', 'Sleep before 11 PM', '☾', 'personal', '#926fc1', 'daily', 7, 'End the day in a way your tomorrow will appreciate.'),
]

export const defaultPreferences: UserPreferences = { theme: 'light', firstDayOfWeek: 1, notificationsEnabled: true }
