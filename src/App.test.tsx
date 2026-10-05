import { StrictMode } from 'react'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import App from './App'
import { useHabitStore } from './features/habits/habitStore'
import type { Habit } from './types'

const habit: Habit = {
  id: 'workout', name: 'Morning workout', icon: 'x', category: 'fitness', color: '#000', frequency: 'daily',
  customDays: [], target: 1, createdAt: '2026-09-01T00:00:00.000Z', paused: false, archived: false,
  completions: { '2026-09-28': false },
}

describe('app Excel synchronization', () => {
  beforeEach(() => {
    useHabitStore.setState({ habits: [habit], onboardingCompleted: true })
    vi.stubGlobal('fetch', vi.fn())
  })

  it('saves a completion made before the initial Excel read finishes', async () => {
    const pendingLoads: Array<(response: Response) => void> = []
    const fetchMock = vi.mocked(fetch)
    fetchMock.mockImplementation((_input, init) => {
      if (init?.method === 'PUT') return Promise.resolve({ ok: true, json: async () => ({ saved: 1, verified: true }) } as Response)
      return new Promise<Response>((resolve) => pendingLoads.push(resolve))
    })

    render(<StrictMode><App /></StrictMode>)
    fireEvent.click(await screen.findByRole('button', { name: 'Complete Morning workout' }))

    for (const resolveLoad of pendingLoads) {
      resolveLoad({ ok: true, json: async () => ({ habits: [habit], initialized: true }) } as Response)
    }

    await waitFor(() => {
      const save = fetchMock.mock.calls.find(([, init]) => init?.method === 'PUT')
      expect(save).toBeDefined()
      expect(JSON.parse(String(save?.[1]?.body)).habits[0].completions['2026-09-28']).toBe(true)
    })
  })
})