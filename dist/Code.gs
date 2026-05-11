// Google Calendar Cleanup, build on 2026-05-11
// Source: https://github.com/scriptPilot/google-calendar-cleanup

function startOfWeek(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  const day = d.getDay()
  const diff = d.getDate() - day + (day === 0 ? -6 : 1)
  d.setDate(diff - (offset * 7))
  return d
}

function startOfMonth(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  d.setMonth(d.getMonth() - offset)
  return d
}

function startOfQuarter(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  const quarterMonth = Math.floor(d.getMonth() / 3) * 3
  d.setMonth(quarterMonth - (offset * 3))
  return d
}

function startOfHalfyear(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  const halfyearMonth = Math.floor(d.getMonth() / 6) * 6
  d.setMonth(halfyearMonth - (offset * 6))
  return d
}

function startOfYear(offset = 0) {
  const d = new Date()
  d.setHours(0, 0, 0, 0)
  d.setDate(1)
  d.setMonth(0)
  d.setFullYear(d.getFullYear() - offset)
  return d
}


// This function resets the script properties (e.g., if a script gets stuck in a waiting state)
function resetScript() {
  PropertiesService.getUserProperties().deleteAllProperties()  
  console.log('Script reset done.')
}


// Automatically deletes events whose end date is older than X days.

/**
 * This function runs the cleanup itself
 * @param {string} calendarName - The name of the calendar to clean up
 * @param {number|string|Date} cutoffDate - Number of days in the past, a "YYYY-MM-DD" string, or a Date object
 */
function runCleanup(calendarName, cutoffDate) {
  // Log cleanup start
  console.info(`Cleanup started for calendar "${calendarName}".`)

  // Get calendar by name
  let calendar = null
  Calendar.CalendarList.list({ showHidden: true }).items.forEach(cal => {
    if (cal.summaryOverride === calendarName || cal.summary === calendarName) calendar = cal
  })
  if (!calendar) throw new Error(`Calendar ${calendarName} not found.`)

  // Calculate cutoff date based on parameter
  let targetDate = cutoffDate
  if (Number.isInteger(targetDate)) {
    const dateObj = new Date()
    dateObj.setHours(0, 0, 0, 0)
    targetDate = new Date(dateObj.setDate(dateObj.getDate() - targetDate))
  } else if (typeof targetDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
    targetDate = new Date(parseInt(targetDate.substr(0, 4)), parseInt(targetDate.substr(5, 2)) - 1, parseInt(targetDate.substr(8, 2)))
  } else if (!(targetDate instanceof Date)) {
    targetDate = new Date(targetDate)
  }
  
  // Lock the script to avoid overlapping executions (up to 30 Min)
  const lock = LockService.getUserLock()
  try {
    lock.waitLock(30 * 60 * 1000)

    // Allow only one waiting script per cleanup job
    const waitingKey = calendar.id + ':cleanup_waiting'
    const waitingValue = PropertiesService.getUserProperties().getProperty(waitingKey)
    if (waitingValue === 'yes') {
      console.info('Script call cancelled because another cleanup script call is already waiting.')
      return
    } else {
      PropertiesService.getUserProperties().setProperty(waitingKey, 'yes')
    }

    // Get events older than cutoff
    let deletedCount = 0
    let pageToken = null
    do {
      const response = Calendar.Events.list(
        calendar.id,
        {
          pageToken,
          showDeleted: false,
          timeMax: targetDate.toISOString(),
          singleEvents: true // Expand recurring events so we can delete past instances
        }
      )
      
      const events = response.items || []
      
      events.forEach(event => {
        if (event.status === 'cancelled') return;

        if (!event.end) return;

        const endDate = event.end.dateTime ? new Date(event.end.dateTime) : new Date(event.end.date)
        if (endDate >= targetDate) return;

        try {
          Calendar.Events.remove(calendar.id, event.id)
          const eventType = event.recurringEventId ? 'recurring event instance' : 'event'
          console.info(`Deleted ${eventType}: "${event.summary || 'Untitled'}" (ended on ${endDate.toISOString()})`)
          deletedCount++
        } catch (error) {
          if (error.details && error.details.code === 410) {
            console.info(`Already deleted ${event.recurringEventId ? 'recurring event instance' : 'event'}: "${event.summary || 'Untitled'}" (ended on ${endDate.toISOString()})`)
          } else {
            console.error(`Failed to delete event: "${event.summary || 'Untitled'}"`, error)
          }
        }
      })

      pageToken = response.nextPageToken
    } while (pageToken !== undefined)

    // Reset waiting script value
    PropertiesService.getUserProperties().setProperty(waitingKey, 'no')

    // Log cleanup end
    console.info(`Cleanup completed. Deleted ${deletedCount} events.`)
  } finally {
    // Always release the lock
    lock.releaseLock()
  }
}
