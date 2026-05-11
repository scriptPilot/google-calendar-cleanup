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
          singleEvents: false // Do not expand recurring events, we want to skip master recurring events
        }
      )
      
      const events = response.items || []
      
      events.forEach(event => {
        let isPastEvent = false;
        let eventType = 'event';
        let logSuffix = '';

        if (event.recurrence) {
          // It's a master recurring event. Check if it has any occurrences on or after the target date.
          try {
            const instancesResponse = Calendar.Events.instances(
              calendar.id, 
              event.id, 
              {
                timeMin: targetDate.toISOString(),
                maxResults: 1
              }
            )
            // If there are no future instances, the series has completely ended in the past
            if (!instancesResponse.items || instancesResponse.items.length === 0) {
              isPastEvent = true;
              eventType = 'recurring event series';
            }
          } catch (error) {
            console.error(`Failed to fetch instances for recurring event: "${event.summary || 'Untitled'}"`, error)
          }
        } else {
          // It's a single event, check its end date
          const endDate = event.end.dateTime ? new Date(event.end.dateTime) : new Date(event.end.date)
          if (endDate < targetDate) {
            isPastEvent = true;
            logSuffix = ` (ended on ${endDate.toISOString()})`;
          }
        }
        
        if (isPastEvent) {
          try {
            Calendar.Events.remove(calendar.id, event.id)
            console.info(`Deleted ${eventType}: "${event.summary || 'Untitled'}"${logSuffix}`)
            deletedCount++
          } catch (error) {
            console.error(`Failed to delete ${eventType}: "${event.summary || 'Untitled'}"`, error)
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
