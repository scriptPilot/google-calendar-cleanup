// Automatically deletes events whose end date is older than X days.

/**
 * This function runs the cleanup itself
 * @param {string} calendarName - The name of the calendar to clean up
 * @param {number|string|Date} cutoffDate - Number of days in the past, a "YYYY-MM-DD" string, or a Date object
 */
function runCleanup(calendarName, cutoffDate) {
  const scriptStartTime = Date.now();

  // Log cleanup start
  console.info(`Cleanup started for calendar "${calendarName}".`)

  // Get calendar by name
  let calendar = null
  Calendar.CalendarList.list({ showHidden: true }).items.forEach(cal => {
    if (cal.summaryOverride === calendarName || cal.summary === calendarName) calendar = cal
  })
  if (!calendar) throw new Error(`Calendar ${calendarName} not found.`)

  const timeZone = calendar.timeZone || Session.getScriptTimeZone();

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

  // Align targetDate to midnight in the calendar's timezone
  const tzFormatted = Utilities.formatDate(targetDate, timeZone, "yyyy-MM-dd'T'00:00:00");
  const tzOffset = Utilities.formatDate(targetDate, timeZone, "Z"); // e.g. "+0200"
  targetDate = new Date(`${tzFormatted}${tzOffset.substr(0,3)}:${tzOffset.substr(3,2)}`);
  
  // Get events older than cutoff
  let deletedCount = 0
  let trimmedCount = 0
  let pageToken = null
    do {
      if (Date.now() - scriptStartTime > 270000) {
        console.warn('Execution time limit approaching (4.5 mins). Stopping early to save progress. The next run will continue cleaning.');
        break;
      }

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
        if (event.eventType === 'birthday') return;

        if (!event.start || !event.end) return;

        const parseAllDay = (dStr) => new Date(`${dStr}T00:00:00${tzOffset.substr(0,3)}:${tzOffset.substr(3,2)}`);
        const startDate = event.start.dateTime ? new Date(event.start.dateTime) : parseAllDay(event.start.date)
        const endDate = event.end.dateTime ? new Date(event.end.dateTime) : parseAllDay(event.end.date)
        
        if (startDate >= targetDate) return;

        const eventType = event.recurringEventId ? 'recurring event instance' : 'event'

        if (endDate <= targetDate) {
          // Event is fully in the past
          try {
            Calendar.Events.remove(calendar.id, event.id)
            console.info(`Deleted ${eventType}: "${event.summary || 'Untitled'}" (ended on ${endDate.toISOString()})`)
            deletedCount++
          } catch (error) {
            if (error.details && error.details.code === 410) {
              console.info(`Already deleted ${eventType}: "${event.summary || 'Untitled'}" (ended on ${endDate.toISOString()})`)
            } else if (error.details && error.details.code === 400 && error.details.message && error.details.message.includes('not valid for this event type')) {
              console.info(`Skipped special event type: "${event.summary || 'Untitled'}"`)
            } else {
              console.error(`Failed to delete event: "${event.summary || 'Untitled'}"`, error)
            }
          }
        } else {
          // Event crosses the cutoff date. Trim its start time.
          try {
            const isAllDay = !!event.start.date;
            const patchObj = { end: event.end }; // Explicitly keep the original end date to prevent auto-shifting
            if (isAllDay) {
              patchObj.start = { date: Utilities.formatDate(targetDate, timeZone, "yyyy-MM-dd") };
            } else {
              patchObj.start = { dateTime: targetDate.toISOString() };
            }

            Calendar.Events.patch(patchObj, calendar.id, event.id);
            console.info(`Trimmed ${eventType}: "${event.summary || 'Untitled'}" (new start: ${patchObj.start.date || patchObj.start.dateTime})`)
            trimmedCount++
          } catch (error) {
            if (error.details && error.details.code === 400 && error.details.message && error.details.message.includes('not valid for this event type')) {
              console.info(`Skipped special event type: "${event.summary || 'Untitled'}"`)
            } else {
              console.error(`Failed to trim event: "${event.summary || 'Untitled'}"`, error)
            }
          }
        }
      })

      pageToken = response.nextPageToken
    } while (pageToken !== undefined)

  // Log cleanup end
  console.info(`Cleanup completed. Deleted ${deletedCount} events. Trimmed ${trimmedCount} events.`)
}
