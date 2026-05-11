# Google Calendar Cleanup

Clean up past Google Calendar events on a regular basis to keep your calendar tidy.

Made with Google Apps Script, related to [Google Calendar Correction](https://github.com/scriptPilot/google-calendar-correction).

## Installation
1. [Backup all Google Calendars](https://calendar.google.com/calendar/u/0/r/settings/export) to be able to restore them if something went wrong.
2. Open [Google Apps Script](https://script.google.com/) and create a new project `Calendar Cleanup`.
3. Replace the `Code.gs` file content with [this code](https://raw.githubusercontent.com/scriptPilot/google-calendar-cleanup/main/dist/Code.gs).
4. Click at the `+` next to `Services`, add `Google Calendar API` `v3` as `Calendar`.

## Usage
The following examples are based on assumed calendars `Work` and `Family`.

### Cleanup
1. Click the `+` next to `Files` to add a new script file `onCalendarCleanup`:

    ```js
    function onCalendarCleanup() {
      // Run cleanup, delete events whose end date was more than 30 days in the past
      runCleanup('Work', 30)
    }
    ```

2. Save the changes and run the `onCalendarCleanup` function manually.

    - Allow the prompt and grant the requested calendar access.

3. On the left menu, select "Triggers" and add a new trigger:

    - Choose which function to run: `onCalendarCleanup`
    - Select event source: `Time-driven`
    - Select type of time based trigger: `Day timer`
    - Select time of day: e.g. `Midnight to 1am`

Now, the script will run automatically every day and delete old events from the `Work` calendar.

### Multiple Calendars
If you want to clean up multiple calendars, simply call the function multiple times.

```js
function onCalendarCleanup() {
  runCleanup('Work', 30)
  runCleanup('Family', 60)
}
```

### Helper Functions
There are a couple of helper function available to support defining the cutoff date.

For past days:

```js
startOfWeek(offset = 0)       
startOfMonth(offset = 0)
startOfQuarter(offset = 0)
startOfHalfyear(offset = 0)
startOfYear(offset = 0)
```

Example - delete events from the beginning of this month:

```js
function onCalendarCleanup() {
  runCleanup('Work', startOfMonth())
}
```

Example - delete events older than the start of last year:

```js
function onCalendarCleanup() {
  runCleanup('Work', startOfYear(1))
}
```

## Update
To update the script version, replace the `Code.gs` file content with [this code](https://raw.githubusercontent.com/scriptPilot/google-calendar-cleanup/main/dist/Code.gs).

## Deinstallation
Remove the Google Apps Script project. This will also remove all triggers.

## Details
- **Recurring Events:** Recurring events are expanded into individual sessions, meaning past sessions are safely deleted without affecting future sessions.
- **Event Trimming:** Events that start before but end after the cutoff date are not deleted. Instead, their start time is trimmed to the cutoff date, ensuring a clean past calendar.

## Support
Feel free to open an [issue](https://github.com/scriptPilot/google-calendar-cleanup/issues) for bugs, feature requests or any other question.
