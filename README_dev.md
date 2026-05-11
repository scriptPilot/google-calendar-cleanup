# Development

This is about the development of this repository.

Feel free to open an issue for bugs, feature requests or any other question.

## Requirements

* [Node.js](https://nodejs.org/) and NPM installed
* [Command Line Apps Script Projects](https://github.com/google/clasp) installed globally

## Installation

1. Clone this repository:

   ```
   git clone https://github.com/scriptPilot/google-calendar-cleanup.git
   ```

2. Login to Google Apps Script CLI:

    ```
    clasp login
    ```

3. Create a new Google Apps Script project:

    ```
    clasp create --type standalone --rootDir lib --title "Google Calendar Cleanup"
    ```

## Workflow

1. Apply changes to the code and documentation.
2. Push the changes to the [Cloud IDE](https://script.google.com/) and open the project:

    ```
    clasp push && clasp open-script
    ```
3. Test the changes in the Cloud IDE according to the documentation.
4. Build the `dist/Code.gs` file:

    ```
    node buildscript.js
    ```

5. Commit and push the changes to GitHub.
