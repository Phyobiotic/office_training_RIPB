# Excel Training Add-in

A simple Office Add-in training project for Excel and Word, with lesson modules loaded into the task pane. By Phyo Ko Ko Thant.

## Quick start

### 1. Prerequisites

Make sure you have:

- Node.js and npm installed
- Microsoft Excel, Word and Powerpoint installed for testing

### 2. Install dependencies

From the project root, run:

```bash
cd excel_training
npm install
```

### 3. Run the add-in

Start the local development server with one of these commands:

```bash
cd excel_training
npm start
npm run start:excel
npm run start:word
npm run start:ppt
```

- `npm start` and `npm run start:excel` launch Excel
- `npm run start:word` launches Word
- `npm run start:ppt` launches PowerPoint

## Project structure

- [manifest.xml](manifest.xml) — the add-in manifest
- [excel_training/src/taskpane/taskpane.html](excel_training/src/taskpane/taskpane.html) — the main task pane shell and lesson list
- [excel_training/assets/lessons](excel_training/assets/lessons) — lesson JavaScript files loaded dynamically at runtime
- [excel_training/src/taskpane/lessons](excel_training/src/taskpane/lessons) — the source location used for lesson modules in the project layout

## Dependencies

### Runtime dependencies

- core-js
- regenerator-runtime

### Development dependencies

- @babel/core
- @babel/preset-env
- @types/office-js
- @types/office-runtime
- acorn
- babel-loader
- copy-webpack-plugin
- eslint-plugin-office-addins
- file-loader
- html-loader
- html-webpack-plugin
- office-addin-cli
- office-addin-debugging
- office-addin-dev-certs
- office-addin-lint
- office-addin-manifest
- office-addin-prettier-config
- os-browserify
- process
- source-map-loader
- webpack
- webpack-cli
- webpack-dev-server

## Adding a new module or lesson

To add a new lesson or module, follow these steps:

1. Create a new lesson JavaScript file in:
   - [excel_training/assets/lessons](excel_training/assets/lessons)
   - Use the naming convention: `assets/lessons/<module-id>.js`

2. Add a new entry to the LESSONS array in:
   - [excel_training/src/taskpane/taskpane.html](excel_training/src/taskpane/taskpane.html)

   Example:

   ```js
   { id: "my-module", title: "My New Module", ready: true }
   ```

3. Reorder lessons by changing the order of entries in the LESSONS array.
   - The first item appears first in the list.
   - The order of the array controls the display order.

4. Change the lesson status by updating the `ready` value:
   - `true` = the lesson is ready and clickable
   - `false` = the lesson appears as "Coming soon" and is not clickable

5. Keep the filename and the `id` value aligned.
   - If the lesson id is `my-module`, the file should be named `my-module.js`.

## Testing the training content

1. Test tutorials to ensure all makes sense and there are no mistakes in instructions.
2. Test assignments to ensure all makes sense and there are no irregularities in scenarios.
3. Look out for any improvements that can be made to syllabus and send to Phyo.

## Notes and caveats

- The task pane loads lesson files dynamically at runtime.
- The file name and the `id` in the LESSONS array should match exactly.
- If you change webpack settings or add new assets, restart the app with `npm start`.
- Live checking is enabled by default, but some host-specific behavior may differ slightly between Excel and Word.
- Fill-color and other formatting checks can be sensitive to Excel API behavior when a value is empty or unset.

