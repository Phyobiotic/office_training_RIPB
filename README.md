# Excel Training Add-in

A simple Office Add-in training project for Excel and Word, with lesson modules loaded into the task pane. By Phyo Ko Ko Thant.

## Quick start

### 1. Prerequisites

Make sure you have:

- Node.js and npm installed
- Microsoft Excel (and optionally Word) installed for testing

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

### Excel

1. Open the pane and confirm the lesson list appears correctly.
2. Open a lesson and try the Tutorial mode first.
3. Then switch to Assignment mode and complete the checklist.
4. Confirm that the live-checking behavior updates as expected while you edit the workbook.

### Word

The manifest can support both Excel and Word. When you run the add-in, you may be asked to choose the host application. That is expected.

Only the Word editing lesson is currently available for Word. It uses the same basic tutorial-and-assignment structure as the Excel lessons, but with Word-specific logic.

### PowerPoint

The manifest now lists Excel, Word, and PowerPoint. Only "Building Your Deck" is built for PowerPoint so far (`lessons/ppt-build.js`), and it is deliberately smaller than the Excel and Word lessons for two practical reasons:

1. No reliable live event exists on this host. Excel has `onChanged` for value and formatting changes, and Word has `onParagraphChanged` and `onParagraphAdded` for content changes. PowerPoint has neither. The related event `Office.EventType.DocumentSelectionChanged` only fires on selection changes, not content edits, and Microsoft has an open issue showing that it is unreliable across recent PowerPoint versions. In practice, live-checking here is a 2.5-second poll doing most of the work rather than a fast event-driven system.
2. The text-reading API is the least certain part of the implementation without real-world testing. `slideHasText()` in `ppt-build.js` searches every shape on a slide for matching text rather than assuming a specific title placeholder, which is safer than guessing shape order but still the part most likely to need adjustment if something does not behave as expected. If a title-text check does not tick, that function is the first place to inspect.

By contrast, `slides.add()` and slide deletion are both well documented and should work reliably.

## Notes and caveats

- The task pane loads lesson files dynamically at runtime.
- The file name and the `id` in the LESSONS array should match exactly.
- If you change webpack settings or add new assets, restart the app with `npm start`.
- Live checking is enabled by default, but some host-specific behavior may differ slightly between Excel and Word.
- Fill-color and other formatting checks can be sensitive to Excel API behavior when a value is empty or unset.

