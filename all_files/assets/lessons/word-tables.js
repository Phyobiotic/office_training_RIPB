// lessons/word-tables.js
// Syllabus refs 2.3.1-2.3.8.
//
// CONTINUITY: unlike free-floating paragraphs, a table is a single object,
// so most steps here just ask "does ANY table in the document currently
// satisfy this content condition" rather than assuming a fixed position —
// the same table naturally carries forward and accumulates changes across
// steps regardless of visit order. The one step that needs a specific
// *problem* pre-loaded (the duplicate row) only injects it if that
// condition doesn't already hold somewhere, exactly like the duplicate-row
// tasks in xl-format.js and founders-day.js.
//
// Pictures use a tiny inline placeholder image (a 1x1 PNG, resized up to
// 100x100pt on insert) rather than anything fetched externally — each
// picture-related step anchors its own picture to its own marker paragraph
// so the "insert", "resize", and "delete" exercises never compete over the
// same picture.

window.LESSON_MODULES = window.LESSON_MODULES || {};

const WORD_TABLES_IMG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

// Ensures at least one table exists somewhere in the body (creating a
// blank 3x3 with an explicit narrow column width baseline if none does),
// and always returns SOME table — never trims or resets an existing one.
async function wordTablesEnsureAnyTable(context) {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  if (tables.items.length > 0) return tables.items[0];
  const t = context.document.body.insertTable(3, 3, Word.InsertLocation.end);
  for (let c = 0; c < 3; c++) t.getCell(0, c).columnWidth = 55;
  await context.sync();
  return t;
}

async function wordTablesEnsureMarkerPicture(context, markerText) {
  const body = context.document.body;
  const paras = body.paragraphs;
  paras.load("items/text");
  await context.sync();
  let marker = paras.items.find(p => p.text.trim().startsWith(markerText));
  if (marker) return marker;
  body.insertParagraph(markerText, Word.InsertLocation.end);
  await context.sync();
  const pic = body.insertInlinePictureFromBase64(WORD_TABLES_IMG_B64, Word.InsertLocation.end);
  pic.width = 100;
  pic.height = 100;
  await context.sync();
  paras.load("items/text");
  await context.sync();
  return paras.items.find(p => p.text.trim().startsWith(markerText));
}

async function wordTablesEnsureMarkerLine(context, markerText) {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  if (!body.text.includes(markerText)) {
    body.insertParagraph(markerText, Word.InsertLocation.end);
    await context.sync();
  }
}

window.LESSON_MODULES["word-tables"] = {
  title: "Tables and Objects",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "Tables organise data a plain paragraph never could, and pictures make a document feel finished. One document the whole way through — nothing you build in an earlier step gets reset as you move on.",

    prepare: async () => Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
    }),

    steps: [
      {
        title: "Create a table",
        teach: `<b>Insert tab → Table</b>, then drag across the grid (or use Insert Table for exact numbers) to set rows and columns.<br><br>Under the heading below, insert a table with at least <b>3 rows and 3 columns</b>.`,
        done: "Insert → Table, drag the grid — rows and columns appear instantly.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureMarkerLine(context, "Supplies Order Form");
        }),
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          if (tables.items.length === 0) return false;
          tables.items.forEach(t => t.load("rowCount,values"));
          await context.sync();
          return tables.items.some(t => t.rowCount >= 3 && t.values[0] && t.values[0].length >= 3);
        })
      },

      {
        title: "Insert and edit data in a table",
        teach: `Click any cell and type — it works exactly like typing anywhere else, just confined to that cell. <span class="keys">Tab</span> moves to the next cell.<br><br>Type these headers into the first row of your table: <code>Item</code>, <code>Qty</code>, <code>Price</code>.`,
        done: "Click a cell, type, Tab to the next one.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureAnyTable(context);
        }),
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          return tables.items.some(t => {
            const row0 = (t.values[0] || []).map(v => (v || "").trim().toLowerCase());
            return row0[0] === "item" && row0[1] === "qty" && row0[2] === "price";
          });
        })
      },

      {
        title: "Insert a row for a missing item",
        teach: `<b>Right-click a row → Insert → Insert Below/Above</b> (or click in the last cell and press <span class="keys">Tab</span> to grow the table by one row at the bottom).<br><br>This order form is missing <b>Napkins</b>. Insert a new row anywhere in the table and add it: <code>Napkins</code>, qty <code>2</code>, price <code>1.50</code>.`,
        done: "Right-click a row → Insert, or Tab from the very last cell to add one at the end.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureAnyTable(context);
        }),
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          return tables.items.some(t => t.values.some(row => (row[0] || "").trim().toLowerCase() === "napkins"));
        })
      },

      {
        title: "Delete a duplicate row",
        teach: `Click any cell in the row, then <b>right-click → Delete Row</b> (not just Delete on the keyboard — that only clears the text, the row stays). Everything below shifts up.<br><br>Somewhere in your table, <code>Cups</code> got entered twice by mistake. Delete one of the two rows so it appears only once.`,
        done: "Right-click → Delete Cells → Delete entire row (plain Delete only empties the text).",
        setup: async () => Word.run(async (context) => {
          const table = await wordTablesEnsureAnyTable(context);
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          const anyDup = tables.items.some(t => t.values.filter(row => (row[0] || "").trim().toLowerCase() === "cups").length >= 2);
          if (!anyDup) {
            tables.items[0].addRows(Word.InsertLocation.end, 2, [["Cups", "10", "0.50"], ["Cups", "10", "0.50"]]);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          return tables.items.some(t => t.values.filter(row => (row[0] || "").trim().toLowerCase() === "cups").length === 1);
        })
      },

      {
        title: "Modify column width and row height",
        teach: `Drag a column's border to resize it, or a row's bottom border to make it taller. Hovering the border turns the cursor into a double-arrow first.<br><br>Widen any column in your table, or make any row taller — either counts.`,
        done: "Drag a border — column borders resize width, row borders resize height.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureAnyTable(context);
        }),
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => { t.load("values"); t.rows.load("items/preferredHeight"); });
          await context.sync();
          // Keep the SAME cell proxy from creation through to reading it back —
          // calling getCell(row,col) again later returns a fresh, unloaded
          // proxy, not the one that was just loaded.
          const perTable = tables.items.map(t => {
            const cols = (t.values[0] || []).length;
            const cells = [];
            for (let c = 0; c < cols; c++) {
              const cell = t.getCell(0, c);
              cell.load("columnWidth");
              cells.push(cell);
            }
            return { table: t, cells };
          });
          await context.sync();
          return perTable.some(({ table, cells }) => {
            const widened = cells.some(cell => cell.columnWidth > 90);
            const heightened = table.rows.items.some(r => r.preferredHeight > 20);
            return widened || heightened;
          });
        })
      },

      {
        title: "Insert a picture",
        teach: `<b>Insert tab → Pictures</b> (This Device, or Stock Images/Online Pictures) drops an image in at your cursor.<br><br>Click right after the marker line below and insert any picture.`,
        done: "Insert → Pictures places an image wherever your cursor is.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureMarkerLine(context, "Insert your photo here:");
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const marker = paras.items.find(p => p.text.trim().startsWith("Insert your photo here:"));
          if (!marker) return false;
          marker.inlinePictures.load("items");
          await context.sync();
          return marker.inlinePictures.items.length >= 1;
        })
      },

      {
        title: "Select and resize a picture",
        teach: `Click a picture once to select it — square handles appear at its corners and edges. Drag a <b>corner</b> handle to resize while keeping its proportions; dragging a side handle stretches it out of shape.<br><br>A placeholder picture is already below. Click it and resize it — bigger or smaller, either counts.`,
        done: "Corner handle = keeps proportions. Side handle = stretches it.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureMarkerPicture(context, "RESIZE-PRACTICE-IMAGE — resize the picture below:");
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const marker = paras.items.find(p => p.text.trim().startsWith("RESIZE-PRACTICE-IMAGE"));
          if (!marker) return false;
          marker.inlinePictures.load("items/width,items/height");
          await context.sync();
          return marker.inlinePictures.items.some(pic => pic.width !== 100 || pic.height !== 100);
        })
      },

      {
        title: "Delete a picture",
        teach: `Click a picture once to select it (not to place your cursor inside text near it), then press <span class="keys">Delete</span>.<br><br>Click the placeholder picture below and delete it.`,
        done: "Click to select the picture itself, then Delete.",
        setup: async () => Word.run(async (context) => {
          await wordTablesEnsureMarkerPicture(context, "DELETE-PRACTICE-IMAGE — delete the picture below:");
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const marker = paras.items.find(p => p.text.trim().startsWith("DELETE-PRACTICE-IMAGE"));
          if (!marker) return false;
          marker.inlinePictures.load("items");
          await context.sync();
          return marker.inlinePictures.items.length === 0;
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Founders' Day Stall Sign-Up Sheet
     ====================================================================== */
  assignment: {
    hook: "The stall sign-up sheet needs to go on the noticeboard tonight — one stall's missing entirely, another got double-booked, the columns are too narrow to read, and there's no photo of last year's layout.",
    brief: "Fix the table (add the missing stall, remove the duplicate, widen it to fit), then add and size a photo for the layout reference.",
    doneMsg: "Sign-up sheet's ready for the noticeboard. Every table and object move from the tutorial, doing real work this time.",

    setup: async () => {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertParagraph("Founders' Day — Stall Sign-Up Sheet", Word.InsertLocation.end);
        const t = body.insertTable(5, 3, Word.InsertLocation.end, [
          ["Stall", "Volunteer", "Time Slot"],
          ["Popcorn", "Aidan Tan", "9am - 12pm"],
          ["Ring Toss", "Priya Sharma", "9am - 12pm"],
          ["Ring Toss", "Priya Sharma", "9am - 12pm"],
          ["Henna Booth", "Rachel Lim", "12pm - 3pm"]
        ]);
        for (let c = 0; c < 3; c++) t.getCell(0, c).columnWidth = 50;
        body.insertParagraph("Layout reference photo:", Word.InsertLocation.end);
        await context.sync();
      });
    },

    tasks: [
      { ref: "2.3.4", text: "The sheet is missing the <b>Cupcake Sale</b> stall. Insert a new row anywhere with <code>Cupcake Sale</code>, <code>Zul Hakim</code>, <code>12pm - 3pm</code>.",
        hint: "Right-click any row → Insert, then type the details into the new row.",
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          return tables.items.some(t => t.values.some(row => (row[0] || "").trim().toLowerCase() === "cupcake sale"));
        }) },

      { ref: "2.3.4", text: "<b>Ring Toss</b> got entered twice with the same volunteer — delete one of the duplicate rows.",
        hint: "Right-click the row number area of one duplicate → Delete Cells → Delete entire row.",
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          return tables.items.some(t => t.values.filter(row => (row[0] || "").trim().toLowerCase() === "ring toss").length === 1);
        }) },

      { ref: "2.3.5", text: "The columns are too narrow to read comfortably — widen at least one column.",
        hint: "Hover the border between two column headers until it becomes a double-arrow, then drag.",
        check: async () => Word.run(async (context) => {
          const tables = context.document.body.tables;
          tables.load("items");
          await context.sync();
          tables.items.forEach(t => t.load("values"));
          await context.sync();
          // Keep the SAME cell proxy from creation through to reading it back —
          // calling getCell(row,col) again later returns a fresh, unloaded proxy.
          const perTable = tables.items.map(t => {
            const cols = (t.values[0] || []).length;
            const cells = [];
            for (let c = 0; c < cols; c++) {
              const cell = t.getCell(0, c);
              cell.load("columnWidth");
              cells.push(cell);
            }
            return cells;
          });
          await context.sync();
          return perTable.some(cells => cells.some(cell => cell.columnWidth > 90));
        }) },

      { ref: "2.3.6", text: "Insert a picture right after \"<b>Layout reference photo:</b>\" for a visual of last year's stall arrangement.",
        hint: "Click after that line, then Insert → Pictures.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const marker = paras.items.find(p => p.text.trim().startsWith("Layout reference photo:"));
          if (!marker) return false;
          marker.inlinePictures.load("items");
          await context.sync();
          return marker.inlinePictures.items.length >= 1;
        }) },

      { ref: "2.3.8", text: "Resize that photo — it's currently either too small to make out or too large for the page.",
        hint: "Click the picture once, then drag a corner handle.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const marker = paras.items.find(p => p.text.trim().startsWith("Layout reference photo:"));
          if (!marker) return false;
          marker.inlinePictures.load("items/width,items/height");
          await context.sync();
          if (marker.inlinePictures.items.length === 0) return false;
          // Any deliberate, non-trivial size counts — the point is that they
          // touched it, not a specific target dimension.
          return marker.inlinePictures.items.every(pic => pic.width > 0 && pic.height > 0);
        }) }
    ]
  }
};
