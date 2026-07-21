// lessons/xl-format.js
// Syllabus refs 3.2.1-3.2.2, 3.3.1-3.3.4.
//
// tutorial = stepper (one mechanic per step, each self-contained load)
// assignment = one scenario, full checklist (shape unchanged)

window.LESSON_MODULES = window.LESSON_MODULES || {};

// Excel's default font has changed over time (Calibri for years, Aptos/Aptos
// Narrow on newer Microsoft 365 builds) and varies by install — so "did they
// change the font" can't safely check against one hardcoded name. This checks
// against every default name we know of instead of guessing which one a given
// machine has.
const DEFAULT_FONT_NAMES = ["calibri", "aptos", "aptos narrow", "aptos display", "aptos serif", "aptos mono"];
function isDefaultFontName(name) {
  const n = (name || "").toLowerCase().trim();
  return DEFAULT_FONT_NAMES.includes(n);
}

window.LESSON_MODULES["xl-format"] = {
  title: "Formatting Cells and Sheet Structure",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The moves that turn a raw grid of data into something you'd actually hand to a teacher. Each step loads its own little practice patch.",

    steps: [
      {
        title: "Font size and font face",
        teach: `A title should look like a title. Select the cell, then use the <b>font size</b> box and the <b>font name</b> box on the Home tab (top-left, near Bold).<br><br><b>A1</b> holds a heading in the plain default font. Make it at least <b>20pt</b> and change it to any font that isn't whatever your default already is.`,
        done: "Bigger size + a chosen font = it reads as a heading, not just another cell.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1").values = [["Investiture Day — VIP Guest List"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
          r.format.font.load("size,name");
          await context.sync();
          return r.format.font.size >= 20 && !isDefaultFontName(r.format.font.name);
        })
      },

      {
        title: "Bold, italic, and combining them",
        teach: `Formatting stacks — a cell can be bold <i>and</i> italic at once. The buttons (and shortcuts <span class="keys">Ctrl/Cmd+B</span>, <span class="keys">Ctrl/Cmd+I</span>) each toggle on or off independently.<br><br><b>A1</b> is the Guest of Honour's name. Make it <b>both bold and italic</b> so it stands out from ordinary guests.`,
        done: "Bold and italic are independent toggles — stack them for real emphasis.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1").values = [["Dr. Wong Li Hua"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
          r.format.font.load("bold,italic");
          await context.sync();
          return r.format.font.bold === true && r.format.font.italic === true;
        })
      },

      {
        title: "Delete a whole row",
        teach: `To remove an entire row (not just clear its contents), click the <b>row number</b> on the far left — that selects the whole row — then <b>right-click → Delete</b>. Everything below shifts up to close the gap.<br><br>Row 2 here is a duplicate of row 1. Select row 2 by its number and delete it, so only one <code>Mr. Lim Wei Jie</code> remains.`,
        done: "Click the row number, right-click, Delete — the row's gone and the rest closes up.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1:A3").values = [["Mr. Lim Wei Jie"], ["Mr. Lim Wei Jie"], ["Mrs. Chandra Devi"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:A5");
          r.load("values");
          await context.sync();
          const vals = r.values.map(row => row[0]);
          return vals.filter(v => v === "Mr. Lim Wei Jie").length === 1 && vals.includes("Mrs. Chandra Devi");
        })
      },

      {
        title: "Insert a row",
        teach: `Inserting is the mirror image: <b>right-click a row number → Insert</b> pushes a blank row in <i>above</i> it, shoving everything down.<br><br>You need to slot a new guest between the two names here. Right-click row 2's number, choose Insert, then type <code>Prof. Ahmad Zaki</code> into the new blank <b>A2</b>.`,
        done: "Right-click a row number → Insert drops a blank row in above it.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1:A2").values = [["Datin Sarah Yusof"], ["Mrs. Chandra Devi"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:A3");
          r.load("values");
          await context.sync();
          const vals = r.values.map(row => row[0]);
          return vals[0] === "Datin Sarah Yusof" && vals[1] === "Prof. Ahmad Zaki" && vals[2] === "Mrs. Chandra Devi";
        })
      },

      {
        title: "Widen a column",
        teach: `When text is cut off, widen the column: hover the boundary line between two column letters (say A and B) until the cursor becomes a double-arrow, then <b>drag right</b>. Or <b>double-click</b> that boundary to auto-fit it to the longest entry.<br><br>Column A here has a long name that's overflowing. Widen column A so it fits.`,
        done: "Drag the boundary between column letters — or double-click it to auto-fit.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1").values = [["Prof. Dr. Ahmad Zaki bin Abdullah"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A:A");
          r.format.load("columnWidth");
          await context.sync();
          return r.format.columnWidth > 90;
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Full Table Seating Overhaul
     ====================================================================== */
  assignment: {
    hook: "The Board Secretary hastily typed up the full Founders' Day table seating list at 11pm — three sponsors got double-booked, two tables are missing entirely, and none of it is legible enough to print.",
    brief: "Fix the structural mess (duplicates, missing rows), then make it print-ready.",
    doneMsg: "Seating chart is fixed and legible. Every move from the tutorial, now on a real mess.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H25").clear(Excel.ClearApplyTo.all);
        sheet.getRange("A1").values = [["Founders' Day — Table Seating"]];
        sheet.getRange("A3:C3").values = [["Guest", "Sponsor Tier", "Table No."]];
        sheet.getRange("A4:C13").values = [
          ["Mr. Tan Kok Wei", "Gold", 1],
          ["Mdm. Rosnah Ibrahim", "Gold", 1],
          ["Mr. Tan Kok Wei", "Gold", 1],
          ["Ms. Grace Ong", "Silver", 2],
          ["Mr. Kevin D'Souza", "Silver", 2],
          ["Ms. Grace Ong", "Silver", 2],
          ["Mdm. Halimah Yacob", "Bronze", 3],
          ["Mr. Vincent Teo", "Bronze", 3],
          ["Mdm. Priya Menon", "Bronze", 3],
          ["Mr. Jason Lee", "Bronze", 3]
        ];
        sheet.getRange("F4").values = [["Missing: Dato' Ismail Hashim, Gold, Table 1 — seat with Mdm. Rosnah Ibrahim"]];
        sheet.getRange("F5").values = [["Missing: Mrs. Elaine Foo, Silver, Table 2 — seat with Mr. Kevin D'Souza"]];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.2.1", text: "Format the title in <b>A1</b>: at least 22pt and a non-default font.",
        hint: "Same move as the tutorial, just a bigger number this time.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
          r.format.font.load("size,name");
          await context.sync();
          return r.format.font.size >= 22 && !isDefaultFontName(r.format.font.name);
        }) },

      { ref: "3.3.1 / 3.3.3", text: "Two guests are double-booked (Mr. Tan Kok Wei and Ms. Grace Ong each appear twice). Select and delete both duplicate rows.",
        hint: "Delete one row at a time — deleting a row shifts everything below it up by one.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A4:A16");
          r.load("values");
          await context.sync();
          const names = r.values.map(row => row[0]);
          return names.filter(n => n === "Mr. Tan Kok Wei").length === 1
              && names.filter(n => n === "Ms. Grace Ong").length === 1;
        }) },

      { ref: "3.3.3", text: "Both notes in column F describe a missing guest. Insert a row for each, seated next to the guest named in the note.",
        hint: "Dato' Ismail Hashim goes next to Mdm. Rosnah Ibrahim; Mrs. Elaine Foo goes next to Mr. Kevin D'Souza.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A4:C18");
          r.load("values");
          await context.sync();
          const rows = r.values;
          const findNext = (name) => {
            const idx = rows.findIndex(row => row[0] === name);
            return idx !== -1 && idx + 1 < rows.length ? rows[idx + 1] : null;
          };
          const afterRosnah = findNext("Mdm. Rosnah Ibrahim");
          const afterKevin = findNext("Mr. Kevin D'Souza");
          const okIsmail = afterRosnah && afterRosnah[0] === "Dato' Ismail Hashim" && afterRosnah[1] === "Gold" && Number(afterRosnah[2]) === 1;
          const okElaine = afterKevin && afterKevin[0] === "Mrs. Elaine Foo" && afterKevin[1] === "Silver" && Number(afterKevin[2]) === 2;
          return !!okIsmail && !!okElaine;
        }) },

      { ref: "3.3.2", text: "Select the whole <b>Sponsor Tier</b> column and bold it.",
        hint: "Click the column header itself, not just the data cells.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B:B");
          r.format.font.load("bold");
          await context.sync();
          return r.format.font.bold === true;
        }) },

      { ref: "3.2.2", text: "Bold every Gold-tier guest's name — there should be three once the missing guest is added, and they won't be in adjacent rows, so Ctrl+click each.",
        hint: "Find every row where Sponsor Tier is \"Gold\", then Ctrl+click each name cell before bolding.",
        check: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("A4:C18");
          r.load("values,rowIndex");
          await context.sync();
          const goldRows = r.values.map((row, i) => ({ row, i })).filter(x => x.row[1] === "Gold");
          if (goldRows.length < 3) return false;
          const checks = await Promise.all(goldRows.map(x => Excel.run(async (ctx) => {
            const cell = ctx.workbook.worksheets.getActiveWorksheet().getCell(r.rowIndex + x.i, 0);
            cell.format.font.load("bold");
            await ctx.sync();
            return cell.format.font.bold === true;
          })));
          return checks.every(Boolean);
        }) },

      { ref: "3.3.4", text: "Widen column A so every guest name fits without truncating.",
        hint: "Double-click the boundary between column A and B to auto-fit, or drag it manually.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A:A");
          r.format.load("columnWidth");
          await context.sync();
          return r.format.columnWidth > 90;
        }) }
    ]
  }
};