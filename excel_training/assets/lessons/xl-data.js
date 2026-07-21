// lessons/xl-data.js
// Syllabus refs 3.1.1-3.1.9.
//
// TWO DIFFERENT SHAPES in this file, on purpose:
//
//   tutorial: { hook, steps: [ {title, teach, setup, check, done}, ... ] }
//     A walkthrough. Each step teaches ONE mechanic in isolation and has
//     its own load state — pressing "Load this step" wipes and re-seeds a
//     tiny focused demo. Steps do NOT build on each other, so a student can
//     jump around and nothing breaks. `teach` is HTML shown in the panel
//     explaining how to do the thing; `check` confirms they did it.
//
//   assignment: { hook, brief, doneMsg, setup, tasks: [ {ref,text,hint,check} ] }
//     One realistic scenario, one load, the full checklist at once. This is
//     the "prove you can do it for real" half. Shape unchanged from before.

window.LESSON_MODULES = window.LESSON_MODULES || {};

window.LESSON_MODULES["xl-data"] = {
  title: "Entering and Managing Data",

  /* ======================================================================
     TUTORIAL — one mechanic per step, each self-contained
     ====================================================================== */
  tutorial: {
    hook: "Five quick things you'll do to almost every sheet. Each one loads its own little practice patch — do them in any order.",

    steps: [
      {
        title: "Autofill a series",
        teach: `Type the first number, then drag the fill handle — the small square at the bottom-right corner of a selected cell — down. Here's the part that trips people up: by default, dragging a <i>single</i> number just <b>copies</b> it, it doesn't count up.<br><br>Once you release the drag, a tiny square icon (<b>Auto Fill Options</b>) appears at the bottom-right of the range you just filled. Click it and choose <b>Fill Series</b> — that's what turns the copies into 1, 2, 3, 4, 5, 6.<br><br><b>A1</b> already has a 1. Drag its corner handle down to <b>A6</b>, then click the Auto Fill Options icon and pick Fill Series.`,
        done: "Drag copies by default — Fill Series (from that little icon) is what actually counts up.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1").values = [[1]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:A6");
          r.load("values");
          await context.sync();
          return r.values.map(row => row[0]).every((v, i) => Number(v) === i + 1);
        })
      },

      {
        title: "Copy a value down",
        teach: `Sometimes you don't want a series — you want the <i>same</i> value repeated. Autofill does that too, and for text this is actually the default (no extra click needed, unlike numbers): drag the fill handle and it just copies straight down.<br><br>Column A lists five prefects on patrol. <b>B1</b> already says <code>Gate</code>. Drag its fill handle down to <b>B5</b> so every one of them is assigned to Gate duty.`,
        done: "Text copies straight down on drag — no series involved, no extra click needed.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1:A5").values = [["Aidan Tan"], ["Priya Sharma"], ["Rachel Lim"], ["Zul Hakim"], ["Farah Aziz"]];
          sheet.getRange("B1").values = [["Gate"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B1:B5");
          r.load("values");
          await context.sync();
          return r.values.every(row => row[0] === "Gate");
        })
      },

      {
        title: "Select a block in one drag",
        teach: `To format or move several cells at once, select them together first. Click the top-left cell and <b>drag to the bottom-right</b> — the whole rectangle highlights.<br><br>Select the block <b>A1:C3</b> (all nine cells) in a single drag, then make them <b>bold</b> — <span class="keys">Ctrl/Cmd + B</span>.`,
        done: "One drag selects a whole block; formatting hits all of it at once.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1:C3").values = [
            ["Aidan", "Priya", "Rachel"],
            ["Zul", "Farah", "Meera"],
            ["Divya", "Kai", "Jason"]
          ];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:C3");
          r.format.font.load("bold");
          await context.sync();
          return r.format.font.bold === true;
        })
      },

      {
        title: "Fix a wrong entry",
        teach: `Editing a cell is just: click it, type the new value, press <span class="keys">Enter</span>. The old contents are replaced.<br><br><b>A1</b> says <code>Priya Sharma</code>, but this duty was actually Rachel's. Click <b>A1</b> and change it to <code>Rachel Lim</code>.`,
        done: "Click, retype, Enter — the cell's replaced.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1").values = [["Priya Sharma"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
          r.load("values");
          await context.sync();
          return r.values[0][0] === "Rachel Lim";
        })
      },

      {
        title: "Find & Replace",
        teach: `When the same mistake appears many times, don't fix them one by one. <b>Home tab → Find &amp; Select → Replace</b> (or <span class="keys">Ctrl/Cmd + H</span>) swaps every match at once.<br><br>Column A has <code>Gaet</code> three times — it should be <code>Gate</code>. Open Replace, find <code>Gaet</code>, replace all with <code>Gate</code>.`,
        done: "Find & Replace fixes every copy of a mistake in one go.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
          sheet.getRange("A1:A4").values = [["Gaet"], ["Library"], ["Gaet"], ["Gaet"]];
          await context.sync();
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A1:A4");
          r.load("values");
          await context.sync();
          const vals = r.values.map(row => row[0]);
          return vals.filter(v => v === "Gate").length === 3 && !vals.includes("Gaet");
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Term 3 CIP Hours Audit (one scenario, full checklist)
     ====================================================================== */
  assignment: {
    hook: "The Board Secretary needs Term 3's Community Involvement Programme hours consolidated before the EXCO meeting — fifteen entries, logged inconsistently, across two terms' worth of paperwork.",
    brief: "Clean up the log, fix a wrong entry, standardise the activity name that's spelled four different ways, flag every session from the beach cleanup day, and set up next term's sheet.",
    doneMsg: "Log's audit-ready. Everything you drilled in the tutorial, now on a real pile of messy data.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H25").clear(Excel.ClearApplyTo.all);
        sheet.getRange("A1").values = [["Prefects' Board — CIP Hours Audit, Term 3"]];
        sheet.getRange("A1").format.font.bold = true;
        sheet.getRange("A1").format.font.size = 16;
        sheet.getRange("A3:E3").values = [["S/N", "Prefect", "Activity", "Date", "Hours"]];
        sheet.getRange("A4").values = [[1]];
        sheet.getRange("B4:E18").values = [
          ["Aidan Tan",       "Beach Cleanup",         "12 Aug", 3],
          ["Priya Sharma",    "Beach clean-up",        "12 Aug", 3],
          ["Rachel Lim",      "Elderly Home Visit",    "15 Aug", 2],
          ["Zul Hakim",       "beach cleanup",          "12 Aug", 3],
          ["Farah Aziz",      "Reading Buddies",        "18 Aug", 1.5],
          ["Meera Nathan",    "Blood Donation Drive",   "20 Aug", 2],
          ["Divya Krishnan",  "Beach  Cleanup",         "12 Aug", 3],
          ["Kai Xuan Ong",    "Elderly Home Visit",     "15 Aug", 2],
          ["Aidan Tan",       "Reading Buddies",        "22 Aug", 1.5],
          ["Priya Sharma",    "Blood Donation Drive",   "20 Aug", 2],
          ["Rachel Lim",      "Beach Cleanup",          "12 Aug", 3],
          ["Zul Hakim",       "Reading Buddies",        "22 Aug", 1.5],
          ["Farah Aziz",      "Elderly Home Visit",     "15 Aug", 2],
          ["Meera Nathan",    "Beach clean-up",         "12 Aug", 3],
          ["Divya Krishnan",  "Blood Donation Drive",   "20 Aug", 2]
        ];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.1.6", text: "Autofill the <b>S/N</b> column, A4 down to A18, so entries are numbered 1 through 15.",
        hint: "Select A4, drag the fill handle down to A18.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A4:A18");
          r.load("values");
          await context.sync();
          return r.values.map(row => row[0]).every((v, i) => Number(v) === i + 1);
        }) },

      { ref: "3.1.4", text: "Kai Xuan Ong's Elderly Home Visit hours (row 11) were logged wrong — it should be <b>2.5</b>, not 2. Fix E11.",
        hint: "Click E11, type 2.5.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("E11");
          r.load("values");
          await context.sync();
          return Number(r.values[0][0]) === 2.5;
        }) },

      { ref: "3.1.8 / 3.1.9", text: "\"Beach Cleanup\" is spelled four different ways in column C (hyphens, lowercase, double spaces). Use Find &amp; Replace to make every instance read exactly <code>Beach Cleanup</code>.",
        hint: "You'll likely need more than one Find & Replace pass to catch every variant — check the whole column afterwards.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("C4:C18");
          r.load("values");
          await context.sync();
          const beachRowIdx = [0, 1, 3, 6, 10, 13];
          return beachRowIdx.every(i => r.values[i][0] === "Beach Cleanup");
        }) },

      { ref: "3.1.3", text: "Every Beach Cleanup entry happened on 12 Aug, but the rows aren't next to each other. Ctrl+click to select all of them at once and bold the Activity cell in each.",
        hint: "Rows 4, 5, 7, 10, 14, and 17 — click the first, then Ctrl+click each of the others before pressing Bold.",
        check: async () => {
          const rows = [4, 5, 7, 10, 14, 17];
          const results = await Promise.all(rows.map(r => Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getRange("C" + r);
            range.format.font.load("bold");
            await context.sync();
            return range.format.font.bold === true;
          })));
          return results.every(Boolean);
        } },

      { ref: "3.1.1 / 3.1.5 / 3.1.7", text: "Set up next term's log: add a new worksheet, name it so it's clearly <b>Term 4</b>, and copy the header row (S/N, Prefect, Activity, Date, Hours) into it.",
        hint: "Right-click a sheet tab → Insert, or use the + button. Rename by double-clicking the tab. Then copy A3:E3 into the new sheet.",
        check: async () => Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          const target = sheets.items.find(s => /term\s*4/i.test(s.name));
          if (!target) return false;
          const header = target.getRange("A3:E3");
          header.load("values");
          await context.sync();
          const expected = ["S/N", "Prefect", "Activity", "Date", "Hours"];
          return expected.every((v, i) => header.values[0][i] === v);
        }) }
    ]
  }
};