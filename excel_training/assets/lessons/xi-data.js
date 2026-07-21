// lessons/xl-data.js
// Syllabus refs 3.1.1–3.1.9. Loaded dynamically by taskpane.html — see the
// comment block at the top of that file for the module shape.
//
// Both scenarios are written as real Prefects' Board admin work, not
// generic "spreadsheet practice" data — the tutorial is the low-stakes
// version of the skill, the assignment is the same skill under the kind
// of messy, denser conditions a board secretary actually deals with.

window.LESSON_MODULES = window.LESSON_MODULES || {};

window.LESSON_MODULES["xl-data"] = {
  title: "Entering and Managing Data",

  /* ======================================================================
     TUTORIAL — Week 1 Duty Roster
     ====================================================================== */
  tutorial: {
    hook: "The Week 1 duty roster has a typo, a duplicate, and a gap — and the serial numbers aren't even filled in yet.",
    brief: "Clean up the roster: number it, fix the duplicate, fill the gap, bold the names, and correct the misspelled post.",
    doneMsg: "Roster's clean. This is the same handful of moves you'll use on every roster from here on.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H20").clear(Excel.ClearApplyTo.all);
        sheet.getRange("A1").values = [["Prefects' Board — Week 1 Duty Roster"]];
        sheet.getRange("A1").format.font.bold = true;
        sheet.getRange("A1").format.font.size = 16;
        sheet.getRange("A3:C3").values = [["Duty No.", "Prefect", "Post"]];
        sheet.getRange("A4").values = [[1]];
        sheet.getRange("B4:C8").values = [
          ["Aidan Tan", "Gate"],
          ["Priya Sharma", "Library"],
          ["Priya Sharma", "Canteen"],   // wrongly duplicated — should be Rachel Lim
          ["Zul Hakim", "Gaet"],          // typo — should be Gate
          ["", "Library"]                 // gap — should be Farah Aziz
        ];
        sheet.getRange("F4").values = [["Note: Farah Aziz covers Friday Library duty."]];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.1.6", text: "Fill in the duty numbers — autofill <b>A4</b> down to <b>A8</b> so they read 1 through 5.",
        hint: "Select A4, drag the fill handle down to A8.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A4:A8");
          r.load("values");
          await context.sync();
          return r.values.map(row => row[0]).every((v, i) => Number(v) === i + 1);
        }) },

      { ref: "3.1.4", text: "Row 6 has \"Priya Sharma\" duplicated by mistake — it should be <b>Rachel Lim</b>. Edit B6.",
        hint: "Click B6, type the correct name, press Enter.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B6");
          r.load("values");
          await context.sync();
          return r.values[0][0] === "Rachel Lim";
        }) },

      { ref: "3.1.5", text: "B8 is missing a name. Check the note in F4, then copy that prefect's name into <b>B8</b>.",
        hint: "The note says who covers Friday Library duty — copy their name in, don't just retype from memory.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B8");
          r.load("values");
          await context.sync();
          return r.values[0][0] === "Farah Aziz";
        }) },

      { ref: "3.1.3", text: "Select the whole <b>Prefect and Post</b> block, <b>B4:C8</b>, in one drag and bold it.",
        hint: "Click B4, drag to C8, then Ctrl/Cmd+B.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B4:C8");
          r.format.font.load("bold");
          await context.sync();
          return r.format.font.bold === true;
        }) },

      { ref: "3.1.8 / 3.1.9", text: "One post is spelled <code>Gaet</code> instead of <code>Gate</code>. Use Find &amp; Replace to fix it.",
        hint: "Home tab → Find & Select → Replace. Find \"Gaet\", replace with \"Gate\".",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("C7");
          r.load("values");
          await context.sync();
          return r.values[0][0] === "Gate";
        }) }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Term 3 CIP Hours Audit
     ====================================================================== */
  assignment: {
    hook: "The Board Secretary needs Term 3's Community Involvement Programme hours consolidated before the EXCO meeting — fifteen entries, logged inconsistently, across two terms' worth of paperwork.",
    brief: "Clean up the log, fix a wrong entry, standardise the activity name that's spelled four different ways, flag every session from the beach cleanup day, and set up next term's sheet.",
    doneMsg: "Log's audit-ready. This is the same editing/selecting/copying skillset as the tutorial — just at the volume a real board secretary actually deals with.",

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
          const beachRowIdx = [0, 1, 3, 6, 10, 13]; // rows 4,5,7,10,14,17 relative to C4
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
        hint: "Right-click a sheet tab → Insert, or use the + button. Rename it by double-clicking the tab. Then copy A3:E3 from this sheet into the new one.",
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