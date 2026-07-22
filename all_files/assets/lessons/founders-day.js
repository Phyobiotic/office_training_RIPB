// lessons/founders-day.js
// The final assignment — Excel portion only for now. Word and PowerPoint
// portions get added the same way once those lessons exist; this file
// will grow an isCapstone-aware multi-app view later, but there's no
// reason to build that scaffolding before there's a second app to plug in.

window.LESSON_MODULES = window.LESSON_MODULES || {};

window.LESSON_MODULES["founders-day"] = {
  title: "Founders' Day — Final Assignment",
  isCapstone: true,

  assignment: {
    hook: "The Board Secretary needs the Founders' Day budget finished before tonight's EXCO meeting — three income streams, six days of takings, and a chart the Treasurer can read without asking questions.",
    brief: "Finish the sheet, total it properly with formulas (not typed numbers), flag the two peak days, and turn it into a chart.",
    doneMsg: "Budget's ready for EXCO. Word and PowerPoint portions of this assignment are still coming.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H25").clear(Excel.ClearApplyTo.all);
        sheet.getRange("A1").values = [["Prefects' Board — Founders' Day Budget"]];
        sheet.getRange("A1").format.font.bold = true;
        sheet.getRange("A1").format.font.size = 18;
        sheet.getRange("A3:E3").values = [["Day", "Drinks", "Snacks", "Raffle", "Total"]];
        sheet.getRange("A4:D8").values = [
          ["Mon", 48, 31, 22],
          ["Tue", 52, 27, 18],
          ["Wed", 39, 44, 25],
          ["Thu", 61, 38, 19],
          ["Fri", 74, 55, 31]
        ];
        sheet.getRange("A11").values = [["Total drinks"]];
        sheet.getRange("A12").values = [["Average drinks"]];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.1.2", text: "The Board forgot Saturday. In row 9, enter <code>Sat</code>, <code>66</code>, <code>49</code>, <code>28</code> for Day/Drinks/Snacks/Raffle.",
        hint: "A9 = Sat, B9 = 66, C9 = 49, D9 = 28.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A9:D9");
          r.load("values");
          await context.sync();
          const v = r.values[0];
          return v[0] === "Sat" && Number(v[1]) === 66 && Number(v[2]) === 49 && Number(v[3]) === 28;
        }) },

      { ref: "3.4.1", text: "In <b>E4</b>, build a formula totalling Monday's three income streams using cell references.",
        hint: "=B4+C4+D4",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("E4");
          r.load("formulas");
          await context.sync();
          const f = String(r.formulas[0][0]);
          return f.startsWith("=") && /B4/i.test(f) && /C4/i.test(f) && /D4/i.test(f);
        }) },

      { ref: "3.1.6", text: "Autofill that formula from <b>E4</b> down to <b>E9</b>.",
        hint: "Select E4, drag the fill handle down to E9.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("E5:E9");
          r.load("formulas");
          await context.sync();
          return r.formulas.every(row => String(row[0]).startsWith("="));
        }) },

      { ref: "3.4.2", text: "In <b>B11</b>, use <code>SUM</code> to total the drinks column.",
        hint: "=SUM(B4:B9)",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B11");
          r.load("formulas,values");
          await context.sync();
          return /SUM/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) > 0;
        }) },

      { ref: "3.4.2", text: "In <b>B12</b>, use <code>AVERAGE</code> for the typical day's drinks.",
        hint: "=AVERAGE(B4:B9)",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B12");
          r.load("formulas,values");
          await context.sync();
          return /AVERAGE/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) > 0;
        }) },

      { ref: "3.2.2", text: "Bold the header row <b>A3:E3</b>.",
        hint: "Select A3:E3, Ctrl/Cmd+B.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("A3:E3");
          r.format.font.load("bold");
          await context.sync();
          return r.format.font.bold === true;
        }) },

      { ref: "3.2.1 (fill)", text: "Give the header row a fill colour.",
        hint: "Select A3:E3, use the fill colour button.",
        check: async () => {
          try {
            return await Excel.run(async (context) => {
              const r = context.workbook.worksheets.getActiveWorksheet().getRange("A3:E3");
              r.format.fill.load("color");
              await context.sync();
              return !!r.format.fill.color;
            });
          } catch (e) { return false; }
        } },

      { ref: "3.1.3", text: "Monday and Friday are the two peak sponsor-visit days — bold both full rows (<b>A4:E4</b> and <b>A8:E8</b>). They're not adjacent, so Ctrl+click to select both before bolding.",
        hint: "Select A4:E4, then hold Ctrl and drag A8:E8, then press Bold.",
        check: async () => {
          const results = await Promise.all(["A4:E4", "A8:E8"].map(addr => Excel.run(async (context) => {
            const r = context.workbook.worksheets.getActiveWorksheet().getRange(addr);
            r.format.font.load("bold");
            await context.sync();
            return r.format.font.bold === true;
          })));
          return results.every(Boolean);
        } },

      { ref: "3.5.1", text: "Insert a chart from the daily totals so the Treasurer can read the week at a glance.",
        hint: "Select A3:A9 and E3:E9 (Ctrl+click for the second range), then Insert → Chart.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("count");
          await context.sync();
          return charts.count > 0;
        }) }
    ]
  }
};