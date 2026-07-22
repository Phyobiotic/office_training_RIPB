// lessons/xl-formulas.js
// Syllabus refs 3.4.1 (arithmetic formulas), 3.4.2 (SUM, AVERAGE, COUNT).
//
// tutorial = stepper, but a CONTINUOUS one — see xl-data.js for the pattern:
// `prepare()` clears the sheet once on entry, each step owns its own
// non-overlapping column band and only seeds it if empty.
// assignment = one scenario, full checklist

window.LESSON_MODULES = window.LESSON_MODULES || {};

window.LESSON_MODULES["xl-formulas"] = {
  title: "Formulas & Functions",

  /* ======================================================================
     TUTORIAL — one function/idea per step
     ====================================================================== */
  tutorial: {
    hook: "The whole point of a spreadsheet: it does the maths for you, and redoes it the moment a number changes. Each step drills one piece in its own column band, so earlier ones stay put as you go.",

    prepare: async () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange("A1:N10").clear(Excel.ClearApplyTo.all);
      await context.sync();
    }),

    steps: [
      {
        title: "A formula starts with =",
        teach: `Every calculation begins with an <b>equals sign</b>. Type <code>=</code>, then the maths. Crucially, refer to <i>cells</i> (like <code>B1</code>) not the numbers inside them — so if the number changes later, your answer updates itself.<br><br><b>B1</b> is 40, <b>B2</b> is 25. In <b>B3</b>, type <code>=B1+B2</code> and press Enter. (Try changing B1 afterwards — watch B3 recompute.)`,
        done: "=B1+B2 refers to the cells, so the answer re-does itself when they change.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("A1:B2");
          r.load("values");
          await context.sync();
          if (!r.values.some(row => row.some(v => v))) {
            sheet.getRange("A1:B2").values = [["Drinks", 40], ["Snacks", 25]];
            sheet.getRange("A3").values = [["Total"]];
            await context.sync();
          }
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B3");
          r.load("formulas,values");
          await context.sync();
          const f = String(r.formulas[0][0]);
          return f.startsWith("=") && /B1/i.test(f) && /B2/i.test(f) && Number(r.values[0][0]) === 65;
        })
      },

      {
        title: "SUM a range",
        teach: `Adding D1+D2+D3+D4+D5 by hand is painful. <code>SUM</code> totals a whole <b>range</b> at once — the colon <code>:</code> means "everything from here to there".<br><br>Column D has five days of takings. In <b>D6</b>, type <code>=SUM(D1:D5)</code> to total them.`,
        done: "=SUM(D1:D5) — the colon covers the whole range in one go.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("D1:D5");
          r.load("values");
          await context.sync();
          if (!r.values.some(row => row[0])) {
            r.values = [[48], [52], [39], [61], [74]];
            sheet.getRange("C6").values = [["Total"]];
            await context.sync();
          }
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("D6");
          r.load("formulas,values");
          await context.sync();
          return /SUM/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) === 274;
        })
      },

      {
        title: "AVERAGE a range",
        teach: `<code>AVERAGE</code> works exactly like SUM but gives you the mean instead of the total — same range syntax.<br><br>Column F has the same five days of takings. Put <code>=AVERAGE(F1:F5)</code> in <b>F6</b> to find the typical day's takings.`,
        done: "AVERAGE(range) — identical syntax to SUM, different answer.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("F1:F5");
          r.load("values");
          await context.sync();
          if (!r.values.some(row => row[0])) {
            r.values = [[48], [52], [39], [61], [74]];
            sheet.getRange("E6").values = [["Average"]];
            await context.sync();
          }
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("F6");
          r.load("formulas,values");
          await context.sync();
          return /AVERAGE/i.test(String(r.formulas[0][0])) && Math.abs(Number(r.values[0][0]) - 54.8) < 0.01;
        })
      },

      {
        title: "COUNT how many entries",
        teach: `<code>COUNT</code> tells you how many cells in a range actually contain <b>numbers</b> (it ignores blanks and text). Handy for "how many stalls reported?" without counting by eye.<br><br>Column H has some numbers and a couple of gaps. In <b>H7</b>, type <code>=COUNT(H1:H6)</code> and see how many real figures are there.`,
        done: "COUNT only tallies cells with numbers — blanks and text don't count.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("H1:H6");
          r.load("values");
          await context.sync();
          if (!r.values.some(row => row[0])) {
            r.values = [[48], [52], [""], [61], [""], [74]];
            sheet.getRange("G7").values = [["How many days reported"]];
            await context.sync();
          }
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("H7");
          r.load("formulas,values");
          await context.sync();
          return /COUNT/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) === 4;
        })
      },

      {
        title: "Autofill a formula down",
        teach: `Write a formula once, then autofill it down — Excel <b>shifts the cell references automatically</b> for each row. A formula <code>=J1+K1</code> becomes <code>=J2+K2</code>, <code>=J3+K3</code>, and so on, all by dragging.<br><br><b>L1</b> already holds <code>=J1+K1</code>. Drag its fill handle down to <b>L4</b> and watch each row total itself.`,
        done: "Dragging a formula down auto-adjusts the row numbers — that's why references beat typed numbers.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const r = sheet.getRange("J1:K4");
          r.load("values");
          await context.sync();
          if (!r.values.some(row => row.some(v => v))) {
            r.values = [[48, 31], [52, 27], [39, 44], [61, 38]];
            sheet.getRange("L1").formulas = [["=J1+K1"]];
            await context.sync();
          }
        }),
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("L1:L4");
          r.load("formulas,values");
          await context.sync();
          const allFormulas = r.formulas.every(row => String(row[0]).startsWith("="));
          const vals = r.values.map(row => Number(row[0]));
          return allFormulas && vals[0] === 79 && vals[1] === 79 && vals[2] === 83 && vals[3] === 99;
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Founders' Day Stall Reconciliation
     ====================================================================== */
  assignment: {
    hook: "Every stall's cash float and closing takings need reconciling tonight — eight stalls, and the Board takes a 10% cut of profit off the top before anything gets banked.",
    brief: "Work out each stall's profit, total it, average it, count how many stalls reported, and calculate the Board's cut.",
    doneMsg: "Reconciliation's done — ready to bank. Every function from the tutorial, doing a real night's books.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H25").clear(Excel.ClearApplyTo.all);
        sheet.getRange("A1").values = [["Founders' Day — Stall Reconciliation"]];
        sheet.getRange("A1").format.font.bold = true;
        sheet.getRange("A1").format.font.size = 16;
        sheet.getRange("A3:D3").values = [["Stall", "Float", "Closing Cash", "Profit"]];
        sheet.getRange("A4:C11").values = [
          ["Face Painting", 30, 145],
          ["Popcorn", 50, 210],
          ["Ring Toss", 40, 168],
          ["Cupcake Sale", 60, 302],
          ["Balloon Darts", 35, 155],
          ["Henna Booth", 45, 190],
          ["Lucky Draw", 25, 340],
          ["Iced Gems", 55, 221]
        ];
        sheet.getRange("A13").values = [["Total profit"]];
        sheet.getRange("A14").values = [["Average profit per stall"]];
        sheet.getRange("A15").values = [["Stalls reporting"]];
        sheet.getRange("A16").values = [["Board's cut (10%)"]];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.4.1", text: "In <b>D4</b>, subtract the float from closing cash to get Face Painting's profit.",
        hint: "=C4-B4",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("D4");
          r.load("formulas,values");
          await context.sync();
          const f = String(r.formulas[0][0]);
          return f.startsWith("=") && /C4/i.test(f) && /B4/i.test(f) && Number(r.values[0][0]) === 115;
        }) },

      { ref: "3.1.6", text: "Autofill that formula from <b>D4</b> down to <b>D11</b> for the rest of the stalls.",
        hint: "Select D4, drag the fill handle down to D11.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("D5:D11");
          r.load("formulas");
          await context.sync();
          return r.formulas.every(row => String(row[0]).startsWith("="));
        }) },

      { ref: "3.4.2", text: "In <b>B13</b>, use <code>SUM</code> to total every stall's profit.",
        hint: "=SUM(D4:D11)",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B13");
          r.load("formulas,values");
          await context.sync();
          return /SUM/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) > 0;
        }) },

      { ref: "3.4.2", text: "In <b>B14</b>, use <code>AVERAGE</code> for the average profit per stall.",
        hint: "=AVERAGE(D4:D11)",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B14");
          r.load("formulas,values");
          await context.sync();
          return /AVERAGE/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) > 0;
        }) },

      { ref: "3.4.2", text: "In <b>B15</b>, use <code>COUNT</code> to confirm how many stalls actually reported profit figures.",
        hint: "=COUNT(D4:D11) — should come out to 8.",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B15");
          r.load("formulas,values");
          await context.sync();
          return /COUNT/i.test(String(r.formulas[0][0])) && Number(r.values[0][0]) === 8;
        }) },

      { ref: "3.4.1", text: "In <b>B16</b>, calculate the Board's 10% cut of the total profit.",
        hint: "=B13*0.1",
        check: async () => Excel.run(async (context) => {
          const r = context.workbook.worksheets.getActiveWorksheet().getRange("B16");
          r.load("formulas,values");
          await context.sync();
          const f = String(r.formulas[0][0]);
          const b13 = context.workbook.worksheets.getActiveWorksheet().getRange("B13");
          b13.load("values");
          await context.sync();
          const expected = Number(b13.values[0][0]) * 0.1;
          return f.startsWith("=") && /B13/i.test(f) && Math.abs(Number(r.values[0][0]) - expected) < 0.5;
        }) }
    ]
  }
};