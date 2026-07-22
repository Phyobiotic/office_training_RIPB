// lessons/founders-day.js
// The final assignment — now spanning all three apps. Which one shows up
// is decided purely by which host the add-in is currently running in
// (Office.context.host), read fresh every time `assignment` is accessed
// via the getter below — the shell (taskpane.html) always reads
// `mod.assignment` for capstone lessons, so this is the one hook point
// that needs to change to make one lesson module serve three different
// scenarios.

window.LESSON_MODULES = window.LESSON_MODULES || {};

const EXCEL_ASSIGNMENT = {
  hook: "The Board Secretary needs the Founders' Day budget finished before tonight's EXCO meeting — three income streams, six days of takings, and a chart the Treasurer can read without asking questions.",
  brief: "Finish the sheet, total it properly with formulas (not typed numbers), flag the two peak days, and turn it into a chart.",
  doneMsg: "Budget's ready for EXCO.",

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
};

const WORD_ASSIGNMENT = {
  hook: "The printed Founders' Day programme needs to be at the shop by 6pm — the schedule table has a hole in it and a double-booked slot, the heading doesn't read like a heading, and it's still stuck in portrait when it needs to be a wide handout.",
  brief: "Fix the wording, sort out the schedule table, format the heading and reminders properly, drop in a cover photo, and get the page orientation right before it goes to print.",
  doneMsg: "Programme's ready for the printer. Every Word skill from this training, on one real document.",

  setup: async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      body.insertParagraph("Founders' Day 2026 — Event Programme", Word.InsertLocation.end);
      body.insertParagraph("We recieve so much support from our community every single occured year.", Word.InsertLocation.end);
      const table = body.insertTable(4, 3, Word.InsertLocation.end, [
        ["Time", "Activity", "Location"],
        ["9:00am", "Opening Remarks", "Main Field"],
        ["9:00am", "Opening Remarks", "Main Field"],
        ["10:30am", "Stall Games", "Courtyard"]
      ]);
      for (let c = 0; c < 3; c++) table.getCell(0, c).columnWidth = 55;
      body.insertParagraph("Bring your own water bottle", Word.InsertLocation.end);
      body.insertParagraph("Wear closed-toe shoes", Word.InsertLocation.end);
      body.insertParagraph("Arrive 15 minutes early", Word.InsertLocation.end);
      body.insertParagraph("Programme cover photo:", Word.InsertLocation.end);
      body.insertParagraph("See you there!", Word.InsertLocation.end);
      await context.sync();
    });
  },

  tasks: [
    { ref: "2.1.5", text: "Fix the two mistakes in the welcome line — <code>recieve</code> should read <code>receive</code>, and <code>occured</code> should read <code>occurred</code>.",
      hint: "Use Find & Replace (Ctrl/Cmd+H) for each.",
      check: async () => Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return !/recieve/i.test(body.text) && !/occured/i.test(body.text);
      }) },

    { ref: "2.2.1", text: "Format the heading <b>Founders' Day 2026 — Event Programme</b>: at least 20pt and a non-default font.",
      hint: "Select the heading, change the font size and font name boxes on the Home tab.",
      check: async () => Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load("items/text");
        await context.sync();
        const p = paras.items.find(x => x.text.trim() === "Founders' Day 2026 — Event Programme");
        if (!p) return false;
        p.font.load("size,name");
        await context.sync();
        const defaults = ["calibri", "aptos", "aptos narrow", "aptos display", "aptos serif", "aptos mono"];
        return p.font.size >= 20 && !defaults.includes((p.font.name || "").toLowerCase().trim());
      }) },

    { ref: "2.3.4", text: "The schedule is missing the closing slot — insert a row for <code>5:00pm</code>, <code>Closing Ceremony</code>, <code>Main Field</code>.",
      hint: "Right-click any row → Insert, then type the details.",
      check: async () => Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();
        tables.items.forEach(t => t.load("values"));
        await context.sync();
        return tables.items.some(t => t.values.some(row => (row[1] || "").trim().toLowerCase() === "closing ceremony"));
      }) },

    { ref: "2.3.4", text: "<b>Opening Remarks</b> at 9:00am got entered twice by mistake — delete one of the duplicate rows.",
      hint: "Right-click one duplicate row's number area → Delete Cells → Delete entire row.",
      check: async () => Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();
        tables.items.forEach(t => t.load("values"));
        await context.sync();
        return tables.items.some(t => t.values.filter(row => (row[1] || "").trim().toLowerCase() === "opening remarks").length === 1);
      }) },

    { ref: "2.2.7", text: "Turn the three reminder lines (water bottle, shoes, arrive early) into a numbered list.",
      hint: "Select all three lines, then click Numbering on the Home tab.",
      check: async () => {
        try {
          return await Word.run(async (context) => {
            const paras = context.document.body.paragraphs;
            paras.load("items/text");
            await context.sync();
            const wanted = ["Bring your own water bottle", "Wear closed-toe shoes", "Arrive 15 minutes early"];
            const targets = paras.items.filter(p => wanted.includes(p.text.trim()));
            if (targets.length < 3) return false;
            targets.forEach(p => {
              p.listItemOrNullObject.load("isNullObject,level");
              p.listOrNullObject.load("isNullObject,levelTypes");
            });
            await context.sync();
            return targets.every(p => {
              if (p.listItemOrNullObject.isNullObject || p.listOrNullObject.isNullObject) return false;
              const level = p.listItemOrNullObject.level;
              return String(p.listOrNullObject.levelTypes[level]) === "Number";
            });
          });
        } catch (e) { return false; }
      } },

    { ref: "2.2.5", text: "Centre the closing line <b>See you there!</b>",
      hint: "Click into the line, Ctrl/Cmd+E.",
      check: async () => Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load("items/text");
        await context.sync();
        const p = paras.items.find(x => x.text.trim() === "See you there!");
        if (!p) return false;
        p.load("alignment");
        await context.sync();
        return p.alignment === Word.Alignment.centered;
      }) },

    { ref: "2.3.6", text: "Insert a picture right after \"<b>Programme cover photo:</b>\"",
      hint: "Click after that line, then Insert → Pictures.",
      check: async () => Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load("items/text");
        await context.sync();
        const marker = paras.items.find(p => p.text.trim().startsWith("Programme cover photo:"));
        if (!marker) return false;
        marker.inlinePictures.load("items");
        await context.sync();
        return marker.inlinePictures.items.length >= 1;
      }) },

    { ref: "2.3.8", text: "Resize that cover photo so it's a deliberate size, not whatever it defaulted to.",
      hint: "Click it once, then drag a corner handle.",
      check: async () => Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load("items/text");
        await context.sync();
        const marker = paras.items.find(p => p.text.trim().startsWith("Programme cover photo:"));
        if (!marker) return false;
        marker.inlinePictures.load("items/width,items/height");
        await context.sync();
        return marker.inlinePictures.items.length >= 1 && marker.inlinePictures.items.every(pic => pic.width > 0 && pic.height > 0);
      }) },

    { ref: "2.4.3", text: "This is a wide handout — switch the page to <b>Landscape</b> orientation.",
      hint: "Layout tab → Orientation → Landscape.",
      check: async () => {
        try {
          return await Word.run(async (context) => {
            const section = context.document.sections.getFirst();
            section.load("pageSetup/orientation");
            await context.sync();
            return section.pageSetup.orientation === Word.PageOrientation.landscape;
          });
        } catch (e) { return false; }
      } }
  ]
};

async function foundersDayPptSlideIndexesWithText(context, needle){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  slides.items.forEach(s => s.shapes.load("items"));
  await context.sync();
  slides.items.forEach(s => s.shapes.items.forEach(sh => { sh.textFrame.load("hasText"); sh.load("type,width,height"); }));
  await context.sync();
  slides.items.forEach(s => s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) sh.textFrame.textRange.load("text"); }));
  await context.sync();
  const norm = needle.toLowerCase();
  const hits = [];
  slides.items.forEach((s, i) => {
    const has = s.shapes.items.some(sh => sh.textFrame.hasText && (sh.textFrame.textRange.text || "").toLowerCase().includes(norm));
    if (has) hits.push(i);
  });
  return hits;
}

const PPT_ASSIGNMENT = {
  hook: "The sponsor announcement deck needs to go out before the EXCO meeting — the Agenda slide got duplicated, Closing Remarks ended up before Sponsor Benefits, the opening slide has no title, there's a spelling mistake, and the call-to-action doesn't stand out at all.",
  brief: "Fix the structure (duplicate, order, title), fix the spelling, add and resize a visual, and make the call-to-action pop.",
  doneMsg: "Announcement deck's ready for EXCO. Every PowerPoint skill from this training, on one real deck.",

  setup: async () => {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      for (let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
      await context.sync();
      slides.items[0].shapes.addTextBox("");
      context.presentation.slides.add();
      context.presentation.slides.add();
      context.presentation.slides.add();
      context.presentation.slides.add();
      await context.sync();
      slides.load("items");
      await context.sync();
      slides.items[1].shapes.addTextBox("Agenda");
      slides.items[2].shapes.addTextBox("Agenda");
      slides.items[3].shapes.addTextBox("Closing Remarks");
      const benefits = slides.items[4];
      const benefitsText = benefits.shapes.addTextBox("We recieve incredible support from sponsors like you every year.");
      benefitsText.left = 40; benefitsText.top = 40; benefitsText.width = 600; benefitsText.height = 80;
      const cta = benefits.shapes.addTextBox("Sponsor us today");
      cta.left = 40; cta.top = 150; cta.width = 400; cta.height = 60;
      await context.sync();
    });
  },

  tasks: [
    { ref: "4.2.1", text: "Type <b>Founders' Day Sponsor Announcement</b> into the empty box on slide 1.",
      hint: "Click the empty box on slide 1 and type.",
      check: async () => PowerPoint.run(async (context) => {
        const hits = await foundersDayPptSlideIndexesWithText(context, "founders' day sponsor announcement");
        return hits.includes(0);
      }) },

    { ref: "4.4.1", text: "The Sponsor Benefits slide has a spelling mistake — <code>recieve</code> should read <code>receive</code>.",
      hint: "Right-click the red underline, or use Find & Replace.",
      check: async () => PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        slides.items.forEach(s => s.shapes.load("items"));
        await context.sync();
        slides.items.forEach(s => s.shapes.items.forEach(sh => sh.textFrame.load("hasText")));
        await context.sync();
        slides.items.forEach(s => s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) sh.textFrame.textRange.load("text"); }));
        await context.sync();
        const allText = slides.items.flatMap(s => s.shapes.items.filter(sh => sh.textFrame.hasText).map(sh => sh.textFrame.textRange.text || "")).join(" ");
        return !/recieve/i.test(allText) && /receive/i.test(allText);
      }) },

    { ref: "4.1.7", text: "The <b>Agenda</b> slide got duplicated — delete the extra copy.",
      hint: "Right-click one of the two Agenda thumbnails → Delete Slide.",
      check: async () => PowerPoint.run(async (context) => {
        const hits = await foundersDayPptSlideIndexesWithText(context, "agenda");
        return hits.length === 1;
      }) },

    { ref: "4.1.6", text: "<b>Closing Remarks</b> ended up before <b>Sponsor Benefits</b> — reorder so Closing Remarks comes after it.",
      hint: "Drag the Closing Remarks thumbnail below the Sponsor Benefits one.",
      check: async () => PowerPoint.run(async (context) => {
        const closing = await foundersDayPptSlideIndexesWithText(context, "closing remarks");
        const benefits = await foundersDayPptSlideIndexesWithText(context, "sponsor us today");
        if (closing.length === 0 || benefits.length === 0) return false;
        return closing[0] > benefits[0];
      }) },

    { ref: "4.3.1", text: "Add a picture or drawn shape to the Sponsor Benefits slide.",
      hint: "Insert tab → Pictures, or Insert tab → Shapes.",
      check: async () => PowerPoint.run(async (context) => {
        const hits = await foundersDayPptSlideIndexesWithText(context, "sponsor us today");
        if (hits.length === 0) return false;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const s = slides.items[hits[0]];
        s.shapes.load("items");
        await context.sync();
        s.shapes.items.forEach(sh => sh.load("type"));
        await context.sync();
        return s.shapes.items.some(sh => /image|geometric/i.test(String(sh.type)));
      }) },

    { ref: "4.3.3", text: "Resize that picture or shape so it's clearly bigger than its default size.",
      hint: "Click it, drag a corner handle outward.",
      check: async () => PowerPoint.run(async (context) => {
        const hits = await foundersDayPptSlideIndexesWithText(context, "sponsor us today");
        if (hits.length === 0) return false;
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const s = slides.items[hits[0]];
        s.shapes.load("items");
        await context.sync();
        s.shapes.items.forEach(sh => sh.load("type,width,height"));
        await context.sync();
        return s.shapes.items.some(sh => /image|geometric/i.test(String(sh.type)) && (sh.width > 100 || sh.height > 100));
      }) },

    { ref: "4.2.6 / 4.2.7", text: "Make \"<b>Sponsor us today</b>\" bold and give it a font colour besides the default so it actually reads as a call-to-action.",
      hint: "Select the text: Ctrl/Cmd+B, then pick a colour from the font colour dropdown.",
      check: async () => {
        try {
          return await PowerPoint.run(async (context) => {
            const hits = await foundersDayPptSlideIndexesWithText(context, "sponsor us today");
            if (hits.length === 0) return false;
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
            const s = slides.items[hits[0]];
            s.shapes.load("items");
            await context.sync();
            s.shapes.items.forEach(sh => sh.textFrame.load("hasText"));
            await context.sync();
            const withText = s.shapes.items.filter(sh => sh.textFrame.hasText);
            withText.forEach(sh => sh.textFrame.textRange.load("text"));
            await context.sync();
            const target = withText.find(sh => /sponsor us today/i.test(sh.textFrame.textRange.text || ""));
            if (!target) return false;
            target.textFrame.textRange.font.load("bold,color");
            await context.sync();
            const font = target.textFrame.textRange.font;
            const c = (font.color || "").toLowerCase();
            return font.bold === true && !!c && c !== "#000000" && c !== "black" && c !== "automatic";
          });
        } catch (e) { return false; }
      } }
  ]
};

window.LESSON_MODULES["founders-day"] = {
  title: "Founders' Day — Final Assignment",
  isCapstone: true,

  get assignment(){
    if (typeof Office !== "undefined" && Office.context && Office.context.host === Office.HostType.Word) return WORD_ASSIGNMENT;
    if (typeof Office !== "undefined" && Office.context && Office.context.host === Office.HostType.PowerPoint) return PPT_ASSIGNMENT;
    return EXCEL_ASSIGNMENT;
  }
};
