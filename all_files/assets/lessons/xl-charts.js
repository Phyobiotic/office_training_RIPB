// lessons/xl-charts.js
// Syllabus refs 3.5.1-3.5.5. 3.6.1/3.6.2 (preview/print) are not auto-checked
// — the Excel JS API can't observe print-preview or printing (host-UI actions,
// not document state). Left as instructor-verified. See note at file end.
//
// tutorial = stepper, but a CONTINUOUS one: `prepare()` resets the working
// area once on first entry into the tutorial; every step's own `setup()`
// after that is purely additive (only fills in a precondition if it's
// missing) and never deletes the chart the student is building. So the
// same chart carries forward from "insert a chart" through title, legend,
// and resize — no re-loading a fresh chart per step.
// assignment = one scenario, full checklist

window.LESSON_MODULES = window.LESSON_MODULES || {};

const XL_CHARTS_DATA = [
  ["Beach Cleanup", 18], ["Elderly Home Visit", 9],
  ["Blood Donation Drive", 6], ["Reading Buddies", 12], ["Litter Patrol", 7]
];

async function xlChartsEnsureData(context, sheet){
  const dataRange = sheet.getRange("A1:B5");
  dataRange.load("values");
  await context.sync();
  const hasData = dataRange.values.some(row => row[0]);
  if(!hasData){
    dataRange.values = XL_CHARTS_DATA;
    await context.sync();
  }
}

// Only auto-creates the chart if none exists yet (e.g. a student jumping
// straight to a later step) — if the student already built one in the
// "Insert a chart" step, this leaves it completely alone.
async function xlChartsEnsureChart(context, sheet){
  await xlChartsEnsureData(context, sheet);
  const charts = sheet.charts;
  charts.load("count");
  await context.sync();
  if(charts.count === 0){
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, sheet.getRange("A1:B5"), Excel.ChartSeriesBy.columns);
    chart.legend.visible = false;
    chart.height = 180;
    chart.width = 300;
    await context.sync();
  }
}

window.LESSON_MODULES["xl-charts"] = {
  title: "Charts and Outputs",

  /* ======================================================================
     TUTORIAL — one chart, built up one piece per step
     ====================================================================== */
  tutorial: {
    hook: "A chart turns a column of numbers into something readable at a glance. Each step builds up one more piece of the same chart.",

    prepare: async () => Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange("A1:D20").clear(Excel.ClearApplyTo.all);
      sheet.charts.load("items");
      await context.sync();
      sheet.charts.items.forEach(c => c.delete());
      await context.sync();
    }),

    steps: [
      {
        title: "Insert a chart",
        teach: `Select the data <i>including</i> its labels, then <b>Insert tab → Charts</b>. "Recommended Charts" is the easy default, or pick a type directly.<br><br>Select <b>A1:B5</b> (activities and their hours), then insert any chart.`,
        done: "Select data + labels, Insert → chart. That's the core move.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          await xlChartsEnsureData(context, sheet);
        }),
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("count");
          await context.sync();
          return charts.count > 0;
        })
      },

      {
        title: "Give it a real title",
        teach: `The default title just says "Chart Title" — useless to anyone reading it cold. <b>Click that placeholder text</b> on the chart and retype something a stranger would understand.<br><br>The chart you just inserted is still here. Click its title and rename it to something meaningful (anything but the default).`,
        done: "A title should explain the chart to someone who's never seen your data.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          await xlChartsEnsureChart(context, sheet);
        }),
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length === 0) return false;
          const chart = charts.items[0];
          chart.title.load("text");
          await context.sync();
          const t = (chart.title.text || "").trim();
          return t.length > 3 && !/^chart\s*title$/i.test(t);
        })
      },

      {
        title: "Turn on the legend",
        teach: `A <b>legend</b> is the little key showing which colour means what. On a single-series chart it's often hidden by default. Click the chart, hit the <b>"+" (Chart Elements)</b> button that appears beside it, and tick <b>Legend</b>.<br><br>Same chart as before — switch its legend on.`,
        done: "The '+' button beside a selected chart toggles elements like the legend on and off.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          await xlChartsEnsureChart(context, sheet);
        }),
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length === 0) return false;
          const chart = charts.items[0];
          chart.legend.load("visible");
          await context.sync();
          return chart.legend.visible === true;
        })
      },

      {
        title: "Resize it",
        teach: `The default chart is small. Click it once to select it (handles appear at the corners), then <b>drag a corner handle</b> outward to enlarge it — corners keep the proportions, edges stretch.<br><br>Make the same chart noticeably bigger.`,
        done: "Drag a corner handle to resize — corners keep proportions, edges distort.",
        setup: async () => Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          await xlChartsEnsureChart(context, sheet);
        }),
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length === 0) return false;
          const chart = charts.items[0];
          chart.load("height,width");
          await context.sync();
          return chart.height > 240 || chart.width > 400;
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — EXCO Report Pack
     ====================================================================== */
  assignment: {
    hook: "Two datasets need to go into tonight's EXCO report: how CIP hours broke down by activity, and how Founders' Day income trended across the week. Different story, different chart type for each.",
    brief: "Build a pie chart for the CIP breakdown (proportions are the point) and a column chart for the income trend (day-by-day comparison is the point) — both need real titles and a visible legend.",
    doneMsg: "Both charts are report-ready. Picking the right chart type for the story is the real skill here.",

    setup: async () => {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1:H30").clear(Excel.ClearApplyTo.all);
        sheet.charts.load("items");
        await context.sync();
        sheet.charts.items.forEach(c => c.delete());
        await context.sync();
        sheet.getRange("A1").values = [["EXCO Report — Term 3"]];
        sheet.getRange("A1").format.font.bold = true;
        sheet.getRange("A1").format.font.size = 16;
        sheet.getRange("A3:B3").values = [["Activity", "Total Hours"]];
        sheet.getRange("A4:B7").values = [
          ["Beach Cleanup", 18], ["Elderly Home Visit", 9],
          ["Blood Donation Drive", 6], ["Reading Buddies", 12]
        ];
        sheet.getRange("D3:E3").values = [["Day", "Income"]];
        sheet.getRange("D4:E8").values = [
          ["Mon", 101], ["Tue", 97], ["Wed", 108], ["Thu", 118], ["Fri", 160]
        ];
        await context.sync();
      });
    },

    tasks: [
      { ref: "3.5.1", text: "Build a <b>pie chart</b> from the CIP Hours data (A3:B7) — proportions are what matter here.",
        hint: "Select A3:B7, Insert → Pie chart specifically.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items/chartType");
          await context.sync();
          return charts.items.some(c => c.chartType === "Pie");
        }) },

      { ref: "3.5.1", text: "Build a <b>column chart</b> from the Income data (D3:E8) — comparing days side by side.",
        hint: "Select D3:E8, Insert → Column chart.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items/chartType");
          await context.sync();
          return charts.items.some(c => String(c.chartType).includes("Column"));
        }) },

      { ref: "3.5.4", text: "Give <b>both</b> charts real titles — not the default placeholder.",
        hint: "Click each chart's title text and retype it.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length < 2) return false;
          charts.items.forEach(c => c.title.load("text"));
          await context.sync();
          return charts.items.every(c => {
            const t = (c.title.text || "").trim();
            return t.length > 3 && !/^chart\s*title$/i.test(t);
          });
        }) },

      { ref: "3.5.5", text: "Make sure the legend is visible on <b>both</b> charts.",
        hint: "Click each chart → \"+\" → tick Legend.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length < 2) return false;
          charts.items.forEach(c => c.legend.load("visible"));
          await context.sync();
          return charts.items.every(c => c.legend.visible === true);
        }) },

      { ref: "3.5.3", text: "Resize at least one chart so it's clearly bigger than the default insert size.",
        hint: "Click a chart, drag a corner handle outward.",
        check: async () => Excel.run(async (context) => {
          const charts = context.workbook.worksheets.getActiveWorksheet().charts;
          charts.load("items");
          await context.sync();
          if (charts.items.length === 0) return false;
          charts.items.forEach(c => c.load("height,width"));
          await context.sync();
          return charts.items.some(c => c.height > 260 || c.width > 420);
        }) }
    ]
  }
};

// A NOTE ON 3.6.1 / 3.6.2 (preview worksheet / print):
// Real syllabus items, deliberately NOT auto-checked. The Excel JS API
// exposes no way to observe "student opened print preview" or "student
// printed" — those are host-application UI actions, not document state
// Excel.run can read back. A faked check would verify nothing, so these
// are left for the instructor to confirm, or as a manual checklist item
// outside this pane.