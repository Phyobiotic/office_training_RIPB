// lessons/word-output.js
// Syllabus refs 2.4.1-2.4.4.
//
// HONEST CAVEAT, two different flavours in this file:
//
//   - Orientation (2.4.3) IS checkable, via Section.pageSetup.orientation —
//     but that property is WordApiDesktop 1.3, meaning DESKTOP WORD ONLY
//     (not Word on the web). Wrapped in try/catch so it fails gracefully
//     rather than throwing on a host where it's unsupported.
//
//   - Preview (2.4.2) and printing with output options (2.4.4) are host-UI
//     actions with no observable document state at all — there is nothing
//     in the Word JS API that reveals whether Print Preview was opened or
//     a document was sent to a printer. Rather than silently dropping these
//     from a lesson titled "Finishing and Printing", they're kept as real
//     steps whose `check` self-certifies (always true) and whose `teach`
//     text says so plainly, matching the honesty already established for
//     Excel's chart lesson (3.6.1/3.6.2, see xl-charts.js's closing note).
//     They're deliberately left OUT of the assignment checklist below,
//     though, so a real scenario's progress bar only reflects what's
//     actually been verified.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function wordOutputCheckLandscape() {
  try {
    return await Word.run(async (context) => {
      const section = context.document.sections.getFirst();
      section.load("pageSetup/orientation");
      await context.sync();
      return section.pageSetup.orientation === Word.PageOrientation.landscape;
    });
  } catch (e) { return false; }
}

window.LESSON_MODULES["word-output"] = {
  title: "Finishing and Printing a Document",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The last mile before a document leaves your hands — catching mistakes and getting the page itself right. One document the whole way through.",

    prepare: async () => Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
    }),

    steps: [
      {
        title: "Spell check",
        teach: `Word underlines suspected mistakes in red as you type. <b>Right-click</b> an underlined word for suggested fixes, or run <b>Review tab → Spelling &amp; Grammar</b> to step through the whole document. Repeated words (like "the the") get flagged too — delete the extra one.<br><br>The line below has two misspellings and one repeated word. Fix all three.`,
        done: "Right-click a red underline for fixes, or Review → Spelling & Grammar to sweep the whole document.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("whole team is thankful")) {
            body.insertParagraph("We recieve so much generous support for this occured event, and the the whole team is thankful.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          const t = body.text;
          return !/recieve/i.test(t) && !/occured/i.test(t) && !/\bthe\s+the\b/i.test(t);
        })
      },

      {
        title: "Preview the document",
        teach: `<b>File → Print</b> opens a preview of exactly how the page will look before anything gets printed — worth a glance every time, since screen layout and printed layout don't always match.<br><br>Open File → Print now and look at the preview pane.<br><br><i>(This step can't be auto-verified — there's no way for this pane to detect that Print Preview was opened. It's marked done automatically; just make sure you actually looked before moving on.)</i>`,
        done: "File → Print shows the preview pane alongside the print options.",
        setup: async () => {},
        check: async () => true
      },

      {
        title: "Change page orientation",
        teach: `<b>Layout tab → Orientation</b> switches the whole page between Portrait (tall) and Landscape (wide).<br><br>Switch this document to <b>Landscape</b>.`,
        done: "Layout tab → Orientation → Landscape.",
        setup: async () => {},
        check: wordOutputCheckLandscape
      },

      {
        title: "Print with output options",
        teach: `<b>File → Print</b> also has the options themselves: print the whole document, just the current page, a page range, or only text you've selected first — plus how many copies. They're all in that same Print pane, just below the preview.<br><br>Open File → Print and look through those options.<br><br><i>(Also can't be auto-verified, for the same reason as Preview — no document state changes when you print. Marked done automatically.)</i>`,
        done: "File → Print → the options for what and how many are right there next to the preview.",
        setup: async () => {},
        check: async () => true
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Founders' Day Notice, Ready to Post
     ====================================================================== */
  assignment: {
    hook: "The Founders' Day notice needs to go up on the noticeboard today — but it's still riddled with typos and it's a wide event poster stuck in portrait orientation.",
    brief: "Clean up the spelling and the repeated word, then switch the page to landscape so it actually fits the poster board.",
    doneMsg: "Notice is clean and correctly oriented — ready to print and post. (Preview and print itself aren't graded here — the JS API genuinely can't see either one — but do check the preview yourself before it goes up.)",

    setup: async () => {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertParagraph("Founders' Day 2026", Word.InsertLocation.end);
        body.insertParagraph("Come celebrate with us! We recieve so much generous support for this occured event every single year, and the the whole school benefits.", Word.InsertLocation.end);
        body.insertParagraph("Saturday, 12 August — Main Field", Word.InsertLocation.end);
        await context.sync();
      });
    },

    tasks: [
      { ref: "2.4.1", text: "Fix both misspellings — <code>recieve</code> and <code>occured</code> should read <code>receive</code> and <code>occurred</code>.",
        hint: "Right-click each red underline, or use Find & Replace.",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !/recieve/i.test(body.text) && !/occured/i.test(body.text);
        }) },

      { ref: "2.4.1", text: "Delete the repeated word — <code>the the</code> should just be <code>the</code>.",
        hint: "Click right before the second \"the\" and delete it (and the extra space).",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !/\bthe\s+the\b/i.test(body.text);
        }) },

      { ref: "2.4.3", text: "This is a wide poster — switch the page to <b>Landscape</b> orientation.",
        hint: "Layout tab → Orientation → Landscape.",
        check: wordOutputCheckLandscape }
    ]
  }
};
