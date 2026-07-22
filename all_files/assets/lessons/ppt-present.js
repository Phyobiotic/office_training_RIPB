// lessons/ppt-present.js
// Syllabus refs 4.4.1-4.4.4.
//
// HONEST CAVEAT: 4.4.2 (print options), 4.4.3 (start/end a slideshow), and
// 4.4.4 (navigate during a slideshow) all run into the same wall — running
// a live slideshow takes over the whole screen in desktop PowerPoint, and
// the task pane isn't usable while that's happening, so there's no moment
// where this pane could even ask the JS API a question, let alone get a
// useful answer back. Combined with the fact that printing has no
// observable document state either (same as word-output.js), these are
// grouped into two self-certifying steps that are honest about why, rather
// than faking a check that would never really run.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function pptPresentResetToSingleBlankSlide(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  for(let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
  await context.sync();
}

async function pptPresentAllShapeTexts(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  slides.items.forEach(s => s.shapes.load("items"));
  await context.sync();
  slides.items.forEach(s => s.shapes.items.forEach(sh => sh.textFrame.load("hasText")));
  await context.sync();
  slides.items.forEach(s => s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) sh.textFrame.textRange.load("text"); }));
  await context.sync();
  const out = [];
  slides.items.forEach(s => s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) out.push(sh.textFrame.textRange.text || ""); }));
  return out;
}

window.LESSON_MODULES["ppt-present"] = {
  title: "Presenting and Printing",

  /* ======================================================================
     TUTORIAL
     ====================================================================== */
  tutorial: {
    hook: "The last checks before a deck goes in front of an audience. It's the same deck the whole way through.",

    prepare: async () => PowerPoint.run(async (context) => {
      await pptPresentResetToSingleBlankSlide(context);
    }),

    steps: [
      {
        title: "Spell check your slides",
        teach: `PowerPoint underlines suspected mistakes in red as you type, same as Word. <b>Right-click</b> an underlined word for suggested fixes, or run <b>Review tab → Spelling</b> to step through the whole deck. Repeated words get flagged too.<br><br>The slide below has two misspellings and one repeated word. Fix all three.`,
        done: "Right-click a red underline for fixes, or Review → Spelling to sweep the whole deck.",
        setup: async () => PowerPoint.run(async (context) => {
          const texts = await pptPresentAllShapeTexts(context);
          if (texts.some(t => t.includes("recieve so much"))) return;
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          slides.items[0].shapes.addTextBox("We recieve so much generous suport, and the the whole school benefits.");
          await context.sync();
        }),
        check: async () => PowerPoint.run(async (context) => {
          const texts = await pptPresentAllShapeTexts(context);
          const joined = texts.join(" ");
          return !/recieve/i.test(joined) && !/suport/i.test(joined) && !/\bthe\s+the\b/i.test(joined);
        })
      },

      {
        title: "Print your presentation",
        teach: `<b>File → Print</b> has the same kind of output options as Word: the whole deck, specific slides, handouts (several slides per page), notes pages, or an outline view of just the text — plus how many copies.<br><br>Open File → Print and look through those options.<br><br><i>(This step can't be auto-verified — printing changes nothing in the document itself for the JS API to read. Marked done automatically.)</i>`,
        done: "File → Print → look for Full Page Slides, Notes Pages, Outline, and Handouts in the settings dropdown.",
        setup: async () => {},
        check: async () => true
      },

      {
        title: "Run the slideshow",
        teach: `<span class="keys">F5</span> starts the slideshow from the very first slide; <span class="keys">Shift+F5</span> starts it from whichever slide you're currently on. Once it's running, <b>arrow keys</b> (or a click) move to the next/previous slide, typing a slide number and pressing Enter jumps straight to it, and <span class="keys">Esc</span> ends the show.<br><br>Try it now — start it, navigate around, then end it.<br><br><i>(This step can't be auto-verified either: a running slideshow takes over the whole screen, so this task pane isn't even active while it's happening — there's no moment for it to check anything. Marked done automatically.)</i>`,
        done: "F5 from the start, Shift+F5 from the current slide, arrow keys to navigate, Esc to end.",
        setup: async () => {},
        check: async () => true
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — EXCO Pitch Deck, Final Proof
     ====================================================================== */
  assignment: {
    hook: "The pitch deck goes in front of EXCO in an hour and nobody's proofread it — two slides still have spelling mistakes, and one has a word typed twice.",
    brief: "Fix every spelling mistake and the repeated word before it goes up on screen.",
    doneMsg: "Deck's clean. (Do still run it through Slide Show view yourself before the real thing — that part genuinely can't be checked from here.)",

    setup: async () => {
      await PowerPoint.run(async (context) => {
        await pptPresentResetToSingleBlankSlide(context);
        context.presentation.slides.add();
        await context.sync();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        slides.items[0].shapes.addTextBox("Thank you for you're continued suport of Founders' Day.");
        slides.items[1].shapes.addTextBox("This has been an incredible incredible year for the Prefects' Board.");
        await context.sync();
      });
    },

    tasks: [
      { ref: "4.4.1", text: "Slide 1 has two mistakes — <code>you're</code> should read <code>your</code>, and <code>suport</code> should read <code>support</code>.",
        hint: "Right-click each red underline, or use Find & Replace.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const s = slides.items[0];
          s.shapes.load("items");
          await context.sync();
          s.shapes.items.forEach(sh => sh.textFrame.load("hasText"));
          await context.sync();
          s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) sh.textFrame.textRange.load("text"); });
          await context.sync();
          const joined = s.shapes.items.filter(sh => sh.textFrame.hasText).map(sh => sh.textFrame.textRange.text).join(" ");
          return /your continued support/i.test(joined);
        }) },

      { ref: "4.4.1", text: "Slide 2 has \"<b>incredible incredible</b>\" — delete the repeated word.",
        hint: "Click right before the second \"incredible\" and delete it (and the extra space).",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const s = slides.items[1];
          s.shapes.load("items");
          await context.sync();
          s.shapes.items.forEach(sh => sh.textFrame.load("hasText"));
          await context.sync();
          s.shapes.items.forEach(sh => { if (sh.textFrame.hasText) sh.textFrame.textRange.load("text"); });
          await context.sync();
          const joined = s.shapes.items.filter(sh => sh.textFrame.hasText).map(sh => sh.textFrame.textRange.text).join(" ");
          return !/\bincredible\s+incredible\b/i.test(joined) && /incredible/i.test(joined);
        }) }
    ]
  }
};
