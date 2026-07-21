// lessons/ppt-build.js
// Syllabus refs 4.1.1-4.1.8 (insert/delete slide, enter text). First
// PowerPoint lesson — deliberately smaller in scope than the Excel/Word
// lessons. Two reasons:
//
//   1. Live-checking has no reliable event on this host (see taskpane.html)
//      — it's poll-driven, so there's naturally less "instant" feedback.
//   2. The PowerPoint JS API for setting/reading shape text is the part
//      I have the least confidence writing correctly without being able
//      to test it live. Checks here search ALL shapes on a slide for
//      matching text rather than assuming which shape is "the title
//      placeholder" — safer, but if it still doesn't behave as expected,
//      this file is where to look first.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function resetToSingleBlankSlide(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  // delete every slide except the first
  for(let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
  await context.sync();
}

async function slideHasText(context, slide, text){
  slide.shapes.load("items");
  await context.sync();
  for(const shape of slide.shapes.items){
    shape.textFrame.load("hasText");
  }
  await context.sync();
  for(const shape of slide.shapes.items){
    if(shape.textFrame.hasText){
      shape.textFrame.textRange.load("text");
    }
  }
  await context.sync();
  const normalized = (text || "").toLowerCase();
  return slide.shapes.items.some(shape =>
    shape.textFrame.hasText && (shape.textFrame.textRange.text || "").toLowerCase().includes(normalized)
  );
}

window.LESSON_MODULES["ppt-build"] = {
  title: "Building Your Deck",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The basic moves for assembling any deck. Each step resets the presentation to a small, focused starting point.",

    steps: [
      {
        title: "Insert a new slide",
        teach: `<b>Home tab → New Slide</b> (or right-click in the slide thumbnail panel on the left → New Slide) adds a blank slide after the current one.<br><br>This presentation currently has just one slide. Add a second one.`,
        done: "New Slide (Home tab, or right-click the thumbnail panel) adds one after the current slide.",
        setup: async () => PowerPoint.run(async (context) => {
          await resetToSingleBlankSlide(context);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return slides.items.length >= 2;
        })
      },

      {
        title: "Delete a slide",
        teach: `Right-click a slide's thumbnail on the left → <b>Delete Slide</b>. Everything after it shifts up to fill the gap.<br><br>This presentation has three slides — the last one was added by mistake. Delete it.`,
        done: "Right-click the thumbnail → Delete Slide.",
        setup: async () => PowerPoint.run(async (context) => {
          await resetToSingleBlankSlide(context);
          context.presentation.slides.add();
          context.presentation.slides.add();
          await context.sync();
        }),
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return slides.items.length === 2;
        })
      },

      {
        title: "Add a title",
        teach: `Click directly into the title placeholder on the slide (the big text box near the top) and type. It works like any text box — click, type, done.<br><br>Type <code>Founders' Day Pitch</code> into this slide's title.`,
        done: "Click the placeholder, type — same as any text box.",
        setup: async () => PowerPoint.run(async (context) => {
          await resetToSingleBlankSlide(context);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if(slides.items.length === 0) return false;
          return await slideHasText(context, slides.items[0], "Founders' Day Pitch");
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Pitch Deck, First Pass
     ====================================================================== */
  assignment: {
    hook: "The EXCO pitch deck needs a first pass before Thursday — one slide was duplicated by accident during setup, and the opening slide still has no title.",
    brief: "Clean up the duplicate slide and title the opening slide properly.",
    doneMsg: "First pass done. Deliberately a small checklist for now — PowerPoint's the newest host in this pane, kept modest on purpose.",

    setup: async () => {
      await PowerPoint.run(async (context) => {
        await resetToSingleBlankSlide(context);
        context.presentation.slides.add();
        context.presentation.slides.add();
        context.presentation.slides.add(); // total: 4 slides, one is the "mistake" to remove
        await context.sync();
      });
    },

    tasks: [
      { ref: "4.1.6", text: "Type <b>Founders' Day Pitch</b> into the title of the first slide.",
        hint: "Click directly into the title placeholder near the top of slide 1 and type.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          if(slides.items.length === 0) return false;
          return await slideHasText(context, slides.items[0], "Founders' Day Pitch");
        }) },

      { ref: "4.1.3", text: "One slide was added by mistake during setup — the deck should have <b>3</b> slides, not 4. Delete the extra one.",
        hint: "Right-click the extra slide's thumbnail on the left → Delete Slide.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return slides.items.length === 3;
        }) }
    ]
  }
};