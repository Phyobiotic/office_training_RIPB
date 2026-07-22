// lessons/ppt-build.js
// Syllabus refs 4.1.2-4.1.8 (4.1.1 "create a new presentation from a
// template" is skipped — the add-in loads into a presentation that's
// already open, same reasoning as skipping 2.1.1 in the Word lessons).
//
// REWRITE NOTE: the original version of this file tested typing a slide
// title, which actually belongs to 4.2.1 (Text and Formatting) — that
// exercise moved to ppt-text.js, and this file now sticks to genuine 4.1.x
// "Getting Started" deck-structure content.
//
// CONTINUITY: slides don't have addressable "ranges" the way cells or
// paragraphs do, so every step here identifies "its" slide by planting a
// unique marker text box on it (via shapes.addTextBox) and searching ALL
// slides for that marker text — never by position/index, since other
// steps keep adding slides of their own. This mirrors the marker-paragraph
// technique in the Word lessons, just applied to slides instead.
//
// HONEST CAVEAT: PowerPoint's JS API has no way to read/set slide
// transitions, whether a slide is hidden, which view (Normal/Slide
// Sorter/Notes/Outline) is showing, or which design theme is applied —
// confirmed absent from the current object model, not an oversight here.
// Those pieces of 4.1.2/4.1.4/4.1.8 are taught in the last step below with
// a self-certifying check, same honesty as the print/preview steps in
// word-output.js.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function pptBuildResetToSingleBlankSlide(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  for(let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
  await context.sync();
}

// Loads every slide's shapes and text in a fixed number of sync round trips
// (independent of slide count), then reports which slide indexes contain
// the given marker text anywhere among their shapes.
async function pptBuildSlideIndexesWithText(context, needle){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  slides.items.forEach(s => s.shapes.load("items"));
  await context.sync();
  slides.items.forEach(s => s.shapes.items.forEach(sh => sh.textFrame.load("hasText")));
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

async function pptBuildAddMarkerSlide(context, markerText){
  context.presentation.slides.add();
  await context.sync();
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const newSlide = slides.items[slides.items.length - 1];
  newSlide.shapes.addTextBox(markerText);
  await context.sync();
  return newSlide;
}

window.LESSON_MODULES["ppt-build"] = {
  title: "Building Your Deck",

  /* ======================================================================
     TUTORIAL — one mechanic per step, same deck all the way through
     ====================================================================== */
  tutorial: {
    hook: "The basic moves for assembling any deck. It's the same presentation the whole way through — what you add, delete, or reorder in one step is still there in the next.",

    prepare: async () => PowerPoint.run(async (context) => {
      await pptBuildResetToSingleBlankSlide(context);
    }),

    steps: [
      {
        title: "Give a slide a different layout",
        teach: `A new slide's layout isn't fixed. Either pick one at insert time — click the little <b>arrow under New Slide</b> (not the icon itself) to open the layout gallery — or change an existing slide's layout afterwards via <b>Home tab → Layout</b>.<br><br>A slide marked "Layout Practice Slide" is in this deck. Give it a layout that's clearly <b>different</b> from slide 1's layout.`,
        done: "New Slide's dropdown arrow picks a layout up front; Home → Layout changes it on any existing slide.",
        setup: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Layout Practice Slide");
          if (hits.length === 0) await pptBuildAddMarkerSlide(context, "Layout Practice Slide");
        }),
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptBuildSlideIndexesWithText(context, "Layout Practice Slide");
              if (hits.length === 0) return false;
              const slides = context.presentation.slides;
              slides.load("items");
              await context.sync();
              const marker = slides.items[hits[0]];
              const original = slides.items[0];
              marker.layout.load("name");
              original.layout.load("name");
              await context.sync();
              return (marker.layout.name || "").toLowerCase() !== (original.layout.name || "").toLowerCase();
            });
          } catch (e) { return false; }
        }
      },

      {
        title: "Delete a slide",
        teach: `Right-click a slide's thumbnail on the left → <b>Delete Slide</b>. Everything after it shifts up to fill the gap.<br><br>Find the slide marked "Delete Me" and delete it.`,
        done: "Right-click the thumbnail → Delete Slide.",
        setup: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Delete Me");
          if (hits.length === 0) await pptBuildAddMarkerSlide(context, "Delete Me");
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Delete Me");
          return hits.length === 0;
        })
      },

      {
        title: "Duplicate a slide",
        teach: `Right-click a slide's thumbnail → <b>Duplicate Slide</b> (or copy it with Ctrl/Cmd+C and paste it back in with Ctrl/Cmd+V). Either way, you end up with two identical slides.<br><br>Find the slide marked "Duplicate Me" and duplicate it, so that text appears on two slides.`,
        done: "Right-click → Duplicate Slide, or copy-paste the thumbnail — same result.",
        setup: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Duplicate Me");
          if (hits.length === 0) await pptBuildAddMarkerSlide(context, "Duplicate Me");
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Duplicate Me");
          return hits.length >= 2;
        })
      },

      {
        title: "Move a slide",
        teach: `Drag a slide's thumbnail up or down in the left-hand panel to reorder it — drop it where you want it to land.<br><br>Three slides are marked "Move-First", "Move-Second", and "Move-Third", in that order. Drag "Move-Third" so it comes <b>before</b> the other two.`,
        done: "Drag a thumbnail in the panel on the left to reorder — drop it wherever it should go.",
        setup: async () => PowerPoint.run(async (context) => {
          // Run these one at a time — they all share one context, and firing
          // concurrent load()/sync() cycles on the same context via
          // Promise.all races the object-tracking state (that's what causes
          // "the property 'items' is not available" errors).
          const first = await pptBuildSlideIndexesWithText(context, "Move-First");
          const second = await pptBuildSlideIndexesWithText(context, "Move-Second");
          const third = await pptBuildSlideIndexesWithText(context, "Move-Third");
          if (first.length === 0 && second.length === 0 && third.length === 0) {
            await pptBuildAddMarkerSlide(context, "Move-First");
            await pptBuildAddMarkerSlide(context, "Move-Second");
            await pptBuildAddMarkerSlide(context, "Move-Third");
          }
        }),
        check: async () => PowerPoint.run(async (context) => {
          const first = await pptBuildSlideIndexesWithText(context, "Move-First");
          const second = await pptBuildSlideIndexesWithText(context, "Move-Second");
          const third = await pptBuildSlideIndexesWithText(context, "Move-Third");
          if (first.length === 0 || second.length === 0 || third.length === 0) return false;
          return third[0] < first[0] && third[0] < second[0];
        })
      },

      {
        title: "Views, themes, and hide/show",
        teach: `Three more everyday moves, worth knowing even though this tutorial has no way to watch you do them:<br><br>
          <b>Switch views</b> — the buttons at the bottom-right (or View tab) switch between Normal, Slide Sorter, Notes Page, and Outline. Slide Sorter is the fastest way to see and reorder a whole deck at once.<br><br>
          <b>Apply a theme</b> — Design tab → click any theme thumbnail to restyle every slide at once.<br><br>
          <b>Hide a slide</b> — right-click a thumbnail → Hide Slide. It stays in the file but gets skipped during a slideshow.<br><br>
          Try all three now.<br><br><i>(None of these change anything the JS API can read back, so this step is marked done automatically — it's here so the lesson doesn't skip real syllabus content.)</i>`,
        done: "Views: bottom-right buttons or View tab. Themes: Design tab. Hide: right-click a thumbnail.",
        setup: async () => {},
        check: async () => true
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — EXCO Pitch Deck, Structure Pass
     ====================================================================== */
  assignment: {
    hook: "The EXCO pitch deck needs a structure pass before Thursday — the Agenda slide got duplicated during setup, Opening Remarks somehow ended up after Closing, the Agenda slide's layout doesn't fit its content, and the opening slide still has no title.",
    brief: "Title the opening slide, remove the duplicate Agenda slide, move Opening Remarks back before Closing, and give the Agenda slide a layout that actually suits it.",
    doneMsg: "Structure's fixed and ready for content. Every deck-building move from the tutorial, cleaning up a real mess this time.",

    setup: async () => {
      await PowerPoint.run(async (context) => {
        await pptBuildResetToSingleBlankSlide(context);
        await pptBuildAddMarkerSlide(context, "Agenda");
        await pptBuildAddMarkerSlide(context, "Agenda");
        await pptBuildAddMarkerSlide(context, "Closing");
        await pptBuildAddMarkerSlide(context, "Opening Remarks");
      });
    },

    tasks: [
      { ref: "4.2.1", text: "Type <b>Founders' Day Pitch</b> into the title of the first slide.",
        hint: "Click directly into the title placeholder near the top of slide 1 and type.",
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Founders' Day Pitch");
          return hits.includes(0);
        }) },

      { ref: "4.1.7", text: "The <b>Agenda</b> slide got duplicated by mistake — delete the extra copy so it appears only once.",
        hint: "Right-click one of the two Agenda thumbnails → Delete Slide.",
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptBuildSlideIndexesWithText(context, "Agenda");
          return hits.length === 1;
        }) },

      { ref: "4.1.6", text: "<b>Opening Remarks</b> should come before <b>Closing</b>, not after — drag it into the right order.",
        hint: "Drag the Opening Remarks thumbnail so it sits above the Closing thumbnail.",
        check: async () => PowerPoint.run(async (context) => {
          const opening = await pptBuildSlideIndexesWithText(context, "Opening Remarks");
          const closing = await pptBuildSlideIndexesWithText(context, "Closing");
          if (opening.length === 0 || closing.length === 0) return false;
          return opening[0] < closing[0];
        }) },

      { ref: "4.1.3 / 4.1.5", text: "Give the <b>Agenda</b> slide a layout different from slide 1's, so its schedule content has room to breathe.",
        hint: "Select the Agenda slide, then Home tab → Layout → pick a different one.",
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptBuildSlideIndexesWithText(context, "Agenda");
              if (hits.length === 0) return false;
              const slides = context.presentation.slides;
              slides.load("items");
              await context.sync();
              const marker = slides.items[hits[0]];
              const original = slides.items[0];
              marker.layout.load("name");
              original.layout.load("name");
              await context.sync();
              return (marker.layout.name || "").toLowerCase() !== (original.layout.name || "").toLowerCase();
            });
          } catch (e) { return false; }
        } }
    ]
  }
};
