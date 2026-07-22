// lessons/ppt-objects.js
// Syllabus refs 4.3.1-4.3.4.
//
// The syllabus explicitly treats "picture" and "drawn object" as equally
// valid for 4.3.1, so this lesson drills both. Chart-specific mentions in
// 4.3.2/4.3.3 (copy/resize a chart) aren't given their own dedicated check
// here — inserting and reading back a PowerPoint chart object is enough of
// an unverified corner of the JS API that it isn't worth the risk; the
// resizing mechanic taught in "Resize an object" is identical for a chart
// in practice (drag a corner handle), and the teach text says so.
//
// HONEST CAVEAT (4.3.4): there is no transitions API anywhere in the
// current PowerPoint object model — confirmed absent, not an oversight —
// so that step self-certifies, same approach as the housekeeping step at
// the end of ppt-build.js.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function pptObjResetToSingleBlankSlide(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  for(let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
  await context.sync();
}

async function pptObjFindSlidesWithText(context, needle){
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
    const shapes = s.shapes.items.map(sh => ({
      shape: sh,
      type: String(sh.type || ""),
      text: sh.textFrame.hasText ? (sh.textFrame.textRange.text || "") : ""
    }));
    if (shapes.some(x => x.text.toLowerCase().includes(norm))) hits.push({ slideIndex: i, shapes });
  });
  return hits;
}

async function pptObjEnsureMarkerSlide(context, markerText){
  const existing = await pptObjFindSlidesWithText(context, markerText);
  if (existing.length > 0) return existing[0];
  context.presentation.slides.add();
  await context.sync();
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const slide = slides.items[slides.items.length - 1];
  const caption = slide.shapes.addTextBox(markerText);
  caption.left = 20; caption.top = 20; caption.width = 500; caption.height = 50;
  await context.sync();
  return null;
}

window.LESSON_MODULES["ppt-objects"] = {
  title: "Pictures, Objects and Transitions",

  /* ======================================================================
     TUTORIAL — one mechanic per step, each on its own dedicated slide
     ====================================================================== */
  tutorial: {
    hook: "Slides are more than text — pictures, shapes, and the transitions between them are what make a deck feel finished. Each step here has its own slide.",

    prepare: async () => PowerPoint.run(async (context) => {
      await pptObjResetToSingleBlankSlide(context);
    }),

    steps: [
      {
        title: "Insert a picture",
        teach: `<b>Insert tab → Pictures</b> (This Device, or Stock Images/Online Pictures) drops an image onto the slide at a default size and position.<br><br>Find the slide marked "Insert Picture Practice" and insert any picture onto it.`,
        done: "Insert → Pictures places an image directly on the current slide.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptObjEnsureMarkerSlide(context, "Insert Picture Practice");
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptObjFindSlidesWithText(context, "Insert Picture Practice");
          if (hits.length === 0) return false;
          return hits[0].shapes.some(s => /image|picture/i.test(s.type));
        })
      },

      {
        title: "Insert a drawn shape",
        teach: `<b>Insert tab → Shapes</b>, pick one, then drag on the slide to draw it — same idea as a picture, just a shape you draw instead of a file you place.<br><br>Find the slide marked "Insert Shape Practice" and draw any shape onto it.`,
        done: "Insert → Shapes, pick one, drag it out on the slide.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptObjEnsureMarkerSlide(context, "Insert Shape Practice");
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptObjFindSlidesWithText(context, "Insert Shape Practice");
          if (hits.length === 0) return false;
          return hits[0].shapes.some(s => /geometric|shape/i.test(s.type) && !/textbox/i.test(s.type));
        })
      },

      {
        title: "Resize an object",
        teach: `Click an object once to select it — handles appear at the corners and edges. Drag a <b>corner</b> handle to resize while keeping its proportions; drag a <b>side</b> handle to stretch it out of shape instead. (A chart on a slide resizes exactly the same way.)<br><br>Find the slide marked "Resize Object Practice" and resize the shape sitting on it.`,
        done: "Corner handle = keeps proportions. Side handle = stretches it. Same for a chart.",
        setup: async () => PowerPoint.run(async (context) => {
          const hit = await pptObjEnsureMarkerSlide(context, "Resize Object Practice");
          if (hit) return;
          const hits = await pptObjFindSlidesWithText(context, "Resize Object Practice");
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const slide = slides.items[hits[0].slideIndex];
          const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
          shape.left = 100; shape.top = 100; shape.width = 100; shape.height = 100;
          await context.sync();
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptObjFindSlidesWithText(context, "Resize Object Practice");
          if (hits.length === 0) return false;
          const shape = hits[0].shapes.find(s => /geometric/i.test(s.type));
          if (!shape) return false;
          shape.shape.load("width,height");
          await context.sync();
          return shape.shape.width !== 100 || shape.shape.height !== 100;
        })
      },

      {
        title: "Copy or move an object to another slide",
        teach: `Select the object, then either <span class="keys">Ctrl/Cmd+C</span> then <span class="keys">Ctrl/Cmd+V</span> on the destination slide to <b>copy</b> it (original stays put), or <span class="keys">Ctrl/Cmd+X</span> then <span class="keys">Ctrl/Cmd+V</span> to <b>move</b> it instead (original disappears). Both are valid here.<br><br>Copy or move the shape from "Object Move Source" onto "Object Move Destination".`,
        done: "Ctrl+C/Ctrl+V copies, Ctrl+X/Ctrl+V moves — pick whichever the situation calls for.",
        setup: async () => PowerPoint.run(async (context) => {
          const srcHit = await pptObjEnsureMarkerSlide(context, "Object Move Source");
          await pptObjEnsureMarkerSlide(context, "Object Move Destination");
          if (srcHit) return;
          const hits = await pptObjFindSlidesWithText(context, "Object Move Source");
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const slide = slides.items[hits[0].slideIndex];
          const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
          shape.left = 100; shape.top = 100; shape.width = 80; shape.height = 80;
          await context.sync();
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptObjFindSlidesWithText(context, "Object Move Destination");
          if (hits.length === 0) return false;
          return hits[0].shapes.some(s => /geometric/i.test(s.type));
        })
      },

      {
        title: "Add a transition between slides",
        teach: `Select a slide, then <b>Transitions tab</b> → click any transition thumbnail (Fade, Push, Wipe, and so on) to apply it. <b>Apply To All</b> puts the same one on every slide.<br><br>Try adding a transition to any slide in this deck.<br><br><i>(This step can't be auto-verified — the PowerPoint JS API has no way to read back which transition, if any, is applied to a slide. Marked done automatically; the point is knowing where to find it.)</i>`,
        done: "Transitions tab → click a thumbnail. Apply To All matches every slide to it.",
        setup: async () => {},
        check: async () => true
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — EXCO Pitch Deck, Visual Pass
     ====================================================================== */
  assignment: {
    hook: "The pitch deck's slides are all text so far — the Board wants at least one visual on the venue slide, and the logo shape that's supposed to open the deck ended up stuck on the wrong slide.",
    brief: "Add a picture or shape to the venue slide, resize it so it isn't tiny, and move the logo shape onto the opening slide where it belongs.",
    doneMsg: "Deck's got visuals now, and the logo's where it should be. Every object move from the tutorial, doing real deck work.",

    setup: async () => {
      await PowerPoint.run(async (context) => {
        await pptObjResetToSingleBlankSlide(context);
        context.presentation.slides.add();
        await context.sync();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const title = slides.items[0].shapes.addTextBox("Founders' Day Pitch");
        title.left = 40; title.top = 40; title.width = 500; title.height = 60;
        const venue = slides.items[1].shapes.addTextBox("Venue: Old Main Hall");
        venue.left = 40; venue.top = 150; venue.width = 500; venue.height = 60;
        const logo = slides.items[1].shapes.addGeometricShape(PowerPoint.GeometricShapeType.hexagon);
        logo.left = 20; logo.top = 20; logo.width = 40; logo.height = 40;
        await context.sync();
      });
    },

    tasks: [
      { ref: "4.3.1", text: "Add a picture or drawn shape to the venue slide so it isn't just text.",
        hint: "Insert tab → Pictures, or Insert tab → Shapes.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const s = slides.items[1];
          s.shapes.load("items");
          await context.sync();
          s.shapes.items.forEach(sh => sh.load("type,width,height"));
          await context.sync();
          return s.shapes.items.filter(sh => /image|geometric/i.test(String(sh.type))).length >= 2;
        }) },

      { ref: "4.3.3", text: "Whatever you just added is probably tiny at its default size — resize it so it's clearly bigger.",
        hint: "Click it, drag a corner handle outward.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const s = slides.items[1];
          s.shapes.load("items");
          await context.sync();
          s.shapes.items.forEach(sh => sh.load("type,width,height"));
          await context.sync();
          return s.shapes.items.some(sh => /image|geometric/i.test(String(sh.type)) && (sh.width > 100 || sh.height > 100));
        }) },

      { ref: "4.3.2", text: "The little hexagon logo shape is stuck on the venue slide — copy or move it onto the opening slide instead.",
        hint: "Select the hexagon, Ctrl+X (or Ctrl+C) on the venue slide, click the opening slide, Ctrl+V.",
        check: async () => PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const s = slides.items[0];
          s.shapes.load("items");
          await context.sync();
          s.shapes.items.forEach(sh => sh.load("type"));
          await context.sync();
          return s.shapes.items.some(sh => String(sh.type).toLowerCase().includes("geometric"));
        }) }
    ]
  }
};
