// lessons/ppt-text.js
// Syllabus refs 4.2.1-4.2.8.
//
// CONTINUITY: each step gets its OWN dedicated slide, identified by a
// caption/marker shape whose text is never the thing being edited — so
// searching for it always finds the right slide regardless of visit
// order, and other steps' slides never interfere. (See ppt-build.js for
// the same marker-slide technique applied to deck-structure exercises.)
//
// HONEST CAVEAT (4.2.6): shadow is real syllabus content (Home tab → Font
// group → Text Shadow) but there is no `shadow` property anywhere in the
// PowerPoint JS API's TextFrame/TextRange/Font object model — confirmed
// absent, not an oversight. It's mentioned in the teach text as something
// to try, but the check only gates on bold/italic/underline, which are
// real, checkable properties.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function pptTextResetToSingleBlankSlide(context){
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  for(let i = slides.items.length - 1; i > 0; i--) slides.items[i].delete();
  await context.sync();
}

// Returns [{slideIndex, shapes:[{shape,text}]}] entries for every slide
// that has at least one shape containing `needle`, with every shape's text
// on those slides loaded and ready to read synchronously afterwards.
async function pptTextFindSlidesWithText(context, needle){
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
    const texts = s.shapes.items.map(sh => ({ shape: sh, text: sh.textFrame.hasText ? (sh.textFrame.textRange.text || "") : "" }));
    if (texts.some(t => t.text.toLowerCase().includes(norm))) hits.push({ slideIndex: i, shapes: texts });
  });
  return hits;
}

async function pptTextEnsureSlide(context, markerText, extraTexts){
  const existing = await pptTextFindSlidesWithText(context, markerText);
  if (existing.length > 0) return;
  context.presentation.slides.add();
  await context.sync();
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const slide = slides.items[slides.items.length - 1];
  // addTextBox() drops every new box at the same default position, so
  // multiple boxes on one slide land on top of each other unless each
  // gets an explicit, distinct position — stack them vertically.
  const caption = slide.shapes.addTextBox(markerText);
  caption.left = 40; caption.top = 40; caption.width = 600; caption.height = 60;
  (extraTexts || []).forEach((t, i) => {
    const box = slide.shapes.addTextBox(t);
    box.left = 40; box.top = 140 + i * 100; box.width = 600; box.height = 60;
  });
  await context.sync();
}

const PPT_TEXT_DEFAULT_FONTS = ["calibri", "calibri light", "aptos", "aptos display", "arial"];
function pptTextIsDefaultFont(name){
  const n = (name || "").toLowerCase().trim();
  return PPT_TEXT_DEFAULT_FONTS.includes(n);
}

window.LESSON_MODULES["ppt-text"] = {
  title: "Slide Text and Formatting",

  /* ======================================================================
     TUTORIAL — one mechanic per step, each on its own dedicated slide
     ====================================================================== */
  tutorial: {
    hook: "Getting words onto a slide is only half of it — editing, moving, and dressing them up is the rest. Each step here has its own slide, so nothing you've already done gets undone as you move around.",

    prepare: async () => PowerPoint.run(async (context) => {
      await pptTextResetToSingleBlankSlide(context);
    }),

    steps: [
      {
        title: "Enter text in a placeholder",
        teach: `Click directly into a text box or placeholder, then type — same as clicking into any text field.<br><br>Find the slide with an empty box under the instruction, and type <code>RIPB</code> into it.`,
        done: "Click the box, type — it works like any text field.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Type your board's motto in the empty box below:", [""]);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptTextFindSlidesWithText(context, "Type your board's motto in the empty box below:");
          if (hits.length === 0) return false;
          const answer = hits[0].shapes.find(s => !s.text.toLowerCase().includes("type your board's motto"));
          return !!answer && answer.text.toLowerCase().includes("ripb");
        })
      },

      {
        title: "Edit existing text",
        teach: `Click into a text box, select the part that's wrong, and retype it — same as editing text anywhere else.<br><br>The session name below is incomplete. It should read "<i>Founders' Day Volunteer Info Session</i>" — edit it to add the missing word.`,
        done: "Click in, select the wrong bit, retype — no need to delete the whole box.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Fix the session name below so it matches this year's event:", ["Founders' Day Info Session"]);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptTextFindSlidesWithText(context, "Fix the session name below so it matches this year's event:");
          if (hits.length === 0) return false;
          const answer = hits[0].shapes.find(s => !s.text.toLowerCase().includes("fix the session name"));
          return !!answer && /founders/i.test(answer.text) && /volunteer/i.test(answer.text);
        })
      },

      {
        title: "Copy text between slides",
        teach: `Select the text, <span class="keys">Ctrl/Cmd+C</span>, click into the box on the other slide, <span class="keys">Ctrl/Cmd+V</span>. The original stays exactly where it was — copy duplicates, it doesn't move.<br><br>Copy the event name from the "Copy Source Slide" onto the empty box on "Copy Destination Slide".`,
        done: "Copy leaves the source slide untouched — that's what makes it copy, not cut.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Copy Source Slide", ["Founders' Day 2026"]);
          await pptTextEnsureSlide(context, "Copy Destination Slide", [""]);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const src = await pptTextFindSlidesWithText(context, "Copy Source Slide");
          const dst = await pptTextFindSlidesWithText(context, "Copy Destination Slide");
          if (src.length === 0 || dst.length === 0) return false;
          const sourceStillHasIt = src[0].shapes.some(s => s.text.includes("Founders' Day 2026"));
          const destNowHasIt = dst[0].shapes.some(s => s.text.includes("Founders' Day 2026") && !s.text.toLowerCase().includes("copy destination slide"));
          return sourceStillHasIt && destNowHasIt;
        })
      },

      {
        title: "Delete text",
        teach: `Select just the part you don't want and press <span class="keys">Delete</span> — the rest of the box is untouched.<br><br>Below, remove " - DRAFT, do not print" so the box just reads the event name.`,
        done: "Select only the unwanted part, Delete — everything else stays.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Delete the extra words below so it just reads the event name:", ["Founders' Day 2026 - DRAFT, do not print"]);
        }),
        check: async () => PowerPoint.run(async (context) => {
          const hits = await pptTextFindSlidesWithText(context, "Delete the extra words below so it just reads the event name:");
          if (hits.length === 0) return false;
          const answer = hits[0].shapes.find(s => !s.text.toLowerCase().includes("delete the extra words"));
          return !!answer && answer.text.includes("Founders' Day 2026") && !/draft/i.test(answer.text);
        })
      },

      {
        title: "Font size and font type",
        teach: `Select the text, then use the <b>font size</b> box and <b>font name</b> box on the Home tab.<br><br>Find the slide reading "Founders' Day 2026" and make it at least <b>28pt</b> in a font that isn't the theme's default.`,
        done: "Bigger size + a chosen font makes text read as a real headline, not leftover default styling.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Founders' Day 2026", []);
        }),
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptTextFindSlidesWithText(context, "Founders' Day 2026");
              if (hits.length === 0) return false;
              const shape = hits[0].shapes.find(s => s.text.includes("Founders' Day 2026")).shape;
              shape.textFrame.textRange.font.load("size,name");
              await context.sync();
              const font = shape.textFrame.textRange.font;
              return font.size >= 28 && !pptTextIsDefaultFont(font.name);
            });
          } catch (e) { return false; }
        }
      },

      {
        title: "Bold, italic, and underline",
        teach: `These three stack — select the text, then apply any combination with the Home tab buttons (or <span class="keys">Ctrl/Cmd+B</span>, <span class="keys">Ctrl/Cmd+I</span>, <span class="keys">Ctrl/Cmd+U</span>). There's also a <b>Text Shadow</b> button in the same area worth trying, though this pane has no way to check for it.<br><br>Find "RSVP by 1 August" and make it bold, italic, and underlined all at once.`,
        done: "Bold, italic, and underline are independent toggles — stack all three.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "RSVP by 1 August", []);
        }),
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptTextFindSlidesWithText(context, "RSVP by 1 August");
              if (hits.length === 0) return false;
              const shape = hits[0].shapes.find(s => s.text.includes("RSVP by 1 August")).shape;
              shape.textFrame.textRange.font.load("bold,italic,underline");
              await context.sync();
              const font = shape.textFrame.textRange.font;
              return font.bold === true && font.italic === true && !!font.underline && String(font.underline) !== "None";
            });
          } catch (e) { return false; }
        }
      },

      {
        title: "Font colour",
        teach: `Select the text, then use the dropdown arrow on the <b>font colour</b> button (Home tab) to pick anything besides the default.<br><br>Find "Limited seats available!" and give it a font colour that makes it pop.`,
        done: "The dropdown arrow next to the font colour button opens the full palette.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "Limited seats available!", []);
        }),
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptTextFindSlidesWithText(context, "Limited seats available!");
              if (hits.length === 0) return false;
              const shape = hits[0].shapes.find(s => s.text.includes("Limited seats available!")).shape;
              shape.textFrame.textRange.font.load("color");
              await context.sync();
              const c = (shape.textFrame.textRange.font.color || "").toLowerCase();
              return !!c && c !== "#000000" && c !== "black" && c !== "automatic";
            });
          } catch (e) { return false; }
        }
      },

      {
        title: "Align text in a text frame",
        teach: `Left, centre, right — the alignment buttons on the Home tab work inside a text box exactly like they do in a document.<br><br>Find "You're Invited!" and centre it.`,
        done: "Same alignment buttons as any other app — they apply to whichever text box is selected.",
        setup: async () => PowerPoint.run(async (context) => {
          await pptTextEnsureSlide(context, "You're Invited!", []);
        }),
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const hits = await pptTextFindSlidesWithText(context, "You're Invited!");
              if (hits.length === 0) return false;
              const shape = hits[0].shapes.find(s => s.text.includes("You're Invited!")).shape;
              shape.textFrame.textRange.paragraphFormat.load("horizontalAlignment");
              await context.sync();
              return shape.textFrame.textRange.paragraphFormat.horizontalAlignment === PowerPoint.ParagraphHorizontalAlignment.center;
            });
          } catch (e) { return false; }
        }
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — EXCO Pitch Deck, Content Pass
     ====================================================================== */
  assignment: {
    hook: "The Agenda slide's text is a mess — a placeholder never got filled in, the venue name is wrong, and the closing call-to-action needs to actually look like one.",
    brief: "Fill in the empty placeholder, fix the venue name, and format the call-to-action so it stands out.",
    doneMsg: "Content's in and it reads properly. Every text move from the tutorial, shaping a slide someone's about to present.",

    setup: async () => {
      await PowerPoint.run(async (context) => {
        await pptTextResetToSingleBlankSlide(context);
        context.presentation.slides.add();
        context.presentation.slides.add();
        await context.sync();
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const label = slides.items[0].shapes.addTextBox("Presenter:");
        label.left = 40; label.top = 40; label.width = 400; label.height = 50;
        const nameBox = slides.items[0].shapes.addTextBox("");
        nameBox.left = 40; nameBox.top = 110; nameBox.width = 400; nameBox.height = 50;
        const venue = slides.items[1].shapes.addTextBox("Venue: Old Main Hal");
        venue.left = 40; venue.top = 40; venue.width = 500; venue.height = 50;
        const cta = slides.items[2].shapes.addTextBox("Sponsor us today");
        cta.left = 40; cta.top = 40; cta.width = 500; cta.height = 50;
        await context.sync();
      });
    },

    tasks: [
      { ref: "4.2.1", text: "Slide 1 has an empty box after \"Presenter:\" — type your own name into it.",
        hint: "Click the empty box and type any name.",
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
          return s.shapes.items.some(sh => sh.textFrame.hasText && (sh.textFrame.textRange.text || "").trim().length > 0 && !sh.textFrame.textRange.text.includes("Presenter:"));
        }) },

      { ref: "4.2.2", text: "Slide 2's venue name is misspelled — \"<code>Old Main Hal</code>\" should read \"<code>Old Main Hall</code>\".",
        hint: "Click into the box, fix the spelling.",
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
          return s.shapes.items.some(sh => sh.textFrame.hasText && /old main hall/i.test(sh.textFrame.textRange.text || ""));
        }) },

      { ref: "4.2.6 / 4.2.7", text: "Slide 3's call-to-action \"<b>Sponsor us today</b>\" needs to stand out — make it bold and give it a font colour besides the default.",
        hint: "Select the text: Ctrl/Cmd+B, then pick a colour from the font colour dropdown.",
        check: async () => {
          try {
            return await PowerPoint.run(async (context) => {
              const slides = context.presentation.slides;
              slides.load("items");
              await context.sync();
              const s = slides.items[2];
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
  }
};
