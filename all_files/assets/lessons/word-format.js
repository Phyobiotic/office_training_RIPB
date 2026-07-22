// lessons/word-format.js
// Syllabus refs 2.2.1-2.2.7.
//
// tutorial = stepper, but a CONTINUOUS one — see word-edit.js for the
// pattern: `prepare()` clears the body once on entry, each step's setup()
// only appends its own marker paragraph(s) if they're not already present,
// so jumping between steps in any order never disturbs earlier work.
// assignment = one scenario, full checklist (shape unchanged)

window.LESSON_MODULES = window.LESSON_MODULES || {};

// Word's default font has changed over time (Calibri for years, Aptos/Aptos
// Narrow on newer Microsoft 365 builds) and varies by install — mirrors the
// same helper used in xl-format.js, kept file-local since lesson modules
// don't share code between each other.
const WORD_DEFAULT_FONT_NAMES = ["calibri", "aptos", "aptos narrow", "aptos display", "aptos serif", "aptos mono"];
function wordFormatIsDefaultFontName(name) {
  const n = (name || "").toLowerCase().trim();
  return WORD_DEFAULT_FONT_NAMES.includes(n);
}

async function wordFormatFindPara(paras, matchFn){
  return paras.items.find(p => matchFn((p.text || "").trim()));
}

window.LESSON_MODULES["word-format"] = {
  title: "Formatting and Paragraphs",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The moves that turn plain typed text into something that reads like a real newsletter. It's one growing document the whole way through — each step adds its own paragraph below what's already there.",

    prepare: async () => Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
    }),

    steps: [
      {
        title: "Font size and font face",
        teach: `A heading should look like a heading. Select it, then use the <b>font size</b> box and the <b>font name</b> box on the Home tab.<br><br>Below is "<i>Founders' Day — Newsletter</i>" in the plain default font. Make it at least <b>20pt</b> and change it to any font that isn't your default.`,
        done: "Bigger size + a chosen font = it reads as a heading, not just another line.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Founders' Day — Newsletter")) {
            body.insertParagraph("Founders' Day — Newsletter", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = await wordFormatFindPara(paras, t => t === "Founders' Day — Newsletter");
          if (!p) return false;
          p.font.load("size,name");
          await context.sync();
          return p.font.size >= 20 && !wordFormatIsDefaultFontName(p.font.name);
        })
      },

      {
        title: "Bold, italic, and underline together",
        teach: `Formatting stacks — text can be bold, italic, <i>and</i> underlined all at once. Each button (or shortcut <span class="keys">Ctrl/Cmd+B</span>, <span class="keys">Ctrl/Cmd+I</span>, <span class="keys">Ctrl/Cmd+U</span>) toggles independently.<br><br>Below, select "<i>RSVP by 1 Aug</i>" and apply all three: bold, italic, and underline.`,
        done: "Bold, italic, and underline are independent toggles — stack all three for real emphasis.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("RSVP by 1 Aug")) {
            body.insertParagraph("RSVP by 1 Aug to secure your spot.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = await wordFormatFindPara(paras, t => t.includes("RSVP by 1 Aug"));
          if (!p) return false;
          p.font.load("bold,italic,underline");
          await context.sync();
          return p.font.bold === true && p.font.italic === true && !!p.font.underline && String(p.font.underline) !== "None";
        })
      },

      {
        title: "Font colour",
        teach: `Select text, then use the <b>font colour</b> button's dropdown arrow on the Home tab to pick any colour besides black/automatic.<br><br>Below, give "<i>Limited seats available!</i>" a font colour that makes it stand out.`,
        done: "The dropdown arrow next to the font colour button opens the full palette.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Limited seats available")) {
            body.insertParagraph("Limited seats available!", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = await wordFormatFindPara(paras, t => t === "Limited seats available!");
          if (!p) return false;
          p.font.load("color");
          await context.sync();
          const c = (p.font.color || "").toLowerCase();
          return c && c !== "#000000" && c !== "automatic";
        })
      },

      {
        title: "Merge two paragraphs into one",
        teach: `Click at the very <b>start</b> of a paragraph and press <span class="keys">Backspace</span> — it merges into the paragraph above, joining both into one.<br><br>Below are two lines that should really be a single sentence: "<i>Founders' Day begins at 9am.</i>" and "<i>Doors open at 8:30am.</i>" Merge them into one paragraph.`,
        done: "Backspace at the start of a paragraph pulls it up into the one above.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Doors open at 8:30am")) {
            body.insertParagraph("Founders' Day begins at 9am.", Word.InsertLocation.end);
            body.insertParagraph("Doors open at 8:30am.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const merged = paras.items.find(p => p.text.includes("Doors open at 8:30am") && p.text.includes("begins at 9am"));
          return !!merged;
        })
      },

      {
        title: "Align text",
        teach: `Left, centre, right, justified — the four alignment buttons on the Home tab (or <span class="keys">Ctrl/Cmd+E</span> for centre).<br><br>Below, centre the line "<i>You're Invited!</i>".`,
        done: "Ctrl/Cmd+E centres the current paragraph — no need to select the whole line first.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("You're Invited!")) {
            body.insertParagraph("You're Invited!", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = await wordFormatFindPara(paras, t => t === "You're Invited!");
          if (!p) return false;
          p.load("alignment");
          await context.sync();
          return p.alignment === Word.Alignment.centered;
        })
      },

      {
        title: "Paragraph and line spacing",
        teach: `<b>Line spacing</b> controls the gap between lines <i>inside</i> a paragraph; <b>spacing after</b> controls the gap <i>between</i> paragraphs. Both live in the Line and Paragraph Spacing button on the Home tab.<br><br>Below, set this paragraph to <b>double</b> line spacing and add some space after it.`,
        done: "Line spacing = inside a paragraph. Space before/after = between paragraphs. Different knobs.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("printed programme")) {
            body.insertParagraph("This paragraph is crammed into the printed programme.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = await wordFormatFindPara(paras, t => t.includes("printed programme"));
          if (!p) return false;
          p.load("lineSpacing,spaceAfter");
          await context.sync();
          return p.lineSpacing >= 22 && p.spaceAfter > 0;
        })
      },

      {
        title: "Bullets, then switch to numbers",
        teach: `<b>Home tab → Bullets</b> turns selected lines into a bulleted list. Click the little arrow next to <b>Numbering</b> instead (or just click Numbering) to switch the <i>same</i> list over to numbers.<br><br>Below are three volunteer shifts. Turn them into a list, then make sure it ends up <b>numbered</b>, not bulleted.`,
        done: "Bullets and Numbering both apply to a selection the same way — flip between them any time.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Registration desk — 8am")) {
            body.insertParagraph("Setup crew — 7am", Word.InsertLocation.end);
            body.insertParagraph("Registration desk — 8am", Word.InsertLocation.end);
            body.insertParagraph("Cleanup team — 5pm", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => {
          try {
            return await Word.run(async (context) => {
              const paras = context.document.body.paragraphs;
              paras.load("items/text");
              await context.sync();
              const wanted = ["Setup crew — 7am", "Registration desk — 8am", "Cleanup team — 5pm"];
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
        }
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Founders' Day Newsletter, Layout Pass
     ====================================================================== */
  assignment: {
    hook: "The Founders' Day newsletter text is all there, but it was pasted in as one flat, undifferentiated block — nothing looks like a heading, nothing stands out, and the schedule reads like a wall of text.",
    brief: "Make the heading actually look like one, emphasise the key line, merge a split sentence, centre the invite line, fix the cramped spacing, and turn the schedule into a numbered list.",
    doneMsg: "Newsletter's ready to print. Every move from the tutorial, now shaping a document someone will actually read.",

    setup: async () => {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertParagraph("Founders' Day — Newsletter", Word.InsertLocation.end);
        body.insertParagraph("You're Invited!", Word.InsertLocation.end);
        body.insertParagraph("Doors open at 8:30am.", Word.InsertLocation.end);
        body.insertParagraph("Founders' Day begins at 9am.", Word.InsertLocation.end);
        body.insertParagraph("Limited seats available!", Word.InsertLocation.end);
        body.insertParagraph("This paragraph is crammed into the printed programme with no breathing room at all.", Word.InsertLocation.end);
        body.insertParagraph("Setup crew — 7am", Word.InsertLocation.end);
        body.insertParagraph("Registration desk — 8am", Word.InsertLocation.end);
        body.insertParagraph("Cleanup team — 5pm", Word.InsertLocation.end);
        await context.sync();
      });
    },

    tasks: [
      { ref: "2.2.1", text: "Format the heading <b>Founders' Day — Newsletter</b>: at least 20pt and a non-default font.",
        hint: "Select it, then change the font size and font name boxes on the Home tab.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = paras.items.find(x => x.text.trim() === "Founders' Day — Newsletter");
          if (!p) return false;
          p.font.load("size,name");
          await context.sync();
          return p.font.size >= 20 && !wordFormatIsDefaultFontName(p.font.name);
        }) },

      { ref: "2.2.2", text: "Make <b>Limited seats available!</b> bold and italic so it grabs attention.",
        hint: "Select the line, Ctrl/Cmd+B, Ctrl/Cmd+I.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = paras.items.find(x => x.text.trim() === "Limited seats available!");
          if (!p) return false;
          p.font.load("bold,italic");
          await context.sync();
          return p.font.bold === true && p.font.italic === true;
        }) },

      { ref: "2.2.4", text: "\"Doors open at 8:30am.\" and \"Founders' Day begins at 9am.\" should be one sentence — merge the two paragraphs.",
        hint: "Click at the very start of one of the two lines and press Backspace.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          return paras.items.some(p => p.text.includes("Doors open at 8:30am") && p.text.includes("begins at 9am"));
        }) },

      { ref: "2.2.5", text: "Centre the line <b>You're Invited!</b>",
        hint: "Click into the line, Ctrl/Cmd+E.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = paras.items.find(x => x.text.trim() === "You're Invited!");
          if (!p) return false;
          p.load("alignment");
          await context.sync();
          return p.alignment === Word.Alignment.centered;
        }) },

      { ref: "2.2.6", text: "The paragraph about the printed programme is cramped — give it double line spacing and some space after it.",
        hint: "Line and Paragraph Spacing button on the Home tab.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const p = paras.items.find(x => x.text.includes("printed programme"));
          if (!p) return false;
          p.load("lineSpacing,spaceAfter");
          await context.sync();
          return p.lineSpacing >= 22 && p.spaceAfter > 0;
        }) },

      { ref: "2.2.7", text: "Turn the three volunteer shifts into a <b>numbered</b> list.",
        hint: "Select all three lines, then click Numbering on the Home tab.",
        check: async () => {
          try {
            return await Word.run(async (context) => {
              const paras = context.document.body.paragraphs;
              paras.load("items/text");
              await context.sync();
              const wanted = ["Setup crew — 7am", "Registration desk — 8am", "Cleanup team — 5pm"];
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
        } }
    ]
  }
};
