// lessons/word-edit.js
// Syllabus refs 2.1.2-2.1.7 (2.1.1 "create a new document from a template"
// is skipped — the add-in loads into a document that's already open, so
// there's nothing in that action for a live check to observe).
//
// REWRITE NOTE: the original version of this file mixed in a bold+italic
// step that actually belongs to 2.2.2 (Formatting and Alignment) — that's
// been moved to word-format.js, and this file now sticks to genuine 2.1.x
// "Getting Started" content: entering text, symbols, selecting at different
// granularities, over-typing, copying, and deleting.
//
// CONTINUITY: a Word document is linear, so steps can't get their own
// non-overlapping "range" the way Excel cells can. Instead, `prepare()`
// clears the body once on entry, and each step's setup() appends its own
// paragraph(s) to the END of the document — only if a stable marker
// substring for that step isn't already present — so earlier steps'
// paragraphs stay put above. Checks locate their paragraph(s) by searching
// for that same marker text rather than assuming a fixed paragraph index,
// since later steps keep growing the document underneath them.

window.LESSON_MODULES = window.LESSON_MODULES || {};

async function wordEditFindParaIndex(paras, matchFn){
  return paras.items.findIndex(p => matchFn((p.text || "").trim()));
}

window.LESSON_MODULES["word-edit"] = {
  title: "Writing and Editing Text",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The everyday moves for working with text in a real document. It's one growing document the whole way through — each step adds its own passage below what's already there.",

    prepare: async () => Word.run(async (context) => {
      context.document.body.clear();
      await context.sync();
    }),

    steps: [
      {
        title: "Enter text into a document",
        teach: `Click where you want the text to start, then just type — Word inserts it right at the cursor.<br><br>Below is the heading "<i>Founders' Day — Volunteer Briefing</i>". Click at the end of it, press Enter for a new line, and type <code>Volunteers should arrive by 8am.</code>`,
        done: "Click, type — text lands wherever your cursor is.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Volunteer Briefing")) {
            body.insertParagraph("Founders' Day — Volunteer Briefing", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return body.text.includes("Volunteers should arrive by 8am");
        })
      },

      {
        title: "Insert a symbol",
        teach: `Special characters like <b>©</b>, <b>®</b>, <b>™</b> aren't on the keyboard — use <b>Insert tab → Symbol</b> to place one, or let AutoCorrect do it: typing <code>(c)</code> becomes © automatically as soon as you type the next character.<br><br>Below, the line "<code>Sponsored content</code>" needs a © right after the word "content". Add it either way.`,
        done: "Insert → Symbol places it directly, or AutoCorrect turns (c) into © for you.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Sponsored content")) {
            body.insertParagraph("Sponsored content", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const idx = await wordEditFindParaIndex(paras, t => t.startsWith("Sponsored content"));
          return idx !== -1 && paras.items[idx].text.includes("©");
        })
      },

      {
        title: "Select a whole paragraph in one move",
        teach: `Click in the blank margin just to the <b>left</b> of a paragraph, or triple-click anywhere inside it — either grabs the <i>entire</i> paragraph in one move, no dragging needed.<br><br>Below is a full instruction paragraph. Select the whole thing in one move and make it <b>bold</b>.`,
        done: "Margin-click or triple-click selects the whole paragraph at once.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("closed-toe shoes")) {
            body.insertParagraph("Please remember to bring your own water bottle and wear closed-toe shoes.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const idx = await wordEditFindParaIndex(paras, t => t.includes("closed-toe shoes"));
          if (idx === -1) return false;
          const p = paras.items[idx];
          p.font.load("bold");
          await context.sync();
          return p.font.bold === true;
        })
      },

      {
        title: "Over-type to replace text",
        teach: `You don't have to backspace a mistake away first — <b>select</b> the wrong word and just type the replacement straight over it. The selection disappears and your new text takes its place.<br><br>Below, <code>you're support</code> should read <code>your support</code>. Select <code>you're</code> and type <code>your</code> directly over it.`,
        done: "Select the wrong bit, type over it — no backspacing required.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("support this year")) {
            body.insertParagraph("We would love you're support this year.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("you're") && body.text.includes("your support");
        })
      },

      {
        title: "Copy text to another spot",
        teach: `Select text, <span class="keys">Ctrl/Cmd+C</span> to copy, click where it should go, <span class="keys">Ctrl/Cmd+V</span> to paste. The original stays exactly where it was — copy duplicates, it doesn't move.<br><br>Below, copy the event name from the "<code>Founders' Day 2026</code>" line and paste it into the empty line just after "<code>Dear Sponsor,</code>", as a closing signature.`,
        done: "Copy duplicates and leaves the original in place — that's the difference from cut.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("Dear Sponsor,")) {
            body.insertParagraph("Founders' Day 2026", Word.InsertLocation.end);
            body.insertParagraph("Dear Sponsor,", Word.InsertLocation.end);
            body.insertParagraph("", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const dearIdx = await wordEditFindParaIndex(paras, t => t === "Dear Sponsor,");
          if (dearIdx === -1) return false;
          return paras.items.slice(dearIdx + 1).some(p => (p.text || "").trim() === "Founders' Day 2026");
        })
      },

      {
        title: "Delete a whole sentence",
        teach: `To remove a sentence, not just a word: click anywhere inside it while holding <span class="keys">Ctrl</span> — Word selects the entire sentence — then press <span class="keys">Delete</span>. (Dragging from its first letter to its last works too.)<br><br>Below, a stray sentence doesn't belong: "<i>This sentence should not be here.</i>" Select just that sentence and delete it, keeping the one before it intact.`,
        done: "Ctrl+click inside a sentence selects the whole thing — Delete removes just that.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          if (!body.text.includes("generous support")) {
            body.insertParagraph("Thank you for your generous support. This sentence should not be here.", Word.InsertLocation.end);
            await context.sync();
          }
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const idx = await wordEditFindParaIndex(paras, t => t.includes("generous support"));
          if (idx === -1) return false;
          const t = paras.items[idx].text;
          return t.includes("generous support") && !t.includes("should not be here");
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Sponsor Letter, First Draft Cleanup
     ====================================================================== */
  assignment: {
    hook: "The Sponsor Letter draft went out to the Board group chat at midnight — wrong year in three places, a whole paragraph pasted twice, a typo, a missing © notice, and no closing line.",
    brief: "Fix the wrong year everywhere, remove the duplicated paragraph, fix the spelling mistake, add the missing copyright symbol, and copy the event name in as a closing line.",
    doneMsg: "Letter's clean and ready for the Advisor to review. Same editing moves as the tutorial, just untangling a real mess this time.",

    setup: async () => {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        body.insertParagraph("Founders' Day 2026 — Sponsorship Drive", Word.InsertLocation.end);
        body.insertParagraph("Dear Community Partner,", Word.InsertLocation.end);
        body.insertParagraph("Founders' Day is our school's biggest fundraiser of the year, and this year it falls on 12 Aug 2026.", Word.InsertLocation.end);
        body.insertParagraph("We would be delighted if your organisation could recieve one of our sponsorship tiers for the 2025 event.", Word.InsertLocation.end);
        body.insertParagraph("We would be delighted if your organisation could recieve one of our sponsorship tiers for the 2025 event.", Word.InsertLocation.end);
        body.insertParagraph("Kindly confirm your participation by 30 Jul 2025.", Word.InsertLocation.end);
        body.insertParagraph("RIPB 2026", Word.InsertLocation.end);
        body.insertParagraph("", Word.InsertLocation.end);
        await context.sync();
      });
    },

    tasks: [
      { ref: "2.1.5", text: "Use Find &amp; Replace to fix the spelling mistake — <code>recieve</code> should read <code>receive</code> everywhere it appears.",
        hint: "Ctrl/Cmd+H, find \"recieve\", replace with \"receive\", Replace All.",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("recieve");
        }) },

      { ref: "2.1.5", text: "This letter is meant for the <b>2026</b> event, but \"2025\" slipped in twice — fix both with Find &amp; Replace.",
        hint: "Ctrl/Cmd+H, find \"2025\", replace with \"2026\", Replace All.",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("2025");
        }) },

      { ref: "2.1.4 / 2.1.7", text: "One whole paragraph got pasted in twice by mistake — select and delete the duplicate so it appears only once.",
        hint: "Triple-click (or margin-click) the entire duplicate paragraph and delete it.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const count = paras.items.filter(p => p.text.trim().startsWith("We would be delighted")).length;
          return count === 1;
        }) },

      { ref: "2.1.3", text: "The sign-off line reads <code>RIPB Prefects Board 2026</code> — it should carry a © right after \"Prefects Board\".",
        hint: "Click right after \"Board\", then Insert → Symbol → ©, or type (c).",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          return paras.items.some(p => /prefects board/i.test(p.text) && p.text.includes("©"));
        }) },

      { ref: "2.1.6", text: "The letter has no closing line. Copy the event name from the very first line and paste it as the last line.",
        hint: "Select \"Founders' Day 2026 — Sponsorship Drive\" from the top, copy it, click the blank last line, paste.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const matches = paras.items.filter(p => p.text.trim() === "Founders' Day 2026 — Sponsorship Drive");
          return matches.length >= 2;
        }) }
    ]
  }
};
