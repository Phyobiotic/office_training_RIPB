// lessons/word-edit.js
// Syllabus refs 2.1.1-2.1.7 (enter/select/edit/delete text, undo, find,
// replace, copy text). First Word lesson — same stepper/checklist shapes
// as the Excel lessons, but every check reads Word.run instead of Excel.run.
//
// HONEST CAVEAT: Word's live-checking relies partly on a poll (see
// taskpane.html) because document.onParagraphChanged tracks paragraph
// CONTENT changes, with no confirmation it also fires for formatting-only
// edits. If a formatting-only step (bold/italic) feels less "instant" than
// the Excel lessons did, that's expected, not a bug — the poll interval is
// what's catching it.

window.LESSON_MODULES = window.LESSON_MODULES || {};

window.LESSON_MODULES["word-edit"] = {
  title: "Writing and Editing Text",

  /* ======================================================================
     TUTORIAL — one mechanic per step
     ====================================================================== */
  tutorial: {
    hook: "The everyday moves for working with text in a real document. Each step loads its own short passage.",

    steps: [
      {
        title: "Bold and italic together",
        teach: `Select text, then <span class="keys">Ctrl/Cmd+B</span> and <span class="keys">Ctrl/Cmd+I</span> — they toggle independently, so you can stack both on the same selection.<br><br>Select the whole sentence below and make it <b>both bold and italic</b>.`,
        done: "Bold and italic stack — select once, apply both.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.clear();
          body.insertParagraph("Founders' Day needs your help.", Word.InsertLocation.end);
          await context.sync();
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items");
          await context.sync();
          const p = paras.items[0];
          p.font.load("bold,italic");
          await context.sync();
          return p.font.bold === true && p.font.italic === true;
        })
      },

      {
        title: "Fix a typo",
        teach: `Click right into the word, delete the wrong bit, type the fix. No different from editing anywhere else — just click and go.<br><br>Below, <code>you're support</code> should read <code>your support</code>. Fix it.`,
        done: "Click into the word, fix it, done — editing text is just that.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.clear();
          body.insertParagraph("We would love you're support this year.", Word.InsertLocation.end);
          await context.sync();
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("you're") && body.text.includes("your support");
        })
      },

      {
        title: "Undo an edit",
        teach: `<span class="keys">Ctrl/Cmd+Z</span> undoes your last action — and you can press it repeatedly to step back further.<br><br>Below is a finished sentence. Type an extra sentence after it, then undo it with Ctrl/Cmd+Z. The paragraph should end up exactly as it started.`,
        done: "Ctrl/Cmd+Z steps backward — press it again to go back further still.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.clear();
          body.insertParagraph("Thank you for your generous support.", Word.InsertLocation.end);
          await context.sync();
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return body.text.trim() === "Thank you for your generous support.";
        })
      },

      {
        title: "Find & Replace",
        teach: `<span class="keys">Ctrl/Cmd+H</span> opens Find &amp; Replace. Type what's wrong, what it should be, and Replace All fixes every instance in one pass — instead of hunting through the document yourself.<br><br>Below, <code>recieve</code> is misspelled. Use Find &amp; Replace to fix it to <code>receive</code>.`,
        done: "Ctrl/Cmd+H — Find & Replace fixes every instance at once.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.clear();
          body.insertParagraph("Please recieve our thanks for your support.", Word.InsertLocation.end);
          await context.sync();
        }),
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("recieve") && body.text.includes("receive");
        })
      },

      {
        title: "Copy text to another spot",
        teach: `Select text, <span class="keys">Ctrl/Cmd+C</span> to copy, click where it should go, <span class="keys">Ctrl/Cmd+V</span> to paste. The original stays put — copy duplicates, it doesn't move.<br><br>Copy the event name from the first line and paste it into the empty third line below, as a closing signature.`,
        done: "Copy duplicates and leaves the original in place — that's the difference from cut.",
        setup: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.clear();
          body.insertParagraph("Founders' Day 2026", Word.InsertLocation.end);
          body.insertParagraph("Dear Sponsor,", Word.InsertLocation.end);
          body.insertParagraph("", Word.InsertLocation.end);
          await context.sync();
        }),
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          return paras.items.length >= 3 && paras.items[2].text.trim() === "Founders' Day 2026";
        })
      }
    ]
  },

  /* ======================================================================
     ASSIGNMENT — Sponsor Letter, First Draft Cleanup
     ====================================================================== */
  assignment: {
    hook: "The Sponsor Letter draft went out to the Board group chat at midnight — wrong year in three places, a whole paragraph pasted twice, a typo, and no closing line.",
    brief: "Fix the wrong year everywhere, remove the duplicated paragraph, fix the spelling mistake, and copy the event name in as a closing line.",
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
        body.insertParagraph("", Word.InsertLocation.end);
        await context.sync();
      });
    },

    tasks: [
      { ref: "2.1.6 / 2.1.7", text: "Use Find &amp; Replace to fix the spelling mistake — <code>recieve</code> should read <code>receive</code> everywhere it appears.",
        hint: "Ctrl/Cmd+H, find \"recieve\", replace with \"receive\", Replace All.",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("recieve");
        }) },

      { ref: "2.1.6 / 2.1.7", text: "This letter is meant for the <b>2026</b> event, but \"2025\" slipped in twice — fix both with Find &amp; Replace.",
        hint: "Ctrl/Cmd+H, find \"2025\", replace with \"2026\", Replace All.",
        check: async () => Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          return !body.text.includes("2025");
        }) },

      { ref: "2.1.3 / 2.1.4", text: "One whole paragraph got pasted in twice by mistake — delete the duplicate so it appears only once.",
        hint: "Select the entire duplicate paragraph (click at its start, Shift+click at its end, or triple-click it) and delete it.",
        check: async () => Word.run(async (context) => {
          const paras = context.document.body.paragraphs;
          paras.load("items/text");
          await context.sync();
          const count = paras.items.filter(p => p.text.trim().startsWith("We would be delighted")).length;
          return count === 1;
        }) },

      { ref: "2.1.5", text: "The letter has no closing line. Copy the event name from the very first line and paste it as the last line.",
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