Office.onReady(() => {
    document.getElementById("initDoc").onclick = initializeDocument;
    document.getElementById("insertToc").onclick = insertTOC;
    document.getElementById("validate").onclick = validateTOC;
  });
  
  async function initializeDocument() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
  
      if (body.text.trim().length > 0) {
        const confirmClear = confirm(
          "Le document contient déjà du contenu. Voulez-vous l'effacer ?"
        );
        if (!confirmClear) return;
        body.clear();
      }
  
      // Page de garde
      body.insertParagraph("Page de garde", Word.InsertLocation.start);
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  
      // Titre "Table des matières"
      body.insertParagraph("Table des matières", Word.InsertLocation.end).style = "Heading 1";
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  
      // Exemple de contenu
      body.insertParagraph("Chapitre 1", Word.InsertLocation.end).style = "Heading 1";
      body.insertParagraph("Section 1.1", Word.InsertLocation.end).style = "Heading 2";
      body.insertParagraph("Sous-section 1.1.1", Word.InsertLocation.end).style = "Heading 3";
  
      body.insertParagraph("Chapitre 2", Word.InsertLocation.end).style = "Heading 1";
      body.insertParagraph("Section 2.1", Word.InsertLocation.end).style = "Heading 2";
  
      // Source
      body.insertParagraph("Source : Wikipedia", Word.InsertLocation.end).italic = true;
  
      await context.sync();
      setStatus("✅ Document initialisé.");
    });
  }
  
  async function insertTOC() {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();
  
      let tocParagraph = null;
      for (let p of paragraphs.items) {
        if (p.text.includes("Table des matières")) {
          tocParagraph = p;
          break;
        }
      }
  
      if (tocParagraph) {
        tocParagraph.insertTableOfContents("rightAlign", true, 1, 3);
        await context.sync();
        setStatus("📑 Table des matières insérée.");
      } else {
        setStatus("⚠️ Pas de titre 'Table des matières' trouvé.");
      }
    });
  }
  
  async function validateTOC() {
    await Word.run(async (context) => {
      const fields = context.document.fields;
      fields.load("items");
      await context.sync();
  
      const tocPresent = fields.items.some(f => f.code.includes("TOC"));
  
      if (tocPresent) {
        setStatus("✅ Validation réussie : Table des matières détectée.");
      } else {
        setStatus("❌ Validation échouée : Table des matières non trouvée.");
      }
    });
  }
  
  function setStatus(msg) {
    document.getElementById("status").innerText = msg;
  }
  