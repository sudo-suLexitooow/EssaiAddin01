function setStatus(message) {
  const statusDiv = document.getElementById("status");
  statusDiv.textContent = message;
}

// ========================
// 1. Initialisation
// ========================
async function initializeDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      // Vérifier si le document contient déjà du texte
      if (body.text.trim().length > 0) {
        const confirmClear = confirm("Le document contient déjà du texte. Voulez-vous l'effacer ?");
        if (!confirmClear) {
          setStatus("⚠️ Initialisation annulée.");
          return;
        }
        body.clear();
      }

      // ==============================
      // Page de garde
      // ==============================
      const title = body.insertParagraph("Mon Rapport de Projet", Word.InsertLocation.start);
      title.alignment = "Center";
      title.style = "Title";

      const subtitle = body.insertParagraph("Analyse et Résultats", Word.InsertLocation.after);
      subtitle.alignment = "Center";
      subtitle.style = "Subtitle";

      const author = body.insertParagraph("Auteur : Votre Nom", Word.InsertLocation.after);
      author.alignment = "Center";
      author.style = "Normal";

      const date = body.insertParagraph("Date : " + new Date().toLocaleDateString(), Word.InsertLocation.after);
      date.alignment = "Center";
      date.style = "Normal";

      await context.sync(); // appliquer les styles

      // Saut de page après la page de garde
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

      // ==============================
      // Titre Table des matières
      // ==============================
      const tocTitle = body.insertParagraph("Table des matières", Word.InsertLocation.end);
      tocTitle.style = "Heading 1";

      body.insertParagraph("", Word.InsertLocation.end);

      // ==============================
      // Contenu structuré
      // ==============================
      const intro = body.insertParagraph("Introduction", Word.InsertLocation.end);
      intro.style = "Heading 1";

      const contexte = body.insertParagraph("Contexte", Word.InsertLocation.end);
      contexte.style = "Heading 2";

      const historique = body.insertParagraph("Historique", Word.InsertLocation.end);
      historique.style = "Heading 3";

      const objectifs = body.insertParagraph("Objectifs", Word.InsertLocation.end);
      objectifs.style = "Heading 2";

      const methode = body.insertParagraph("Méthodologie", Word.InsertLocation.end);
      methode.style = "Heading 1";

      const analyse = body.insertParagraph("Analyse", Word.InsertLocation.end);
      analyse.style = "Heading 2";

      const resultats = body.insertParagraph("Résultats", Word.InsertLocation.end);
      resultats.style = "Heading 3";

      const conclusion = body.insertParagraph("Conclusion", Word.InsertLocation.end);
      conclusion.style = "Heading 1";

      await context.sync(); // appliquer les styles finaux

      setStatus("📄 Document initialisé avec une vraie page de garde.");
    });
  } catch (error) {
    console.error("Erreur lors de l'initialisation :", error);
    setStatus("❌ Erreur pendant l'initialisation du document.");
  }
}

// ========================
// 2. Insertion de la TOC
// ========================
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
      // Insérer la TOC juste après le paragraphe "Table des matières"
      const range = tocParagraph.getRange(Word.RangeLocation.after);
      range.insertTableOfContents(false, true, 1, 3); // (hidePageNumbers, rightAlign, minHeading, maxHeading)

      await context.sync();
      setStatus("📑 Table des matières insérée.");
    } else {
      setStatus("⚠️ Pas de titre 'Table des matières' trouvé.");
    }
  });
}

// ========================
// 3. Validation de la TOC
// ========================
async function validateTOC() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    const tocs = body.tablesOfContents;
    tocs.load("items");
    await context.sync();

    // Vérifier la présence du titre "Table des matières"
    let tocParagraph = null;
    for (let p of paragraphs.items) {
      if (p.text.includes("Table des matières")) {
        tocParagraph = p;
        break;
      }
    }

    if (!tocParagraph) {
      setStatus("❌ Le titre 'Table des matières' est manquant.");
      return;
    }

    if (tocs.items.length === 0) {
      setStatus("❌ Aucune table des matières trouvée.");
      return;
    }

    // Vérifier que la TOC est après le titre
    const tocRange = tocs.items[0].getRange();
    const tocText = tocRange.text;

    const afterTitleRange = tocParagraph.getRange(Word.RangeLocation.after);
    afterTitleRange.load("text");
    await context.sync();

    if (afterTitleRange.text.includes(tocText.trim().split("\n")[0])) {
      setStatus("✅ La table des matières est correctement insérée.");
    } else {
      setStatus("❌ La table des matières n'est pas au bon emplacement.");
    }
  });
}

// ========================
// 4. Liaisons des boutons
// ========================
window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("initDoc").addEventListener("click", initializeDocument);
  document.getElementById("insertToc").addEventListener("click", insertTOC);
  document.getElementById("validate").addEventListener("click", validateTOC);
});
