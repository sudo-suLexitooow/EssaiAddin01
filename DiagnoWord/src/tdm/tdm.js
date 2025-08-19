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
      range.insertTableOfContents("rightAlign", true, 1, 3);

      await context.sync();
      setStatus("📑 Table des matières insérée.");
    } else {
      setStatus("⚠️ Pas de titre 'Table des matières' trouvé.");
    }
  });
}
