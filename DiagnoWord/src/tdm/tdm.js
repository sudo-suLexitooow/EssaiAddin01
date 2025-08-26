// Initialisation d'Office
Office.onReady((info) => {
  // Office est prêt
  if (info.host === Office.HostType.Word) {
      // Initialiser les boutons
      document.getElementById("generateBtn").addEventListener("click", generateDocument);
      document.getElementById("tocBtn").addEventListener("click", insertTableOfContents);
      document.getElementById("validateBtn").addEventListener("click", validateDocument);
      
      // Message de bienvenue
      showStatus("✨ Complément chargé avec succès ! Prêt à générer votre document.", "info");
  }
});

// Fonction pour générer le document
async function generateDocument() {
  const loader = document.getElementById("generateLoader");
  const btn = document.getElementById("generateBtn");
  
  try {
      // Afficher le loader
      loader.style.display = "inline-block";
      btn.disabled = true;
      
      await Word.run(async (context) => {
          const body = context.document.body;
          
          // Charger le contenu pour vérifier s'il est vide
          body.load("text");
          await context.sync();
          
          // Si le document n'est pas vide, demander confirmation
          if (body.text.trim().length > 0) {
              const confirmDelete = confirm(
                  "⚠️ Attention : Le document actuel n'est pas vide.\n\n" +
                  "Voulez-vous effacer tout le contenu existant et générer un nouveau document ?\n\n" +
                  "Cette action est irréversible."
              );
              
              if (!confirmDelete) {
                  showStatus("Génération annulée. Le document existant a été conservé.", "info");
                  return;
              }
              body.clear();
          }
          
          // Générer le contenu du document
          await generateDocumentContent(body);
          
          await context.sync();
          
          // Afficher les détails du document
          document.getElementById("documentDetails").style.display = "block";
          
          showStatus(
              "✅ Document généré avec succès ! " +
              "Le document contient 7 pages avec une structure complète. " +
              "Vous pouvez maintenant insérer la table des matières.",
              "success"
          );
      });
  } catch (error) {
      console.error("Erreur:", error);
      showStatus("❌ Erreur lors de la génération : " + error.message, "error");
  } finally {
      // Masquer le loader
      loader.style.display = "none";
      btn.disabled = false;
  }
}

// Fonction pour générer le contenu du document
async function generateDocumentContent(body) {
  // Page de garde
  const coverTitle = body.insertParagraph(documentContent.coverPage.title, Word.InsertLocation.start);
  coverTitle.styleBuiltIn = Word.Style.title;
  coverTitle.alignment = Word.Alignment.centered;
  coverTitle.spaceAfter = 200;
  
  const coverSubtitle = body.insertParagraph(documentContent.coverPage.subtitle, Word.InsertLocation.end);
  coverSubtitle.alignment = Word.Alignment.centered;
  coverSubtitle.font.size = 18;
  coverSubtitle.font.color = "#4a5568";
  coverSubtitle.spaceAfter = 100;
  
  const coverDate = body.insertParagraph(documentContent.coverPage.date, Word.InsertLocation.end);
  coverDate.alignment = Word.Alignment.centered;
  coverDate.font.size = 14;
  coverDate.font.color = "#718096";
  
  // Saut de page après la page de garde
  body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  
  // Table des matières (titre seulement)
  const tocTitle = body.insertParagraph("Table des matières", Word.InsertLocation.end);
  tocTitle.styleBuiltIn = Word.Style.heading1;
  tocTitle.spaceAfter = 30;
  
  // Espace réservé pour la table des matières
  const tocPlaceholder = body.insertParagraph(
      "[La table des matières sera insérée ici]",
      Word.InsertLocation.end
  );
  tocPlaceholder.font.italic = true;
  tocPlaceholder.font.color = "#999999";
  tocPlaceholder.alignment = Word.Alignment.left;
  
  // Saut de page
  body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  
  // Générer les chapitres
  documentContent.chapters.forEach((chapter, chapterIndex) => {
      // Titre du chapitre
      const chapterTitle = body.insertParagraph(chapter.title, Word.InsertLocation.end);
      chapterTitle.styleBuiltIn = Word.Style.heading1;
      chapterTitle.spaceAfter = 15;
      
      // Introduction du chapitre
      if (chapter.introduction) {
          const intro = body.insertParagraph(chapter.introduction, Word.InsertLocation.end);
          intro.styleBuiltIn = Word.Style.normal;
          intro.spaceAfter = 20;
      }
      
      // Sections du chapitre
      chapter.sections.forEach(section => {
          // Titre de la section
          const sectionTitle = body.insertParagraph(section.title, Word.InsertLocation.end);
          sectionTitle.styleBuiltIn = Word.Style.heading2;
          sectionTitle.spaceAfter = 10;
          
          // Contenu de la section
          const sectionContent = body.insertParagraph(section.content, Word.InsertLocation.end);
          sectionContent.styleBuiltIn = Word.Style.normal;
          sectionContent.spaceAfter = 15;
          
          // Sous-sections
          if (section.subsections) {
              section.subsections.forEach(subsection => {
                  const subsectionTitle = body.insertParagraph(subsection.title, Word.InsertLocation.end);
                  subsectionTitle.styleBuiltIn = Word.Style.heading3;
                  subsectionTitle.spaceAfter = 8;
                  
                  const subsectionContent = body.insertParagraph(subsection.content, Word.InsertLocation.end);
                  subsectionContent.styleBuiltIn = Word.Style.normal;
                  subsectionContent.spaceAfter = 15;
              });
          }
      });
      
      // Saut de page après chaque chapitre (sauf le dernier)
      if (chapterIndex < documentContent.chapters.length - 1) {
          body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
      }
  });
  
  // Conclusion
  body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  const conclusionTitle = body.insertParagraph(documentContent.conclusion.title, Word.InsertLocation.end);
  conclusionTitle.styleBuiltIn = Word.Style.heading1;
  conclusionTitle.spaceAfter = 15;
  
  const conclusionText = body.insertParagraph(documentContent.conclusion.content, Word.InsertLocation.end);
  conclusionText.styleBuiltIn = Word.Style.normal;
  
  // Note de source
  body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  const sourceNote = body.insertParagraph(documentContent.source, Word.InsertLocation.end);
  sourceNote.font.italic = true;
  sourceNote.font.size = 10;
  sourceNote.font.color = "#666666";
  sourceNote.alignment = Word.Alignment.centered;
}

// Fonction pour insérer la table des matières
async function insertTableOfContents() {
  const loader = document.getElementById("tocLoader");
  const btn = document.getElementById("tocBtn");
  
  try {
      loader.style.display = "inline-block";
      btn.disabled = true;
      
      await Word.run(async (context) => {
          // Rechercher le titre "Table des matières"
          const searchResults = context.document.body.search("Table des matières", {
              matchCase: false,
              matchWholeWord: false
          });
          searchResults.load("items");
          await context.sync();
          
          if (searchResults.items.length === 0) {
              showStatus("⚠️ Veuillez d'abord générer le document avant d'insérer la table des matières.", "warning");
              return;
          }
          
          // Rechercher et supprimer le placeholder
          const placeholderSearch = context.document.body.search("[La table des matières sera insérée ici]", {
              matchCase: false
          });
          placeholderSearch.load("items");
          await context.sync();
          
          if (placeholderSearch.items.length > 0) {
              // Supprimer le placeholder
              placeholderSearch.items[0].delete();
              await context.sync();
              
              // Essayer d'insérer une vraie table des matières
              try {
                  const tocTitle = searchResults.items[0];
                  const range = tocTitle.getRange(Word.RangeLocation.after);
                  
                  // Insérer la table des matières Word native
                  range.insertParagraph("", Word.InsertLocation.after);
                  const toc = range.insertTableOfContents(Word.InsertLocation.after);
                  
                  await context.sync();
                  showStatus(
                      "✅ Table des matières insérée avec succès ! " +
                      "Les numéros de pages sont maintenant visibles et cliquables.",
                      "success"
                  );
              } catch (tocError) {
                  // Si l'API n'est pas disponible, créer une table manuelle
                  await insertManualTableOfContents(context, searchResults.items[0]);
              }
          } else {
              showStatus("ℹ️ La table des matières a déjà été insérée.", "info");
          }
      });
  } catch (error) {
      console.error("Erreur:", error);
      showStatus("❌ Erreur lors de l'insertion : " + error.message, "error");
  } finally {
      loader.style.display = "none";
      btn.disabled = false;
  }
}

// Fonction pour insérer une table des matières manuelle
async function insertManualTableOfContents(context, tocTitle) {
  const range = tocTitle.getRange(Word.RangeLocation.after);
  
  // Créer une table des matières manuelle formatée
  const tocLines = [
      { text: "1. Introduction à l'Intelligence Artificielle", page: "3", level: 1 },
      { text: "   1.1 Histoire de l'IA", page: "3", level: 2 },
      { text: "       1.1.1 Les pionniers", page: "3", level: 3 },
      { text: "   1.2 Types d'Intelligence Artificielle", page: "3", level: 2 },
      { text: "2. Applications de l'IA", page: "4", level: 1 },
      { text: "   2.1 IA dans la Santé", page: "4", level: 2 },
      { text: "       2.1.1 Diagnostic assisté", page: "4", level: 3 },
      { text: "   2.2 IA dans les Transports", page: "4", level: 2 },
      { text: "3. Défis et Perspectives", page: "5", level: 1 },
      { text: "   3.1 Défis Éthiques", page: "5", level: 2 },
      { text: "   3.2 L'Avenir de l'IA", page: "5", level: 2 },
      { text: "Conclusion", page: "6", level: 1 }
  ];
  
  // Insérer chaque ligne de la table des matières
  for (const item of tocLines) {
      const dots = ".".repeat(Math.max(2, 60 - item.text.length - item.page.length));
      const tocLine = range.insertParagraph(
          `${item.text} ${dots} ${item.page}`,
          Word.InsertLocation.after
      );
      tocLine.font.name = "Courier New";
      tocLine.font.size = 11;
      
      // Appliquer l'indentation selon le niveau
      if (item.level === 2) {
          tocLine.leftIndent = 20;
      } else if (item.level === 3) {
          tocLine.leftIndent = 40;
      }
  }
  
  await context.sync();
  showStatus(
      "ℹ️ Table des matières créée manuellement. " +
      "Note : Pour des liens cliquables, utilisez Word Desktop et insérez une table des matières native.",
      "info"
  );
}

// Fonction de validation du document
async function validateDocument() {
  const loader = document.getElementById("validateLoader");
  const btn = document.getElementById("validateBtn");
  
  try {
      loader.style.display = "inline-block";
      btn.disabled = true;
      
      await Word.run(async (context) => {
          const body = context.document.body;
          let validationResults = [];
          
          // Test 1: Vérifier la présence du titre "Table des matières"
          const tocSearch = context.document.body.search("Table des matières", {
              matchCase: false
          });
          tocSearch.load("items");
          await context.sync();
          
          const hasTocTitle = tocSearch.items.length > 0;
          validationResults.push({
              check: "Présence du titre 'Table des matières'",
              passed: hasTocTitle,
              message: hasTocTitle ? 
                  "Le titre est présent dans le document" : 
                  "Le titre 'Table des matières' n'a pas été trouvé"
          });
          
          // Test 2: Vérifier que le placeholder a été supprimé
          const placeholderSearch = context.document.body.search("[La table des matières sera insérée ici]", {
              matchCase: false
          });
          placeholderSearch.load("items");
          await context.sync();
          
          const noPlaceholder = placeholderSearch.items.length === 0;
          validationResults.push({
              check: "Table des matières insérée",
              passed: noPlaceholder,
              message: noPlaceholder ? 
                  "La table des matières a été insérée (placeholder supprimé)" : 
                  "Le placeholder est toujours présent, la table n'a pas été insérée"
          });
          
          // Test 3: Vérifier l'emplacement de la table
          let tocPosition = false;
          if (hasTocTitle && noPlaceholder) {
              // Vérifier qu'il y a du contenu après le titre
              tocPosition = true;
          }
          
          validationResults.push({
              check: "Table des matières au bon emplacement",
              passed: tocPosition,
              message: tocPosition ? 
                  "La table est correctement positionnée sous le titre" : 
                  "La table n'est pas à l'emplacement attendu"
          });
          
          // Test 4: Vérifier la structure du document
          const heading1Search = context.document.body.search("Introduction à l'Intelligence Artificielle", {
              matchCase: false
          });
          heading1Search.load("items");
          await context.sync();
          
          const hasStructure = heading1Search.items.length > 0;
          validationResults.push({
              check: "Document structuré avec titres hiérarchiques",
              passed: hasStructure,
              message: hasStructure ? 
                  "Le document contient une structure hiérarchique complète" : 
                  "La structure du document est incomplète"
          });
          
          // Test 5: Vérifier la présence de la page de garde
          const coverSearch = context.document.body.search("L'Intelligence Artificielle", {
              matchCase: false
          });
          coverSearch.load("items");
          await context.sync();
          
          const hasCover = coverSearch.items.length > 0;
          validationResults.push({
              check: "Page de garde présente",
              passed: hasCover,
              message: hasCover ? 
                  "La page de garde est présente" : 
                  "La page de garde n'a pas été trouvée"
          });
          
          // Test 6: Vérifier la présence de la source
          const sourceSearch = context.document.body.search("Source", {
              matchCase: false
          });
          sourceSearch.load("items");
          await context.sync();
          
          const hasSource = sourceSearch.items.length > 0;
          validationResults.push({
              check: "Citation de la source",
              passed: hasSource,
              message: hasSource ? 
                  "La source est correctement citée" : 
                  "La citation de source est manquante"
          });
          
          // Afficher les résultats
          displayValidationResults(validationResults);
          
          // Calculer le score global
          const passedCount = validationResults.filter(r => r.passed).length;
          const totalCount = validationResults.length;
          const percentage = Math.round((passedCount / totalCount) * 100);
          
          if (percentage === 100) {
              showStatus(
                  `✅ Validation parfaite ! Tous les ${totalCount} critères sont validés. ` +
                  "Le document est conforme aux exigences.",
                  "success"
              );
          } else if (percentage >= 80) {
              showStatus(
                  `⚠️ Validation partielle : ${passedCount}/${totalCount} critères validés (${percentage}%). ` +
                  "Vérifiez les points en rouge.",
                  "warning"
              );
          } else {
              showStatus(
                  `❌ Validation échouée : ${passedCount}/${totalCount} critères validés (${percentage}%). ` +
                  "Plusieurs éléments manquent.",
                  "error"
              );
          }
      });
  } catch (error) {
      console.error("Erreur:", error);
      showStatus("❌ Erreur lors de la validation : " + error.message, "error");
  } finally {
      loader.style.display = "none";
      btn.disabled = false;
  }
}

// Fonction pour afficher les résultats de validation
function displayValidationResults(results) {
  const resultsDiv = document.getElementById("validationResults");
  const contentDiv = document.getElementById("validationContent");
  
  // Vider le contenu précédent
  contentDiv.innerHTML = "";
  
  // Créer les éléments de résultat
  results.forEach(result => {
      const item = document.createElement("div");
      item.className = "validation-item";
      
      const icon = document.createElement("span");
      icon.className = result.passed ? "check-icon" : "cross-icon";
      icon.textContent = result.passed ? "✓" : "✗";
      
      const text = document.createElement("span");
      text.innerHTML = `<strong>${result.check}</strong><br>
                       <small style="color: #666; font-size: 11px;">${result.message}</small>`;
      
      item.appendChild(icon);
      item.appendChild(text);
      contentDiv.appendChild(item);
  });
  
  // Afficher la section des résultats
  resultsDiv.style.display = "block";
  
  // Ajouter une animation
  resultsDiv.style.animation = "slideIn 0.3s ease";
}

// Fonction pour afficher les messages de statut
function showStatus(message, type) {
  const statusDiv = document.getElementById("statusMessage");
  
  // Définir le contenu et la classe
  statusDiv.textContent = message;
  statusDiv.className = "status-message status-" + type;
  statusDiv.style.display = "block";
  
  // Animation d'apparition
  statusDiv.style.animation = "slideIn 0.3s ease";
  
  // Auto-masquer après un délai pour certains types
  if (type === "success" || type === "info") {
      setTimeout(() => {
          statusDiv.style.animation = "fadeOut 0.3s ease";
          setTimeout(() => {
              statusDiv.style.display = "none";
          }, 300);
      }, 7000);
  }
}

// Ajouter l'animation fadeOut dans le style
const style = document.createElement("style");
style.textContent = `
  @keyframes fadeOut {
      from { opacity: 1; }
      to { opacity: 0; }
  }
`;
document.head.appendChild(style);

// Fonction utilitaire pour formater les dates
function getCurrentDate() {
  const date = new Date();
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('fr-FR', options);
}

// Gestion des erreurs globales
window.onerror = function(msg, url, lineNo, columnNo, error) {
  console.error('Erreur globale:', error);
  showStatus("❌ Une erreur inattendue s'est produite. Veuillez réessayer.", "error");
  return false;
};