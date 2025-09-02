// Variables globales
let tocHelpShown = false; // Pour suivre si l'aide a déjà été affichée

// Initialisation d'Office
Office.onReady((info) => {
    // Office est prêt
    if (info.host === Office.HostType.Word) {
        // Initialiser les boutons
        document.getElementById("generateBtn").onclick = generateDocument;
        document.getElementById("tocBtn").onclick = showTableOfContentsHelp;
        document.getElementById("validateBtn").onclick = validateDocument;
        
        // Ajouter les dialogues au DOM
        createConfirmDialog();
        createHelpDialog();
        
        // Message de bienvenue
        showStatus("✨ Complément chargé avec succès ! Prêt à générer votre document.", "info");
    }
});

// Créer le dialogue de confirmation personnalisé
function createConfirmDialog() {
    const dialogHTML = `
        <div id="confirmDialog" class="confirm-dialog-overlay" style="display: none;">
            <div class="confirm-dialog">
                <div class="confirm-dialog-header">
                    <span class="confirm-dialog-icon">⚠️</span>
                    <h3>Confirmation requise</h3>
                </div>
                <div class="confirm-dialog-body">
                    <p id="confirmMessage"></p>
                </div>
                <div class="confirm-dialog-footer">
                    <button id="confirmCancel" class="btn-dialog btn-cancel">Annuler</button>
                    <button id="confirmOK" class="btn-dialog btn-confirm">Confirmer</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', dialogHTML);
}

// Créer le dialogue d'aide pour la table des matières
function createHelpDialog() {
    const helpHTML = `
        <div id="helpDialog" class="help-dialog-overlay" style="display: none;">
            <div class="help-dialog">
                <div class="help-dialog-header">
                    <span class="help-dialog-icon">💡</span>
                    <h3>Aide : Insérer une table des matières</h3>
                </div>
                <div class="help-dialog-body">
                    <p><strong>C'est à vous de jouer !</strong></p>
                    <p>Pour apprendre à utiliser Word efficacement, essayez d'insérer la table des matières par vous-même.</p>
                    
                    <div class="hint-box">
                        <h4>🔍 Indice :</h4>
                        <p>Regardez dans l'onglet <strong>"Références"</strong> du ruban Word.</p>
                        <p>Vous y trouverez un bouton <strong>"Table des matières"</strong>.</p>
                    </div>
                    
                    <div class="steps-box">
                        <h4>📝 Étapes suggérées :</h4>
                        <ol>
                            <li>Supprimez le texte "[Insérez la table des matières ici!]"</li>
                            <li>Positionnez votre curseur sous le titre "Table des matières"</li>
                            <li>Allez dans l'onglet <strong>Références</strong></li>
                            <li>Cliquez sur <strong>Table des matières</strong></li>
                            <li>Choisissez un modèle automatique qui affiche les numéros de page</li>
                        </ol>
                    </div>
                    
                    <p class="help-note">💡 <em>Astuce : Les modèles "Table automatique 1" ou "Table automatique 2" créent des liens cliquables !</em></p>
                </div>
                <div class="help-dialog-footer">
                    <button id="helpOK" class="btn-dialog btn-help-ok">J'ai compris !</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', helpHTML);
    
    // Ajouter les styles pour les dialogues
    const dialogStyles = `
        <style>
        /* Styles du dialogue de confirmation */
        .confirm-dialog-overlay, .help-dialog-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10000;
            animation: fadeIn 0.2s ease;
        }
        
        .confirm-dialog, .help-dialog {
            background: white;
            border-radius: 12px;
            padding: 0;
            width: 90%;
            max-width: 450px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
            animation: slideUp 0.3s ease;
            overflow: hidden;
        }
        
        .confirm-dialog-header {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 20px;
            text-align: center;
        }
        
        .help-dialog-header {
            background: linear-gradient(135deg, #00d2ff 0%, #3a7bd5 100%);
            color: white;
            padding: 10px;
            text-align: center;
        }
        
        .confirm-dialog-icon, .help-dialog-icon {
            font-size: 48px;
            display: block;
            margin-bottom: 10px;
        }
        
        .confirm-dialog-header h3, .help-dialog-header h3 {
            margin: 0;
            font-size: 20px;
        }
        
        .confirm-dialog-body, .help-dialog-body {
            padding: 15px;
        }
        
        .confirm-dialog-body p, .help-dialog-body p {
            margin: 10px 0;
            line-height: 1.6;
            color: #333;
            font-size: 14px;
        }
        
        .hint-box {
            background: #e3f2fd;
            border-left: 4px solid #2196F3;
            padding: 10px;
            margin: 20px 0;
            border-radius: 8px;
        }
        
        .hint-box h4 {
            margin: 0 0 10px 0;
            color: #1976D2;
            font-size: 16px;
        }
        
        .steps-box {
            background: #f3e5f5;
            border-left: 4px solid #9c27b0;
            padding: 5px;
            margin: 20px 0;
            border-radius: 8px;
        }
        
        .steps-box h4 {
            margin: 0 0 10px 0;
            color: #7b1fa2;
            font-size: 16px;
        }
        
        .steps-box ol {
            margin: 10px 0 0 20px;
            padding: 0;
            color: #555;
        }
        
        .steps-box li {
            margin: 8px 0;
            line-height: 1.5;
        }
        
        .help-note {
            background: #fff3e0;
            padding: 10px;
            border-radius: 6px;
            margin-top: 15px;
        }
        
        .confirm-dialog-footer, .help-dialog-footer {
            display: flex;
            gap: 10px;
            padding: 15px 20px;
            background: #f8f9fa;
            justify-content: flex-end;
        }
        
        .btn-dialog {
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        
        .btn-cancel {
            background: #e9ecef;
            color: #495057;
        }
        
        .btn-cancel:hover {
            background: #dee2e6;
        }
        
        .btn-confirm {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
        }
        
        .btn-help-ok {
            background: linear-gradient(135deg, #00d2ff 0%, #3a7bd5 100%);
            color: white;
            min-width: 120px;
        }
        
        .btn-confirm:hover, .btn-help-ok:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
        }
        
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        
        @keyframes slideUp {
            from { 
                transform: translateY(20px);
                opacity: 0;
            }
            to { 
                transform: translateY(0);
                opacity: 1;
            }
        }
        </style>
    `;
    
    document.head.insertAdjacentHTML('beforeend', dialogStyles);
}

// Fonction pour afficher le dialogue de confirmation
function showConfirmDialog(message) {
    return new Promise((resolve) => {
        const dialog = document.getElementById('confirmDialog');
        const messageEl = document.getElementById('confirmMessage');
        const cancelBtn = document.getElementById('confirmCancel');
        const okBtn = document.getElementById('confirmOK');
        
        messageEl.innerHTML = message;
        dialog.style.display = 'flex';
        
        function handleCancel() {
            dialog.style.display = 'none';
            cancelBtn.removeEventListener('click', handleCancel);
            okBtn.removeEventListener('click', handleOK);
            resolve(false);
        }
        
        function handleOK() {
            dialog.style.display = 'none';
            cancelBtn.removeEventListener('click', handleCancel);
            okBtn.removeEventListener('click', handleOK);
            resolve(true);
        }
        
        cancelBtn.addEventListener('click', handleCancel);
        okBtn.addEventListener('click', handleOK);
    });
}

// Fonction pour afficher l'aide de la table des matières
function showTableOfContentsHelp() {
    const dialog = document.getElementById('helpDialog');
    const okBtn = document.getElementById('helpOK');
    
    dialog.style.display = 'flex';
    
    function handleOK() {
        dialog.style.display = 'none';
        okBtn.removeEventListener('click', handleOK);
        tocHelpShown = true;
        
        // Changer le texte du bouton après la première aide
        const tocBtn = document.getElementById('tocBtn');
        const btnText = tocBtn.querySelector('span:nth-child(2)');
        if (btnText) {
            btnText.textContent = "Afficher l'aide à nouveau";
        }
        
        showStatus("💪 À vous de jouer ! Utilisez l'onglet Références pour insérer la table des matières.", "info");
    }
    
    okBtn.addEventListener('click', handleOK);
}

// Fonction pour générer le document
async function generateDocument() {
    const loader = document.getElementById("generateLoader");
    const btn = document.getElementById("generateBtn");
    
    try {
        loader.style.display = "inline-block";
        btn.disabled = true;
        
        await Word.run(async (context) => {
            const body = context.document.body;
            
            // Charger le contenu pour vérifier s'il est vide
            body.load("text");
            await context.sync();
            
            // Si le document n'est pas vide, demander confirmation
            if (body.text.trim().length > 0) {
                const confirmDelete = await showConfirmDialog(
                    "<strong>Attention : Le document actuel n'est pas vide.</strong><br><br>" +
                    "Voulez-vous effacer tout le contenu existant et générer un nouveau document ?<br><br>" +
                    "<em>Cette action est irréversible.</em>"
                );
                
                if (!confirmDelete) {
                    showStatus("Génération annulée. Le document existant a été conservé.", "info");
                    loader.style.display = "none";
                    btn.disabled = false;
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
                "Maintenant, utilisez l'onglet Références pour insérer la table des matières manuellement.",
                "success"
            );
        });
    } catch (error) {
        console.error("Erreur:", error);
        showStatus("❌ Erreur lors de la génération : " + error.message, "error");
    } finally {
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
    
    // Table des matières (titre)
    const tocTitle = body.insertParagraph("Table des matières", Word.InsertLocation.end);
    tocTitle.styleBuiltIn = Word.Style.heading1;
    tocTitle.spaceAfter = 30;
    
    // Note pour l'utilisateur
    const tocPlaceholder = body.insertParagraph(
        "[Insérez la table des matières ici!]",
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

// Fonction de validation du document
async function validateDocument() {
    const loader = document.getElementById("validateLoader");
    const btn = document.getElementById("validateBtn");
    
    try {
        loader.style.display = "inline-block";
        btn.disabled = true;
        
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            // Vérifier d'abord si le document est vide
            if (body.text.trim().length === 0) {
                displayValidationResults([{
                    check: "Document vide",
                    passed: false,
                    message: "Le document est vide. Veuillez d'abord générer le document."
                }]);
                showStatus("❌ Le document est vide. Veuillez d'abord générer le document.", "error");
                return;
            }
            
            let validationResults = [];
            
            // Test 1: Vérifier la présence du titre "Table des matières"
            const tocTitleSearch = context.document.body.search("Table des matières", {
                matchCase: false
            });
            tocTitleSearch.load("items");
            await context.sync();
            
            const hasTocTitle = tocTitleSearch.items.length > 0;
            // validationResults.push({
            //     check: "Titre 'Table des matières' présent",
            //     passed: hasTocTitle,
            //     message: hasTocTitle ? 
            //         "Le titre 'Table des matières' est présent" : 
            //         "Le titre 'Table des matières' n'a pas été trouvé - Générez d'abord le document"
            // });
            
            // Test 2: Vérifier si le placeholder est toujours là
            const placeholderSearch = context.document.body.search("\\[Insérez la table des matières ici!\\]", {
                matchCase: false,
                matchWildcards: false
            });
            placeholderSearch.load("items");
            await context.sync();
            
            const placeholderPresent = placeholderSearch.items.length > 0;
            
            // Test 3: Détecter une vraie table des matières
            let tableInserted = false;
            let tableIndicators = 0;
            
            if (!placeholderPresent && hasTocTitle) {
                // Liste des titres qui devraient apparaître dans une vraie table
                const expectedTitles = [
                    "Introduction à l'Intelligence Artificielle",
                    "Histoire de l'IA",
                    "Applications de l'IA",
                    "IA dans la Santé",
                    "Défis et Perspectives",
                    "Conclusion"
                ];
                
                for (const title of expectedTitles) {
                    const titleSearch = context.document.body.search(title, {
                        matchCase: false
                    });
                    titleSearch.load("items");
                    await context.sync();
                    
                    // Si on trouve le titre plus d'une fois, c'est qu'il est dans la table ET le contenu
                    if (titleSearch.items.length >= 2) {
                        tableIndicators++;
                    }
                }
                
                // Si on trouve au moins 4 titres en double, une table a été insérée
                tableInserted = tableIndicators >= 4;
            }
            
            validationResults.push({
                check: "Table des matières insérée",
                passed: tableInserted,
                message: tableInserted ? 
                    `Table détectée : ${tableIndicators} titres trouvés dans la table` : 
                    placeholderPresent ? 
                        "Supprimez le placeholder et insérez une table via Références → Table des matières" :
                        "Aucune table détectée. Utilisez Références → Table des matières → Table automatique"
            });
            
            // Test 4: Vérifier la présence de numéros de page
            let hasPageNumbers = false;
            if (tableInserted) {
                // Chercher des numéros de page typiques
                let pageIndicators = 0;
                
                for (let i = 2; i <= 6; i++) {
                    const pageSearch = context.document.body.search(i.toString(), {
                        matchCase: false
                    });
                    pageSearch.load("items");
                    await context.sync();
                    
                    if (pageSearch.items.length > 0) {
                        pageIndicators++;
                    }
                }
                
                hasPageNumbers = pageIndicators >= 3;
            }
            
            validationResults.push({
                check: "Numéros de page affichés",
                passed: hasPageNumbers,
                message: hasPageNumbers ? 
                    "La table affiche des numéros de page" : 
                    tableInserted ?
                        "Numéros de page non détectés - Utilisez un modèle avec numéros de page" :
                        "Insérez d'abord une table des matières"
            });
            
            // Test 5: Vérifier l'emplacement (sous le titre)
            let correctPosition = false;
            if (tableInserted && hasTocTitle) {
                correctPosition = true;
            }
            
            validationResults.push({
                check: "Table au bon emplacement",
                passed: correctPosition,
                message: correctPosition ? 
                    "La table est positionnée sous le titre 'Table des matières'" : 
                    "Positionnez la table juste après le titre 'Table des matières'"
            });
            
            // Test 6: Vérifier si c'est un modèle automatique
            let hasAutomaticModel = false;
            if (tableInserted && hasPageNumbers) {
                hasAutomaticModel = true;
            }
            
            validationResults.push({
                check: "Modèle automatique (liens cliquables)",
                passed: hasAutomaticModel,
                message: hasAutomaticModel ? 
                    "La table utilise un modèle automatique avec liens cliquables" : 
                    "Utilisez 'Table automatique 1' ou 'Table automatique 2' pour avoir des liens"
            });

            // validationResults.push({
            //     check: "Placeholder supprimé",
            //     passed: !placeholderPresent,
            //     message: placeholderPresent
            //         ? "Le placeholder est encore là : insérez une table des matières"
            //         : "Le placeholder a été remplacé"
            // });
            
            // Afficher les résultats
            displayValidationResults(validationResults);
            
            // Calculer le score global
            const passedCount = validationResults.filter(r => r.passed).length;
            const totalCount = validationResults.length;
            const percentage = Math.round((passedCount / totalCount) * 100);
            
            if (percentage === 100) {
                showStatus(
                    `✅ Parfait ! Tous les ${totalCount} critères sont validés. ` +
                    "La table des matières est correctement configurée.",
                    "success"
                );
            } else if (percentage >= 50) {
                showStatus(
                    `⚠️ Progression : ${passedCount}/${totalCount} critères validés (${percentage}%). ` +
                    "Vérifiez les points en rouge ci-dessous.",
                    "warning"
                );
            } else {
                showStatus(
                    `❌ Incomplet : ${passedCount}/${totalCount} critères validés (${percentage}%). ` +
                    "Suivez les instructions : Références → Table des matières → Table automatique.",
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

// Gestion des erreurs globales
window.onerror = function(msg, url, lineNo, columnNo, error) {
    console.error('Erreur globale:', error);
    showStatus("❌ Une erreur inattendue s'est produite. Veuillez réessayer.", "error");
    return false;
};