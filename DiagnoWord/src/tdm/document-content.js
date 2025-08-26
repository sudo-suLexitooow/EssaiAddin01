// Contenu structuré du document sur l'Intelligence Artificielle
const documentContent = {
    // Page de garde
    coverPage: {
        title: "L'Intelligence Artificielle",
        subtitle: "Une révolution technologique du XXIe siècle",
        date: "2024"
    },
    
    // Chapitres du document
    chapters: [
        {
            title: "1. Introduction à l'Intelligence Artificielle",
            introduction: "L'intelligence artificielle (IA) représente l'une des avancées technologiques les plus significatives de notre époque. Elle désigne l'ensemble des théories et techniques mises en œuvre en vue de réaliser des machines capables de simuler l'intelligence humaine. Cette technologie transforme progressivement tous les aspects de notre société, de la médecine à l'industrie, en passant par l'éducation et les services. L'IA ne se limite plus à la science-fiction ; elle est devenue une réalité tangible qui influence notre quotidien de manière croissante.",
            sections: [
                {
                    title: "1.1 Histoire de l'IA",
                    content: "L'histoire de l'intelligence artificielle remonte aux années 1950, marquée par les travaux visionnaires d'Alan Turing. Son célèbre test de Turing, proposé en 1950, pose la question fondamentale : une machine peut-elle penser ? Cette interrogation a lancé des décennies de recherche intense. La conférence de Dartmouth en 1956 est considérée comme l'acte de naissance officiel de l'IA en tant que discipline académique. Depuis lors, le domaine a connu des périodes d'optimisme intense alternant avec des 'hivers de l'IA', périodes durant lesquelles les financements et l'intérêt diminuaient face aux difficultés techniques rencontrées.",
                    subsections: [
                        {
                            title: "1.1.1 Les pionniers",
                            content: "Les pères fondateurs de l'intelligence artificielle ont posé les bases théoriques et pratiques de ce domaine révolutionnaire. Alan Turing, avec sa machine universelle et son test d'intelligence, a établi les fondements conceptuels. John McCarthy, qui a inventé le terme 'intelligence artificielle', a développé le langage de programmation LISP, encore utilisé aujourd'hui. Marvin Minsky a cofondé le laboratoire d'IA du MIT et a contribué aux réseaux de neurones. Claude Shannon a apporté la théorie de l'information, essentielle pour le traitement des données. Ces visionnaires ont créé un nouveau champ scientifique qui continue d'évoluer et de nous surprendre."
                        }
                    ]
                },
                {
                    title: "1.2 Types d'Intelligence Artificielle",
                    content: "L'intelligence artificielle se décline en plusieurs catégories selon ses capacités et son niveau de sophistication. L'IA faible ou étroite (Narrow AI) est spécialisée dans des tâches spécifiques : reconnaissance d'images, traduction automatique, jeu d'échecs. C'est le type d'IA que nous utilisons quotidiennement. L'IA générale (AGI) représente une intelligence capable de comprendre, apprendre et appliquer ses connaissances à n'importe quel problème, comme le ferait un humain. Cette forme d'IA reste théorique. La superintelligence artificielle (ASI) dépasserait l'intelligence humaine dans tous les domaines. Cette perspective soulève des questions éthiques et philosophiques profondes sur l'avenir de l'humanité."
                }
            ]
        },
        {
            title: "2. Applications de l'IA",
            introduction: "L'intelligence artificielle s'est infiltrée dans pratiquement tous les secteurs de l'économie et de la société. Des assistants vocaux sur nos smartphones aux systèmes de recommandation des plateformes de streaming, en passant par les diagnostics médicaux assistés et les véhicules autonomes, l'IA transforme notre manière de vivre, de travailler et d'interagir. Cette révolution technologique offre des opportunités sans précédent pour améliorer l'efficacité, la précision et créer de nouvelles possibilités autrefois inimaginables.",
            sections: [
                {
                    title: "2.1 IA dans la Santé",
                    content: "Le secteur de la santé connaît une transformation radicale grâce à l'intelligence artificielle. Les algorithmes d'apprentissage profond analysent des millions d'images médicales pour détecter des cancers, des maladies cardiaques ou des anomalies avec une précision souvent supérieure aux médecins expérimentés. L'IA accélère la découverte de nouveaux médicaments en simulant des millions de combinaisons moléculaires. La médecine personnalisée utilise l'IA pour adapter les traitements au profil génétique unique de chaque patient. Les robots chirurgicaux assistés par IA permettent des opérations d'une précision microscopique.",
                    subsections: [
                        {
                            title: "2.1.1 Diagnostic assisté",
                            content: "Les systèmes de diagnostic assisté par intelligence artificielle révolutionnent la pratique médicale. Grâce au deep learning et aux réseaux de neurones convolutifs, ces systèmes analysent radiographies, IRM, scanners et autres images médicales avec une rapidité et une précision remarquables. Par exemple, l'IA peut détecter des signes précoces de rétinopathie diabétique, identifier des tumeurs de quelques millimètres, ou prédire le risque de maladies cardiovasculaires. Ces outils ne remplacent pas les médecins mais augmentent leurs capacités diagnostiques, réduisant les erreurs et permettant une détection plus précoce des pathologies."
                        }
                    ]
                },
                {
                    title: "2.2 IA dans les Transports",
                    content: "L'industrie des transports est en pleine mutation grâce à l'intelligence artificielle. Les véhicules autonomes, équipés de capteurs sophistiqués et d'algorithmes d'apprentissage automatique, promettent de réduire drastiquement les accidents de la route causés par l'erreur humaine. L'IA optimise les flux de trafic urbain en temps réel, réduisant les embouteillages et la pollution. Dans l'aviation, les systèmes de pilotage automatique deviennent de plus en plus sophistiqués. Le transport maritime utilise l'IA pour optimiser les routes et réduire la consommation de carburant. Les applications de mobilité prédisent les besoins de transport et proposent des solutions multimodales personnalisées."
                }
            ]
        },
        {
            title: "3. Défis et Perspectives",
            introduction: "Malgré ses avancées spectaculaires et ses promesses enthousiasmantes, l'intelligence artificielle soulève des défis considérables. Les questions éthiques, juridiques, sociales et techniques doivent être abordées pour assurer un développement responsable et bénéfique de cette technologie. La société doit naviguer entre les opportunités extraordinaires offertes par l'IA et les risques potentiels qu'elle représente, trouvant un équilibre qui maximise les bénéfices tout en minimisant les dangers.",
            sections: [
                {
                    title: "3.1 Défis Éthiques",
                    content: "Les implications éthiques de l'intelligence artificielle sont au cœur des débats contemporains. Le biais algorithmique peut perpétuer et amplifier les discriminations existantes si les données d'entraînement reflètent des préjugés sociétaux. La protection de la vie privée est menacée par la capacité de l'IA à analyser d'énormes quantités de données personnelles. La question de la responsabilité se pose : qui est responsable quand une IA prend une décision erronée ? L'impact sur l'emploi inquiète, avec l'automatisation qui pourrait remplacer de nombreux métiers. La transparence des décisions algorithmiques reste un défi majeur, particulièrement avec les réseaux de neurones profonds dont le fonctionnement est opaque."
                },
                {
                    title: "3.2 L'Avenir de l'IA",
                    content: "L'avenir de l'intelligence artificielle s'annonce à la fois passionnant et incertain. Les progrès vers l'IA générale pourraient révolutionner notre compréhension de l'intelligence elle-même. L'informatique quantique promet d'accélérer exponentiellement les capacités de calcul, ouvrant de nouvelles frontières pour l'IA. La collaboration homme-machine évoluera vers une symbiose où l'IA augmente les capacités humaines plutôt que de les remplacer. Les interfaces cerveau-ordinateur pourraient permettre une interaction directe avec l'IA. Les questions de gouvernance mondiale de l'IA deviennent cruciales pour assurer un développement équitable et sûr. L'émergence possible d'une conscience artificielle soulève des questions philosophiques fondamentales sur la nature de la conscience et les droits des entités artificielles."
                }
            ]
        }
    ],
    
    // Conclusion
    conclusion: {
        title: "Conclusion",
        content: "L'intelligence artificielle représente sans conteste l'une des technologies les plus transformatrices de notre époque. Elle redéfinit les limites du possible dans pratiquement tous les domaines de l'activité humaine, de la médecine à l'éducation, de l'industrie aux arts. Alors que nous naviguons dans cette révolution technologique sans précédent, il est impératif de développer l'IA de manière responsable et éthique. Les défis sont nombreux : biais algorithmiques, protection de la vie privée, impact sur l'emploi, questions de sécurité et de contrôle. Cependant, les opportunités sont immenses : amélioration de la santé, résolution de problèmes complexes, augmentation des capacités humaines, découvertes scientifiques accélérées. L'avenir sera façonné par notre capacité collective à exploiter le potentiel extraordinaire de l'IA tout en gérant ses risques avec sagesse et prévoyance. La collaboration entre chercheurs, entreprises, gouvernements et société civile sera essentielle pour garantir que l'IA bénéficie à l'humanité dans son ensemble."
    },
    
    // Source
    source: "Source : Contenu généré par intelligence artificielle à des fins pédagogiques - Basé sur des connaissances générales sur l'intelligence artificielle compilées jusqu'en 2024. Ce document est créé dans un but éducatif pour l'apprentissage de la création de compléments Word avec table des matières."
};

// Fonction utilitaire pour obtenir tout le texte du document (pour recherche)
function getAllDocumentText() {
    let fullText = [];
    
    // Page de garde
    fullText.push(documentContent.coverPage.title);
    fullText.push(documentContent.coverPage.subtitle);
    
    // Chapitres
    documentContent.chapters.forEach(chapter => {
        fullText.push(chapter.title);
        if (chapter.introduction) {
            fullText.push(chapter.introduction);
        }
        
        chapter.sections.forEach(section => {
            fullText.push(section.title);
            fullText.push(section.content);
            
            if (section.subsections) {
                section.subsections.forEach(subsection => {
                    fullText.push(subsection.title);
                    fullText.push(subsection.content);
                });
            }
        });
    });
    
    // Conclusion
    fullText.push(documentContent.conclusion.title);
    fullText.push(documentContent.conclusion.content);
    
    return fullText.join('\n');
}

// Fonction pour obtenir la structure de la table des matières
function getTableOfContentsStructure() {
    let toc = [];
    
    documentContent.chapters.forEach((chapter, chapterIndex) => {
        // Ajouter le chapitre
        toc.push({
            level: 1,
            title: chapter.title,
            page: chapterIndex + 3 // Les chapitres commencent page 3
        });
        
        // Ajouter les sections
        chapter.sections.forEach((section, sectionIndex) => {
            toc.push({
                level: 2,
                title: section.title,
                page: chapterIndex + 3
            });
            
            // Ajouter les sous-sections
            if (section.subsections) {
                section.subsections.forEach(subsection => {
                    toc.push({
                        level: 3,
                        title: subsection.title,
                        page: chapterIndex + 3
                    });
                });
            }
        });
    });
    
    // Ajouter la conclusion
    toc.push({
        level: 1,
        title: "Conclusion",
        page: 6
    });
    
    return toc;
}

// Export pour utilisation dans d'autres scripts
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        documentContent,
        getAllDocumentText,
        getTableOfContentsStructure
    };
}