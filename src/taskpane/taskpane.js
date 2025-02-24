/* global Office, Word, XLSX */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initialize();
  }
});

let phrasesData = [];

function initialize() {
  document.getElementById("loadButton").addEventListener("click", async () => {
    const fileInput = document.getElementById("fileInput");
    if (!fileInput.files.length) {
      alert("Sélectionnez un fichier Excel.");
      return;
    }
    await loadLocalExcelFile(fileInput.files[0]);
  });
  document.getElementById("searchInput").addEventListener("input", searchPhrases);
}

async function loadLocalExcelFile(file) {
  try {
    const data = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // Prépare les données pour la recherche
    phrasesData = buildSearchData(sheetData);

    // Construit le menu accordéon
    const structuredData = buildStructureFromData(sheetData);
    displayDataInMenu(structuredData);
  } catch (err) {
    console.error("Erreur lecture Excel :", err);
    alert("Erreur lors de la lecture du fichier Excel.");
  }
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

function buildSearchData(sheetData) {
  let list = [];
  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    if (!row || row.length < 5) continue;
    list.push({
      title: row[2] || "Sans titre",
      text: row[4] || "",
      avis: row[3] || "",
    });
  }
  return list;
}

function buildStructureFromData(sheetData) {
  let structuredData = {};
  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    if (!row || row.length < 5) continue;
    const [travail, theme, situation, avis, phrase] = row;

    let categories = (travail || "")
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
    let subThemes = (theme || "")
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);

    if (!categories.length) categories = ["Sans catégorie"];
    if (!subThemes.length) subThemes = ["Sans sous-catégorie"];

    for (let cat of categories) {
      if (!structuredData[cat]) structuredData[cat] = {};
      for (let st of subThemes) {
        if (!structuredData[cat][st]) structuredData[cat][st] = [];
        structuredData[cat][st].push({
          title: situation,
          text: phrase,
          avis,
        });
      }
    }
  }
  return structuredData;
}

function displayDataInMenu(data) {
  const menuContainer = document.getElementById("menu-container");
  menuContainer.innerHTML = "";

  for (const category in data) {
    // Catégorie
    let catDiv = document.createElement("div");
    catDiv.className = "category-header";
    catDiv.innerHTML = `<span>${category}</span><span class="arrow">▶</span>`;

    // Contenu de la catégorie
    let catContent = document.createElement("div");
    catContent.className = "category-content";
    catContent.style.display = "block";

    // Toggle sur la catégorie
    catDiv.onclick = () => toggleCategory(catDiv, catContent);

    // Sous-catégories
    for (const subCategory in data[category]) {
      let subCatDiv = document.createElement("div");
      subCatDiv.className = "subcategory-header";
      // On ajoute la flèche pour éviter le null
      subCatDiv.innerHTML = `<span>${subCategory}</span><span class="arrow">▶</span>`;

      let phraseList = document.createElement("div");
      phraseList.className = "subcategory-content"; // caché par défaut

      // Toggle sur la sous-catégorie
      subCatDiv.onclick = () => toggleCategory(subCatDiv, phraseList);

      data[category][subCategory].forEach((item) => {
        let btn = document.createElement("button");
        btn.className = "phrase-button";
        let displayTitle = item.title || "Sans titre";

        // Ajout de l'icône + classe (couleur + hover)
        if (item.avis === "Bon") {
          displayTitle = "✅ " + displayTitle;
          btn.classList.add("bon");
        } else if (item.avis === "Mauvais") {
          displayTitle = "⛔ " + displayTitle;
          btn.classList.add("mauvais");
        }
        btn.textContent = displayTitle;

        btn.onclick = () => insertPhraseInWord(item.text);
        phraseList.appendChild(btn);
      });

      catContent.appendChild(subCatDiv);
      catContent.appendChild(phraseList);
    }

    menuContainer.appendChild(catDiv);
    menuContainer.appendChild(catContent);
  }
}

function searchPhrases() {
  const query = document.getElementById("searchInput").value.toLowerCase().trim();
  const resultsContainer = document.getElementById("search-results");
  resultsContainer.innerHTML = "";

  if (!query) return;

  const words = query.split(/\s+/);

  // Filtrer les résultats basés sur le texte ET le titre
  const results = phrasesData.filter((item) =>
    words.every(
      (w) => item.text.toLowerCase().includes(w) || item.title.toLowerCase().includes(w) // 🔍 Recherche aussi dans les titres
    )
  );

  if (!results.length) {
    resultsContainer.innerHTML = '<p class="no-results">Aucun résultat trouvé</p>';
  } else {
    results.forEach((item) => {
      let btn = document.createElement("button");
      btn.className = "result-button";
      let displayTitle = item.title;

      // Appliquer les styles selon l'avis
      if (item.avis === "Bon") {
        displayTitle = "✅ " + displayTitle;
        btn.style.backgroundColor = "#BFEDC6";
        btn.onmouseover = () => {
          btn.style.backgroundColor = "#A7DFAF";
        };
        btn.onmouseout = () => {
          btn.style.backgroundColor = "#BFEDC6";
        };
      } else if (item.avis === "Mauvais") {
        displayTitle = "⛔ " + displayTitle;
        btn.style.backgroundColor = "#FFC2C3";
        btn.onmouseover = () => {
          btn.style.backgroundColor = "#F5B4B5";
        };
        btn.onmouseout = () => {
          btn.style.backgroundColor = "#FFC2C3";
        };
      }

      btn.textContent = displayTitle;
      btn.onclick = () => insertPhraseInWord(item.text);
      resultsContainer.appendChild(btn);
    });
  }
}

/**
 * Toggle la catégorie ou sous-catégorie, + rotation de la flèche
 */
function toggleCategory(headerElement, contentElement) {
  if (!contentElement) return;
  const arrow = headerElement.querySelector(".arrow");
  if (contentElement.style.display === "block") {
    contentElement.style.display = "none";
    if (arrow) arrow.style.transform = "rotate(0deg)";
  } else {
    contentElement.style.display = "block";
    if (arrow) arrow.style.transform = "rotate(90deg)";
  }
}

async function insertPhraseInWord(text) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text + "\n", Word.InsertLocation.replace);
    const range = selection.getRange(Word.RangeLocation.after);
    range.select();
    await context.sync();
  });
}
