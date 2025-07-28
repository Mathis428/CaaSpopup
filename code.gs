/**
 * Fonction principale pour ouvrir la fenêtre popup de saisie
 */
function ouvrirFenetreCalculateur() {
  // Créer la boîte de dialogue HTML
  var htmlOutput = HtmlService.createTemplateFromFile('popup');
  
  // Passer les données existantes si nécessaire
  htmlOutput.donnees = getDonneesExistantes();
  
  // Ajouter la fonction pour formater l'index (résout le problème "idx is not defined")
  htmlOutput.formatIndex = function(i) {
    return (i < 10) ? '0' + i : '' + i;
  };
  
  var html = htmlOutput.evaluate()
    .setWidth(1200)
    .setHeight(900)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  // Afficher la popup
  SpreadsheetApp.getUi().showModalDialog(html, 'Calculateur CaaS - Saisie des données');
}

/**
 * Fonction pour créer un bouton dans le menu - s'exécute automatiquement à l'ouverture
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calculateur CaaS')
    .addItem('Ouvrir la fenêtre de saisie', 'ouvrirFenetreCalculateur')
    .addSeparator()
    .addItem('Configurer la feuille d\'aide', 'creerBoutonUI')
    .addItem('Initialiser les en-têtes', 'initialiserEnTetes')
    .addSeparator()
    .addItem('Optimiser les chillers', 'optimiserSelectionChillers')
    .addItem('Calculer moyenne demande', 'calculerMoyenneDemande')
    .addItem('Installer formule load minimum', 'installerFormuleLoadMinimum')
    .addItem('Tester optimisation seule', 'testerOptimisationChillers')
    .addItem('Tester processus complet', 'testerProcessusComplet')
    .addSeparator()
    .addItem('Installer les déclencheurs', 'installerDeclencheurs')
    .addToUi();
    
  // Créer une interface utilisateur plus fiable avec des instructions claires
  creerBoutonUI();
  
  // Pour l'ouverture automatique, on doit utiliser une approche différente
  // Si l'événement est une ouverture simple (par un utilisateur)
  if (e && e.authMode === ScriptApp.AuthMode.FULL) {
    // Afficher une boîte de dialogue demandant à l'utilisateur s'il souhaite ouvrir le calculateur
    var reponse = ui.alert(
      'Calculateur CaaS',
      'Souhaitez-vous ouvrir le calculateur CaaS maintenant ?',
      ui.ButtonSet.YES_NO
    );
    
    if (reponse === ui.Button.YES) {
      ouvrirFenetreCalculateur();
    }
  }
}

/**
 * Fonction pour ouvrir directement la fenêtre sans demander
 * Cette fonction peut être utilisée avec un bouton de la feuille ou un lien
 * 
 * IMPORTANT : Pour créer un lien manuel dans une cellule, utilisez la formule :
 * =HYPERLINK("https://script.google.com/macros/d/{ID_SCRIPT}/exec?functionName=ouvrirFenetreDirecte", "OUVRIR CALCULATEUR")
 * Où {ID_SCRIPT} est l'identifiant de votre script, visible dans l'URL de l'éditeur de script
 */
function ouvrirFenetreDirecte() {
  ouvrirFenetreCalculateur();
}

/**
 * Fonction utilitaire pour nettoyer les déclencheurs temporaires
 */
function nettoyerDeclencheursTmp() {
  // Obtenir tous les déclencheurs
  var declencheurs = ScriptApp.getProjectTriggers();
  
  // Parcourir les déclencheurs et supprimer les déclencheurs temporaires
  for (var i = 0; i < declencheurs.length; i++) {
    var declencheur = declencheurs[i];
    
    // Si c'est un déclencheur temporaire (timeBased) pour la fonction ouvrirFenetreAutomatique, le supprimer
    if (declencheur.getEventType() === ScriptApp.EventType.CLOCK && 
        declencheur.getHandlerFunction() === 'ouvrirFenetreAutomatique') {
      ScriptApp.deleteTrigger(declencheur);
    }
  }
}

/**
 * Fonction pour créer un bouton dans la cellule I11 de la feuille "0 - Read me"
 */
function creerBoutonCellule() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('0 - Read me');
    
    if (sheet) {
      // Sélectionner la cellule I11
      var cellule = sheet.getRange("I11");
      
      // Créer une formule avec un lien HYPERLINK qui exécute une fonction Apps Script
      var scriptURL = ScriptApp.getService().getUrl();
      var formuleBouton = '=HYPERLINK("' + scriptURL + '?functionName=ouvrirFenetreDirecte", "OUVRIR CALCULATEUR")';
      cellule.setFormula(formuleBouton);
      
      // Appliquer une mise en forme au style de bouton
      cellule.setBackground("#4285F4");
      cellule.setFontColor("white");
      cellule.setFontWeight("bold");
      cellule.setHorizontalAlignment("center");
      cellule.setVerticalAlignment("middle");
      
      // Ajouter une bordure pour ressembler à un bouton
      var bordures = SpreadsheetApp.BorderStyle.SOLID;
      cellule.setBorder(true, true, true, true, false, false, "#3367D6", bordures);
      
      // Ajouter une note pour expliquer comment utiliser le bouton
      cellule.setNote("Cliquez sur ce bouton pour ouvrir le calculateur CaaS");
      
      // Ajouter un commentaire pour l'instruction d'utilisation
      sheet.getRange("I12").setValue("Cliquez sur ce bouton bleu pour ouvrir le calculateur");
      sheet.getRange("I12").setFontStyle("italic");
      sheet.getRange("I12").setFontSize(10);
      
      // Ajouter l'URL complète dans une cellule pour faciliter la copie
      sheet.getRange("I13").setValue("URL du script: " + scriptURL);
      sheet.getRange("I13").setFontSize(8);
      sheet.getRange("I13").setFontColor("#666666");
      sheet.getRange("I13").setNote("Copiez cette URL si vous devez créer le lien manuellement");
      
      // Créer un vrai bouton dessiné sur la feuille comme alternative
      var image = SpreadsheetApp.newCellImage()
        .setAltTextTitle("Ouvrir Calculateur")
        .setAltTextDescription("Bouton pour ouvrir le calculateur CaaS")
        .build();
        
      sheet.getRange("K11").setValue("Ou cliquez ici →");
      sheet.getRange("L11").setValue("OUVRIR CALCULATEUR");
      sheet.getRange("L11").setBackground("#4CAF50");
      sheet.getRange("L11").setFontColor("white");
      sheet.getRange("L11").setFontWeight("bold");
      sheet.getRange("L11").setNote("Assigné à la fonction 'ouvrirFenetreCalculateur'");
      
      // Créer une assignation pour ce bouton (ceci n'est pas possible par programmation, mais c'est une instruction)
      sheet.getRange("K12:L12").merge();
      sheet.getRange("K12").setValue("(Assignez ce bouton à la fonction 'ouvrirFenetreCalculateur' via le menu 'Insertion > Dessins')");
      sheet.getRange("K12").setFontStyle("italic");
      sheet.getRange("K12").setFontSize(10);
    }
  } catch (error) {
    console.log("Erreur lors de la création du bouton dans la cellule: " + error);
  }
}

/**
 * Fonction pour récupérer les données existantes de la feuille "1 - MAIN"
 */
function getDonneesExistantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('1 - MAIN');
  
  // Si la feuille n'existe pas, créer un message d'erreur
  if (!sheet) {
    console.error("La feuille '1 - MAIN' n'a pas été trouvée");
    // Utiliser la feuille active comme fallback
    sheet = ss.getActiveSheet();
  }
  
  // Structure de données avec tous les paramètres selon le nouveau mapping
  var donnees = {
    // Mode de calcul (on peut le déterminer à partir des données présentes)
    modeCalcul: 'avance', // Par défaut en mode avancé
    
    // Informations générales
    nomProjet: '',
    
    // Type de bâtiment depuis H4 (corrigé de H5)
    typeBatiment: sheet.getRange('H4').getValue() || 'Hotels',
    
    // Paramètres géographiques
    hemisphere: sheet.getRange('H7').getValue() || 'South',
    cddBase: sheet.getRange('H10').getValue() || 18,
    
    // Paramètres ECMs - Night Reduction depuis J4 et K5-M7
    nightReduction: sheet.getRange('J4').getValue() === 'TRUE',
    startTimeWeekday: sheet.getRange('K5').getValue() || 22,
    startTimeFriday: sheet.getRange('L5').getValue() || 22,
    startTimeWeekend: sheet.getRange('M5').getValue() || 22,
    endTimeWeekday: sheet.getRange('K6').getValue() || 9,
    endTimeFriday: sheet.getRange('L6').getValue() || 9,
    endTimeWeekend: sheet.getRange('M6').getValue() || 9,
    nightDeltaT: sheet.getRange('M7').getValue() || 8,
    
    // Sélection des Chillers depuis H12, H13, H14
    chiller1: sheet.getRange('H12').getValue() || '',
    chiller2: sheet.getRange('H13').getValue() || '',
    chiller3: sheet.getRange('H14').getValue() || '',
    
    // Données des équipements - B19-E25
    // Chillers
    chillers_units: sheet.getRange('B19').getValue() || '3',
    chillers_load: sheet.getRange('C19').getValue() || '2600',
    chillers_cop: sheet.getRange('E19').getValue() || '5.2',
    
    // Configuration individuelle des chillers (déterminée en comparant les valeurs)
    individual_chillers: (function() {
      var load1 = sheet.getRange('C19').getValue() || 0;
      var load2 = sheet.getRange('F20').getValue() || 0;
      var load3 = sheet.getRange('G20').getValue() || 0;
      var cop1 = sheet.getRange('D19').getValue() || 0;
      var cop2 = sheet.getRange('F22').getValue() || 0;
      var cop3 = sheet.getRange('G22').getValue() || 0;
      
      // Si toutes les valeurs sont identiques ou vides, ce n'est pas individuel
      return (load1 !== load2 || load1 !== load3 || cop1 !== cop2 || cop1 !== cop3) && 
             (load2 !== 0 || load3 !== 0 || cop2 !== 0 || cop3 !== 0);
    })(),
    
    // Données individuelles des chillers
    chiller_1_load: sheet.getRange('C19').getValue() || '2600',  // Load 1 -> C19
    chiller_1_cop: sheet.getRange('D19').getValue() || '5.2',   // COP 1 -> D19
    chiller_2_load: sheet.getRange('F20').getValue() || '2600', // Load 2 -> F20
    chiller_2_cop: sheet.getRange('F22').getValue() || '5.2',   // COP 2 -> F22
    chiller_3_load: sheet.getRange('G20').getValue() || '2600', // Load 3 -> G20
    chiller_3_cop: sheet.getRange('G22').getValue() || '5.2',   // COP 3 -> G22
    
    // Chillers 4 et 5 (interface seulement, pas sauvegardés)
    chiller_4_load: '',
    chiller_4_cop: '',
    chiller_5_load: '',
    chiller_5_cop: '',
    
    // Pompes primaires
    prim_pumps_units: sheet.getRange('B21').getValue() || '3',
    prim_pumps_flow: sheet.getRange('C21').getValue() || '330.0',
    prim_pumps_n_flow: sheet.getRange('D21').getValue() || '330',
    prim_pumps_n_load: sheet.getRange('E21').getValue() || '22.06',
    
    // Pompes secondaires
    sec_pumps_units: sheet.getRange('B23').getValue() || '2',
    sec_pumps_flow: sheet.getRange('C23').getValue() || '250.0',
    sec_pumps_n_flow: sheet.getRange('D23').getValue() || '250',
    sec_pumps_n_load: sheet.getRange('E23').getValue() || '50',
    
    // Tours de refroidissement
    cl_towers_units: sheet.getRange('B25').getValue() || '6',
    cl_towers_flow: sheet.getRange('C25').getValue() || '128.3',
    
    // Données d'énergie - extraites des cellules A5-A16 et C5-C16
    date_01: '', electricity_01: '',
    date_02: '', electricity_02: '',
    date_03: '', electricity_03: '',
    date_04: '', electricity_04: '',
    date_05: '', electricity_05: '',
    date_06: '', electricity_06: '',
    date_07: '', electricity_07: '',
    date_08: '', electricity_08: '',
    date_09: '', electricity_09: '',
    date_10: '', electricity_10: '',
    date_11: '', electricity_11: '',
    date_12: '', electricity_12: '',
    
    // Paramètres météo - extraits des cellules A17-A18
    includeWeatherData: false,
    locationName: '',
    geocode: ''
  };
  
  // Extraire les informations générales du projet depuis les cellules A1
  try {
    var a1Value = sheet.getRange('A1').getValue();
    if (a1Value && a1Value.toString().includes('Nom Projet:')) {
      donnees.nomProjet = a1Value.toString().replace('Nom Projet: ', '');
    }
  } catch (error) {
    console.log('Erreur lors de la lecture des informations générales:', error);
  }
  
  // Extraire les données d'énergie (dates et électricité) depuis A5-A16 et C5-C16
  try {
    // Lecture des dates (A5-A16) et valeurs d'électricité (C5-C16)
    for (var i = 1; i <= 12; i++) {
      var idx = (i < 10) ? '0' + i : '' + i;
      var row = 4 + i; // Les données commencent à la ligne 5 (index 4+1)
      
      // Lire la date
      var dateValue = sheet.getRange('A' + row).getValue();
      if (dateValue) {
        donnees['date_' + idx] = dateValue.toString();
      }
      
      // Lire la valeur d'électricité
      var elecValue = sheet.getRange('C' + row).getValue();
      if (elecValue) {
        donnees['electricity_' + idx] = elecValue.toString();
      }
    }
  } catch (error) {
    console.log('Erreur lors de la lecture des données d\'énergie:', error);
  }
  
  // Extraire les données météo depuis les cellules A17-A18
  try {
    var a17Value = sheet.getRange('A17').getValue();
    if (a17Value && a17Value.toString().includes('Coordonnées GPS:')) {
      donnees.geocode = a17Value.toString().replace('Coordonnées GPS: ', '');
    }
    
    var a18Value = sheet.getRange('A18').getValue();
    if (a18Value && a18Value.toString().includes('Lieu:')) {
      donnees.locationName = a18Value.toString().replace('Lieu: ', '');
    }
  } catch (error) {
    console.log('Erreur lors de la lecture des données météo:', error);
  }
  
  return donnees;
}

/**
 * Fonction pour récupérer les résultats techniques du calcul depuis la feuille "1 - MAIN"
 */
function getResultatsTechniques() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('1 - MAIN');
    
    // Si la feuille n'existe pas, créer un message d'erreur
    if (!sheet) {
      console.error("La feuille '1 - MAIN' n'a pas été trouvée");
      return {error: "La feuille '1 - MAIN' n'a pas été trouvée"};
    }
    
    var resultats = {
      // CHILLERS
      chillers_current_consumption: sheet.getRange('R4').getValue() || '0',
      chillers_new_consumption: sheet.getRange('R5').getValue() || '0',
      chillers_savings: sheet.getRange('R6').getValue() || '0',
      chiller1_startups_runtime: sheet.getRange('R7').getValue() || '0 / 0 hours',
      chiller2_startups_runtime: sheet.getRange('R8').getValue() || '0 / 0 hours',
      chiller3_startups_runtime: sheet.getRange('R9').getValue() || '0 / 0 hours',
      
      // SEC PUMPS
      sec_pumps_current_consumption: sheet.getRange('R10').getValue() || '0',
      sec_pumps_new_consumption: sheet.getRange('R11').getValue() || '0',
      sec_pumps_savings: sheet.getRange('R12').getValue() || '0',
      
      // TOTAL
      total_savings: sheet.getRange('R13').getValue() || '0',
      total_percent_cooling: sheet.getRange('R14').getValue() || '0',
      total_percent_cspt: sheet.getRange('R15').getValue() || '0'
    };
    
    return resultats;
  } catch (error) {
    console.log('Erreur lors de la récupération des résultats techniques:', error);
    return {
      error: true, 
      message: 'Erreur lors de la récupération des résultats: ' + error.toString()
    };
  }
}

/**
 * Fonction pour sauvegarder les données dans la feuille "1 - MAIN"
 */
/**
 * Récupération des données météo depuis l'API Veolia
 * @param {Object} weatherParams - Les paramètres de la demande météo (coordonnées, dates, etc.)
 * @returns {Object} Les données météo récupérées
 */
function getWeatherData(weatherParams) {
  try {
    console.log('Récupération des données météo avec les paramètres:', JSON.stringify(weatherParams));
    
    // Dans un environnement réel, cette fonction appellerait l'API Veolia
    // Ici nous simulons la récupération de données pour démontrer le flux
    
    // Simulation de l'appel à l'API Veolia en utilisant les informations d'authentification OAuth2.0
    var mockWeatherData = {
      location: {
        name: weatherParams.locationName || "Location",
        latitude: weatherParams.latitude,
        longitude: weatherParams.longitude,
        geocode: weatherParams.geocode
      },
      period: {
        startDate: weatherParams.startDate,
        endDate: weatherParams.endDate
      },
      data: generateMockWeatherData(weatherParams.startDate, weatherParams.endDate),
      source: "Veolia Weather API (simulated)",
      timestamp: new Date().toISOString()
    };
    
    // Sauvegarde des données dans la feuille Weather Data
    saveWeatherDataToSheet(mockWeatherData);
    
    return mockWeatherData;
  } catch (error) {
    console.error('Erreur lors de la récupération des données météo:', error);
    throw new Error('Failed to retrieve weather data: ' + error.message);
  }
}

/**
 * Génère des données météo simulées pour la période demandée
 */
function generateMockWeatherData(startDate, endDate) {
  var data = [];
  
  // Convertir les dates en objets Date
  var start = new Date(startDate);
  var end = new Date(endDate || startDate);
  
  // Limiter à 30 jours maximum pour la simulation
  var maxDays = 30;
  var currentDate = new Date(start);
  var dayCount = 0;
  
  while (currentDate <= end && dayCount < maxDays) {
    // Température entre 5 et 30 degrés
    var temperature = 5 + Math.random() * 25;
    
    // Humidité entre 30 et 90%
    var humidity = 30 + Math.random() * 60;
    
    // Vent entre 0 et 30 km/h
    var windSpeed = Math.random() * 30;
    
    data.push({
      date: new Date(currentDate).toISOString().split('T')[0],
      temperature: temperature.toFixed(1),
      humidity: humidity.toFixed(1),
      windSpeed: windSpeed.toFixed(1),
      precipitation: (Math.random() * 10).toFixed(1)
    });
    
    // Passer au jour suivant
    currentDate.setDate(currentDate.getDate() + 1);
    dayCount++;
  }
  
  return data;
}

/**
 * Télécharge et sauvegarde les données météo Veolia dans la feuille "Weather Data"
 */
function downloadAndSaveVeoliaWeatherData(weatherParams) {
  try {
    console.log('Téléchargement des données météo Veolia...', weatherParams);
    
    // 1. Récupérer les données depuis l'API Veolia
    const veoliaData = getVeoliaWeatherData(weatherParams);
    
    if (!veoliaData.success) {
      throw new Error(veoliaData.message || 'Erreur lors du téléchargement des données météo');
    }
    
    console.log('Données Veolia récupérées avec succès');
    
    // 2. Sauvegarder dans la feuille Weather Data
    saveVeoliaWeatherDataToSheet(veoliaData, weatherParams);
    
    return {
      success: true,
      message: 'Données météo téléchargées et sauvegardées avec succès',
      data: veoliaData
    };
    
  } catch (error) {
    console.error('Erreur lors du téléchargement des données météo:', error);
    return {
      success: false,
      message: 'Erreur lors du téléchargement: ' + error.toString(),
      error: error.toString()
    };
  }
}

/**
 * Sauvegarde les données météo Veolia dans la feuille "Weather Data" avec format professionnel
 */
function saveVeoliaWeatherDataToSheet(veoliaData, weatherParams) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Weather Data');
  
  // Créer la feuille si elle n'existe pas
  if (!sheet) {
    sheet = ss.insertSheet('Weather Data');
  }
  
  // Effacer tout le contenu existant
  sheet.clear();
  
  // Créer les métadonnées d'en-tête avec format professionnel
  const timestamp = new Date();
  const location = weatherParams.locationName || weatherParams.locationAddress || 
                  (weatherParams.latitude + ', ' + weatherParams.longitude);
  
  // === SECTION 1: WEATHER DATA PLATFORM OVERVIEW ===
  let currentRow = 1;
  
  // En-tête principal
  sheet.getRange(currentRow, 1).setValue('WEATHER DATA PLATFORM').setFontWeight('bold').setFontSize(16);
  sheet.getRange(currentRow, 1, 1, 8).merge().setHorizontalAlignment('center');
  sheet.getRange(currentRow, 1, 1, 8).setBackground('#1f4e79').setFontColor('white');
  currentRow += 2;
  
  // Informations de téléchargement
  sheet.getRange(currentRow, 1).setValue('Downloaded:').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setValue(timestamp.toLocaleString('en-US'));
  sheet.getRange(currentRow, 4).setValue('Location:').setFontWeight('bold');
  sheet.getRange(currentRow, 5).setValue(location);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue('Coordinates:').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setValue(weatherParams.latitude + ', ' + weatherParams.longitude);
  sheet.getRange(currentRow, 4).setValue('Data Type:').setFontWeight('bold');
  sheet.getRange(currentRow, 5).setValue(weatherParams.dataType.toUpperCase());
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue('API Endpoint:').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setValue('https://api.veolia.com/weather/v1/' + weatherParams.dataType);
  sheet.getRange(currentRow, 4).setValue('Interval:').setFontWeight('bold');
  sheet.getRange(currentRow, 5).setValue(weatherParams.timeInterval || 'N/A');
  currentRow++;
  
  if (weatherParams.startDate && weatherParams.endDate) {
    sheet.getRange(currentRow, 1).setValue('Period:').setFontWeight('bold');
    sheet.getRange(currentRow, 2).setValue(weatherParams.startDate + ' to ' + weatherParams.endDate);
    currentRow++;
  }
  
  // Mise en forme de l'en-tête
  sheet.getRange(2, 1, currentRow - 2, 8).setBackground('#e8f4fd');
  sheet.getRange(2, 1, currentRow - 2, 1).setBackground('#cce7ff');
  sheet.getRange(2, 4, currentRow - 2, 1).setBackground('#cce7ff');
  currentRow += 2;
  
  // === SECTION 2: DETAILED WEATHER DATA ===
  sheet.getRange(currentRow, 1).setValue('DETAILED WEATHER DATA').setFontWeight('bold').setFontSize(14);
  sheet.getRange(currentRow, 1, 1, 8).merge().setHorizontalAlignment('center');
  sheet.getRange(currentRow, 1, 1, 8).setBackground('#4472c4').setFontColor('white');
  currentRow += 2;
  
  // Traiter selon le type de données reçues
  if (veoliaData.data) {
    const rawData = veoliaData.data;
    
    if (weatherParams.dataType === 'current' && rawData.observations) {
      // === DONNÉES ACTUELLES ===
      const headers = [
        'DateTime', 'Temperature (°C)', 'Humidity (%)', 'Wind Speed (km/h)',
        'Wind Direction (°)', 'Pressure (hPa)', 'Precipitation (mm)', 'Visibility (km)',
        'Cloud Cover (%)', 'Dew Point (°C)', 'UV Index', 'Weather Description'
      ];
      
      // Ajouter les en-têtes avec style professionnel
      for (let i = 0; i < headers.length; i++) {
        const cell = sheet.getRange(currentRow, i + 1);
        cell.setValue(headers[i]);
        cell.setFontWeight('bold');
        cell.setBackground('#d9e1f2');
        cell.setBorder(true, true, true, true, false, false);
      }
      currentRow++;
      
      // Ajouter les données avec formatage
      rawData.observations.forEach(obs => {
        const rowData = [
          obs.dateTime || obs.date || new Date().toISOString(),
          obs.temperature || obs.temp || '',
          obs.humidity || obs.relativeHumidity || '',
          obs.windSpeed || obs.windSpeedKmH || '',
          obs.windDirection || obs.windDirectionDegree || '',
          obs.pressure || obs.pressureHPa || '',
          obs.precipitation || obs.precipMM || '',
          obs.visibility || obs.visibilityKm || '',
          obs.cloudCover || obs.cloudCoverPercent || '',
          obs.dewPoint || obs.dewPointC || '',
          obs.uvIndex || '',
          obs.weatherDescription || obs.description || ''
        ];
        
        for (let i = 0; i < rowData.length; i++) {
          const cell = sheet.getRange(currentRow, i + 1);
          cell.setValue(rowData[i]);
          cell.setBorder(true, true, true, true, false, false);
          
          // Formatage spécial pour les nombres
          if (typeof rowData[i] === 'number' && i > 0) {
            cell.setNumberFormat('0.0');
          }
        }
        
        // Alternance de couleurs pour lisibilité
        if (currentRow % 2 === 0) {
          sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f2f2f2');
        }
        currentRow++;
      });
      
    } else if (weatherParams.dataType === 'forecast') {
      // === DONNÉES DE PRÉVISION ===
      const headers = [
        'DateTime', 'Temperature (°C)', 'Temp Min (°C)', 'Temp Max (°C)',
        'Humidity (%)', 'Wind Speed (km/h)', 'Precipitation (mm)',
        'Precip Probability (%)', 'Weather Description', 'UV Index', 'Pressure (hPa)'
      ];
      
      // Ajouter les en-têtes
      for (let i = 0; i < headers.length; i++) {
        const cell = sheet.getRange(currentRow, i + 1);
        cell.setValue(headers[i]);
        cell.setFontWeight('bold');
        cell.setBackground('#d9e1f2');
        cell.setBorder(true, true, true, true, false, false);
      }
      currentRow++;
      
      // Ajouter les données
      const forecasts = rawData.forecasts || rawData;
      if (Array.isArray(forecasts)) {
        forecasts.forEach(forecast => {
          const rowData = [
            forecast.dateTime || forecast.date || '',
            forecast.temperature || forecast.temp || '',
            forecast.temperatureMin || forecast.tempMin || '',
            forecast.temperatureMax || forecast.tempMax || '',
            forecast.humidity || forecast.relativeHumidity || '',
            forecast.windSpeed || forecast.windSpeedKmH || '',
            forecast.precipitation || forecast.precipMM || '',
            forecast.precipitationProbability || forecast.precipProb || '',
            forecast.description || forecast.weather || '',
            forecast.uvIndex || '',
            forecast.pressure || forecast.pressureHPa || ''
          ];
          
          for (let i = 0; i < rowData.length; i++) {
            const cell = sheet.getRange(currentRow, i + 1);
            cell.setValue(rowData[i]);
            cell.setBorder(true, true, true, true, false, false);
            
            if (typeof rowData[i] === 'number' && i > 0) {
              cell.setNumberFormat('0.0');
            }
          }
          
          if (currentRow % 2 === 0) {
            sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f2f2f2');
          }
          currentRow++;
        });
      }
      
    } else if (weatherParams.dataType === 'history') {
      // === DONNÉES HISTORIQUES ===
      const headers = [
        'Date', 'Avg Temperature (°C)', 'Min Temperature (°C)', 'Max Temperature (°C)',
        'Avg Humidity (%)', 'Avg Wind Speed (km/h)', 'Precipitation (mm)',
        'Sunshine Hours', 'Avg Pressure (hPa)', 'Weather Summary'
      ];
      
      // Ajouter les en-têtes
      for (let i = 0; i < headers.length; i++) {
        const cell = sheet.getRange(currentRow, i + 1);
        cell.setValue(headers[i]);
        cell.setFontWeight('bold');
        cell.setBackground('#d9e1f2');
        cell.setBorder(true, true, true, true, false, false);
      }
      currentRow++;
      
      // Traiter les données historiques
      const historyData = rawData.history || rawData.data || [rawData];
      historyData.forEach(day => {
        const rowData = [
          day.date || '',
          day.temperatureAvg || day.avgTemp || '',
          day.temperatureMin || day.minTemp || '',
          day.temperatureMax || day.maxTemp || '',
          day.humidityAvg || day.avgHumidity || '',
          day.windSpeedAvg || day.avgWindSpeed || '',
          day.precipitation || day.precipMM || '',
          day.sunshine || day.sunshineHours || '',
          day.pressureAvg || day.avgPressure || '',
          day.weatherSummary || day.summary || ''
        ];
        
        for (let i = 0; i < rowData.length; i++) {
          const cell = sheet.getRange(currentRow, i + 1);
          cell.setValue(rowData[i]);
          cell.setBorder(true, true, true, true, false, false);
          
          if (typeof rowData[i] === 'number' && i > 0) {
            cell.setNumberFormat('0.0');
          }
        }
        
        if (currentRow % 2 === 0) {
          sheet.getRange(currentRow, 1, 1, headers.length).setBackground('#f2f2f2');
        }
        currentRow++;
      });
    }
  }
  
  // === SECTION 3: ANALYSIS & RECOMMENDATIONS ===
  if (veoliaData.recommendations && veoliaData.recommendations.length > 0) {
    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue('ANALYSIS & RECOMMENDATIONS').setFontWeight('bold').setFontSize(14);
    sheet.getRange(currentRow, 1, 1, 8).merge().setHorizontalAlignment('center');
    sheet.getRange(currentRow, 1, 1, 8).setBackground('#70ad47').setFontColor('white');
    currentRow += 2;
    
    veoliaData.recommendations.forEach((rec, index) => {
      sheet.getRange(currentRow, 1).setValue((index + 1) + '.');
      sheet.getRange(currentRow, 2).setValue(rec);
      sheet.getRange(currentRow, 1, 1, 8).setBackground('#e2efda');
      currentRow++;
    });
  }
  
  // === SECTION 4: TECHNICAL INFORMATION ===
  currentRow += 2;
  sheet.getRange(currentRow, 1).setValue('TECHNICAL INFORMATION').setFontWeight('bold').setFontSize(14);
  sheet.getRange(currentRow, 1, 1, 8).merge().setHorizontalAlignment('center');
  sheet.getRange(currentRow, 1, 1, 8).setBackground('#7030a0').setFontColor('white');
  currentRow += 2;
  
  // Informations techniques
  const techInfo = [
    ['API Base URL:', 'https://api.veolia.com/weather/v1'],
    ['Authentication:', 'OAuth 2.0 Bearer Token'],
    ['Response Format:', 'JSON'],
    ['Units:', 'Metric (°C, km/h, mm, hPa)'],
    ['Language:', 'English (en-US)'],
    ['Timestamp:', new Date().toISOString()],
    ['Data Source:', 'Veolia Weather Data Platform'],
    ['Generated by:', 'CaaS Calculator Weather Integration']
  ];
  
  techInfo.forEach(info => {
    sheet.getRange(currentRow, 1).setValue(info[0]).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setValue(info[1]);
    sheet.getRange(currentRow, 1, 1, 8).setBackground('#f2f2f2');
    currentRow++;
  });
  
  // === SECTION 5: RAW DATA (JSON) ===
  currentRow += 2;
  sheet.getRange(currentRow, 1).setValue('RAW API RESPONSE (JSON)').setFontWeight('bold').setFontSize(12);
  sheet.getRange(currentRow, 1, 1, 8).merge().setHorizontalAlignment('center');
  sheet.getRange(currentRow, 1, 1, 8).setBackground('#d9d9d9');
  currentRow++;
  
  const jsonData = JSON.stringify(veoliaData.data, null, 2);
  sheet.getRange(currentRow, 1).setValue(jsonData);
  sheet.getRange(currentRow, 1, 1, 8).merge();
  sheet.getRange(currentRow, 1).setVerticalAlignment('top').setWrap(true);
  sheet.getRange(currentRow, 1).setBackground('#f9f9f9');
  
  // Ajuster la largeur des colonnes
  sheet.autoResizeColumns(1, 12);
  
  // Définir la largeur minimale pour les colonnes principales
  for (let col = 1; col <= 12; col++) {
    if (sheet.getColumnWidth(col) < 100) {
      sheet.setColumnWidth(col, 100);
    }
  }
  
  // Figer les en-têtes
  sheet.setFrozenRows(1);
  
  console.log('Données météo sauvegardées dans la feuille Weather Data avec format professionnel');
}

/**
 * Sauvegarde les données avec téléchargement automatique des données météo Veolia
 */
function sauvegarderDonneesAvecMeteo(donnees, weatherParams) {
  try {
    console.log('Début sauvegarde avec données météo Veolia...');
    
    // 1. D'abord sauvegarder les données normales du projet
    var result = sauvegarderDonnees(donnees);
    
    if (!result.success) {
      return result;
    }
    
    // 2. Si les paramètres météo sont fournis, télécharger et sauvegarder les données Veolia
    if (weatherParams && weatherParams.latitude && weatherParams.longitude) {
      console.log('Téléchargement des données météo Veolia...', weatherParams);
      
      // Préparer les paramètres pour l'API Veolia
      const veoliaParams = {
        latitude: weatherParams.latitude,
        longitude: weatherParams.longitude,
        locationName: weatherParams.locationName || weatherParams.locationAddress,
        locationAddress: weatherParams.locationAddress,
        dataType: 'current', // Par défaut, données actuelles
        timeInterval: 'hourly',
        dataSource: 'cfsr',
        startDate: weatherParams.startDate,
        endDate: weatherParams.endDate
      };
      
      // Si des dates sont spécifiées, utiliser les données historiques
      if (weatherParams.startDate && weatherParams.endDate) {
        veoliaParams.dataType = 'history';
      }
      
      // Télécharger et sauvegarder les données météo
      const weatherResult = downloadAndSaveVeoliaWeatherData(veoliaParams);
      
      if (weatherResult.success) {
        result.weatherDataDownloaded = true;
        result.message += ' Données météo Veolia téléchargées et sauvegardées.';
        result.weatherInfo = weatherResult.data;
      } else {
        console.warn('Erreur lors du téléchargement des données météo:', weatherResult.message);
        result.weatherDataDownloaded = false;
        result.weatherError = weatherResult.message;
        result.message += ' (Attention: Erreur lors du téléchargement des données météo: ' + weatherResult.message + ')';
      }
    } else {
      console.log('Pas de paramètres météo fournis, sauvegarde normale uniquement.');
      result.weatherDataDownloaded = false;
    }
    
    return result;
    
  } catch (error) {
    console.error('Erreur dans sauvegarderDonneesAvecMeteo:', error);
    return {
      success: false,
      message: 'Erreur lors de la sauvegarde avec données météo: ' + error.toString(),
      error: error.toString()
    };
  }
}

function sauvegarderDonnees(donnees) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('1 - MAIN');
    
    // Debug: Log des données reçues
    console.log('Données reçues pour sauvegarde:', JSON.stringify(donnees));
    
    // Si la feuille n'existe pas, créer un message d'erreur
    if (!sheet) {
      console.error("La feuille '1 - MAIN' n'a pas été trouvée");
      return {success: false, message: "Erreur: La feuille '1 - MAIN' n'a pas été trouvée"};
    }
    
    // Si on est en mode simple, on sauvegarde uniquement le mode
    if (donnees.modeCalcul === 'simple') {
      sheet.getRange('H5').setValue("Mode Simple actif");
      return {success: true, message: 'Mode simple sauvegardé avec succès!'};
    }
    
    // Mode avancé - Mapping des données vers les cellules du Google Sheet selon votre tableau
    
    // 1. Type de bâtiment -> H4 (colonne Type) - corrigé de H5
    if (donnees.typeBatiment) {
      sheet.getRange('H4').setValue(donnees.typeBatiment);
    }
    
    // 2. Hémisphère -> H7
    if (donnees.hemisphere) {
      sheet.getRange('H7').setValue(donnees.hemisphere);
    }
    
    // 3. CDD Base -> H10
    if (donnees.cddBase) {
      sheet.getRange('H10').setValue(parseInt(donnees.cddBase));
    }
    
    // 4. Night Reduction -> J4 (TRUE/FALSE)
    sheet.getRange('J4').setValue(donnees.nightReduction ? 'TRUE' : 'FALSE');
    
    // 5. Si Night Reduction est activée, sauvegarder les paramètres de temps
    if (donnees.nightReduction) {
      // Start times -> K5 (Weekday), L5 (Friday), M5 (Weekend)
      if (donnees.startTimeWeekday) sheet.getRange('K5').setValue(parseInt(donnees.startTimeWeekday));
      if (donnees.startTimeFriday) sheet.getRange('L5').setValue(parseInt(donnees.startTimeFriday));
      if (donnees.startTimeWeekend) sheet.getRange('M5').setValue(parseInt(donnees.startTimeWeekend));
      
      // End times -> K6 (Weekday), L6 (Friday), M6 (Weekend)
      if (donnees.endTimeWeekday) sheet.getRange('K6').setValue(parseInt(donnees.endTimeWeekday));
      if (donnees.endTimeFriday) sheet.getRange('L6').setValue(parseInt(donnees.endTimeFriday));
      if (donnees.endTimeWeekend) sheet.getRange('M6').setValue(parseInt(donnees.endTimeWeekend));
      
      // Night Delta T -> M7
      if (donnees.nightDeltaT) sheet.getRange('M7').setValue(parseInt(donnees.nightDeltaT));
    }
    
    // 6. Capacités des chillers -> H12 (Chiller 1), H13 (Chiller 2), H14 (Chiller 3)
    if (donnees.chiller1) {
      sheet.getRange('H12').setValue(parseInt(donnees.chiller1));
    }
    if (donnees.chiller2) {
      sheet.getRange('H13').setValue(parseInt(donnees.chiller2));
    }
    if (donnees.chiller3) {
      sheet.getRange('H14').setValue(parseInt(donnees.chiller3));
    }
    
    // 6b. Sauvegarder les données des équipements dans les cellules B19-E25
    // Chillers (B19-E19)
    if (donnees.chillers_units) sheet.getRange('B19').setValue(parseInt(donnees.chillers_units));
    
    // Vérifier si la configuration individuelle des chillers est activée
    if (donnees.individual_chillers) {
      // Configuration individuelle - envoyer les valeurs spécifiques aux cellules demandées
      if (donnees.chiller_1_load && donnees.chiller_1_load.trim() !== '') {
        sheet.getRange('C19').setValue(parseFloat(donnees.chiller_1_load.replace(',', '.')));  // Load 1 -> C19
      }
      if (donnees.chiller_1_cop && donnees.chiller_1_cop.trim() !== '') {
        sheet.getRange('D19').setValue(parseFloat(donnees.chiller_1_cop.replace(',', '.')));     // COP 1 -> D19
      }
      if (donnees.chiller_2_load && donnees.chiller_2_load.trim() !== '') {
        sheet.getRange('F20').setValue(parseFloat(donnees.chiller_2_load.replace(',', '.')));   // Load 2 -> F20
      }
      if (donnees.chiller_2_cop && donnees.chiller_2_cop.trim() !== '') {
        sheet.getRange('F22').setValue(parseFloat(donnees.chiller_2_cop.replace(',', '.')));     // COP 2 -> F22
      }
      if (donnees.chiller_3_load && donnees.chiller_3_load.trim() !== '') {
        sheet.getRange('G20').setValue(parseFloat(donnees.chiller_3_load.replace(',', '.')));   // Load 3 -> G20
      }
      if (donnees.chiller_3_cop && donnees.chiller_3_cop.trim() !== '') {
        sheet.getRange('G22').setValue(parseFloat(donnees.chiller_3_cop.replace(',', '.')));     // COP 3 -> G22
      }
    } else {
      // Configuration globale - envoyer les mêmes valeurs dans les 6 cellules
      var globalLoad = null;
      var globalCop = null;
      
      if (donnees.chillers_load && donnees.chillers_load.trim() !== '') {
        globalLoad = parseFloat(donnees.chillers_load.replace(',', '.'));
      }
      if (donnees.chillers_cop && donnees.chillers_cop.trim() !== '') {
        globalCop = parseFloat(donnees.chillers_cop.replace(',', '.'));
      }
      
      if (globalLoad) {
        sheet.getRange('C19').setValue(globalLoad);  // Load 1 -> C19
        sheet.getRange('F20').setValue(globalLoad);  // Load 2 -> F20
        sheet.getRange('G20').setValue(globalLoad);  // Load 3 -> G20
      }
      
      if (globalCop) {
        sheet.getRange('D19').setValue(globalCop);   // COP 1 -> D19
        sheet.getRange('F22').setValue(globalCop);   // COP 2 -> F22
        sheet.getRange('G22').setValue(globalCop);   // COP 3 -> G22
      }
      
      // Maintenir la compatibilité avec l'ancienne cellule E19 pour le COP global
      if (globalCop) sheet.getRange('E19').setValue(globalCop);
    }
    
    // Pompes primaires (B21-E21)
    if (donnees.prim_pumps_units) sheet.getRange('B21').setValue(parseInt(donnees.prim_pumps_units));
    if (donnees.prim_pumps_flow) sheet.getRange('C21').setValue(parseFloat(donnees.prim_pumps_flow));
    if (donnees.prim_pumps_n_flow) sheet.getRange('D21').setValue(parseInt(donnees.prim_pumps_n_flow));
    if (donnees.prim_pumps_n_load) sheet.getRange('E21').setValue(parseFloat(donnees.prim_pumps_n_load));
    
    // Pompes secondaires (B23-E23)
    if (donnees.sec_pumps_units) sheet.getRange('B23').setValue(parseInt(donnees.sec_pumps_units));
    if (donnees.sec_pumps_flow) sheet.getRange('C23').setValue(parseFloat(donnees.sec_pumps_flow));
    if (donnees.sec_pumps_n_flow) sheet.getRange('D23').setValue(parseInt(donnees.sec_pumps_n_flow));
    if (donnees.sec_pumps_n_load) sheet.getRange('E23').setValue(parseFloat(donnees.sec_pumps_n_load));
    
    // Tours de refroidissement (B25-C25)
    if (donnees.cl_towers_units) sheet.getRange('B25').setValue(parseInt(donnees.cl_towers_units));
    if (donnees.cl_towers_flow) sheet.getRange('C25').setValue(parseFloat(donnees.cl_towers_flow));
    
    // 7. Sauvegarder les données d'énergie -> A5-A16 (dates) et C5-C16 (électricité)
    for (var i = 1; i <= 12; i++) {
      var idx = (i < 10) ? '0' + i : '' + i;
      var row = 4 + i; // Les données commencent à la ligne 5 (index 4+1)
      
      // Sauvegarder la date
      var dateKey = 'date_' + idx;
      if (donnees[dateKey]) {
        sheet.getRange('A' + row).setValue(donnees[dateKey]);
      }
      
      // Sauvegarder la valeur d'électricité
      var elecKey = 'electricity_' + idx;
      if (donnees[elecKey]) {
        // Convertir la valeur en nombre si possible (en gérant les séparateurs de milliers)
        var elecValue = donnees[elecKey].replace(/,/g, '');
        var numValue = parseFloat(elecValue);
        if (!isNaN(numValue)) {
          sheet.getRange('C' + row).setValue(numValue);
        } else {
          sheet.getRange('C' + row).setValue(donnees[elecKey]);
        }
      }
    }
    
    // 8. Données dans les colonnes A et C pour les dates/données d'énergie si nécessaire
    // On peut utiliser les premières lignes libres (A5-A16, C5-C16) pour d'autres données du projet
    
    // Informations générales du projet dans les colonnes disponibles
    if (donnees.nomProjet) {
      // Nom du projet dans une cellule disponible
      sheet.getRange('A1').setValue('Nom Projet: ' + donnees.nomProjet);
    }
    
    // Données météo si disponibles
    if (donnees.includeWeatherData && donnees.geocode) {
      sheet.getRange('A17').setValue('Coordonnées GPS: ' + donnees.geocode);
      if (donnees.locationName) {
        sheet.getRange('A18').setValue('Lieu: ' + donnees.locationName);
      }
    }
    
    // Déclencher le recalcul si nécessaire
    SpreadsheetApp.flush();
    
    // ÉTAPE CRITIQUE : Optimiser les chillers avant de calculer les résultats
    console.log("=== DÉBUT DE L'OPTIMISATION DES CHILLERS ===");
    try {
      // Vérifier d'abord que nous avons des données de consommation
      var donneesConsommation = false;
      for (var i = 5; i <= 16; i++) {
        var valeur = sheet.getRange('C' + i).getValue();
        if (valeur && valeur > 0) {
          donneesConsommation = true;
          break;
        }
      }
      
      if (!donneesConsommation) {
        console.warn("ATTENTION: Aucune donnée de consommation électrique détectée dans C5:C16");
      }
      
      // Calculer la moyenne de la demande avant l'optimisation
      var moyenneCalculee = calculerMoyenneDemande();
      console.log("Moyenne de la demande calculée:", moyenneCalculee);
      
      var optimisationResult = optimiserSelectionChillers();
      
      if (optimisationResult) {
        console.log("Optimisation des chillers réussie:", optimisationResult);
        console.log("Chillers optimisés - K10:", optimisationResult.chiller1, "K11:", optimisationResult.chiller2, "K12:", optimisationResult.chiller3);
        console.log("Capacité totale:", optimisationResult.totalCapacite);
        console.log("Consommation optimisée:", optimisationResult.consommationOptimisee);
      } else {
        console.warn("L'optimisation des chillers a échoué, utilisation des valeurs existantes");
      }
    } catch (optimError) {
      console.error("Erreur lors de l'optimisation des chillers:", optimError);
      // Continuer même si l'optimisation échoue
    }
    
    // Installer la formule pour le load minimum en F25
    try {
      installerFormuleLoadMinimum();
      console.log("Formule F25 installée avec succès");
    } catch (formuleError) {
      console.error("Erreur lors de l'installation de la formule F25:", formuleError);
    }
    
    // Attendre un moment pour que les calculs et l'optimisation se fassent dans la feuille
    Utilities.sleep(3000); // Attendre 3 secondes pour l'optimisation
    
    // Forcer le recalcul après optimisation
    SpreadsheetApp.flush();
    
    // Récupérer les résultats techniques (qui utilisent maintenant les chillers optimisés)
    var resultats = getResultatsTechniques();
    console.log("Résultats techniques récupérés:", resultats);
    
    // Vérifier si nous avons des résultats valides
    var hasValidResults = false;
    if (resultats && !resultats.error) {
      // Vérifier si au moins un des résultats principaux n'est pas zéro
      var totalSavings = parseFloat(resultats.total_savings) || 0;
      var chillersNewConsumption = parseFloat(resultats.chillers_new_consumption) || 0;
      
      if (totalSavings > 0 || chillersNewConsumption > 0) {
        hasValidResults = true;
      }
    }
    
    console.log("Résultats valides détectés:", hasValidResults);
    
    // Afficher un message de confirmation et renvoyer les résultats
    var message = 'Données sauvegardées avec succès dans le Google Sheet!';
    if (hasValidResults) {
      message += ' Optimisation réussie et résultats calculés.';
    } else {
      message += ' Note: Vérifiez que toutes les données nécessaires sont remplies pour voir les résultats d\'optimisation.';
    }
    
    return {
      success: true, 
      message: message,
      resultats: resultats,
      hasValidResults: hasValidResults
    };
    
  } catch (error) {
    console.error('Erreur lors de la sauvegarde:', error);
    return {success: false, message: 'Erreur lors de la sauvegarde: ' + error.toString()};
  }
}

/**
 * Fonction pour calculer automatiquement certaines valeurs
 */
function calculerCouts(donnees) {
  try {
    var couts = {
      coutMensuel: 0,
      coutSetup: 0,
      coutTotal: 0
    };
    
    // Sélectionner le mode de calcul approprié
    if (donnees.modeCalcul === 'simple') {
      return calculModeSimple(donnees);
    } else if (donnees.modeCalcul === 'avance') {
      // Mode avancé avec les nouveaux paramètres
      return calculModeAvance(donnees);
    } else {
      // Mode par défaut (ancienne logique)
      return calculModeDefault(donnees);
    }
    
    return couts;
    
  } catch (error) {
    console.error('Erreur lors du calcul:', error);
    return {coutMensuel: 0, coutSetup: 0, coutTotal: 0};
  }
}

/**
 * Fonction pour calculer selon le mode simple
 * Basé sur la logique du calculateur Veolia CaaS online
 */
function calculModeSimple(donnees) {
  var couts = {
    coutMensuel: 0,
    coutSetup: 0,
    coutTotal: 0
  };
  
  // Tarifs de base pour le mode simple (à adapter selon votre modèle)
  var tarifs = {
    Basic: { base: 99, utilisateur: 9 },
    Standard: { base: 199, utilisateur: 12 },
    Premium: { base: 299, utilisateur: 15 }
  };
  
  // Récupération du tarif selon le type de service
  var tarifService = tarifs[donnees.typeService] || tarifs.Basic;
  
  // Calcul du coût mensuel
  couts.coutMensuel = tarifService.base;
  
  // Ajout du coût par utilisateur supplémentaire
  if (donnees.nombreUtilisateurs > 1) {
    couts.coutMensuel += (parseInt(donnees.nombreUtilisateurs) - 1) * tarifService.utilisateur;
  }
  
  // Ajout des options
  if (donnees.supportTechnique) couts.coutMensuel += 99;
  if (donnees.monitoring) couts.coutMensuel += 79;
  if (donnees.backup) couts.coutMensuel += 49;
  if (donnees.securiteAvancee) couts.coutMensuel += 129;
  
  // Calcul du coût d'installation (setup)
  couts.coutSetup = couts.coutMensuel * 0.75; // 75% du coût mensuel
  
  // Calcul du coût total sur la durée du projet
  if (donnees.dureeProjet) {
    couts.coutTotal = (couts.coutMensuel * parseInt(donnees.dureeProjet)) + couts.coutSetup;
  }
  
  return couts;
}

/**
 * Fonction pour calculer selon le mode par défaut
 */
/**
 * Fonction pour calculer en mode avancé avec tous les paramètres complexes
 */
function calculModeAvance(donnees) {
  var couts = {
    coutMensuel: 0,
    coutSetup: 0,
    coutTotal: 0,
    weatherInfo: null,
    weatherImpact: null
  };
  
  // Facteurs de base selon le type de bâtiment
  var facteurTypeBatiment = {
    'Hotels': 1.2,
    'Airport': 1.5,
    'Hospital': 1.4,
    'Mall': 1.3,
    'Offices': 1.0,
    'Universities': 1.1,
    'CUSTOM': 1.25
  };
  
  // Facteur selon l'hémisphère (peut affecter les calculs CDD)
  var facteurHemisphere = {
    'North': 1.0,
    'South': 1.05
  };
  
  // Facteur CDD base
  var facteurCDD = {
    '15': 1.1,
    '16': 1.08,
    '17': 1.05,
    '18': 1.0,
    '19': 0.95,
    '20': 0.9
  };
  
  // Prix de base pour chaque type de service
  var prixBaseService = {
    'Basic': 150,
    'Standard': 300,
    'Premium': 450,
    'Enterprise': 600
  };
  
  // 1. Calcul du prix de base selon le type de service
  var prixBase = prixBaseService[donnees.typeService] || prixBaseService.Basic;
  
  // 2. Application du facteur selon le type de bâtiment
  var facteurBatiment = facteurTypeBatiment[donnees.typeBatiment] || 1.0;
  
  // 3. Application du facteur selon l'hémisphère
  var facteurHemi = facteurHemisphere[donnees.hemisphere] || 1.0;
  
  // 4. Application du facteur CDD
  var facteurCddBase = facteurCDD[donnees.cddBase] || 1.0;
  
  // 5. Récupération des données météo et calcul de l'impact
  var facteurMeteo = 1.0;
  if (donnees.includeWeatherData === true) {
    try {
      var weatherData = getWeatherImpactData(donnees);
      if (weatherData && !weatherData.error) {
        couts.weatherInfo = weatherData;
        facteurMeteo = weatherData.weatherImpact.overallFactor;
        couts.weatherImpact = weatherData.weatherImpact;
      }
    } catch (error) {
      console.log('Erreur lors de la récupération des données météo:', error);
      // Continuer sans les données météo
    }
  }
  
  // 6. Calcul de l'impact de la réduction nocturne si activée
  var facteurNightReduction = 1.0;
  if (donnees.nightReductionActive === true || donnees.nightReduction === true) {
    // Calcul basé sur le delta T et les heures
    var deltaT = parseFloat(donnees.nightReductionDeltaT || donnees.nightDeltaT || 8);
    facteurNightReduction = 1.0 - (deltaT * 0.01); // Réduction de 1% par degré de deltaT
  }
  
  // 7. Prise en compte des chillers (capacité totale)
  var capaciteTotaleChillers = 0;
  if (donnees.chiller1Capacity) capaciteTotaleChillers += parseFloat(donnees.chiller1Capacity);
  if (donnees.chiller2Capacity) capaciteTotaleChillers += parseFloat(donnees.chiller2Capacity);
  if (donnees.chiller3Capacity) capaciteTotaleChillers += parseFloat(donnees.chiller3Capacity);
  if (donnees.chiller1) capaciteTotaleChillers += parseFloat(donnees.chiller1);
  if (donnees.chiller2) capaciteTotaleChillers += parseFloat(donnees.chiller2);
  if (donnees.chiller3) capaciteTotaleChillers += parseFloat(donnees.chiller3);
  
  var facteurChillers = 1.0;
  if (capaciteTotaleChillers > 0) {
    // Plus la capacité est grande, plus le coût mensuel par kW diminue (économies d'échelle)
    facteurChillers = Math.max(0.7, 1.0 - (Math.log(capaciteTotaleChillers) / 20));
  }
  
  // 8. Calcul du coût mensuel de base avec le facteur météo
  couts.coutMensuel = prixBase * facteurBatiment * facteurHemi * facteurCddBase * facteurNightReduction * facteurChillers * facteurMeteo;
  
  // 9. Ajustement selon le nombre d'utilisateurs
  if (donnees.nombreUtilisateurs) {
    couts.coutMensuel += (parseInt(donnees.nombreUtilisateurs) - 1) * 12;
  }
  
  // 10. Ajout des services additionnels
  if (donnees.supportTechnique) couts.coutMensuel += 120;
  if (donnees.monitoring) couts.coutMensuel += 90;
  if (donnees.backup) couts.coutMensuel += 75;
  if (donnees.securiteAvancee) couts.coutMensuel += 150;
  
  // 11. Calcul du coût de setup (85% du coût mensuel en mode avancé)
  couts.coutSetup = couts.coutMensuel * 0.85;
  
  // 12. Calcul du coût total sur la durée du projet
  if (donnees.dureeProjet) {
    couts.coutTotal = (couts.coutMensuel * parseInt(donnees.dureeProjet)) + couts.coutSetup;
  }
  
  // Arrondir les valeurs pour plus de lisibilité
  couts.coutMensuel = Math.round(couts.coutMensuel);
  couts.coutSetup = Math.round(couts.coutSetup);
  couts.coutTotal = Math.round(couts.coutTotal);
  
  return couts;
}

function calculModeDefault(donnees) {
  var couts = {
    coutMensuel: 0,
    coutSetup: 0,
    coutTotal: 0
  };
  
  // Prix de base
  var basePrice = 100;
  
  // Calcul selon le type de service
  switch(donnees.typeService) {
    case 'Basic':
      couts.coutMensuel = basePrice;
      break;
    case 'Standard':
      couts.coutMensuel = basePrice * 2;
      break;
    case 'Premium':
      couts.coutMensuel = basePrice * 3;
      break;
    default:
      couts.coutMensuel = basePrice;
  }
  
  // Ajustement selon le nombre d'utilisateurs
  if (donnees.nombreUtilisateurs) {
    couts.coutMensuel += (parseInt(donnees.nombreUtilisateurs) - 1) * 10;
  }
  
  // Calcul du coût de setup
  couts.coutSetup = couts.coutMensuel * 0.5;
  
  // Calcul du coût total sur la durée
  if (donnees.dureeProjet) {
    couts.coutTotal = (couts.coutMensuel * parseInt(donnees.dureeProjet)) + couts.coutSetup;
  }
  
  return couts;
}

/**
 * Fonction pour inclure des fichiers CSS/JS dans le HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fonction pour initialiser les en-têtes de colonnes
 */
function initialiserEnTetes() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Création d'un tableau pour les en-têtes
    var entetes = [
      ["Paramètre", "Valeur", "Description"],
      ["Mode de calcul", "", "Simple ou Avancé"],
      ["Nom du projet", "", "Nom du projet CaaS"],
      ["Type de service", "", "Basic, Standard, Premium"],
      ["Durée du projet (mois)", "", "Durée du contrat"],
      ["Nombre d'utilisateurs", "", "Nombre d'utilisateurs du service"],
      ["Coût mensuel", "", "Coût mensuel calculé"],
      ["Coût de setup", "", "Coût d'installation initial"],
      ["Support technique", "", "Support technique 24/7"],
      ["", "", ""],
      ["Type de bâtiment", "", "Hotels, Office, etc."],
      ["Night Reduction active", "", "TRUE ou FALSE"],
      ["Night Start (Sem.)", "", "Heure de début semaine"],
      ["Night End (Sem.)", "", "Heure de fin semaine"],
      ["Night Start (Ven.)", "", "Heure de début vendredi"],
      ["Night End (Ven.)", "", "Heure de fin vendredi"],
      ["Night Start (WE)", "", "Heure de début weekend"],
      ["Night End (WE)", "", "Heure de fin weekend"],
      ["Night ΔT", "", "Delta T de nuit en °C"],
      ["Hémisphère", "", "North ou South"],
      ["CDD base", "", "Température de base en °C"],
      ["Chiller 1 (kWc)", "", "Capacité du chiller 1"],
      ["Chiller 2 (kWc)", "", "Capacité du chiller 2"],
      ["Chiller 3 (kWc)", "", "Capacité du chiller 3"],
      ["Monitoring", "", "Monitoring avancé"],
      ["Backup", "", "Sauvegarde automatique"],
      ["Sécurité avancée", "", "Sécurité et conformité"],
      ["Météo activée", "", "Utilisation des données météo"],
      ["Nom du lieu", "", "Nom du lieu pour les données météo"],
      ["Coordonnées GPS", "", "Latitude,Longitude pour l'API météo"]
    ];
    
    // Application des en-têtes
    sheet.getRange(1, 1, entetes.length, 3).setValues(entetes);
    
    // Mise en forme
    sheet.getRange(1, 1, 1, 3).setBackground("#4285F4").setFontColor("white").setFontWeight("bold");
    sheet.getRange(2, 1, entetes.length-1, 1).setBackground("#E8F0FE");
    sheet.getRange(10, 1, 1, 3).setBackground("#F1F3F4").setFontWeight("bold");
    
    // Ajuster la largeur des colonnes
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250);
    
    return {success: true, message: "En-têtes initialisés avec succès!"};
  } catch (error) {
    console.error("Erreur lors de l'initialisation des en-têtes:", error);
    return {success: false, message: "Erreur: " + error.toString()};
  }
}

/**
 * Fonction pour récupérer les données météo Veolia avec les vraies clés API
 */
function getVeoliaWeatherData(params) {
  try {
    console.log('Appel API Veolia avec les paramètres:', params);
    
    // Utiliser les vraies clés API fournies
    const API_KEY = 'IOd00pXJRxXcpFFqUySU8tyIGkdSeFoC';
    const PRIVATE_KEY = 'LWGYE563U0cqIZa7';
    
    // Construire le geocode à partir des coordonnées
    const geocode = params.latitude + ',' + params.longitude;
    
    // 1. Obtenir le token d'authentification
    const token = getVeoliaTokenBearer(API_KEY, PRIVATE_KEY);
    if (!token) {
      throw new Error('Impossible d\'obtenir le token d\'authentification');
    }
    
    console.log('Token obtenu avec succès');
    
    // 2. Construire l'URL selon le type de données demandé
    let apiUrl = '';
    const baseUrl = 'https://api.veolia.com/weather/v1';
    
    switch (params.dataType) {
      case 'current':
        apiUrl = baseUrl + '/current/ondemand';
        break;
      case 'forecast':
        if (params.timeInterval === 'hourly') {
          apiUrl = baseUrl + '/forecast/hourly';
        } else if (params.timeInterval === 'daily') {
          apiUrl = baseUrl + '/forecast/daily';
        } else {
          apiUrl = baseUrl + '/forecast/15minutes';
        }
        break;
      case 'history':
        apiUrl = baseUrl + '/history/general';
        break;
      case 'degreedays':
        apiUrl = baseUrl + '/history/degreedays/calcul';
        break;
      default:
        apiUrl = baseUrl + '/current/ondemand';
    }
    
    // 3. Ajouter les paramètres de requête
    const queryParams = [
      'geocode=' + encodeURIComponent(geocode),
      'units=m',
      'language=en-US',  // Changé de fr-FR à en-US pour éviter certaines erreurs
      'format=json'
    ];
    
    // Ajouter les dates pour les données historiques - CORRECTION DES DATES
    if (params.dataType === 'history' && params.startDate && params.endDate) {
      // Vérifier que les dates ne sont pas trop anciennes ou trop éloignées
      const startDate = new Date(params.startDate);
      const endDate = new Date(params.endDate);
      const now = new Date();
      
      // Limiter la période à maximum 1 an
      const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
      const maxEndDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
      
      // Ajuster les dates si nécessaire
      let adjustedStartDate = startDate;
      let adjustedEndDate = endDate;
      
      if (startDate < oneYearAgo) {
        adjustedStartDate = oneYearAgo;
        console.log('Date de début ajustée à un an maximum:', adjustedStartDate.toISOString().split('T')[0]);
      }
      
      if (endDate > maxEndDate) {
        adjustedEndDate = maxEndDate;
        console.log('Date de fin ajustée à hier maximum:', adjustedEndDate.toISOString().split('T')[0]);
      }
      
      // Vérifier que la période ne dépasse pas 1 an
      const timeDiff = adjustedEndDate.getTime() - adjustedStartDate.getTime();
      const daysDiff = timeDiff / (1000 * 3600 * 24);
      
      if (daysDiff > 365) {
        adjustedStartDate = new Date(adjustedEndDate.getTime() - (365 * 24 * 60 * 60 * 1000));
        console.log('Période réduite à 365 jours maximum');
      }
      
      queryParams.push('startDate=' + adjustedStartDate.toISOString().split('T')[0]);
      queryParams.push('endDate=' + adjustedEndDate.toISOString().split('T')[0]);
      
      console.log('Dates utilisées:', {
        startDate: adjustedStartDate.toISOString().split('T')[0],
        endDate: adjustedEndDate.toISOString().split('T')[0],
        daysDiff: Math.round(daysDiff)
      });
    }
    
    const fullUrl = apiUrl + '?' + queryParams.join('&');
    console.log('URL de l\'API:', fullUrl);
    
    // 4. Effectuer l'appel API avec gestion d'erreur améliorée
    const headers = {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    };
    
    const options = {
      'method': 'GET',
      'headers': headers,
      'muteHttpExceptions': true  // Ajouté pour voir la réponse complète en cas d'erreur
    };
    
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('Code de réponse:', responseCode);
    console.log('Réponse complète:', responseText);
    
    if (responseCode !== 200) {
      // Gestion spécifique des erreurs 400
      if (responseCode === 400) {
        let errorDetails = 'Erreur 400 - Paramètres invalides';
        try {
          const errorData = JSON.parse(responseText);
          if (errorData.reason) {
            errorDetails += ': ' + errorData.reason;
          }
          if (errorData.message) {
            errorDetails += ' - ' + errorData.message;
          }
        } catch (e) {
          errorDetails += ': ' + responseText.substring(0, 200);
        }
        throw new Error(errorDetails);
      } else {
        throw new Error('Erreur API: ' + responseCode + ' - ' + responseText);
      }
    }
    
    const data = JSON.parse(responseText);
    
    // 5. Traiter et retourner les données
    const result = {
      success: true,
      dataType: params.dataType,
      timeInterval: params.timeInterval,
      location: {
        latitude: params.latitude,
        longitude: params.longitude,
        name: params.locationName || params.locationAddress
      },
      data: data,
      recommendations: generateWeatherRecommendations(data, params),
      timestamp: new Date().toISOString()
    };
    
    console.log('Données météo traitées avec succès');
    return result;
    
  } catch (error) {
    console.error('Erreur lors de l\'appel API Veolia:', error);
    return {
      success: false,
      error: error.toString(),
      message: 'Erreur lors de la récupération des données météo: ' + error.toString()
    };
  }
}

/**
 * Fonction pour obtenir le token d'authentification Veolia
 */
function getVeoliaTokenBearer(apiKey, privateKey) {
  try {
    const TOKEN_URL = 'https://api.veolia.com/security/v2/oauth/token';
    
    // Encoder les identifiants en Base64
    const credentials = Utilities.base64Encode(apiKey + ':' + privateKey);
    
    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': 'Basic ' + credentials
    };
    
    const payload = 'grant_type=client_credentials';
    
    const options = {
      'method': 'POST',
      'headers': headers,
      'payload': payload
    };
    
    const response = UrlFetchApp.fetch(TOKEN_URL, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && responseData.access_token) {
      return responseData.access_token;
    } else {
      console.error('Erreur lors de l\'obtention du token:', responseData);
      return null;
    }
    
  } catch (error) {
    console.error('Erreur lors de l\'authentification:', error);
    return null;
  }
}

/**
 * Fonction utilitaire pour tester le téléchargement des données météo Veolia
 * Peut être appelée manuellement depuis l'éditeur de script Google Apps Script
 */
function testVeoliaWeatherDownload() {
  try {
    console.log('=== TEST TÉLÉCHARGEMENT DONNÉES MÉTÉO VEOLIA ===');
    
    // Paramètres de test (Paris par défaut)
    const testParams = {
      latitude: '48.8566',
      longitude: '2.3522',
      locationName: 'Paris, France',
      locationAddress: 'Paris, France',
      dataType: 'current',
      timeInterval: 'hourly'
    };
    
    console.log('Paramètres de test:', testParams);
    
    // Tester le téléchargement
    const result = downloadAndSaveVeoliaWeatherData(testParams);
    
    if (result.success) {
      console.log('✅ TEST RÉUSSI - Données téléchargées avec succès');
      console.log('Message:', result.message);
      
      // Afficher quelques recommandations si disponibles
      if (result.data && result.data.recommendations) {
        console.log('📊 Recommandations:');
        result.data.recommendations.forEach((rec, i) => {
          console.log(`${i + 1}. ${rec}`);
        });
      }
      
      // Vérifier que la feuille a été créée
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const weatherSheet = ss.getSheetByName('Weather Data');
      if (weatherSheet) {
        console.log('✅ Feuille "Weather Data" créée avec succès');
        console.log(`📊 Données dans ${weatherSheet.getLastRow()} lignes`);
      } else {
        console.log('❌ Feuille "Weather Data" non trouvée');
      }
      
    } else {
      console.log('❌ TEST ÉCHOUÉ:', result.message);
      console.log('Erreur:', result.error);
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ ERREUR LORS DU TEST:', error);
    return {
      success: false,
      message: 'Erreur lors du test: ' + error.toString(),
      error: error.toString()
    };
  }
}

/**
 * Fonction pour tester différents types de données météo
 */
function testAllWeatherDataTypes() {
  const locations = [
    { name: 'Paris, France', lat: '48.8566', lng: '2.3522' },
    { name: 'New York, USA', lat: '40.7128', lng: '-74.0060' },
    { name: 'Tokyo, Japan', lat: '35.6762', lng: '139.6503' }
  ];
  
  const dataTypes = ['current', 'forecast'];
  
  console.log('=== TEST COMPLET DONNÉES MÉTÉO VEOLIA ===');
  
  locations.forEach(location => {
    dataTypes.forEach(dataType => {
      console.log(`\n--- Test ${location.name} - Type: ${dataType} ---`);
      
      const params = {
        latitude: location.lat,
        longitude: location.lng,
        locationName: location.name,
        dataType: dataType,
        timeInterval: 'hourly'
      };
      
      try {
        const result = getVeoliaWeatherData(params);
        if (result.success) {
          console.log('✅ Succès pour', location.name, '-', dataType);
        } else {
          console.log('❌ Échec pour', location.name, '-', dataType, ':', result.message);
        }
      } catch (error) {
        console.log('❌ Erreur pour', location.name, '-', dataType, ':', error.toString());
      }
    });
  });
}

/**
 * Fonction pour nettoyer la feuille Weather Data (utile pour les tests)
 */
function clearWeatherDataSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Weather Data');
    
    if (sheet) {
      sheet.clear();
      console.log('✅ Feuille "Weather Data" nettoyée');
      return { success: true, message: 'Feuille nettoyée avec succès' };
    } else {
      console.log('⚠️ Feuille "Weather Data" non trouvée');
      return { success: false, message: 'Feuille non trouvée' };
    }
  } catch (error) {
    console.error('❌ Erreur lors du nettoyage:', error);
    return { success: false, message: 'Erreur: ' + error.toString() };
  }
}
function generateWeatherRecommendations(weatherData, params) {
  const recommendations = [];
  
  try {
    // Informations générales sur l'API et les données
    recommendations.push('Data retrieved from Veolia Weather Data Platform API');
    recommendations.push('Base URL: https://api.veolia.com/weather/v1/' + params.dataType);
    recommendations.push('Authentication: OAuth 2.0 Bearer Token');
    recommendations.push('Units: Metric system (°C, km/h, mm, hPa)');
    
    // Analyser les données selon le type
    if (params.dataType === 'current' && weatherData.observations && weatherData.observations.length > 0) {
      const obs = weatherData.observations[0];
      if (obs) {
        recommendations.push('Current weather conditions analysis:');
        
        if (obs.temperature > 30) {
          recommendations.push('• High temperature (' + obs.temperature + '°C): Increased cooling demand expected');
        } else if (obs.temperature < 5) {
          recommendations.push('• Low temperature (' + obs.temperature + '°C): Reduced cooling demand expected');
        } else {
          recommendations.push('• Moderate temperature (' + obs.temperature + '°C): Normal cooling demand expected');
        }
        
        if (obs.humidity > 70) {
          recommendations.push('• High humidity (' + obs.humidity + '%): HVAC systems may experience increased load');
        } else if (obs.humidity < 30) {
          recommendations.push('• Low humidity (' + obs.humidity + '%): Optimal conditions for HVAC efficiency');
        }
        
        if (obs.windSpeed > 20) {
          recommendations.push('• High wind speed (' + obs.windSpeed + ' km/h): May affect building heat loss calculations');
        }
        
        if (obs.precipitation > 0) {
          recommendations.push('• Precipitation detected (' + obs.precipitation + ' mm): Consider moisture impact on building envelope');
        }
      }
    }
    
    if (params.dataType === 'forecast') {
      const forecasts = weatherData.forecasts || weatherData;
      if (Array.isArray(forecasts) && forecasts.length > 0) {
        // Analyser les tendances
        const temps = forecasts.map(f => f.temperature || f.temp).filter(t => t !== undefined && t !== '');
        const humidities = forecasts.map(f => f.humidity || f.relativeHumidity).filter(h => h !== undefined && h !== '');
        
        if (temps.length > 0) {
          const avgTemp = temps.reduce((sum, t) => sum + parseFloat(t), 0) / temps.length;
          const maxTemp = Math.max(...temps.map(t => parseFloat(t)));
          const minTemp = Math.min(...temps.map(t => parseFloat(t)));
          
          recommendations.push('Forecast analysis for ' + forecasts.length + ' periods:');
          recommendations.push('• Average temperature: ' + avgTemp.toFixed(1) + '°C');
          recommendations.push('• Temperature range: ' + minTemp.toFixed(1) + '°C to ' + maxTemp.toFixed(1) + '°C');
          
          if (maxTemp > 28) {
            recommendations.push('• Peak cooling demand expected (max: ' + maxTemp.toFixed(1) + '°C)');
          }
          
          if (minTemp < 10) {
            recommendations.push('• Potential heating requirements (min: ' + minTemp.toFixed(1) + '°C)');
          }
        }
        
        if (humidities.length > 0) {
          const avgHumidity = humidities.reduce((sum, h) => sum + parseFloat(h), 0) / humidities.length;
          recommendations.push('• Average humidity: ' + avgHumidity.toFixed(1) + '%');
          
          if (avgHumidity > 65) {
            recommendations.push('• High humidity period: Plan for increased dehumidification needs');
          }
        }
      }
    }
    
    if (params.dataType === 'history') {
      recommendations.push('Historical weather data analysis completed');
      if (params.startDate && params.endDate) {
        recommendations.push('Period analyzed: ' + params.startDate + ' to ' + params.endDate);
        
        // Calculer la durée de la période
        const start = new Date(params.startDate);
        const end = new Date(params.endDate);
        const daysDiff = Math.ceil((end - start) / (1000 * 60 * 60 * 24));
        recommendations.push('Analysis period: ' + daysDiff + ' days of historical data');
        
        if (daysDiff > 300) {
          recommendations.push('• Long-term trend analysis suitable for annual energy planning');
        } else if (daysDiff > 30) {
          recommendations.push('• Medium-term analysis suitable for seasonal planning');
        } else {
          recommendations.push('• Short-term analysis suitable for operational optimization');
        }
      }
    }
    
    // Recommandations sur la source des données
    const dataSource = params.dataSource || 'cfsr';
    if (dataSource === 'cfsr') {
      recommendations.push('Data source: CFSR virtual grid point (0.25° resolution, ~25km accuracy)');
      recommendations.push('• Grid-based data optimal for building energy calculations');
    } else {
      recommendations.push('Data source: Metar/Airport weather station data');
      recommendations.push('• Point-specific data with high local accuracy');
    }
    
    // Recommandations générales sur l'utilisation
    recommendations.push('Weather Data Platform Integration:');
    recommendations.push('• Data automatically integrated with CaaS cooling calculations');
    recommendations.push('• Use temperature data for degree-day calculations');
    recommendations.push('• Consider humidity impact on apparent temperature and comfort');
    recommendations.push('• Monitor wind conditions for natural ventilation opportunities');
    
    // Recommandations techniques
    recommendations.push('Technical recommendations:');
    recommendations.push('• Refresh weather data daily for current conditions');
    recommendations.push('• Use historical data for baseline energy consumption modeling');
    recommendations.push('• Combine forecast data with building schedules for optimal planning');
    
  } catch (error) {
    console.error('Erreur lors de la génération des recommandations:', error);
    recommendations.push('Error in weather analysis: ' + error.toString());
    recommendations.push('Raw data available in JSON section for manual analysis');
  }
  
  return recommendations;
}

/**
 * Fonction pour récupérer les données météo actuelles
 */
function getWeatherData(geocode, client_id, client_secret) {
  try {
    // Obtenir le token d'authentification
    var token = getTokenBearer(client_id, client_secret);
    if (!token) {
      return { error: 'Impossible d\'obtenir le token d\'authentification' };
    }
    
    // URL de l'API météo
    var API_URL = "https://api.veolia.com/weather/v1/current/ondemand?geocode=" + geocode + "&units=m&language=fr-FR&format=json";
    
    // Définir les en-têtes
    var header_weather = {
      "Authorization": "Bearer " + token
    };
    
    // Définir les options
    var options_weather = {
      'method': 'GET',
      'headers': header_weather
    };
    
    // Effectuer l'appel API
    var response = UrlFetchApp.fetch(API_URL, options_weather);
    var data = JSON.parse(response.getContentText());
    
    return data;
  } catch (error) {
    console.error('Erreur lors de la récupération des données météo:', error);
    return { error: 'Erreur lors de la récupération des données météo: ' + error.toString() };
  }
}

/**
 * Fonction pour récupérer les données de prévision horaire
 */
function getWeatherForecast(geocode, client_id, client_secret) {
  try {
    // Obtenir le token d'authentification
    var token = getTokenBearer(client_id, client_secret);
    if (!token) {
      return { error: 'Impossible d\'obtenir le token d\'authentification' };
    }
    
    // URL de l'API météo pour les prévisions
    var API_URL = "https://api.veolia.com/weather/v1/forecast/hourly?geocode=" + geocode + "&units=m&language=fr-FR&format=json";
    
    // Définir les en-têtes
    var header_weather = {
      "Authorization": "Bearer " + token
    };
    
    // Définir les options
    var options_weather = {
      'method': 'GET',
      'headers': header_weather
    };
    
    // Effectuer l'appel API
    var response = UrlFetchApp.fetch(API_URL, options_weather);
    var data = JSON.parse(response.getContentText());
    
    return data;
  } catch (error) {
    console.error('Erreur lors de la récupération des prévisions météo:', error);
    return { error: 'Erreur lors de la récupération des prévisions météo: ' + error.toString() };
  }
}

/**
 * Fonction pour calculer les Cooling Degree Days à partir des données météo
 */
function calculateCoolingDegreeDays(weatherData, baseTemp) {
  if (!weatherData || weatherData.error) {
    return 0;
  }
  
  try {
    var cdd = 0;
    var baseTemperature = parseFloat(baseTemp) || 18;
    
    // Si on a des données de prévision
    if (weatherData.forecasts && weatherData.forecasts.length > 0) {
      weatherData.forecasts.forEach(function(forecast) {
        var temp = parseFloat(forecast.temperature) || 0;
        if (temp > baseTemperature) {
          cdd += (temp - baseTemperature);
        }
      });
    }
    // Si on a des données actuelles
    else if (weatherData.observations && weatherData.observations.length > 0) {
      var currentTemp = parseFloat(weatherData.observations[0].temperature) || 0;
      if (currentTemp > baseTemperature) {
        cdd = currentTemp - baseTemperature;
      }
    }
    
    return Math.round(cdd * 10) / 10; // Arrondi à 1 décimale
  } catch (error) {
    console.error('Erreur lors du calcul des CDD:', error);
    return 0;
  }
}

/**
 * Fonction pour récupérer les données météo et calculer les impacts
 */
function getWeatherImpactData(donnees) {
  try {
    // Coordonnées par défaut (Paris) si pas de géocode fourni
    var geocode = donnees.geocode || "48.8566,2.3522";
    
    // Identifiants API (à configurer dans les propriétés du script)
    var client_id = PropertiesService.getScriptProperties().getProperty('VEOLIA_CLIENT_ID');
    var client_secret = PropertiesService.getScriptProperties().getProperty('VEOLIA_CLIENT_SECRET');
    
    if (!client_id || !client_secret) {
      return {
        error: 'Identifiants API non configurés',
        message: 'Veuillez configurer VEOLIA_CLIENT_ID et VEOLIA_CLIENT_SECRET dans les propriétés du script'
      };
    }
    
    // Récupérer les données météo actuelles
    var currentWeather = getWeatherData(geocode, client_id, client_secret);
    
    // Récupérer les prévisions
    var forecast = getWeatherForecast(geocode, client_id, client_secret);
    
    // Calculer les CDD
    var cdd = calculateCoolingDegreeDays(forecast, donnees.cddBase);
    
    // Calculer l'impact météo sur le coût
    var weatherImpact = calculateWeatherImpact(currentWeather, cdd, donnees);
    
    return {
      currentWeather: currentWeather,
      forecast: forecast,
      cdd: cdd,
      weatherImpact: weatherImpact
    };
  } catch (error) {
    console.error('Erreur lors de la récupération des données météo:', error);
    return {
      error: 'Erreur lors de la récupération des données météo: ' + error.toString()
    };
  }
}

/**
 * Fonction pour calculer l'impact météo sur les coûts
 */
function calculateWeatherImpact(weatherData, cdd, donnees) {
  try {
    var impact = {
      temperatureFactor: 1.0,
      humidityFactor: 1.0,
      overallFactor: 1.0,
      recommendations: []
    };
    
    if (weatherData && weatherData.observations && weatherData.observations.length > 0) {
      var obs = weatherData.observations[0];
      var temp = parseFloat(obs.temperature) || 20;
      var humidity = parseFloat(obs.humidity) || 50;
      
      // Facteur de température
      if (temp > 30) {
        impact.temperatureFactor = 1.3;
        impact.recommendations.push("Température élevée : augmentation de 30% des besoins de refroidissement");
      } else if (temp > 25) {
        impact.temperatureFactor = 1.15;
        impact.recommendations.push("Température modérée : augmentation de 15% des besoins de refroidissement");
      } else if (temp > 20) {
        impact.temperatureFactor = 1.0;
      } else {
        impact.temperatureFactor = 0.85;
        impact.recommendations.push("Température fraîche : réduction de 15% des besoins de refroidissement");
      }
      
      // Facteur d'humidité
      if (humidity > 80) {
        impact.humidityFactor = 1.2;
        impact.recommendations.push("Humidité élevée : augmentation de 20% de la charge de climatisation");
      } else if (humidity > 60) {
        impact.humidityFactor = 1.1;
        impact.recommendations.push("Humidité modérée : augmentation de 10% de la charge de climatisation");
      } else {
        impact.humidityFactor = 1.0;
      }
      
      // CDD impact
      var cddImpact = 1.0;
      if (cdd > 10) {
        cddImpact = 1.25;
        impact.recommendations.push("CDD élevé (" + cdd + ") : augmentation significative des besoins de refroidissement");
      } else if (cdd > 5) {
        cddImpact = 1.1;
        impact.recommendations.push("CDD modéré (" + cdd + ") : augmentation modérée des besoins de refroidissement");
      }
      
      // Facteur global
      impact.overallFactor = impact.temperatureFactor * impact.humidityFactor * cddImpact;
    }
    
    return impact;
  } catch (error) {
    console.error('Erreur lors du calcul de l\'impact météo:', error);
    return {
      temperatureFactor: 1.0,
      humidityFactor: 1.0,
      overallFactor: 1.0,
      recommendations: ['Erreur lors du calcul de l\'impact météo']
    };
  }
}

/**
 * Fonction pour convertir les dates au format MM-YYYY en format YYYY-MM pour l'élément input type="month"
 */
function convertirEnFormatMois(dateStr) {
  if (!dateStr) return null;
  
  // Essayons différents formats possibles
  
  // Format MM-YYYY
  if (/^\d{1,2}-\d{4}$/.test(dateStr)) {
    const [mois, annee] = dateStr.split('-');
    const moisPadded = mois.padStart(2, '0');
    return `${annee}-${moisPadded}`;
  }
  
  // Format déjà au format YYYY-MM
  if (/^\d{4}-\d{1,2}$/.test(dateStr)) {
    const [annee, mois] = dateStr.split('-');
    const moisPadded = mois.padStart(2, '0');
    return `${annee}-${moisPadded}`;
  }
  
  // Si c'est une date avec jour (DD-MM-YYYY ou YYYY-MM-DD)
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(dateStr)) {
    const parts = dateStr.split('-');
    return `${parts[2]}-${parts[1].padStart(2, '0')}`;
  }
  
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(dateStr)) {
    const parts = dateStr.split('-');
    return `${parts[0]}-${parts[1].padStart(2, '0')}`;
  }
  
  // Si c'est un objet Date
  if (dateStr instanceof Date) {
    const annee = dateStr.getFullYear();
    const mois = (dateStr.getMonth() + 1).toString().padStart(2, '0');
    return `${annee}-${mois}`;
  }
  
  // Par défaut, retourner null
  return null;
}

/**
 * Fonction pour convertir les valeurs d'électricité avec séparateurs en nombres purs
 */
function convertirEnNombre(valeurStr) {
  if (!valeurStr) return 0;
  
  // Supprimer les espaces, les virgules et autres séparateurs
  const valeurNettoyee = valeurStr.toString().replace(/[^\d.-]/g, '');
  
  // Convertir en nombre
  const nombre = parseFloat(valeurNettoyee);
  
  // Vérifier si la conversion a réussi
  if (isNaN(nombre)) {
    return 0;
  }
  
  return nombre;
}

/**
 * Fonction pour gérer le clic sur les cellules de bouton
 * Cette fonction sera associée à un déclencheur d'édition
 */
function onEdit(e) {
  try {
    if (e && e.range && e.source) {
      var sheet = e.source.getActiveSheet();
      var range = e.range;
      
      // Vérifier si l'édition concerne la cellule L11 de la feuille "0 - Read me"
      if (range.getA1Notation() === "L11" && sheet.getName() === "0 - Read me") {
        // Animation de clic sur le bouton
        var cellule = range;
        var valeurOriginale = cellule.getValue();
        var couleurOriginale = cellule.getBackground();
        
        // Effet visuel de clic
        cellule.setValue("CHARGEMENT...");
        cellule.setBackground("#388E3C"); // Vert plus foncé
        
        // Créer un déclencheur de feuille de calcul pour l'utilisateur actif
        afficherMessageOuverture();
        
        // Attendre un court instant pour l'effet visuel
        Utilities.sleep(500);
        
        // Remettre la valeur et la couleur originales
        cellule.setValue(valeurOriginale || "OUVRIR CALCULATEUR");
        cellule.setBackground(couleurOriginale || "#4CAF50");
        
        // Créer une feuille cachée pour stocker un indicateur
        var indiquerOuverture = true;
        try {
          // Tenter de créer une feuille cachée pour stocker l'information
          var feuilleIndicateur = e.source.insertSheet("_indicateurOuverture");
          feuilleIndicateur.hideSheet();
          feuilleIndicateur.getRange("A1").setValue("ouvrir");
        } catch (errFeuille) {
          // La feuille existe probablement déjà
          try {
            var feuilleIndicateur = e.source.getSheetByName("_indicateurOuverture");
            if (feuilleIndicateur) {
              feuilleIndicateur.getRange("A1").setValue("ouvrir");
            }
          } catch (errAcces) {
            // Ignorer les erreurs d'accès
          }
        }
      }
    }
  } catch (error) {
    console.log("Erreur dans le gestionnaire d'événement onEdit: " + error);
  }
}

/**
 * Fonction pour afficher un message indiquant comment ouvrir le calculateur
 */
function afficherMessageOuverture() {
  try {
    var ui = SpreadsheetApp.getUi();
    var reponse = ui.alert(
      'Calculateur CaaS',
      'Vous avez cliqué sur le bouton d\'ouverture du calculateur.\n\n' +
      'En raison des limitations de sécurité de Google Sheets, nous ne pouvons pas ouvrir directement la fenêtre à partir d\'un clic sur une cellule.\n\n' +
      'Pour ouvrir le calculateur, veuillez utiliser le menu "Calculateur CaaS" > "Ouvrir la fenêtre de saisie" en haut de l\'écran.',
      ui.ButtonSet.OK
    );
    
    if (reponse === ui.Button.OK) {
      // Vérifier si l'utilisateur a l'autorisation d'exécuter la fonction
      try {
        // Tentative d'ouverture directe avec un petit délai
        Utilities.sleep(500);
        ouvrirFenetreCalculateur();
      } catch (errAuth) {
        // Les restrictions de sécurité ont empêché l'ouverture automatique
        console.log("Impossible d'ouvrir automatiquement la fenêtre: " + errAuth);
      }
    }
  } catch (error) {
    console.log("Erreur lors de l'affichage du message: " + error);
  }
}

/**
 * Fonction pour vérifier périodiquement s'il faut ouvrir le calculateur
 * Cette fonction est appelée régulièrement par un déclencheur temporel
 */
function verifierOuvertureRequise() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Vérifier si la feuille d'indicateur existe
    var feuilleIndicateur = ss.getSheetByName("_indicateurOuverture");
    if (feuilleIndicateur) {
      var valeur = feuilleIndicateur.getRange("A1").getValue();
      
      if (valeur === "ouvrir") {
        // Réinitialiser l'indicateur
        feuilleIndicateur.getRange("A1").setValue("");
        
        // Tenter d'ouvrir le calculateur
        try {
          ouvrirFenetreCalculateur();
        } catch (errOuverture) {
          console.log("Impossible d'ouvrir automatiquement la fenêtre via le vérificateur: " + errOuverture);
        }
      }
    }
  } catch (error) {
    console.log("Erreur lors de la vérification d'ouverture requise: " + error);
  }
}

/**
 * Fonction pour installer les déclencheurs nécessaires
 * À exécuter manuellement une fois pour configurer les déclencheurs
 * Maintenant avec une interface utilisateur pour l'installation
 */
function installerDeclencheurs() {
  try {
    // Supprimer les déclencheurs existants pour éviter les doublons
    var declencheurs = ScriptApp.getProjectTriggers();
    for (var i = 0; i < declencheurs.length; i++) {
      ScriptApp.deleteTrigger(declencheurs[i]);
    }
    
    // 1. Créer un déclencheur pour la fonction onEdit (pour les interactions cellules)
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    // 2. Configurer l'interface utilisateur d'accueil qui s'affiche à chaque ouverture
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
    
    // 3. Créer un déclencheur pour l'ouverture automatique à l'ouverture de la feuille
    // Cette fonction s'exécutera à chaque ouverture du document
    ScriptApp.newTrigger('ouvrirPopupAutomatique')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
    
    // 4. Créer un déclencheur qui vérifie périodiquement s'il faut ouvrir le calculateur
    ScriptApp.newTrigger('verifierOuvertureRequise')
      .timeBased()
      .everyMinutes(1)
      .create();
    
    // Afficher un message de confirmation à l'utilisateur
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'Installation réussie',
      'Les déclencheurs ont été installés avec succès.\n\n' +
      'Maintenant, à chaque ouverture de la feuille de calcul, le calculateur s\'ouvrira automatiquement.\n\n' +
      'Vous pouvez également l\'ouvrir à tout moment via le menu "Calculateur CaaS" > "Ouvrir la fenêtre de saisie".',
      ui.ButtonSet.OK
    );
    
    // Ouvrir directement la fenêtre après l'installation des déclencheurs
    ouvrirFenetreCalculateur();
    
    return {success: true, message: "Tous les déclencheurs ont été installés avec succès"};
  } catch (error) {
    console.log("Erreur lors de l'installation des déclencheurs: " + error);
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'Erreur',
      'Une erreur s\'est produite lors de l\'installation des déclencheurs: ' + error.toString() + '\n\n' +
      'Veuillez réessayer ou contacter l\'administrateur.',
      ui.ButtonSet.OK
    );
    return {success: false, message: "Erreur: " + error.toString()};
  }
}

/**
 * Fonction spécifique pour ouvrir le popup automatiquement à l'ouverture du document
 * Cette fonction est appelée par un déclencheur onOpen configuré via installerDeclencheurs()
 */
function ouvrirPopupAutomatique() {
  try {
    // Déterminer si l'utilisateur a déjà autorisé le script
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    
    if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      // L'utilisateur n'a pas encore autorisé le script, on ne peut pas ouvrir le popup automatiquement
      // On lui propose d'installer les autorisations
      var ui = SpreadsheetApp.getUi();
      var reponse = ui.alert(
        'Calculateur CaaS - Autorisation requise',
        'Pour permettre l\'ouverture automatique du calculateur CaaS, ' +
        'vous devez autoriser ce script à s\'exécuter.\n\n' +
        'Souhaitez-vous autoriser le script maintenant ?',
        ui.ButtonSet.YES_NO
      );
      
      if (reponse === ui.Button.YES) {
        // Tenter d'ouvrir le calculateur pour déclencher l'autorisation
        ouvrirFenetreCalculateur();
      }
    } else {
      // L'utilisateur a déjà autorisé le script, on peut ouvrir le popup directement
      // Ajouter un délai pour s'assurer que l'interface est prête
      Utilities.sleep(1000);
      ouvrirFenetreCalculateur();
    }
  } catch (error) {
    console.log("Erreur lors de l'ouverture automatique du popup: " + error);
    // En cas d'erreur, ne pas afficher de message à l'utilisateur pour ne pas perturber l'expérience
  }
}

/**
 * Fonction utilitaire pour obtenir et afficher l'URL complète du script
 * Exécutez cette fonction pour voir l'URL à utiliser dans la formule HYPERLINK
 */
function afficherURLScript() {
  var scriptURL = ScriptApp.getService().getUrl();
  console.log("URL du script à utiliser dans la formule HYPERLINK :");
  console.log(scriptURL + "?functionName=ouvrirFenetreDirecte");
  
  // Créer également une boîte de dialogue avec l'information
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'URL du script pour HYPERLINK',
    'Copiez cette URL pour votre formule HYPERLINK :\n\n' +
    scriptURL + '?functionName=ouvrirFenetreDirecte\n\n' +
    'Formule complète :\n' +
    '=HYPERLINK("' + scriptURL + '?functionName=ouvrirFenetreDirecte", "OUVRIR CALCULATEUR")',
    ui.ButtonSet.OK
  );
  
  return scriptURL + "?functionName=ouvrirFenetreDirecte";
}

/**
 * Fonction pour créer un bouton d'interface utilisateur réel qui ouvre la popup
 */
function creerBoutonUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('0 - Read me');
  
  if (!sheet) {
    sheet = ss.insertSheet('0 - Read me');
  }
  
  // Ajouter une instruction pour utiliser le bouton
  sheet.getRange("B5").setValue("Comment ouvrir le calculateur CaaS :");
  sheet.getRange("B5").setFontWeight("bold");
  sheet.getRange("B5").setFontSize(14);
  
  // Méthode 1: Par le menu
  sheet.getRange("B7").setValue("Méthode 1: Via le menu en haut");
  sheet.getRange("B7").setFontWeight("bold");
  sheet.getRange("B8").setValue("Cliquez sur le menu \"Calculateur CaaS\" puis sur \"Ouvrir la fenêtre de saisie\"");
  
  // Méthode 2: Via les boutons de dessins insérés manuellement
  sheet.getRange("B10").setValue("Méthode 2: Via les boutons ci-dessous");
  sheet.getRange("B10").setFontWeight("bold");
  sheet.getRange("B11").setValue("Cliquez sur l'un des boutons ci-dessous (vous devrez autoriser le script)");
  
  // Instructions pour insérer des boutons de dessin
  sheet.getRange("B14").setValue("Pour ajouter des boutons fonctionnels qui ouvrent le calculateur:");
  sheet.getRange("B15").setValue("1. Menu \"Insertion\" > \"Dessin\"");
  sheet.getRange("B16").setValue("2. Dessinez un bouton et ajoutez-y le texte \"OUVRIR CALCULATEUR\"");
  sheet.getRange("B17").setValue("3. Cliquez sur les trois points du bouton et choisissez \"Attribuer un script\"");
  sheet.getRange("B18").setValue("4. Saisissez \"ouvrirFenetreCalculateur\" comme nom de fonction");
  
  // Mettre en forme les instructions
  sheet.getRange("B14:B18").setFontStyle("italic");
  sheet.getRange("B14:B18").setFontColor("#666666");
}

/**
 * Fonction pour créer un bouton image directement dans la feuille
 * Cette fonction peut être utilisée pour créer un bouton qui déclenche l'ouverture du calculateur
 */
function creerBoutonImage() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('0 - Read me');
    
    if (sheet) {
      // Créer un bouton coloré avec du texte
      // Note: Cette méthode ne peut pas directement associer un script au bouton,
      // mais elle peut servir d'instruction visuelle pour l'utilisateur
      
      // Définir les propriétés du bouton
      var btnWidth = 200;
      var btnHeight = 40;
      var btnX = 100; // Position X en pixels
      var btnY = 200; // Position Y en pixels
      
      // Créer un bouton rectangle avec texte
      var bouton = '<svg width="' + btnWidth + '" height="' + btnHeight + '">' +
                   '<rect width="100%" height="100%" rx="8" ry="8" fill="#4285F4" />' +
                   '<text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" fill="white" font-family="Arial" font-size="14" font-weight="bold">OUVRIR CALCULATEUR</text>' +
                   '</svg>';
      
      // Convertir le SVG en base64
      var encodedBouton = Utilities.base64Encode(bouton);
      
      // Insérer l'image à partir des données Base64
      // Cette partie est commentée car elle ne fonctionnera pas directement dans Google Sheets
      // via Google Apps Script à cause des limitations
      // sheet.insertImage('data:image/svg+xml;base64,' + encodedBouton, 2, 20, 0, 0);
      
      // Au lieu de cela, ajouter des instructions claires
      sheet.getRange("C20").setValue("IMPORTANT: Pour créer un bouton fonctionnel, utilisez le menu \"Insertion > Dessin\"");
      sheet.getRange("C21").setValue("puis associez le script \"ouvrirFenetreCalculateur\" au bouton via le menu contextuel.");
      sheet.getRange("C20:C21").setFontColor("red");
      sheet.getRange("C20:C21").setFontWeight("bold");
      
      return true;
    }
    return false;
  } catch (error) {
    console.log("Erreur lors de la création du bouton image: " + error);
    return false;
  }
}

/**
 * Fonction pour calculer la moyenne de la demande énergétique
 * Calcule la moyenne des valeurs G5:G16 et l'affiche en G25
 */
function calculerMoyenneDemande() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('1 - MAIN');
    
    if (!sheet) {
      throw new Error("Feuille '1 - MAIN' non trouvée");
    }
    
    // Calculer la moyenne de G5:G16
    var demandes = sheet.getRange('G5:G16').getValues();
    var somme = 0;
    var count = 0;
    
    for (var i = 0; i < demandes.length; i++) {
      var valeur = parseFloat(demandes[i][0]);
      if (!isNaN(valeur) && valeur > 0) {
        somme += valeur;
        count++;
      }
    }
    
    var moyenne = count > 0 ? somme / count : 0;
    
    // Afficher la moyenne en G25
    sheet.getRange('G25').setValue(moyenne);
    sheet.getRange('G25').setNumberFormat('#,##0.00');
    
    console.log("Moyenne de la demande calculée (G5:G16):", moyenne, "affichée en G25");
    return moyenne;
  } catch (error) {
    console.error("Erreur lors du calcul de la moyenne de demande:", error);
    return 0;
  }
}

/**
 * Fonction pour optimiser la sélection des chillers
 * Trouve la combinaison optimale de chillers pour minimiser la consommation
 */
function optimiserSelectionChillers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('1 - MAIN');
    
    if (!sheet) {
      throw new Error("Feuille '1 - MAIN' non trouvée");
    }
    
    // Capacités disponibles des chillers
    var capacitesDisponibles = [400, 600, 800, 1000, 1200, 2600];
    
    // Lire le maximum de la demande (colonne G) depuis E25
    var maximumDemande = parseFloat(sheet.getRange('E25').getValue()) || 0;
    console.log("Maximum de la demande pour optimisation (E25):", maximumDemande);
    
    // Si le maximum est trop faible, utiliser une valeur minimale de sécurité
    if (maximumDemande < 100) {
      console.warn("Maximum de demande très faible (" + maximumDemande + "), utilisation d'une valeur minimale de 1000 kWc");
      maximumDemande = 1000;
    }
    
    // Lire la consommation actuelle minimale (F25) pour comparaison
    var consommationActuelle = parseFloat(sheet.getRange('F25').getValue()) || 0;
    console.log("Consommation actuelle (F25):", consommationActuelle);
    
    var meilleureCombinaison = null;
    var meilleureConsommation = Infinity;
    var meilleurTotal = 0;
    
    // Tester toutes les combinaisons possibles de 3 chillers
    console.log("Test de", Math.pow(capacitesDisponibles.length, 3), "combinaisons possibles...");
    
    for (var i = 0; i < capacitesDisponibles.length; i++) {
      for (var j = 0; j < capacitesDisponibles.length; j++) {
        for (var k = 0; k < capacitesDisponibles.length; k++) {
          var chiller1 = capacitesDisponibles[i];
          var chiller2 = capacitesDisponibles[j];
          var chiller3 = capacitesDisponibles[k];
          var totalCapacite = chiller1 + chiller2 + chiller3;
          
          // Vérifications de contraintes
          // 1. La capacité totale doit couvrir au moins le maximum de la demande
          if (totalCapacite < maximumDemande) {
            continue;
          }
          
          // 2. En cas de panne d'un chiller, les 2 autres doivent couvrir au moins 100% du maximum
          var capaciteSecours1 = chiller2 + chiller3;
          var capaciteSecours2 = chiller1 + chiller3;
          var capaciteSecours3 = chiller1 + chiller2;
          var seuilSecours = maximumDemande * 1;
          
          if (capaciteSecours1 < seuilSecours || 
              capaciteSecours2 < seuilSecours || 
              capaciteSecours3 < seuilSecours) {
            continue;
          }
          
          // Calculer la consommation estimée pour cette combinaison
          var consommationEstimee = calculerConsommationChillers(chiller1, chiller2, chiller3, maximumDemande);
          
          // Vérifier si c'est la meilleure combinaison (consommation la plus faible)
          if (consommationEstimee < meilleureConsommation) {
            meilleureConsommation = consommationEstimee;
            meilleureCombinaison = [chiller1, chiller2, chiller3];
            meilleurTotal = totalCapacite;
            
            console.log("Nouvelle meilleure combinaison trouvée:");
            console.log("- Chillers:", chiller1, chiller2, chiller3, "kWc");
            console.log("- Total:", totalCapacite, "kWc");
            console.log("- Consommation:", consommationEstimee);
          }
        }
      }
    }
    
    console.log("Analyse terminée. Meilleure consommation trouvée:", meilleureConsommation);
    
    // Appliquer la meilleure combinaison trouvée
    if (meilleureCombinaison) {
      // Mettre à jour seulement K10, K11, K12 (pas K13 ni K14)
      sheet.getRange('K10').setValue(meilleureCombinaison[0]);
      sheet.getRange('K11').setValue(meilleureCombinaison[1]);
      sheet.getRange('K12').setValue(meilleureCombinaison[2]);
      
      // Calculer le total pour le retour de fonction (mais ne pas l'écrire dans K13)
      var totalOptimise = meilleureCombinaison[0] + meilleureCombinaison[1] + meilleureCombinaison[2];
      
      // Calculer la différence pour le retour de fonction (mais ne pas l'écrire dans K14)
      var difference = Math.max(0, maximumDemande - totalOptimise);
      
      console.log("=== OPTIMISATION TERMINÉE ===");
      console.log("Chiller 1:", meilleureCombinaison[0], "kWc (cellule K10)");
      console.log("Chiller 2:", meilleureCombinaison[1], "kWc (cellule K11)");
      console.log("Chiller 3:", meilleureCombinaison[2], "kWc (cellule K12)");
      console.log("Total calculé:", totalOptimise, "kWc (non écrit dans K13)");
      console.log("Missing kWhc calculé:", difference, "kWc (non écrit dans K14)");
      console.log("Consommation optimisée:", meilleureConsommation);
      console.log("Maximum demande:", maximumDemande, "kWc (cellule E25)");
      
      return {
        chiller1: meilleureCombinaison[0],
        chiller2: meilleureCombinaison[1],
        chiller3: meilleureCombinaison[2],
        total: totalOptimise,
        consommation: meilleureConsommation,
        maximumDemande: maximumDemande,
        difference: difference
      };
    } else {
      console.error("Aucune combinaison optimale trouvée pour une demande de", maximumDemande, "kWc");
      throw new Error("Aucune combinaison optimale trouvée - Vérifiez que la demande maximum est cohérente");
    }
    
  } catch (error) {
    console.error("Erreur lors de l'optimisation des chillers:", error);
    return null;
  }
}

/**
 * Fonction pour calculer la consommation estimée d'une combinaison de chillers
 */
function calculerConsommationChillers(chiller1, chiller2, chiller3, demande) {
  // Formule simplifiée de calcul de consommation
  // Vous pouvez ajuster cette formule selon vos besoins spécifiques
  var totalCapacite = chiller1 + chiller2 + chiller3;
  var facteurUtilisation = Math.min(demande / totalCapacite, 1.0);
  
  // Efficacité différente selon la taille (les plus gros chillers sont plus efficaces)
  var efficaciteChiller1 = getEfficaciteChiller(chiller1);
  var efficaciteChiller2 = getEfficaciteChiller(chiller2);
  var efficaciteChiller3 = getEfficaciteChiller(chiller3);
  
  // Répartition de la charge (priorité aux chillers les plus efficaces)
  var chillers = [
    {capacite: chiller1, efficacite: efficaciteChiller1},
    {capacite: chiller2, efficacite: efficaciteChiller2},
    {capacite: chiller3, efficacite: efficaciteChiller3}
  ].sort((a, b) => b.efficacite - a.efficacite);
  
  var chargeRestante = demande * facteurUtilisation;
  var consommationTotale = 0;
  
  for (var i = 0; i < chillers.length && chargeRestante > 0; i++) {
    var chargeChiller = Math.min(chargeRestante, chillers[i].capacite);
    var consommationChiller = chargeChiller / chillers[i].efficacite;
    consommationTotale += consommationChiller;
    chargeRestante -= chargeChiller;
  }
  
  return consommationTotale;
}

/**
 * Fonction pour obtenir l'efficacité d'un chiller selon sa capacité
 */
function getEfficaciteChiller(capacite) {
  // COP (Coefficient of Performance) selon la taille
  // Les plus gros chillers ont généralement un meilleur COP
  if (capacite >= 2600) return 4.5;
  if (capacite >= 1200) return 4.2;
  if (capacite >= 1000) return 4.0;
  if (capacite >= 800) return 3.8;
  if (capacite >= 600) return 3.6;
  return 3.4; // Pour 400 kWc
}

/**
 * Fonction pour installer la formule de calcul du load minimum (F25)
 */
function installerFormuleLoadMinimum() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('1 - MAIN');
    
    if (!sheet) {
      throw new Error("Feuille '1 - MAIN' non trouvée");
    }
    
    // Formule pour calculer le minimum des loads des 3 chillers
    var formule = '=MIN(C19,F20,G20)';
    sheet.getRange('F25').setFormula(formule);
    
    console.log("Formule installée en F25:", formule);
    return true;
  } catch (error) {
    console.error("Erreur lors de l'installation de la formule:", error);
    return false;
  }
}

/**
 * Fonction de test pour l'optimisation des chillers
 * Peut être appelée manuellement pour tester l'optimisation
 */
function testerOptimisationChillers() {
  try {
    console.log("=== DÉBUT DU TEST D'OPTIMISATION ===");
    
    // Calculer d'abord la moyenne
    var moyenne = calculerMoyenneDemande();
    console.log("Moyenne calculée:", moyenne);
    
    // Installer la formule F25
    var formuleOK = installerFormuleLoadMinimum();
    console.log("Installation formule F25:", formuleOK);
    
    // Optimiser les chillers
    var resultat = optimiserSelectionChillers();
    
    if (resultat) {
      console.log("=== RÉSULTATS DE L'OPTIMISATION ===");
      console.log("Succès! Chillers optimisés:");
      console.log("- Chiller 1:", resultat.chiller1, "kWc");
      console.log("- Chiller 2:", resultat.chiller2, "kWc");
      console.log("- Chiller 3:", resultat.chiller3, "kWc");
      console.log("- Total:", resultat.total, "kWc");
      console.log("- Consommation:", resultat.consommation);
      console.log("- Moyenne demande:", resultat.moyenneDemande, "kWc");
      return resultat;
    } else {
      console.error("Échec de l'optimisation");
      return null;
    }
  } catch (error) {
    console.error("Erreur lors du test d'optimisation:", error);
    return null;
  }
}

/**
 * Fonction de test complet : optimisation + résultats
 * Simule le processus complet du popup
 */
function testerProcessusComplet() {
  try {
    console.log("=== DÉBUT DU TEST COMPLET ===");
    
    // 1. Optimiser les chillers
    console.log("1. Optimisation des chillers...");
    var optimisationResult = optimiserSelectionChillers();
    
    if (optimisationResult) {
      console.log("✅ Optimisation réussie");
      console.log("Chillers sélectionnés - K10:", optimisationResult.chiller1, "K11:", optimisationResult.chiller2, "K12:", optimisationResult.chiller3);
    } else {
      console.warn("⚠️ Optimisation échouée");
    }
    
    // 2. Attendre et forcer le recalcul
    console.log("2. Attente et recalcul...");
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    
    // 3. Récupérer les résultats
    console.log("3. Récupération des résultats techniques...");
    var resultats = getResultatsTechniques();
    
    if (resultats) {
      console.log("✅ Résultats obtenus:");
      console.log("- Chillers current consumption:", resultats.chillers_current_consumption);
      console.log("- Chillers new consumption:", resultats.chillers_new_consumption);
      console.log("- Total savings:", resultats.total_savings);
    } else {
      console.error("❌ Échec de récupération des résultats");
    }
    
    console.log("=== FIN DU TEST COMPLET ===");
    return {
      optimisation: optimisationResult,
      resultats: resultats
    };
    
  } catch (error) {
    console.error("Erreur lors du test complet:", error);
    return null;
  }
}

/**
 * Fonction pour obtenir le token Veolia OAuth2.0
 */
function getVeoliaAccessToken(client_id, client_secret) {
  try {
    const TOKEN_URL = 'https://api.veolia.com/security/v2/oauth/token';
    
    // Encoder les credentials en Base64
    const credentials = client_id + ':' + client_secret;
    const encodedCredentials = Utilities.base64Encode(credentials);
    
    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': 'Basic ' + encodedCredentials
    };
    
    const payload = 'grant_type=client_credentials';
    
    const options = {
      'method': 'POST',
      'headers': headers,
      'payload': payload
    };
    
    const response = UrlFetchApp.fetch(TOKEN_URL, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.access_token) {
      console.log('Token Veolia obtenu avec succès');
      return responseData.access_token;
    } else {
      console.error('Erreur dans la réponse du token:', responseData);
      return null;
    }
    
  } catch (error) {
    console.error('Erreur lors de l\'obtention du token Veolia:', error);
    return null;
  }
}

/**
 * Fonction pour obtenir des suggestions d'adresse via l'API Veolia Weather Data
 * Utilise le géocodage inverse et la recherche de lieux
 */
function getAddressSuggestions(query) {
  try {
    console.log('Recherche Veolia pour:', query);
    
    // Identifiants Veolia (à remplacer par vos vrais identifiants)
    // Pour obtenir vos identifiants :
    // 1. Allez sur https://veglobal.service-now.com/weather
    // 2. Demandez un accès à l'API Weather Data
    // 3. Remplacez les valeurs ci-dessous par vos vrais identifiants
    const VEOLIA_CLIENT_ID = 'YOUR_VEOLIA_CLIENT_ID'; // Remplacez par votre client_id
    const VEOLIA_CLIENT_SECRET = 'YOUR_VEOLIA_CLIENT_SECRET'; // Remplacez par votre client_secret
    
    // Obtenir le token d'authentification
    const token = getVeoliaAccessToken(VEOLIA_CLIENT_ID, VEOLIA_CLIENT_SECRET);
    if (!token) {
      console.error('Impossible d\'obtenir le token Veolia');
      return getFallbackSuggestions(query);
    }
    
    const suggestions = [];
    
    // 1. Essayer la recherche par coordonnées si la query ressemble à des coordonnées
    const coordsMatch = query.match(/(-?\d+\.?\d*),?\s*(-?\d+\.?\d*)/);
    if (coordsMatch) {
      const lat = parseFloat(coordsMatch[1]);
      const lng = parseFloat(coordsMatch[2]);
      
      if (lat >= -90 && lat <= 90 && lng >= -180 && lng <= 180) {
        const coordSuggestion = getLocationByCoordinates(token, lat, lng);
        if (coordSuggestion) {
          suggestions.push(coordSuggestion);
        }
      }
    }
    
    // 2. Rechercher dans les principales villes mondiales via l'API météo
    const citySearchResults = searchCitiesViaWeatherAPI(token, query);
    suggestions.push(...citySearchResults);
    
    // 3. Si pas assez de résultats, utiliser la base de données de fallback
    if (suggestions.length < 3) {
      const fallbackResults = getFallbackSuggestions(query);
      suggestions.push(...fallbackResults);
    }
    
    // Supprimer les doublons et limiter à 8 résultats
    const uniqueSuggestions = removeDuplicateSuggestions(suggestions);
    return uniqueSuggestions.slice(0, 8);
    
  } catch (error) {
    console.error('Erreur lors de la recherche Veolia:', error);
    return getFallbackSuggestions(query);
  }
}

/**
 * Recherche des villes via l'API météo Veolia
 */
function searchCitiesViaWeatherAPI(token, query) {
  const suggestions = [];
  
  try {
    // Liste des principales villes avec leurs coordonnées exactes
    const majorCities = [
      { name: "Paris", address: "Paris, France", lat: 48.8566, lng: 2.3522 },
      { name: "London", address: "London, UK", lat: 51.5074, lng: -0.1278 },
      { name: "New York", address: "New York, NY, USA", lat: 40.7128, lng: -74.0060 },
      { name: "Tokyo", address: "Tokyo, Japan", lat: 35.6762, lng: 139.6503 },
      { name: "Berlin", address: "Berlin, Germany", lat: 52.5200, lng: 13.4050 },
      { name: "Sydney", address: "Sydney, NSW, Australia", lat: -33.8688, lng: 151.2093 },
      { name: "Toronto", address: "Toronto, ON, Canada", lat: 43.6532, lng: -79.3832 },
      { name: "Dubai", address: "Dubai, UAE", lat: 25.2048, lng: 55.2708 },
      { name: "Singapore", address: "Singapore", lat: 1.3521, lng: 103.8198 },
      { name: "Hong Kong", address: "Hong Kong", lat: 22.3193, lng: 114.1694 },
      
      // Aéroports majeurs
      { name: "Charles de Gaulle Airport", address: "Paris CDG Airport, France", lat: 49.0097, lng: 2.5479 },
      { name: "Heathrow Airport", address: "London Heathrow Airport, UK", lat: 51.4700, lng: -0.4543 },
      { name: "JFK Airport", address: "John F. Kennedy Airport, NY, USA", lat: 40.6413, lng: -73.7781 },
      { name: "Narita Airport", address: "Tokyo Narita Airport, Japan", lat: 35.7720, lng: 140.3929 },
      { name: "Frankfurt Airport", address: "Frankfurt am Main Airport, Germany", lat: 50.0379, lng: 8.5622 },
      { name: "Dubai International", address: "Dubai International Airport, UAE", lat: 25.2532, lng: 55.3657 }
    ];
    
    const queryLower = query.toLowerCase().trim();
    
    // Filtrer et valider via l'API météo Veolia
    for (const city of majorCities) {
      if (city.name.toLowerCase().includes(queryLower) || 
          city.address.toLowerCase().includes(queryLower)) {
        
        // Valider que la ville a des données météo disponibles
        const weatherValidation = validateLocationWithWeatherAPI(token, city.lat, city.lng);
        if (weatherValidation.isValid) {
          suggestions.push({
            name: city.name,
            address: city.address,
            lat: city.lat,
            lng: city.lng,
            verified: true,
            weatherAvailable: true
          });
        }
        
        if (suggestions.length >= 5) break; // Limiter les appels API
      }
    }
    
  } catch (error) {
    console.error('Erreur lors de la recherche via API météo:', error);
  }
  
  return suggestions;
}

/**
 * Valider qu'un lieu a des données météo disponibles via l'API Veolia
 */
function validateLocationWithWeatherAPI(token, lat, lng) {
  try {
    const geocode = lat + ',' + lng;
    const API_URL = 'https://api.veolia.com/weather/v1/current/ondemand?geocode=' + encodeURIComponent(geocode) + '&units=m&language=en-US&format=json';
    
    const headers = {
      'Authorization': 'Bearer ' + token
    };
    
    const options = {
      'method': 'GET',
      'headers': headers
    };
    
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseData = JSON.parse(response.getContentText());
    
    // Si la réponse contient des données météo, le lieu est valide
    return {
      isValid: !responseData.error && responseData.length > 0,
      data: responseData
    };
    
  } catch (error) {
    console.error('Erreur validation météo pour', lat, lng, ':', error);
    return { isValid: false, error: error.toString() };
  }
}

/**
 * Obtenir les informations d'un lieu par ses coordonnées
 */
function getLocationByCoordinates(token, lat, lng) {
  try {
    // Valider avec l'API météo
    const validation = validateLocationWithWeatherAPI(token, lat, lng);
    
    if (validation.isValid) {
      return {
        name: `Location ${lat.toFixed(4)}, ${lng.toFixed(4)}`,
        address: `Coordinates: ${lat.toFixed(6)}, ${lng.toFixed(6)}`,
        lat: lat,
        lng: lng,
        verified: true,
        weatherAvailable: true
      };
    }
    
    return null;
    
  } catch (error) {
    console.error('Erreur lors de la recherche par coordonnées:', error);
    return null;
  }
}

/**
 * Base de données de fallback mondiale en cas d'échec de l'API OpenCage
 */
function getFallbackSuggestions(query) {
  const fallbackLocations = [
    // Europe
    { name: "Paris", address: "Paris, France", lat: 48.8566, lng: 2.3522, country: "France" },
    { name: "London", address: "London, United Kingdom", lat: 51.5074, lng: -0.1278, country: "United Kingdom" },
    { name: "Berlin", address: "Berlin, Germany", lat: 52.5200, lng: 13.4050, country: "Germany" },
    { name: "Madrid", address: "Madrid, Spain", lat: 40.4168, lng: -3.7038, country: "Spain" },
    { name: "Rome", address: "Rome, Italy", lat: 41.9028, lng: 12.4964, country: "Italy" },
    
    // North America
    { name: "New York", address: "New York, NY, USA", lat: 40.7128, lng: -74.0060, country: "United States" },
    { name: "Los Angeles", address: "Los Angeles, CA, USA", lat: 34.0522, lng: -118.2437, country: "United States" },
    { name: "Toronto", address: "Toronto, ON, Canada", lat: 43.6532, lng: -79.3832, country: "Canada" },
    { name: "Mexico City", address: "Mexico City, Mexico", lat: 19.4326, lng: -99.1332, country: "Mexico" },
    
    // Asia
    { name: "Tokyo", address: "Tokyo, Japan", lat: 35.6762, lng: 139.6503, country: "Japan" },
    { name: "Beijing", address: "Beijing, China", lat: 39.9042, lng: 116.4074, country: "China" },
    { name: "Singapore", address: "Singapore", lat: 1.3521, lng: 103.8198, country: "Singapore" },
    { name: "Mumbai", address: "Mumbai, India", lat: 19.0760, lng: 72.8777, country: "India" },
    { name: "Seoul", address: "Seoul, South Korea", lat: 37.5665, lng: 126.9780, country: "South Korea" },
    
    // Oceania
    { name: "Sydney", address: "Sydney, NSW, Australia", lat: -33.8688, lng: 151.2093, country: "Australia" },
    { name: "Auckland", address: "Auckland, New Zealand", lat: -36.8485, lng: 174.7633, country: "New Zealand" },
    
    // South America
    { name: "São Paulo", address: "São Paulo, Brazil", lat: -23.5558, lng: -46.6396, country: "Brazil" },
    { name: "Buenos Aires", address: "Buenos Aires, Argentina", lat: -34.6118, lng: -58.3960, country: "Argentina" },
    
    // Africa
    { name: "Cairo", address: "Cairo, Egypt", lat: 30.0444, lng: 31.2357, country: "Egypt" },
    { name: "Lagos", address: "Lagos, Nigeria", lat: 6.5244, lng: 3.3792, country: "Nigeria" },
    { name: "Cape Town", address: "Cape Town, South Africa", lat: -33.9249, lng: 18.4241, country: "South Africa" }
  ];
  
  const queryLower = query.toLowerCase().trim();
  const matches = [];
  
  for (const location of fallbackLocations) {
    if (location.name.toLowerCase().includes(queryLower) || 
        location.address.toLowerCase().includes(queryLower) ||
        location.country.toLowerCase().includes(queryLower)) {
      
      const virtualGridPoint = calculateClosestVirtualGridPoint(location.lat, location.lng);
      
      matches.push({
        ...location,
        weatherLat: virtualGridPoint.lat,
        weatherLng: virtualGridPoint.lng,
        gridDistance: virtualGridPoint.distance,
        verified: false,
        weatherAvailable: false,
        confidence: 6 // Score modéré pour les fallbacks
      });
    }
  }
  
  return matches.slice(0, 8); // Limite à 8 résultats
}

/**
 * Supprimer les suggestions en double
 */
function removeDuplicateSuggestions(suggestions) {
  const seen = new Set();
  const unique = [];
  
  for (const suggestion of suggestions) {
    const key = `${suggestion.lat.toFixed(4)},${suggestion.lng.toFixed(4)}`;
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(suggestion);
    }
  }
  
  return unique;
}

/* =====================================================
 * INTÉGRATION OPENCAGE GEOCODER API
 * ===================================================== */

/**
 * Configuration de l'API OpenCage Geocoder
 */
const OPENCAGE_API_KEY = 'd4698a8893784dcda05d8b9544aa460c'; // Votre clé API OpenCage
const OPENCAGE_BASE_URL = 'https://api.opencagedata.com/geocode/v1/json';

/**
 * Fonction principale pour obtenir les suggestions d'adresses avec OpenCage
 * Remplace l'ancienne fonction getAddressSuggestions
 */
function getAddressSuggestions(query) {
  console.log('OpenCage - Recherche de suggestions pour:', query);
  
  if (!query || query.trim().length < 2) {
    return [];
  }
  
  try {
    // Recherche avec OpenCage API
    const suggestions = searchAddressesWithOpenCage(query.trim());
    
    // Vérification météo pour chaque suggestion avec points de grille virtuels
    const enhancedSuggestions = suggestions.map(suggestion => {
      const weatherVerification = validateLocationWithWeatherAPI(
        suggestion.weatherLat, // Utiliser les coordonnées de grille pour la météo
        suggestion.weatherLng
      );
      return {
        ...suggestion,
        verified: weatherVerification.verified,
        weatherAvailable: weatherVerification.weatherAvailable,
        finalGridDistance: weatherVerification.gridDistance,
        weatherGridInfo: weatherVerification.weatherGridPoint,
        isAlternativeGrid: weatherVerification.isAlternativeGrid || false
      };
    });
    
    console.log('OpenCage - Suggestions enrichies:', enhancedSuggestions.length);
    return enhancedSuggestions;
    
  } catch (error) {
    console.error('Erreur lors de la recherche OpenCage:', error);
    // Fallback vers des suggestions statiques en cas d'erreur
    return getFallbackSuggestions(query);
  }
}

/**
 * Recherche d'adresses avec l'API OpenCage Geocoder - MONDIALE
 */
function searchAddressesWithOpenCage(query) {
  try {
    // Recherche mondiale sans restriction de pays pour coverage complète
    const url = `${OPENCAGE_BASE_URL}?key=${OPENCAGE_API_KEY}&q=${encodeURIComponent(query)}&limit=12&language=en&no_annotations=1&no_record=1&min_confidence=5`;
    
    console.log('OpenCage - URL de requête mondiale:', url);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'VeoliaCaaSCalculator/1.0'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenCage API error: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    console.log('OpenCage - Réponse brute:', JSON.stringify(data, null, 2));
    
    if (!data.results || data.results.length === 0) {
      console.log('OpenCage - Aucun résultat trouvé');
      return [];
    }
    
    // Transformer les résultats OpenCage en format standardisé avec virtual grid point
    const suggestions = data.results.map(result => {
      // Calculer le point de grille virtuel le plus proche pour la météo
      const virtualGridPoint = calculateClosestVirtualGridPoint(
        parseFloat(result.geometry.lat), 
        parseFloat(result.geometry.lng)
      );
      
      return {
        name: extractLocationName(result),
        address: result.formatted || 'Address not available',
        lat: parseFloat(result.geometry.lat),
        lng: parseFloat(result.geometry.lng),
        // Coordonnées du point de grille virtuel pour l'API météo
        weatherLat: virtualGridPoint.lat,
        weatherLng: virtualGridPoint.lng,
        gridDistance: virtualGridPoint.distance,
        confidence: result.confidence || 0,
        country: result.components.country || '',
        countryCode: result.components.country_code || '',
        state: result.components.state || result.components.province || '',
        city: result.components.city || result.components.town || result.components.village || '',
        postcode: result.components.postcode || '',
        continent: result.components.continent || '',
        timezone: result.annotations?.timezone?.name || '',
        verified: false, // Sera mis à jour par la vérification météo
        weatherAvailable: false // Sera mis à jour par la vérification météo
      };
    });
    
    // Filtrer et trier par pertinence avec couverture mondiale
    const filteredSuggestions = suggestions
      .filter(s => s.confidence >= 5) // Seuil de confiance minimum réduit pour couverture mondiale
      .sort((a, b) => {
        // Trier par pertinence : confiance + disponibilité météo + distance de grille
        const scoreA = a.confidence + (a.weatherAvailable ? 2 : 0) - (a.gridDistance * 0.1);
        const scoreB = b.confidence + (b.weatherAvailable ? 2 : 0) - (b.gridDistance * 0.1);
        return scoreB - scoreA;
      })
      .slice(0, 10); // Augmenter à 10 suggestions pour plus de choix mondial
    
    console.log('OpenCage - Suggestions filtrées:', filteredSuggestions.length);
    return filteredSuggestions;
    
  } catch (error) {
    console.error('Erreur OpenCage API:', error);
    throw error;
  }
}

/**
 * Calcul du point de grille virtuel le plus proche pour l'API météo Veolia
 * Utilise une grille mondiale avec des points espacés de ~25km pour une couverture optimale
 */
function calculateClosestVirtualGridPoint(lat, lng) {
  // Grille virtuelle mondiale avec résolution de 0.25° (~25km)
  const GRID_RESOLUTION = 0.25;
  
  // Calculer les coordonnées de grille les plus proches
  const gridLat = Math.round(lat / GRID_RESOLUTION) * GRID_RESOLUTION;
  const gridLng = Math.round(lng / GRID_RESOLUTION) * GRID_RESOLUTION;
  
  // Calculer la distance entre le point réel et le point de grille (en km approximatif)
  const distance = calculateDistanceBetweenPoints(lat, lng, gridLat, gridLng);
  
  // Assurer les limites géographiques
  const finalGridLat = Math.max(-90, Math.min(90, gridLat));
  const finalGridLng = Math.max(-180, Math.min(180, gridLng));
  
  return {
    lat: parseFloat(finalGridLat.toFixed(4)),
    lng: parseFloat(finalGridLng.toFixed(4)),
    distance: parseFloat(distance.toFixed(2)),
    resolution: GRID_RESOLUTION
  };
}

/**
 * Calcul de distance entre deux points géographiques (formule de Haversine)
 */
function calculateDistanceBetweenPoints(lat1, lng1, lat2, lng2) {
  const R = 6371; // Rayon de la Terre en km
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
           Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
           Math.sin(dLng/2) * Math.sin(dLng/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

/**
 * Obtention de grilles météo alternatives pour les zones avec couverture limitée
 */
function getAlternativeWeatherGridPoints(lat, lng, radius = 1.0) {
  const alternatives = [];
  const steps = [-1, -0.5, 0, 0.5, 1];
  
  for (const latStep of steps) {
    for (const lngStep of steps) {
      if (latStep === 0 && lngStep === 0) continue; // Skip le point central
      
      const altLat = lat + (latStep * radius);
      const altLng = lng + (lngStep * radius);
      
      // Vérifier les limites géographiques
      if (altLat >= -90 && altLat <= 90 && altLng >= -180 && altLng <= 180) {
        const gridPoint = calculateClosestVirtualGridPoint(altLat, altLng);
        alternatives.push({
          ...gridPoint,
          originalDistance: calculateDistanceBetweenPoints(lat, lng, altLat, altLng)
        });
      }
    }
  }
  
  // Trier par distance et retourner les 3 meilleurs
  return alternatives
    .sort((a, b) => a.originalDistance - b.originalDistance)
    .slice(0, 3);
}
function extractLocationName(result) {
  const components = result.components;
  
  // Priorité : ville > village > quartier > lieu-dit
  if (components.city) return components.city;
  if (components.town) return components.town;
  if (components.village) return components.village;
  if (components.suburb) return components.suburb;
  if (components.neighbourhood) return components.neighbourhood;
  if (components.hamlet) return components.hamlet;
  
  // Sinon, extraire le premier élément significatif de l'adresse formatée
  const formatted = result.formatted || '';
  const parts = formatted.split(',');
  return parts[0] ? parts[0].trim() : 'Lieu inconnu';
}

/**
 * Géocodage inverse - Convertir les coordonnées en adresse (MONDIAL)
 */
function reverseGeocodeWithOpenCage(lat, lng) {
  try {
    // Recherche mondiale avec support multi-langue
    const url = `${OPENCAGE_BASE_URL}?key=${OPENCAGE_API_KEY}&q=${lat},${lng}&language=en&no_annotations=1&no_record=1`;
    
    console.log('OpenCage - Géocodage inverse mondial:', lat, lng);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'VeoliaCaaSCalculator/1.0'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenCage API error: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.results || data.results.length === 0) {
      return {
        name: 'Coordinates ' + lat + ', ' + lng,
        address: 'Address not found',
        lat: lat,
        lng: lng,
        country: 'Unknown',
        weatherGridPoint: calculateClosestVirtualGridPoint(lat, lng)
      };
    }
    
    const result = data.results[0];
    const virtualGridPoint = calculateClosestVirtualGridPoint(lat, lng);
    
    return {
      name: extractLocationName(result),
      address: result.formatted || 'Address not available',
      lat: lat,
      lng: lng,
      country: result.components.country || 'Unknown',
      countryCode: result.components.country_code || '',
      confidence: result.confidence || 0,
      weatherGridPoint: virtualGridPoint,
      gridDistance: virtualGridPoint.distance
    };
    
  } catch (error) {
    console.error('Erreur géocodage inverse OpenCage:', error);
    const virtualGridPoint = calculateClosestVirtualGridPoint(lat, lng);
    return {
      name: 'Coordinates ' + lat + ', ' + lng,
      address: 'Geocoding error',
      lat: lat,
      lng: lng,
      country: 'Unknown',
      weatherGridPoint: virtualGridPoint,
      gridDistance: virtualGridPoint.distance
    };
  }
}

/**
 * Validation des coordonnées avec vérification météo Veolia via virtual grid points
 */
function validateLocationWithWeatherAPI(lat, lng) {
  try {
    // Vérifier si les coordonnées sont valides
    if (!isValidCoordinate(lat, lng)) {
      return { verified: false, weatherAvailable: false, gridDistance: null };
    }
    
    // Calculer le point de grille virtuel le plus proche
    const virtualGridPoint = calculateClosestVirtualGridPoint(lat, lng);
    
    // Utiliser les coordonnées de grille pour l'API météo
    const weatherGeocode = `${virtualGridPoint.lat},${virtualGridPoint.lng}`;
    
    console.log(`Validation météo - Point réel: ${lat},${lng} → Grille virtuelle: ${weatherGeocode} (distance: ${virtualGridPoint.distance}km)`);
    
    // Configuration pour test de connectivité avec le point de grille
    const testParams = {
      geocode: weatherGeocode,
      dataType: 'current',
      interval: 'hourly'
    };
    
    // Essayer d'obtenir des données météo pour le point de grille
    try {
      const weatherTest = getVeoliaWeatherData(testParams);
      
      if (weatherTest && !weatherTest.error) {
        return { 
          verified: true, 
          weatherAvailable: true, 
          gridDistance: virtualGridPoint.distance,
          weatherGridPoint: virtualGridPoint
        };
      } else {
        // Si le point de grille principal échoue, essayer des alternatives
        const alternatives = getAlternativeWeatherGridPoints(lat, lng);
        
        for (const alt of alternatives) {
          try {
            const altParams = {
              geocode: `${alt.lat},${alt.lng}`,
              dataType: 'current',
              interval: 'hourly'
            };
            const altWeatherTest = getVeoliaWeatherData(altParams);
            
            if (altWeatherTest && !altWeatherTest.error) {
              return { 
                verified: true, 
                weatherAvailable: true, 
                gridDistance: alt.originalDistance,
                weatherGridPoint: alt,
                isAlternativeGrid: true
              };
            }
          } catch (altError) {
            continue; // Essayer le point alternatif suivant
          }
        }
        
        // Aucun point de grille ne fonctionne
        return { 
          verified: true, 
          weatherAvailable: false, 
          gridDistance: virtualGridPoint.distance,
          weatherGridPoint: virtualGridPoint
        };
      }
      
    } catch (weatherError) {
      console.log('Point de grille valide mais météo indisponible:', weatherError);
      return { 
        verified: true, 
        weatherAvailable: false, 
        gridDistance: virtualGridPoint.distance,
        weatherGridPoint: virtualGridPoint
      };
    }
    
  } catch (error) {
    console.error('Erreur validation coordonnées avec grille virtuelle:', error);
    return { verified: false, weatherAvailable: false, gridDistance: null };
  }
}

/**
 * Validation des coordonnées géographiques
 */
function isValidCoordinate(lat, lng) {
  const latitude = parseFloat(lat);
  const longitude = parseFloat(lng);
  
  return !isNaN(latitude) && 
         !isNaN(longitude) && 
         latitude >= -90 && 
         latitude <= 90 && 
         longitude >= -180 && 
         longitude <= 180;
}

/**
 * Recherche d'adresses par coordonnées approximatives (pour la carte)
 */
function searchNearbyLocations(lat, lng, radius = 0.01) {
  try {
    // Recherche dans un rayon approximatif
    const queries = [
      `${lat},${lng}`,
      `${lat + radius},${lng}`,
      `${lat - radius},${lng}`,
      `${lat},${lng + radius}`,
      `${lat},${lng - radius}`
    ];
    
    const results = [];
    
    for (const query of queries) {
      try {
        const location = reverseGeocodeWithOpenCage(
          parseFloat(query.split(',')[0]), 
          parseFloat(query.split(',')[1])
        );
        if (location && location.name !== 'Lieu inconnu') {
          results.push(location);
        }
      } catch (e) {
        // Ignorer les erreurs individuelles
        continue;
      }
    }
    
    // Supprimer les doublons et retourner les résultats uniques
    return removeDuplicateSuggestions(results).slice(0, 3);
    
  } catch (error) {
    console.error('Erreur recherche proximité:', error);
    return [];
  }
}
