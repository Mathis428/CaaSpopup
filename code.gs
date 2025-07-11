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
 * Fonction pour créer un bouton dans le menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calculateur CaaS')
    .addItem('Ouvrir la fenêtre de saisie', 'ouvrirFenetreCalculateur')
    .addSeparator()
    .addItem('Initialiser les en-têtes', 'initialiserEnTetes')
    .addToUi();
}

/**
 * Fonction pour récupérer les données existantes de la feuille
 */
function getDonneesExistantes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
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
    chillers_min_load: sheet.getRange('D19').getValue() || '520',
    chillers_cop: sheet.getRange('E19').getValue() || '5.2',
    
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
      donnees.includeWeatherData = true;
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
 * Fonction pour récupérer les résultats techniques du calcul
 */
function getResultatsTechniques() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
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
 * Fonction pour sauvegarder les données dans la feuille
 */
function sauvegarderDonnees(donnees) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
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
    if (donnees.chillers_load) sheet.getRange('C19').setValue(parseFloat(donnees.chillers_load));
    if (donnees.chillers_min_load) sheet.getRange('D19').setValue(parseInt(donnees.chillers_min_load));
    if (donnees.chillers_cop) sheet.getRange('E19').setValue(parseFloat(donnees.chillers_cop));
    
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
    
    // Attendre un moment pour que les calculs se fassent dans la feuille
    Utilities.sleep(2000); // Attendre 2 secondes
    
    // Récupérer les résultats techniques
    var resultats = getResultatsTechniques();
    
    // Afficher un message de confirmation et renvoyer les résultats
    return {
      success: true, 
      message: 'Données sauvegardées avec succès dans le Google Sheet!',
      resultats: resultats
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
 * Fonction pour obtenir le token d'authentification OAuth2
 */
function getTokenBearer(client_id, client_secret) {
  var TOKEN_URL = 'https://api.veolia.com/security/v2/oauth/token';
  
  // Définir les en-têtes
  var header = {
    "Content-Type": "application/x-www-form-urlencoded",
    "Authorization": "Basic " + Utilities.base64Encode(client_id + ":" + client_secret)
  };
  
  // Définir les données
  var data = {
    'grant_type': 'client_credentials',
  };
  
  // Ajuster les options
  var options = {
    'method': 'POST',
    'headers': header,
    'payload': data
  };
  
  try {
    // Récupérer le token
    var response = UrlFetchApp.fetch(TOKEN_URL, options);
    var responseData = JSON.parse(response.getContentText());
    return responseData.access_token;
  } catch (error) {
    console.error('Erreur lors de l\'obtention du token:', error);
    return null;
  }
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
