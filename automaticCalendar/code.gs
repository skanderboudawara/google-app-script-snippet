// Traitement pour récupérer les absences déclarées dans les calendriers sous la forme d'une Absence du bureau.
// Les absences trouvées sont renseignées dans l'onglet 'absences'
// Le traitement log les informations du traitement dans l'onglet 'log'
DAY_MILLIS = 24 * 60 * 60 * 1000;

var logSheet;
var rowLog = 1;
var absencesSheet;
var rowAbsence = 1;

var begin;
var end;

var allCalendars = [];


// fonction principale : 
// - recherche de la plage de date depuis la feuille de calcul
// - recherche des calendriers à scanner
// - récupération des absences dans chaque calendrier
function synchronizeOutOfOfficeFromCalendar() {
  spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // préparation onglet 'log' (raz)
  logSheet = spreadSheet.getSheetByName("log");
  logSheet.clear();
  sheetLog("info", "Process started");
  
  // préparation onglet 'ooo' (raz)
  absencesSheet = spreadSheet.getSheetByName("oOo");
  absencesSheet.clear();
  calSheet = spreadSheet.getSheetByName("calendar");
  
  // récupération des calendriers du user exécutant le script
  // le but est de déterminer les calendriers auxquels s'abonner dans la suite du traitement
  //
  // Fonctionnement à vérifier : les calendriers des collaborateurs ne sont pas pris en compte systématiquement.
  // Alors que la doc Google explicite que tous les calendriers auxquel à souscrit l'utilisateur. (Bug? Délai de propagation de l'info?)
  launcherAllCalendars = CalendarApp.getAllCalendars();
  for(var a = 0; a < launcherAllCalendars.length; a ++) {
    allCalendars.push(launcherAllCalendars[a].getId());
  }
  
  // recherche de la date min et max de la ligne 4 (ligne des dates)
  var plage = calSheet.getRange("4:4")
    .getValues()[0]
    .filter(function(obj) {return obj instanceof Date})
    .sort(function(a, b) {return a.getTime() - b.getTime()} );
  begin = plage[0];
  end = plage[plage.length - 1];

  // liste des adresses mails dans la colonne C. Une adresse est simplement caractérisée par un @
  var ids = calSheet.getRange("B:B")
    .getValues()
    .filter(function(obj) {
            if (!(typeof obj[0] === 'string')) return false;
            if (obj[0].indexOf('@') === -1) return false;
            return true})
    .sort();

  // Boucle sur les calendriers des adresses mail trouvées
  for(var a = 0; a < ids.length; a ++ ) {
    lookUpOneCalendar(ids[a]);
  }

  sheetLog("info", "Process ended");
}

// Traitement d'un Calendier
// Remarque : pour faire une recherche dans un calendrier il faut d'abord s'y abonner
function lookUpOneCalendar(id) {
  sheetLog("info", "Start of calendar process " + id);
  var cal;
  var unsubscribe = false;
  
  // on vérifie si l'utilisateur est déjà ou non abonné à un calendrier
  if (! containsObject(id, allCalendars)) {
    // s'il n'est pas abonné, on s'y abonne
    sheetLog("info", "calendar " + id + " not subscripted");
    try {
      cal = CalendarApp.subscribeToCalendar(id);
      unsubscribe = true;
    } catch (err) {
      // cas où il n'existe pas de calendrier publique correspondant à l'adresse mail
      sheetLog("alert", "calendar " + id + " unfound");
      sheetLog("info", "End of calendar process " + id);
      return;
    }
  } else {
    //si l'utilisateur est déjà abonné, on récupère directement le calendrier
    sheetLog("info", "calendar " + id + " subcripted");
    cal = CalendarApp.getCalendarById(id);
  }

  // recherche des événement sur le critère 'Absent du bureau' (titre en dur des événements
  // créés dans l'agenda avec le type 'Absent du bureau')
  var events= cal.getEvents(begin, end, {search: 'Absent du bureau'});
  var events= events.concat(cal.getEvents(begin, end, {search: ('Out of Office')}));
  
  var event;
  var dates;
  var fraction;
  var duree;
  
  // boucle sur les événements trouvés
  for(var i = 0; i < events.length; i ++) {
    event = events[i];

    // on ne traite que les événements dont le titre est 'Absent du bureau'
    // (les fonctions Javascript de Google ne semble pas proposer mieux pour le moment)
    // --> probablement lié à la nouveauté de la fonctionnalité 'out of office'. Possible
    // que Google propose plus précis par la suite
  // 28/10/20 - DEBUT modif EDARCO + 29/06 Ajout Absent du bureau Congés
    if((event.getTitle() == 'Absent du bureau')  || (event.getTitle() == 'Out of office') ){
      
      // création de la liste de jour entre la date de début (incluse) et la date de fin (exclue i.e - 1 milliseconde)
      dates = createDateSpan(event.getStartTime(), new Date(event.getEndTime().getTime() - 1));

      // pour chaque jour de la plage
      for(var j = 0; j < dates.length; j ++) {
        // en colonne 1 : l'adresse mail
        absencesSheet.getRange(rowAbsence, 1).setValue(cal.getId());
        // en colonne 2 : la date du jour
        absencesSheet.getRange(rowAbsence, 2).setValue(dateTrunc(dates[j]));
        // la colonne 3 est une concaténation des colonnes 1 et 2 pour pouvoir être utilisée
        // par une formule VLOOKUP dans la feuille de calcul
        absencesSheet.getRange(rowAbsence, 3).setFormulaR1C1('=CONCAT(INDIRECT("R[0]C[-2]";FALSE);INDIRECT("R[0]C[-1]";FALSE))');
  // 28/10/20 - EDARCO
    
        // Si Absence ou Out of Office 
        if (event.getTitle() != 'Télétravail') 
        {
        // en colonne 4 : une fraction de journée
        // 1 pour indiquer une journée complète
        // 0,5 pour les journée non complète
        fraction = 1;
        if(j === 0 && (event.getStartTime().getHours() * 60 + event.getStartTime().getMinutes()) > 0 ) {
          // EDARCO 29/06 -  fraction = 0.5;
          duree = (event.getEndTime().getHours() * 60 + event.getEndTime().getMinutes()) - (event.getStartTime().getHours() * 60 + event.getStartTime().getMinutes())

          if (duree > 4*60) {fraction = 1}
          else if (duree > 2*60) {fraction = 0.5}
          else break

        }
        if(j === (dates.length - 1) && (event.getEndTime().getHours() * 60 + event.getEndTime().getMinutes()) > 0 ) {
          // EDARCO 29/06 -  fraction = 0.5;
          duree = (event.getEndTime().getHours() * 60 + event.getEndTime().getMinutes()) - (event.getStartTime().getHours() * 60 + event.getStartTime().getMinutes())

          if (duree > 4*60) {fraction = 1}
          else if (duree > 3*60) {fraction = 0.5}
          else break
          
          
        }
        absencesSheet.getRange(rowAbsence, 4).setValue(fraction);
        }
        rowAbsence ++;
      }
    }
  }
  


  // désabonnement pour revenir à la situation de départ (pour pas que l'utilisateur
  // se retrouve avec tous les calendriers scannés dans son interface utilisateur)
  // ne marche pas toujours dans le cas de collaborateurs présents dans l'interface utilisateur
  // Calendar avant le passage du traitrement (ils sont alors désabonnés)
  // --> bug ou mauvais fonctionnement de la fonction Google getAllCalendars()
  if (unsubscribe) {
    cal.unsubscribeFromCalendar();
  }
  sheetLog("info", "End of calendar process " + id);
}

// fonction utilitaire pour créer une liste de dates comprises entre une date de début
// et une date de fin
function createDateSpan(startDate, endDate) {
  if (startDate.getTime() > endDate.getTime()) {
    throw Error('startDate > endDate');
  }

  var dates = [];

  var curDate = new Date(startDate.getTime());
  while (!dateCompare(curDate, endDate)) {
    dates.push(curDate);
    curDate = new Date(curDate.getTime() + DAY_MILLIS);
  }
  dates.push(endDate);
  return dates;
}

// fonction utilitaire de comparaison de 2 dates (ne tient pas compte des heures, minutes, secondes, millis)
function dateCompare(a, b) {
  return a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

// fonction utilitaire pour retourner une date sans heure, minute, seconde et milliseconde
function dateTrunc(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// fonction utilitaire pour écrire dans l'onglet 'log'
function sheetLog(level, msg, cptl) {
  logSheet.getRange(rowLog, 1).setValue(Utilities.formatDate(new Date(), spreadSheet.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss.SSS"));
  if(level != undefined) {
    logSheet.getRange(rowLog, 2).setValue(level);
  }
  if(msg != undefined) {
    logSheet.getRange(rowLog, 3).setValue(msg);
  }
  if(cptl != undefined) {
    logSheet.getRange(rowLog, 4).setValue(cptl);
  }
  rowLog ++;
}

// fonction utilitaire pour chercher si la représentation en chaîne de caractères d'un objet
// est dans une liste
function containsObject(obj, list) {
    var i;
    for (i = 0; i < list.length; i++) {
        if (list[i].toString() === obj.toString()) {
            return true;
        }
    }

    return false;
}

// fonction pour ajouter la fonction principal dans le menu de la feuille de calcul
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: 'Synchro Calendar', functionName: 'synchronizeOutOfOfficeFromCalendar'} ];
  sheet.addMenu('Fonctions', menuEntries);
}


