/**
 * OLD VERSION
 * @fileoverview Script per gestire la sincronizzazione delle date dei laboratori
 * tra un Google Sheet e un Firebase Realtime Database.
 * * Funzionalità principali:
 * 1. Menu personalizzato per la sincronizzazione manuale completa.
 * 2. Trigger automatico per l'invio di NUOVE date tramite Google Form.
 * 3. Funzione manuale per colorare le celle in base allo stato delle assegnazioni lette da Firebase.
 * * Versione aggiornata per supportare fino a 25 colonne di date.
 */

// =====================================================================================
// CONFIGURAZIONE
// =====================================================================================

/**
 * URL del tuo Firebase Realtime Database.
 */
const FIREBASE_ROOT_URL = 'https://bimbiunito-test2.europe-west1.firebasedatabase.app/';

/**
 * Fuso orario per la formattazione delle date.
 */
const SCRIPT_TIMEZONE = 'Europe/Rome';

/**
 * Nodi Firebase per la colorazione.
 */
// Nodo dove vengono salvati i "SI" / "NO" confermati dagli utenti (per i VERDI)
const ASSEGNAZIONI_CONFERMATE_NODE = 'assegnazioni_primarie'; 
// Nodo dove l'AI salva i risultati (per i GIALLI)
const RISULTATI_AI_NODE = 'risultati_assegnazione_primarie'; 


// =====================================================================================
// CREAZIONE MENU PERSONALIZZATO
// =====================================================================================

/**
 * Trigger che si attiva all'apertura del foglio.
 * Crea un menu personalizzato nella UI di Google Sheets per lanciare le funzioni manuali.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Strumenti Firebase') // Questo sarà il nome del menu nella barra in alto
    .addItem('Sincronizza Dati Completi (Foglio -> Firebase)', 'sincronizzaDatiCompleti')
    .addSeparator() // Aggiunge una linea di separazione nel menu
    .addItem('Reimposta Trigger del Form', 'setupDateTrigger') // Un'opzione utile da avere a portata di mano
    .addSeparator()
    .addItem('Colora Celle Assegnazioni (da Firebase)', 'coloraCelleDaFirebase') // NUOVA VOCE DI MENU
    .addToUi();
}


// =====================================================================================
// FUNZIONE MANUALE PRINCIPALE (LANCIATA DAL MENU)
// =====================================================================================

/**
 * Esegue una sincronizzazione completa tra il Google Sheet e Firebase.
 * Il foglio è la fonte della verità: le date e il punto di incontro su Firebase
 * verranno resi IDENTICI a quelli aggregati dal foglio.
 * QUESTA VERSIONE PERMETTE DATE DUPLICATE PER LO STESSO LABORATORIO.
 */
function sincronizzaDatiCompleti() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'ATTENZIONE: Sincronizzazione Completa',
    'Questa operazione SOSTITUIRÀ completamente le date disponibili e il punto di incontro su Firebase con i dati di questo foglio. Le date non più presenti nel foglio verranno CANCELLATE da Firebase. Sei assolutamente sicuro di voler procedere?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Operazione annullata.');
    return;
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    // --- Passo 1: Scarica tutti i dati da Firebase per riferimento ---
    Logger.log('Download di tutti i laboratori da Firebase...');
    const allLabsFromFirebase = getAllLabs();
    if (!allLabsFromFirebase) {
      throw new Error('Impossibile scaricare i dati dei laboratori da Firebase.');
    }
    
    const firebaseTitleMap = {};
    for (const labId in allLabsFromFirebase) {
      const lab = allLabsFromFirebase[labId];
      if (lab.titolo) {
        firebaseTitleMap[normalizeString(lab.titolo)] = labId;
      }
    }
    Logger.log('Mappa dei laboratori Firebase creata.');

    // --- Passo 2: Raggruppa tutti i dati dal Foglio Google per laboratorio ---
    Logger.log('Aggregazione dei dati dal Google Sheet...');
    const sheetLabsData = {};
    const labTitleColIndex = headers.indexOf('Scegli il tuo laboratorio');
    const puntoIncontroColIndex = headers.indexOf('Punto di incontro con la scolaresca');
    const dateColsIndexes = [];
    for (let i = 1; i <= 10; i++) {
        const colIndex = headers.indexOf(`${i}° data`);
        if (colIndex !== -1) dateColsIndexes.push(colIndex);
    }
    // Aggiungo supporto fino a 25° data  <-- *** MODIFICA 1 DI 3 ***
    for (let i = 11; i <= 25; i++) {
        const colIndex = headers.indexOf(`${i}° data`);
        if (colIndex !== -1) dateColsIndexes.push(colIndex);
    }


    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const labTitle = row[labTitleColIndex];
      if (!labTitle) continue; 

      const normalizedTitle = normalizeString(labTitle);
      
      if (!sheetLabsData[normalizedTitle]) {
        sheetLabsData[normalizedTitle] = {
          punto_incontro: '',
          dates: [] // MODIFICA 1: Sostituito new Set() con un Array vuoto []
        };
      }
      
      const puntoIncontro = row[puntoIncontroColIndex];
      if (puntoIncontro && String(puntoIncontro).trim() !== '') {
        sheetLabsData[normalizedTitle].punto_incontro = String(puntoIncontro).trim();
      }

      dateColsIndexes.forEach(colIndex => {
        let dateValue = row[colIndex];
        if (dateValue) {
          if (dateValue instanceof Date) {
            dateValue = Utilities.formatDate(dateValue, 'Europe/Rome', 'dd/MM/yyyy HH:mm:ss');
          }
          // MODIFICA 2: Sostituito .add() con .push() per aggiungere all'array
          sheetLabsData[normalizedTitle].dates.push(String(dateValue).trim());
        }
      });
    }
    Logger.log(`Dati aggregati per ${Object.keys(sheetLabsData).length} laboratori unici dal foglio.`);

    // --- Passo 3: Itera sui dati aggregati e aggiorna Firebase ---
    let updatedLabsCount = 0;
    for (const normalizedTitle in sheetLabsData) {
      const labId = firebaseTitleMap[normalizedTitle];
      
      if (!labId) {
        Logger.log(`ATTENZIONE: Il laboratorio "${normalizedTitle}" presente nel foglio non è stato trovato su Firebase. Verrà ignorato.`);
        continue;
      }
      
      const labDataFromSheet = sheetLabsData[normalizedTitle];
      
      // MODIFICA 3: Rimosso Array.from() perché 'dates' è già un array
      const newDateArray = labDataFromSheet.dates.map(datetime => ({
        datetime: datetime,
        status: 'disponibile',
        assignedClassId: ''
      }));
      
      const datiDaAggiornare = {
        punto_incontro: labDataFromSheet.punto_incontro,
        date_disponibili: newDateArray,
        timestamp_aggiornamento_date: new Date().toISOString()
      };
      
      Logger.log(`Aggiornamento di Firebase per il laboratorio: ${normalizedTitle} (ID: ${labId})...`);
      updateLabData(labId, datiDaAggiornare);
      updatedLabsCount++;
    }

    const successMessage = `Sincronizzazione completata! Sono stati aggiornati ${updatedLabsCount} laboratori su Firebase.`;
    Logger.log(successMessage);
    ui.alert(successMessage);

  } catch (error) {
    Logger.log(`ERRORE FATALE durante la sincronizzazione: ${error.toString()}`);
    ui.alert('Errore Fatale', `Si è verificato un errore critico durante la sincronizzazione: ${error.message}`, ui.ButtonSet.OK);
  }
}


// =====================================================================================
// FUNZIONI AUTOMATICHE (TRIGGER ON FORM SUBMIT)
// =====================================================================================

/**
 * Funzione trigger che si attiva all'invio di un NUOVO form.
 * @param {Object} e L'oggetto evento del trigger onFormSubmit.
 */
function onDateFormSubmit(e) {
  if (!e || !e.namedValues) {
    Logger.log('Evento non valido o senza valori. Trigger ignorato.');
    return;
  }
  
  Logger.log('Nuova risposta ricevuta dal form...');
  try {
    const titoloLaboratorio = e.namedValues['Scegli il tuo laboratorio'][0];
    const labId = findLabIdByTitleDirect(titoloLaboratorio);
    
    if (!labId) {
      throw new Error(`Laboratorio con titolo "${titoloLaboratorio}" non trovato.`);
    }

    const puntoIncontro = e.namedValues['Punto di incontro con la scolaresca'][0];
    const dateDisponibili = [];
    for (let i = 1; i <= 10; i++) {
        const key = `${i}° data`;
        if (e.namedValues[key] && e.namedValues[key][0]) {
            let dateValue = e.namedValues[key][0];
            if (dateValue instanceof Date) {
                dateValue = Utilities.formatDate(dateValue, 'Europe/Rome', 'dd/MM/yyyy HH:mm:ss');
            }
            dateDisponibili.push({ datetime: String(dateValue).trim(), status: 'disponibile', assignedClassId: '' });
        }
    }
    // Supporto fino a 25° data  <-- *** MODIFICA 2 DI 3 ***
    for (let i = 11; i <= 25; i++) {
        const key = `${i}° data`;
        if (e.namedValues[key] && e.namedValues[key][0]) {
            let dateValue = e.namedValues[key][0];
            if (dateValue instanceof Date) {
                dateValue = Utilities.formatDate(dateValue, 'Europe/Rome', 'dd/MM/yyyy HH:mm:ss');
            }
            dateDisponibili.push({ datetime: String(dateValue).trim(), status: 'disponibile', assignedClassId: '' });
        }
    }


    if (dateDisponibili.length === 0) throw new Error('Nessuna data valida trovata.');

    const datiDaAggiornare = {
      punto_incontro: puntoIncontro.trim(),
      date_disponibili: dateDisponibili,
      timestamp_aggiornamento_date: new Date().toISOString()
    };
    
    updateLabData(labId, datiDaAggiornare);
    Logger.log(`✅ Nuova risposta per il lab ${labId} processata con successo.`);
    
  } catch (error) {
    Logger.log(`❌ Errore nel processare la nuova risposta dal form: ${error.toString()}`);
  }
}

/**
 * Funzione di utilità per (ri)configurare il trigger onFormSubmit.
 * Eseguire una volta per attivare l'automazione.
 */
function setupDateTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onDateFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onDateFormSubmit').forSpreadsheet(sheet).onFormSubmit().create();
  Logger.log('Trigger "onDateFormSubmit" configurato/aggiornato con successo.');
  SpreadsheetApp.getUi().alert('Il trigger per l\'invio di nuovi form è stato configurato correttamente.');
}


// =====================================================================================
// FUNZIONI DI SUPPORTO (HELPER FUNCTIONS)
// =====================================================================================

/**
 * Pulisce una stringa per il confronto: rimuove spazi extra, converte in minuscolo.
 */
function normalizeString(str) {
  if (typeof str !== 'string') return '';
  return str.toLowerCase().trim().replace(/\s+/g, ' ').replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"');
}

/**
 * NUOVO HELPER: Pulisce una stringa per un confronto robusto (spazi e maiuscole).
 */
function cleanStringForCompare(str) {
    if (typeof str !== 'string') return '';
    // Converte in minuscolo, rimuove spazi iniziali/finali e sostituisce spazi multipli con uno singolo
    return str.toLowerCase().trim().replace(/\s+/g, ' '); 
}

/**
 * NUOVO HELPER: Normalizza una stringa di data (con o senza secondi) nel formato 'dd/MM/yyyy HH:mm:ss'
 */
function normalizeDateString(str) {
    if (!str) return null;
    str = String(str).trim();
    if (str.startsWith("'")) { // Rimuove l'apostrofo iniziale se presente
        str = str.substring(1);
    }
    
    // Regex per data e ora, con secondi opzionali
    const parts = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
    if (parts) {
        const [, day, month, year, hours, minutes, seconds = '00'] = parts;
        // Ricostruisce nel formato standard
        return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year} ${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}:${seconds.padStart(2, '0')}`;
    }
    return null; // Formato non riconosciuto
}


function getFirebaseSecret() {
  const secret = PropertiesService.getScriptProperties().getProperty('FIREBASE_SECRET');
  if (!secret) throw new Error('FIREBASE_SECRET non trovato nelle proprietà dello script.');
  return secret;
}

/**
 * NUOVO HELPER: Funzione generica per leggere un percorso da Firebase
 */
function firebaseGet(path) {
  const secret = getFirebaseSecret();
  const url = `${FIREBASE_ROOT_URL}/${path}.json?auth=${secret}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  
  const responseCode = response.getResponseCode();
  const content = response.getContentText();
  
  if (responseCode === 200) {
    return JSON.parse(content);
  }
  if (responseCode === 404 || content === 'null') {
    Logger.log(`Dati non trovati (404 o null) su Firebase al percorso: ${path}`);
    return null;
  }
  Logger.log(`Errore ${responseCode} leggendo da Firebase: ${content}`);
  throw new Error(`Errore leggendo da Firebase: ${responseCode}`);
}


function getAllLabs() {
  // Riadattato per usare il nuovo helper
  return firebaseGet('laboratori');
}

function findLabIdByTitleDirect(titolo) {
  const secret = getFirebaseSecret();
  const titoloNormalizzato = normalizeString(titolo);

  const url = `${FIREBASE_ROOT_URL}/laboratori.json?orderBy="titolo"&auth=${secret}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  
  if (response.getResponseCode() !== 200) {
    throw new Error('Impossibile leggere i dati da Firebase per la ricerca diretta.');
  }
  
  const allLabs = JSON.parse(response.getContentText());
  for (const labId in allLabs) {
    if (normalizeString(allLabs[labId].titolo) === titoloNormalizzato) {
      return labId;
    }
  }
  return null;
}

function updateLabData(labId, dataToUpdate) {
  const secret = getFirebaseSecret();
  const updateUrl = `${FIREBASE_ROOT_URL}/laboratori/${labId}.json?auth=${secret}`;
  const options = {
    method: 'PATCH',
    contentType: 'application/json',
    payload: JSON.stringify(dataToUpdate),
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(updateUrl, options);
  if (response.getResponseCode() >= 400) {
    throw new Error(`Errore PATCH Firebase: ${response.getContentText()}`);
  }
}


// ====================================================================
// ====================================================================
// --- INIZIO NUOVA SEZIONE: FUNZIONE DI COLORAZIONE (DA FIREBASE) ---
// ====================================================================
// ====================================================================

/**
 * Legge lo stato delle assegnazioni direttamente da Firebase (nodi AI e conferme)
 * e applica i colori (giallo/verde) al foglio "Date ricercatori" ATTIVO.
 * Chiamata dal menu "Colora Celle Assegnazioni (da Firebase)".
 * * CORREZIONE: Utilizza getDisplayValues() e normalizza TUTTE le date/nomi
 * prima del confronto.
 * * MODIFICA: La colorazione (setBackgrounds) viene applicata solo dalla riga 2 in giù,
 * per non toccare mai l'intestazione (riga 1).
 */
function coloraCelleDaFirebase() {
    Logger.log("--- AVVIO COLORAZIONE CELLE DA FIREBASE ---");
    const ui = SpreadsheetApp.getUi();
    ui.alert('Avvio colorazione', 'Il processo di lettura da Firebase e colorazione è iniziato. Potrebbero volerci alcuni secondi. Clicca OK per continuare.', ui.ButtonSet.OK);

    const COLOR_GREEN = '#90ee90'; // Verde chiaro
    const COLOR_YELLOW = '#ffffe0'; // Giallo chiaro
    const COLOR_WHITE = '#ffffff'; // Bianco

    // La mappa degli stati conterrà le priorità: "SI" (verde) sovrascrive "proposta" (giallo)
    // CHIAVE: "nomelabpulito|datanormalizzata" (es. "robotiamo|30/10/2025 09:00:00")
    // VALORE: "SI" o "proposta da elaborare"
    const statusMap = {}; 

    try {
        // --- 1. Leggi PASSO 1 (VERDI e GIALLI confermati): Assegnazioni presenti nel nodo principale ---
        try {
            const confermatiData = firebaseGet(ASSEGNAZIONI_CONFERMATE_NODE);
            if (confermatiData) {
                let processedCount = 0;
                for (const key in confermatiData) {
                    const record = confermatiData[key];
                    // CORREZIONE QUI: Non filtriamo solo "SI". Leggiamo anche "proposta da elaborare"
                    // che potrebbe essere finita qui se l'utente ha salvato i dati o sono stati migrati.
                    if (record && record.proposta_accettata) {
                        const labName = record.nome_lab;
                        const dateVal = record.data_lab; // Questa è una data (oggetto o stringa)
                        const status = record.proposta_accettata;

                        if (labName && dateVal) {
                             const normalizedDate = normalizeDateString(dateVal);
                             if (normalizedDate) {
                                const cleanLab = cleanStringForCompare(labName);
                                const mapKey = `${cleanLab}|${normalizedDate}`;
                                
                                if (status === 'SI') {
                                    statusMap[mapKey] = "SI"; // Il SI vince sempre
                                } else if (status === 'proposta da elaborare') {
                                    // Imposta giallo solo se non è già verde (SI)
                                    if (statusMap[mapKey] !== "SI") {
                                        statusMap[mapKey] = "proposta da elaborare";
                                    }
                                }
                                processedCount++;
                             } else {
                                Logger.log(`Attenzione: Impossibile parsare data per record: ${dateVal}.`);
                             }
                        }
                    }
                }
                Logger.log(`Processati ${processedCount} record dal nodo ${ASSEGNAZIONI_CONFERMATE_NODE}.`);
            } else {
                 Logger.log(`Nessun dato trovato nel nodo ${ASSEGNAZIONI_CONFERMATE_NODE}.`);
            }
        } catch (e) {
            Logger.log(`Errore durante la lettura delle assegnazioni confermate: ${e.message}`);
            // Non bloccante, continuiamo con i gialli dell'AI
        }

        // --- 2. Leggi PASSO 2 (GIALLI AI): Proposte AI (PIÙ RECENTI) ---
        // Questo serve per colorare le proposte appena uscite dall'AI che magari non sono ancora
        // state consolidate nel nodo principale (anche se idealmente dovrebbero coincidere).
        try {
            const proposteDataParent = firebaseGet(RISULTATI_AI_NODE);
            if (proposteDataParent) {
                const allTimestamps = Object.keys(proposteDataParent).sort();
                if (allTimestamps.length > 0) {
                    const latestTimestampKey = allTimestamps[allTimestamps.length - 1];
                    const aiProposals = proposteDataParent[latestTimestampKey];
                    Logger.log(`Letti risultati AI dal timestamp: ${latestTimestampKey}`);
                    
                    let yellowCount = 0;
                    for (const id in aiProposals) {
                        const record = aiProposals[id];
                        // L'output dell'AI ha 'labAssegnato' e 'dataAssegnata'
                        if (record && record.labAssegnato && record.dataAssegnata) {
                            
                            const cleanLab = cleanStringForCompare(record.labAssegnato);
                            // La data dell'AI è già stringa "dd/MM/yyyy HH:mm:ss", ma normalizziamola per sicurezza
                            const cleanDate = normalizeDateString(record.dataAssegnata);
                            
                            if (cleanDate) {
                                const mapKey = `${cleanLab}|${cleanDate}`;
                                // Aggiungi solo se non è già presente (quindi non sovrascriviamo "SI" o dati consolidati)
                                if (!statusMap[mapKey]) { 
                                    statusMap[mapKey] = "proposta da elaborare";
                                    yellowCount++;
                                }
                            }
                        }
                    }
                    Logger.log(`Trovate ${yellowCount} nuove 'proposte da elaborare' (gialle) dall'AI (non presenti nel nodo principale).`);
                } else {
                    Logger.log(`Nodo ${RISULTATI_AI_NODE} vuoto, nessun risultato AI da processare.`);
                }
            } else {
                 Logger.log(`Nessun dato (o errore) leggendo il nodo ${RISULTATI_AI_NODE}.`);
            }
        } catch (e) {
            Logger.log(`Errore durante la lettura dei risultati AI: ${e.message}`);
            // Non bloccante
        }

        // --- 3. Applica i colori al foglio ATTIVO "Date ricercatori" ---
        const ricercatoriSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        
        // Leggiamo tutti i valori per trovare le intestazioni e il numero di righe
        const allValuesRange = ricercatoriSheet.getDataRange();
        const ricercatoriValues = allValuesRange.getDisplayValues(); // getDisplayValues() legge le date come TESTO
        
        const ricercatoriHeaders = ricercatoriValues[0];

        const colLabRicercatore = ricercatoriHeaders.indexOf("Scegli il tuo laboratorio");
        const colDataInizio = ricercatoriHeaders.indexOf("1° data");
        
        // <-- *** MODIFICA 3 DI 3 ***
        let colDataFine = ricercatoriHeaders.lastIndexOf("25° data"); 
        if (colDataFine === -1) { // Fallback se "25° data" non esiste
            Logger.log("Attenzione: Colonna '25° data' non trovata, uso l'ultima colonna come limite.");
            colDataFine = ricercatoriHeaders.length - 1;
        }

        if (colLabRicercatore === -1 || colDataInizio === -1) {
            Logger.log("ERRORE: Colonne mancanti nel foglio 'Date ricercatori' attivo: 'Scegli il tuo laboratorio' o '1° data'.");
            ui.alert("Errore: Colonne mancanti nel foglio attivo.");
            return;
        }

        let celleColorate = 0;
        let matchFound = false;

        // Calcoliamo il numero di righe di dati (totale - 1 per l'intestazione)
        const numDataRows = ricercatoriValues.length - 1;
        const numCols = ricercatoriSheet.getLastColumn();

        // Procediamo solo se ci sono effettivamente righe di dati da colorare
        if (numDataRows > 0) {
            
            // Definiamo il range che include SOLO i dati (partendo da riga 2, colonna 1)
            const dataOnlyRange = ricercatoriSheet.getRange(2, 1, numDataRows, numCols);
            
            // Leggiamo gli sfondi correnti SOLO di questo range di dati
            const dataBackgrounds = dataOnlyRange.getBackgrounds();

            // Iteriamo i VALORI (che includono ancora l'intestazione, quindi partiamo da i = 1)
            for (let i = 1; i < ricercatoriValues.length; i++) { // i = 1 è la prima riga di dati nei 'ricercatoriValues'
                
                // L'indice per dataBackgrounds (che è 0-based e non ha l'intestazione)
                const dataRowIndex = i - 1; 
                
                const ricercatoreLabName = ricercatoriValues[i][colLabRicercatore]; // Valore da riga i
                if (!ricercatoreLabName) continue;
                
                const cleanRicercatoreLab = cleanStringForCompare(ricercatoreLabName);

                for (let j = colDataInizio; j <= colDataFine; j++) { // Itera colonne date
                    const cellDateString = ricercatoriValues[i][j]; // È una stringa
                    
                    // Resetta il colore sull'array 'dataBackgrounds'
                    dataBackgrounds[dataRowIndex][j] = COLOR_WHITE; // Resetta colore

                    if (cellDateString) { // Se la cella non è vuota
                        try {
                            // Normalizza la data letta come TESTO dal foglio
                            const cleanRicercatoreDate = normalizeDateString(cellDateString); 
                            
                            if (!cleanRicercatoreDate) continue; // Stringa non valida, salta
                            
                            let foundStatus = null;
                            
                            // Cerchiamo la corrispondenza
                            for (const mapKey in statusMap) {
                                const [firebaseCleanLab, firebaseCleanDate] = mapKey.split('|');
                                
                                // 1. Confronta le date (Stringa Normalizzata vs. Stringa Normalizzata)
                                if (cleanRicercatoreDate === firebaseCleanDate) {
                                    
                                    // 2. Confronta i nomi (Pulito e Parziale)
                                    // Controlla se il nome del ricercatore (lungo) CONTIENE il nome di firebase (corto)
                                    if (cleanRicercatoreLab.includes(firebaseCleanLab)) {
                                        foundStatus = statusMap[mapKey];
                                        
                                        if (!matchFound) {
                                            Logger.log(`DEBUG: Primo match trovato! Firebase: [${firebaseCleanLab}|${firebaseCleanDate}] | Ricercatori: [${cleanRicercatoreLab}|${cleanRicercatoreDate}]`);
                                            matchFound = true;
                                        }
                                        break; 
                                    }
                                }
                            }

                            // Applica il colore in base allo stato trovato
                            if (foundStatus === "SI") {
                                dataBackgrounds[dataRowIndex][j] = COLOR_GREEN;
                                celleColorate++;
                            } else if (foundStatus === "proposta da elaborare") {
                                dataBackgrounds[dataRowIndex][j] = COLOR_YELLOW;
                                celleColorate++;
                            }
                        } catch (e) {
                             Logger.log(`Attenzione: Errore (Ricercatori) riga ${i+1}, col ${j+1}: ${cellDateString}. Errore: ${e.message}`);
                        }
                    }
                }
            }

            // Applica tutti i colori in una volta SOLTANTO al range dei dati
            // L'intestazione (riga 1) non viene toccata.
            dataOnlyRange.setBackgrounds(dataBackgrounds);
        
        } else {
            Logger.log("Nessuna riga di dati trovata (solo intestazione), nessuna colorazione applicata.");
        }


        if (!matchFound && celleColorate === 0 && Object.keys(statusMap).length > 0) {
             Logger.log("ATTENZIONE: 0 celle colorate. Nessuna corrispondenza trovata tra i dati Firebase e le date/nomi di questo foglio.");
        }
        
        const successMsg = `Colorazione da Firebase completata. ${celleColorate} celle colorate (giallo/verde). L'intestazione non è stata modificata.`;
        Logger.log(`✅ ${successMsg}`);
        ui.alert(successMsg);

    } catch (e) {
        Logger.log(`🛑 ERRORE CRITICO nell'aggiornamento colori: ${e.toString()}\nStack: ${e.stack}`);
        ui.alert(`Errore Critico: ${e.message}`);
    }
}
// ====================================================================
// --- FINE NUOVA SEZIONE ---
// ====================================================================