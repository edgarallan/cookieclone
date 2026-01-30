// 1. CONFIGURAZIONE
const FIREBASE_BASE_URL = 'https://bimbiunito-test2.europe-west1.firebasedatabase.app/';
const FIREBASE_NODE = 'laboratori'; 
const API_KEY = 'AIzaSyBbi2LE2CMxwrz44v3qyuO-XkW6ediefJE';

// CONFIGURAZIONE FOGLIO DI CALCOLO
// !!! ASSICURATI CHE SIA IL NOME ESATTO DEL FOGLIO CON LE RISPOSTE DEI RICERCATORI !!!
const SHEET_NAME = 'Risposte del modulo 1'; // Modifica se il nome del foglio è diverso
const BATCH_SIZE = 50; // Invia dati in blocchi (60 proposte sono poche, quindi 50 va bene)


// ====================================================================
// MENU E FUNZIONE DI POPOLAMENTO
// ====================================================================

/**
 * Crea un menu personalizzato per eseguire il popolamento.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Firebase Admin (Laboratori)')
    .addItem('Popola Firebase da Foglio (CANCELLA E RICARICA)', 'repopulateFirebaseFromSheet')
    .addToUi();
}

/**
 * Funzione MASTER per cancellare e ripopolare il nodo /laboratori.
 * DA ESEGUIRE UNA SOLA VOLTA.
 */
function repopulateFirebaseFromSheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'ATTENZIONE!',
    'Questa operazione CANCELLERÀ tutti i dati nel nodo "' + FIREBASE_NODE + '" e li sostituirà con i dati di questo foglio. Sei sicuro di voler procedere?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Operazione annullata.');
    return;
  }
  
  try {
    Logger.log('--- AVVIO POPOLAMENTO COMPLETO (LABORATORI) ---');

    Logger.log('Cancellazione dei dati esistenti in Firebase...');
    deleteAllDataFromFirebase();
    Logger.log('Dati esistenti cancellati.');

    const sheetData = getSheetDataAsObjects();
    if (!sheetData || sheetData.length === 0) {
      ui.alert('Nessun dato trovato nel foglio di calcolo.');
      return;
    }
    Logger.log(`Trovate ${sheetData.length} righe (laboratori) da importare.`);
    
    const allPayloads = [];
    for (const row of sheetData) {
      const payload = createFirebasePayload(row);
      if (payload && Object.keys(payload).length > 1) {
        allPayloads.push(payload);
      }
    }

    if (allPayloads.length > 0) {
      Logger.log(`Invio di ${allPayloads.length} record in batch...`);
      
      const batchPayload = {};
      allPayloads.forEach(payload => {
        const uniqueId = generateUniqueId();
        batchPayload[uniqueId] = payload;
      });

      sendBatchDataToFirebase(batchPayload);
      
      ui.alert(`Popolamento completato! Inviati ${allPayloads.length} laboratori a Firebase.`);
    } else {
      ui.alert('Nessun record valido trovato nel foglio da inviare.');
    }

  } catch (error) {
    Logger.log('ERRORE CRITICO durante il popolamento: ' + error.toString());
    ui.alert('Si è verificato un errore critico. Controllare i log (Estensioni > Apps Script > Esecuzioni).');
  }
}


// ====================================================================
// FUNZIONE CENTRALE DI MAPPATURA DEI DATI
// ====================================================================

/**
 * Mappa i dati grezzi in un oggetto JSON pulito, secondo la struttura desiderata.
 */
function createFirebasePayload(submittedData) {
  const payload = {};

  const getValue = (question) => {
    const value = submittedData[question];
    if (value) {
      const firstValue = Array.isArray(value) ? value[0] : value;
      return (typeof firstValue === 'string') ? firstValue.trim() : firstValue;
    }
    return null;
  };
  
  // Mappa le intestazioni del CSV alle chiavi JSON
  payload.timestamp = getValue("Informazioni cronologiche");
  payload.email = getValue("Indirizzo email");
  payload.privacy = getValue("Ho preso visione dell'informativa sul trattamento dei dati personali e della nostra privacy policy");
  payload.nome = getValue("Nome");
  payload.cognome = getValue("Cognome");
  payload.dipartimento = getValue("Dipartimento/Struttura di afferenza");
  payload.telefono = getValue("Contatto telefonico");
  payload.telefono_interno = getValue("Numero di telefono interno");
  payload.collaboratori = getValue("Nome, cognome ed email istituzionale di eventuali collaboratori");
  payload.titolo = getValue("Titolo della proposta");
  payload.descrizione = getValue("Descrizione");
  payload.obiettivi_metodi = getValue("Obiettivi e Metodi");
  payload.area_tematica = getValue("Area tematica prevalente");
  payload.tipologia_attivita = getValue("Tipologia attività");
  payload.destinatari = getValue("Destinatari");
  
  // Gestione dei campi multi-selezione (li trasformiamo in array)
  const primaria = getValue("Classi di scuola Primaria");
  if (primaria) payload.primaria = primaria.split(',').map(item => item.trim());

  const secondaria = getValue("Classi di scuola Secondaria di I grado");
  if (secondaria) payload.secondaria_1 = secondaria.split(',').map(item => item.trim());
  
  payload.sede = getValue("Sede di svolgimento dell’attività");
  payload.indirizzo_unito = getValue("Indirizzi sedi Unito");
  payload.repliche = getValue("Quante repliche del tuo laboratorio puoi tenere in questa edizione di UGAU?");
  
  const autunnale = getValue("Sessione autunnale");
  if (autunnale) payload.sessione_autunnale = autunnale.split(',').map(item => item.trim());
  
  const primaverile = getValue("Sessione primaverile");
  if (primaverile) payload.sessione_primaverile = primaverile.split(',').map(item => item.trim());

  payload.n_incontri = getValue("Numero incontri previsti dal laboratorio per singolo gruppo di partecipanti:");
  payload.durata_incontro = getValue("Durata approssimativa del singolo incontro:");
  payload.disabilita_possibile = getValue("Possibilità di accogliere alunne/i con disabilità, in base alle caratteristiche dell'attività (non della location in cui si svolgerà)");
  
  const tipologieDisabilita = getValue("Selezionare le tipologie ammesse");
  if (tipologieDisabilita) payload.tipologie_disabilita = tipologieDisabilita.split(',').map(item => item.trim());

  // Aggiunge lo stato come richiesto dal vecchio script
  payload.stato = 'nuovo';

  // Rimuove eventuali campi nulli o vuoti
  for (const key in payload) {
    if (payload[key] === null || payload[key] === '') {
      delete payload[key];
    }
  }

  return payload;
}


// ====================================================================
// FUNZIONI DI UTILITÀ E COMUNICAZIONE CON FIREBASE
// ====================================================================

function getSheetDataAsObjects() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Foglio non trovato: "${SHEET_NAME}"`);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  Logger.log('Intestazioni lette dal foglio: ' + JSON.stringify(headers));
  return data.map(row => {
    const namedValues = {};
    headers.forEach((header, index) => { namedValues[header.trim()] = [row[index]]; });
    return namedValues;
  });
}

function deleteAllDataFromFirebase() {
  const fullUrl = `${FIREBASE_BASE_URL}/${FIREBASE_NODE}.json?key=${API_KEY}`;
  const options = { 'method': 'delete', 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(fullUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`Errore cancellazione dati Firebase (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function sendBatchDataToFirebase(batchPayload) {
  // Per l'importazione totale, usiamo PUT per sovrascrivere tutto in un colpo solo
  const fullUrl = `${FIREBASE_BASE_URL}/${FIREBASE_NODE}.json?key=${API_KEY}`;
  const options = {
    'method': 'put', 
    'contentType': 'application/json',
    'payload': JSON.stringify(batchPayload), 
    'muteHttpExceptions': true
  };
  const response = UrlFetchApp.fetch(fullUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`Errore invio batch Firebase (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function generateUniqueId() {
    const PUSH_CHARS = '-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz';
    let lastPushTime = 0, lastRandChars = [];
    let now = new Date().getTime(), duplicateTime = (now === lastPushTime);
    lastPushTime = now;
    let timeStampChars = new Array(8);
    for (let i = 7; i >= 0; i--) { timeStampChars[i] = PUSH_CHARS.charAt(now % 64); now = Math.floor(now / 64); }
    let id = timeStampChars.join('');
    if (!duplicateTime) { for (let i = 0; i < 12; i++) { lastRandChars[i] = Math.floor(Math.random() * 64); } } 
    else { for (let i = 11; i >= 0 && lastRandChars[i] === 63; i--) { lastRandChars[i] = 0; } lastRandChars[i]++; }
    for (let i = 0; i < 12; i++) { id += PUSH_CHARS.charAt(lastRandChars[i]); }
    return id;
}