/**
 * @fileoverview Servizio per la sincronizzazione dei dati dal Foglio Google
 * (risposte modulo) verso Firebase.
 * VERSIONE AGGIORNATA: Determina la destinazione in base al NOME DEL FILE.
 */

var SheetSyncService = (function() {

  /**
   * Logica di onFormSubmitHandler.
   * (Legge dal foglio attivo/originale)
   * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e L'oggetto evento.
   */
  function handleFormSubmit(e) {
    // Questa funzione usa il foglio a cui è legato lo script
    const submittedData = e.namedValues;
    Logger.log('NUOVA RISPOSTA - Chiavi ricevute: ' + JSON.stringify(Object.keys(submittedData)));
    const firebasePayload = _createFirebasePayload(submittedData);
    
    if (Object.keys(firebasePayload).length > 1) {
      
      // --- MODIFICA: Determina destinazione dal Nome del File (Spreadsheet) ---
      const currentSpreadsheetName = e.range.getSheet().getParent().getName();
      const destinazione = _getDestinazione(currentSpreadsheetName);
      // -----------------------------------------------------------------------

      let node = CONFIG.REQUESTS_NODE_PRIMARY; // Default
      if (destinazione === CONFIG.SECONDARY_DESTINATION) {
        node = CONFIG.REQUESTS_NODE_SECONDARY;
      }
      
      FirebaseService.firebasePost(node, firebasePayload);
      Logger.log(`Nuova risposta inviata con successo a ${node} (Source: ${currentSpreadsheetName}): ${JSON.stringify(firebasePayload)}`);
    }
  }

  /**
   * MODIFICATO: Logica per popolare gli ID Firebase, legge da un file specifico.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Il file da cui leggere.
   * @returns {number} Il numero di ID aggiornati.
   */
  function syncFirebaseIdsToSheet(spreadsheet) {
    Logger.log(`--- Avvio popolamento Firebase IDs da: ${spreadsheet.getName()} ---`);
    const sheet = spreadsheet.getSheetByName(CONFIG.REPOPULATE_SOURCE_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Foglio "${CONFIG.REPOPULATE_SOURCE_SHEET_NAME}" non trovato in "${spreadsheet.getName()}"`);
    }
    
    const dataRange = sheet.getDataRange();
    const sheetData = dataRange.getValues();
    const headers = sheetData.shift();
    const idColIdx = headers.indexOf('firebase_id');
    const tsColIdx = headers.indexOf('Informazioni cronologiche');
    const emailColIdx = headers.indexOf('Indirizzo email');
    if ([idColIdx, tsColIdx, emailColIdx].includes(-1)) {
      throw new Error('Mancano colonne obbligatorie (firebase_id, Informazioni cronologiche, Indirizzo email).');
    }

    Logger.log('Lettura di tutti i record da Firebase (Primarie e Secondarie)...');
    const dataPrimarie = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_PRIMARY) || {};
    const dataSecondarie = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_SECONDARY) || {};
    const firebaseData = { ...dataPrimarie, ...dataSecondarie };
    if (Object.keys(firebaseData).length === 0) {
      Logger.log('Nessun dato trovato su Firebase.');
      return 0;
    }

    const firebaseMap = new Map();
    for (const id in firebaseData) {
      const record = firebaseData[id];
      if (record && record.timestamp && record.email) {
        const recordTimestamp = new Date(record.timestamp).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
        const key = `${recordTimestamp}_${record.email.trim()}`;
        if (!firebaseMap.has(key)) {
          firebaseMap.set(key, id);
        }
      }
    }

    let updatedCount = 0;
    const updates = [];
    sheetData.forEach((row, index) => {
      const firebaseId = row[idColIdx];
      if (!firebaseId) {
        const sheetTimestampRaw = row[tsColIdx];
        const email = row[emailColIdx] ? String(row[emailColIdx]).trim() : null;
        
        if (sheetTimestampRaw && email) {
          const sheetTimestamp = new Date(sheetTimestampRaw).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
          const key = `${sheetTimestamp}_${email}`;
          const foundId = firebaseMap.get(key);
          
          if (foundId) {
            updates.push({ row: index + 2, col: idColIdx + 1, value: foundId });
            updatedCount++;
          }
        }
      }
    });
    if (updates.length > 0) {
      Logger.log(`Trovati ${updates.length} ID da scrivere sul foglio...`);
      updates.forEach(update => {
        sheet.getRange(update.row, update.col).setValue(update.value);
      });
    }
    return updatedCount;
  }

 /**
   * Aggiorna le proposte leggendo dal file, gestendo l'incremento del contatore NO
   * basandosi sul valore massimo registrato nel gruppo di assegnazioni per lo stesso ID.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Il file da cui leggere.
   * @param {string} targetAssegnazioniNode Il nodo Firebase a cui scrivere.
   * @returns {{successCount: number, errorCount: number}} Conteggio delle operazioni.
   */
  function syncProposalsFromFile(spreadsheet, targetAssegnazioniNode) {
    Logger.log(`--- Avvio aggiornamento assegnazioni da: ${spreadsheet.getName()} ---`);
    const sheet = spreadsheet.getSheetByName(CONFIG.REPOPULATE_SOURCE_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Foglio "${CONFIG.REPOPULATE_SOURCE_SHEET_NAME}" non trovato in "${spreadsheet.getName()}"`);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); 

    const idColIdx = headers.indexOf('firebase_id');
    const acceptedColIdx = headers.indexOf('Proposta accettata');
    const labNameColIdx = headers.indexOf('Nome laboratorio proposto/accettato');
    const dateTimeColIdx = headers.indexOf('Data e ora proposta/accettata');
    const consiglioAiColIdx = headers.indexOf('consiglio AI');
    if ([idColIdx, acceptedColIdx, labNameColIdx, dateTimeColIdx, consiglioAiColIdx].includes(-1)) {
      throw new Error(`Mancano colonne necessarie in ${spreadsheet.getName()}: firebase_id, Proposta accettata, Nome laboratorio..., Data e ora..., consiglio AI`);
    }

    Logger.log('Lettura dati assegnazioni esistenti (Primarie e Secondarie)...');
    const assignmentsDataPrimarie = FirebaseService.firebaseGet(CONFIG.ASSEGNAZIONI_PRIMARIE_NODE) || {};
    const assignmentsDataSecondarie = FirebaseService.firebaseGet(CONFIG.ASSEGNAZIONI_SECONDARIE_NODE) || {};
    // Uniamo i dati per avere una visione globale dello storico dell'ID
    const assignmentsData = { ...assignmentsDataPrimarie, ...assignmentsDataSecondarie };
    
    const existingNoOrToProcessRecords = new Set();
    const maxCounterPerFirebaseId = new Map();

    // 1. SCANSIONE STORICO: Trova il valore massimo del contatore per ogni gruppo (firebase_id)
    for (const id in assignmentsData) {
      const record = assignmentsData[id];
      const fId = record.id_firebase;
      if (!fId) continue;
      
      const proposalStatus = (record.proposta_accettata || '').toUpperCase();
      const duplicationKey = `${fId}_${proposalStatus}_${record.nome_lab}_${record.data_lab}`;
      if (['NO', 'SI', 'SÌ', 'PROPOSTA DA ELABORARE'].includes(proposalStatus)) {
        existingNoOrToProcessRecords.add(duplicationKey);
      }

      const storedCounter = parseInt(record.contatore_no || '0', 10);
      const currentMax = maxCounterPerFirebaseId.get(fId) || 0;
      
      if (storedCounter > currentMax) {
        maxCounterPerFirebaseId.set(fId, storedCounter);
      } else if (!maxCounterPerFirebaseId.has(fId)) {
        maxCounterPerFirebaseId.set(fId, 0);
      }
    }

    let successCount = 0;
    let errorCount = 0;
    const assignmentBatch = {};

    // 2. ELABORAZIONE FILE
    data.forEach((row) => {
      const firebaseId = row[idColIdx];
      const labName = String(row[labNameColIdx]).trim();
      const labDateTime = row[dateTimeColIdx];
      const isAcceptedRaw = String(row[acceptedColIdx]).trim();
      const isAccepted = isAcceptedRaw.toUpperCase();

      let acceptedStatus = null;
      if (isAccepted === 'NO') acceptedStatus = 'NO';
      else if (['SÌ', 'SI', 'PROPOSTA DA ELABORARE'].includes(isAccepted)) acceptedStatus = 'SI_OR_TO_PROCESS';

      if (firebaseId && acceptedStatus) {
        let currentMaxCount = maxCounterPerFirebaseId.get(firebaseId) || 0;

        // --- CASO 1: SI o PROPOSTA DA ELABORARE ---
        if (acceptedStatus === 'SI_OR_TO_PROCESS') {
          const rawKeyUpper = isAcceptedRaw.toUpperCase();
          const duplicationKey = `${firebaseId}_${rawKeyUpper}_${labName}_${labDateTime}`;
          
          if (existingNoOrToProcessRecords.has(duplicationKey)) {
             return;
          }
          
          const payload = {
            id_firebase: firebaseId,
            proposta_accettata: isAcceptedRaw,
            nome_lab: labName,
            data_lab: labDateTime,
            contatore_no: currentMaxCount 
          };
          assignmentBatch[Utilities_.generateUniqueId()] = payload;
          existingNoOrToProcessRecords.add(duplicationKey);
          successCount++;
        }

        // --- CASO 2: NO ---
        else if (acceptedStatus === 'NO') {
          const duplicationKey = `${firebaseId}_NO_${labName}_${labDateTime}`;
          if (existingNoOrToProcessRecords.has(duplicationKey)) {
            return;
          }
          
          const consiglioAi = String(row[consiglioAiColIdx] || '').trim().toUpperCase();
          let newCount;
          
          // Logic aggiornata per gestione fasi esplicite
          const standardPhases = ['ALTA DOMANDA', 'BASSA DOMANDA', 'EXTRA (RIPIEGO)', 'EXTRA (TEMATICO)'];
          
          if (standardPhases.includes(consiglioAi)) {
              newCount = 1;
          } else if (consiglioAi.includes('FASE 1')) {
              newCount = 2;
          } else if (consiglioAi.includes('FASE 2')) {
              newCount = 3;
          } else {
              newCount = currentMaxCount + 1;
          }

          maxCounterPerFirebaseId.set(firebaseId, newCount);

          const payload = {
            id_firebase: firebaseId,
            proposta_accettata: 'NO',
            nome_lab: labName, 
            data_lab: labDateTime,
            contatore_no: newCount 
          };
          assignmentBatch[Utilities_.generateUniqueId()] = payload;
          existingNoOrToProcessRecords.add(duplicationKey);
          successCount++;
        }
      }
    });

    // Scrittura batch finale su Firebase
    try {
      if (Object.keys(assignmentBatch).length > 0) {
        Logger.log(`Invio di ${Object.keys(assignmentBatch).length} assegnazioni in batch a ${targetAssegnazioniNode}...`);
        FirebaseService.firebasePatch(targetAssegnazioniNode, assignmentBatch);
      } else {
        Logger.log('Nessuna nuova assegnazione da aggiornare.');
      }
    } catch (e) {
      Logger.log(`🛑 ERRORE durante l'invio batch: ${e.toString()}`);
      errorCount = successCount; 
      successCount = 0;
    }

    return { successCount, errorCount };
  }

  /**
   * Ripopola un singolo nodo Firebase da uno Spreadsheet specifico.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet L'oggetto Spreadsheet da cui leggere.
   * @param {string} targetNode Il nodo Firebase da cancellare e ripopolare.
   * @returns {number} Il numero di record inviati.
   */
  function repopulateNodeFromSpreadsheet(spreadsheet, targetNode) {
    Logger.log(`--- AVVIO POPOLAMENTO per nodo: ${targetNode} ---`);
    Logger.log(`Lettura da file: ${spreadsheet.getName()}`);

    // 1. Cancella il nodo di destinazione
    Logger.log(`Cancellazione nodo: ${targetNode}`);
    FirebaseService.firebaseDelete(targetNode);
    // 2. Leggi i dati dal foglio specificato
    const sheet = spreadsheet.getSheetByName(CONFIG.REPOPULATE_SOURCE_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Foglio "${CONFIG.REPOPULATE_SOURCE_SHEET_NAME}" non trovato nel file "${spreadsheet.getName()}"`);
    }
    
    const sheetData = _getSheetDataAsObjects(sheet);
    if (!sheetData || sheetData.length === 0) {
      Logger.log('Nessun dato trovato nel foglio di calcolo.');
      return 0;
    }
    
    // 3. Crea i payload
    const allPayloads = sheetData
      .map(_createFirebasePayload) 
      .filter(p => p && p.classe_livello && p.laboratori_richiesti);
    if (allPayloads.length === 0) {
      Logger.log('Nessun record valido trovato nel foglio dopo il parsing. Controlla i dati sorgente.');
      return 0;
    }

    // 4. Prepara il batch
    const batchPayload = {};
    allPayloads.forEach(payload => {
      const uniqueId = Utilities_.generateUniqueId();
      batchPayload[uniqueId] = payload;
    });
    // 5. Invia il batch
    _sendBatchData(targetNode, batchPayload);
    Logger.log(`Blocco da ${allPayloads.length} record inviato a ${targetNode}.`);

    return allPayloads.length;
  }

  /**
   * Invia dati batch a Firebase, gestendo la suddivisione in chunk.
   * @param {string} node Il nodo di destinazione.
   * @param {Object} batchPayload L'oggetto payload completo.
   */
  function _sendBatchData(node, batchPayload) {
    const keys = Object.keys(batchPayload);
    for (let i = 0; i < keys.length; i += CONFIG.BATCH_SIZE) {
      const chunkKeys = keys.slice(i, i + CONFIG.BATCH_SIZE);
      const chunkPayload = {};
      chunkKeys.forEach(key => {
        chunkPayload[key] = batchPayload[key];
      });
      FirebaseService.firebasePatch(node, chunkPayload);
      Logger.log(`Chunk da ${chunkKeys.length} record inviato a ${node}.`);
    }
  }


  /**
   * MODIFICATO: Legge i dati dal foglio fornito e li converte in oggetti.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet L'oggetto foglio da cui leggere.
   * @returns {Array<Object>} Array di oggetti "namedValues".
   */
  function _getSheetDataAsObjects(sheet) {
    if (!sheet) throw new Error("_getSheetDataAsObjects: sheet non è valido.");
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    Logger.log('Intestazioni lette dal foglio: ' + JSON.stringify(headers));
    return data.map(row => {
      const namedValues = {};
      headers.forEach((header, index) => {
        namedValues[header] = [row[index]]; // Mantiene il formato array di e.namedValues
      });
      return namedValues;
    });
  }

  /**
   * Crea il payload JSON per Firebase da una riga/risposta.
   * @private
   * @param {Object} submittedData Un oggetto stile e.namedValues.
   * @returns {Object} Il payload JSON pulito.
   */
  function _createFirebasePayload(submittedData) {
    const payload = {};
    const getValue = (question) => {
      const value = submittedData[question];
      if (value) {
        const firstValue = Array.isArray(value) ? value[0] : value;
        if (typeof firstValue === 'string') return firstValue.trim();
        return firstValue;
      }
      return null;
    };
    
    const COL_CLASSE = "Classe per cui si manifesta l'interesse";
    const LISTA_COLONNE_LAB = [
      "Laboratori d'interesse per le Classi I (selezionarne un massimo di 4)",
      "Laboratori d'interesse per le Classi II (selezionarne un massimo di 4)",
      "Laboratori d'interesse per le Classi III (selezionarne un massimo di 4)",
      "Laboratori d'interesse per le Classi IV (selezionarne un massimo di 4)",
      "Laboratori d'interesse per le Classi V (selezionarne un massimo di 4)"
    ];

    payload.timestamp = getValue("Informazioni cronologiche") || new Date().toISOString();
    payload.email = getValue("Indirizzo email");
    const consensoInformativa = getValue("Ho preso visione e compreso l’informativa relativa al trattamento dei dati personali (clicca per visualizzare e scaricare)");
    if (consensoInformativa) payload.consenso_informativa = (consensoInformativa === "Sì");
    
    const consensoNewsletter = getValue("Presto il CONSENSO a ricevere informazioni su altri progetti di public engagement per la scuola");
    if (consensoNewsletter) payload.consenso_newsletter = (consensoNewsletter === "Sì");
    
    payload.istituto_nome = getValue("Nome dell'Istituto");
    payload.istituto_codice_mecc = getValue("Codice meccanografico dell'Istituto (Es.:TO1A005001)");
    payload.istituto_circoscrizione = getValue("Circoscrizione del Comune di Torino in cui si trova l'Istituto (solo il numero)");
    payload.plesso_nome = getValue("Plesso di appartenenza della classe");
    payload.plesso_indirizzo = getValue("Indirizzo del plesso");
    
    const nome = getValue("Nome insegnante referente");
    const cognome = getValue("Cognome insegnante referente");
    if (nome || cognome) {
      payload.insegnante_referente = `${nome || ''} ${cognome || ''}`.trim();
    }
    
    payload.insegnante_cellulare = getValue("Numero cellulare insegnante referente");
    payload.classe_sezione = getValue("Sezione della classe per cui si manifesta l'interesse");
    payload.classe_studenti_numero = getValue("Numero degli studenti della sezione");
    
    const classeValue = getValue(COL_CLASSE);
    if (!classeValue) {
       return payload;
    }
    payload.classe_livello = classeValue;

    for (const colName of LISTA_COLONNE_LAB) {
        const labValue = getValue(colName);
        if (labValue) { 
            payload.laboratori_richiesti = labValue.split(',').map(item => item.trim());
            break;
        }
    }
    
    if (!payload.laboratori_richiesti) {
        Logger.log(`ATTENZIONE: Nessun dato laboratorio trovato per la riga con classe "${classeValue}" (Email: ${payload.email}). La riga verrà saltata dal filtro.`);
    }

    const disabilitaPresente = getValue("Nella classe sono presenti persone con disabilità?");
    if (disabilitaPresente) payload.disabilita_presente = (disabilitaPresente === "Sì");
    
    const tipiDisabilita = getValue("Selezionare le tipologie presenti");
    if (tipiDisabilita) {
      payload.disabilita_tipologie = tipiDisabilita.split(',').map(item => item.trim());
    }

    // Pulisce valori nulli o vuoti
    for (const key in payload) {
      if (payload[key] === null || payload[key] === '') {
        delete payload[key];
      }
    }
    return payload;
  }

  /**
   * Determina la destinazione (Primaria/Secondaria) in base al NOME DEL FILE.
   * @private
   * @param {string} fileName Il nome del file (spreadsheet) da cui arrivano i dati.
   * @returns {string} CONFIG.PRIMARY_DESTINATION o CONFIG.SECONDARY_DESTINATION
   */
  function _getDestinazione(fileName) {
    const normalizedName = String(fileName || '').toUpperCase();
    // Se il nome del file contiene "SECONDO GRADO" o "SECONDARIE"
    if (normalizedName.includes('SECONDO GRADO') || normalizedName.includes('SECONDARIE')) {
      return CONFIG.SECONDARY_DESTINATION;
    }
    return CONFIG.PRIMARY_DESTINATION; // Default
  }
  
  /**
   * Sincronizza le NUOVE righe dal foglio a Firebase.
   * Legge il foglio, verifica se la riga esiste già su Firebase (tramite Timestamp+Email),
   * e se è nuova la inserisce nel nodo appropriato (Primarie/Secondarie),
   * scrivendo poi il nuovo ID sul foglio.
   * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Il file da cui leggere.
   * @returns {number} Il numero di nuove righe aggiunte.
   */
  function syncNewRequestsToFirebase(spreadsheet) {
    Logger.log(`--- AVVIO SYNC NUOVE RIGHE da: ${spreadsheet.getName()} ---`);
    
    const sheet = spreadsheet.getSheetByName(CONFIG.REPOPULATE_SOURCE_SHEET_NAME);
    if (!sheet) throw new Error(`Foglio "${CONFIG.REPOPULATE_SOURCE_SHEET_NAME}" non trovato.`);

    const rawSheetObjects = _getSheetDataAsObjects(sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColIdx = headers.indexOf('firebase_id');
    if (idColIdx === -1) throw new Error("Colonna 'firebase_id' mancante.");

    Logger.log('Lettura dati esistenti su Firebase...');
    const fbPrim = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_PRIMARY) || {};
    const fbSec = FirebaseService.firebaseGet(CONFIG.REQUESTS_NODE_SECONDARY) || {};
    
    const existingMap = new Map();
    const populateMap = (data) => {
      for (const id in data) {
        const r = data[id];
        if (r.timestamp && r.email) {
          const ts = new Date(r.timestamp).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
          existingMap.set(`${ts}_${r.email.trim()}`, id);
        }
      }
    };
    populateMap(fbPrim);
    populateMap(fbSec);

    const batchPrim = {};
    const batchSec = {};
    const sheetUpdates = [];
    let addedCount = 0;

    rawSheetObjects.forEach((rowObj, index) => {
      const sheetRow = index + 2;
      const currentSheetId = rowObj['firebase_id'] ? String(rowObj['firebase_id'][0]).trim() : '';
      if (currentSheetId) return;

      const payload = _createFirebasePayload(rowObj);
      if (!payload.email || !payload.timestamp || !payload.classe_livello) return;

      const tsSheet = new Date(payload.timestamp).toLocaleString('it-IT', { timeZone: CONFIG.SCRIPT_TIMEZONE });
      const dupKey = `${tsSheet}_${payload.email.trim()}`;

      if (existingMap.has(dupKey)) {
        const existingId = existingMap.get(dupKey);
        Logger.log(`Riga ${sheetRow}: Trovata su Firebase (${existingId}) ma vuota su foglio. Aggiorno foglio.`);
        sheetUpdates.push({ row: sheetRow, col: idColIdx + 1, val: existingId });
        return;
      }

      // --- MODIFICA: Determina destinazione dal nome del File ---
      const destinazione = _getDestinazione(spreadsheet.getName());
      // --------------------------------------------------------

      const newId = Utilities_.generateUniqueId();
      if (destinazione === CONFIG.SECONDARY_DESTINATION) {
        batchSec[newId] = payload;
      } else {
        batchPrim[newId] = payload;
      }

      sheetUpdates.push({ row: sheetRow, col: idColIdx + 1, val: newId });
      addedCount++;
    });

    // 4. Invio a Firebase (Batch)
    if (Object.keys(batchPrim).length > 0) {
      Logger.log(`Inserimento ${Object.keys(batchPrim).length} nuove righe in Primarie...`);
      FirebaseService.firebasePatch(CONFIG.REQUESTS_NODE_PRIMARY, batchPrim);
    }
    if (Object.keys(batchSec).length > 0) {
      Logger.log(`Inserimento ${Object.keys(batchSec).length} nuove righe in Secondarie...`);
      FirebaseService.firebasePatch(CONFIG.REQUESTS_NODE_SECONDARY, batchSec);
    }

    // 5. Aggiornamento Foglio
    if (sheetUpdates.length > 0) {
      Logger.log(`Scrittura di ${sheetUpdates.length} ID sul foglio...`);
      sheetUpdates.forEach(u => {
        sheet.getRange(u.row, u.col).setValue(u.val);
      });
    }

    Logger.log(`--- Completato. Aggiunte ${addedCount} nuove richieste. ---`);
    return addedCount;
  }

  // Interfaccia pubblica
  return {
    handleFormSubmit: handleFormSubmit,
    syncFirebaseIdsToSheet: syncFirebaseIdsToSheet,
    syncProposalsFromFile: syncProposalsFromFile,
    repopulateNodeFromSpreadsheet: repopulateNodeFromSpreadsheet,
    syncNewRequestsToFirebase: syncNewRequestsToFirebase
  };
})();