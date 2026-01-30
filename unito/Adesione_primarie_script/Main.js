/**
 * @fileoverview Funzioni principali, gestione del menu UI e trigger.
 */

/**
 * Crea un menu personalizzato all'apertura del foglio di calcolo.
 */
function onOpen() {
  Utilities_.getUi()
    .createMenu('Firebase Admin')
    // MODIFICATO: Diviso in due menu
    
    //.addItem('1b. Popola IDs (SECONDARIE)', 'uiPopulateFirebaseIds_Secondarie')
    //.addSeparator()
    // MODIFICATO: Diviso in due menu
    .addItem('1. Aggiorna Proposte', 'uiUpdateAcceptedProposals_Primarie')
    //.addItem('2b. Aggiorna Proposte (SECONDARIE)', 'uiUpdateAcceptedProposals_Secondarie')
    .addSeparator()
    //.addItem('3. Run Assegnazioni AI (SOLO SECONDARIE)', 'uiRunMatchingSecondarie')
    .addItem('2. Assegnazioni AI', 'uiRunMatchingPrimarie')
    // --- AMMINISTRAZIONE ---
    .addSeparator()
    .addItem('Aggiungi le nuove richieste', 'uiSyncNewRequests_Primarie') // <--- NUOVO
    //.addItem('[Admin] Sync NUOVE Righe (Secondarie)', 'uiSyncNewRequests_Secondarie') // <--- NUOVO
    .addSeparator()
    .addItem('[Admin] 1. Ricarica da File', 'uiRepopulatePrimarie')
    .addItem('[Admin] 2. Popola IDs ', 'uiPopulateFirebaseIds_Primarie')
    //.addItem('PERICOLOSO: Ricarica SECONDARIE da File', 'uiRepopulateSecondarie')
    .addToUi();
}

/**
 * Gestisce il trigger onFormSubmit.
 * (Questa funzione è l'UNICA che usa il foglio 'ORIGINAL_FORM_SHEET_NAME')
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e L'oggetto evento.
 */
function onFormSubmitHandler(e) {
  try {
    SheetSyncService.handleFormSubmit(e);
  } catch (error) {
    Logger.log('ERRORE CRITICO in onFormSubmitHandler: ' + error.toString());
  }
}

// ====================================================================
// WRAPPER UI (Funzioni chiamate dal menu)
// ====================================================================

/**
 * NUOVA FUNZIONE: Legge dal file sorgente PRIMARIE e popola gli ID.
 */
function uiPopulateFirebaseIds_Primarie() {
  const ui = Utilities_.getUi();
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_PRIMARY;
  
  if (ui.alert(`Stai per leggere dal file sorgente "${fileName}" e popolare gli ID Firebase mancanti. Continuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    const ssPrimarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssPrimarie) {
      throw new Error(`File sorgente Primarie non trovato: ${fileName}`);
    }
    
    Logger.log('--- Popolamento ID per Primarie ---');
    const countP = SheetSyncService.syncFirebaseIdsToSheet(ssPrimarie);
    
    Logger.log(`Completato. Trovati ${countP} ID da popolare.`);
    ui.alert(`Operazione completata per le Primarie.\n\n${countP} ID popolati.`);
  } catch (e) {
    Logger.log(`ERRORE in uiPopulateFirebaseIds_Primarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}

/**
 * NUOVA FUNZIONE: Legge dal file sorgente SECONDARIE e popola gli ID.
 */
function uiPopulateFirebaseIds_Secondarie() {
  const ui = Utilities_.getUi();
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_SECONDARY;

  if (ui.alert(`Stai per leggere dal file sorgente "${fileName}" e popolare gli ID Firebase mancanti. Continuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    const ssSecondarie = Utilities_.findSpreadsheetByName(fileName);
     if (!ssSecondarie) {
      throw new Error(`File sorgente Secondarie non trovato: ${fileName}`);
    }
    
    Logger.log('--- Popolamento ID per Secondarie ---');
    const countS = SheetSyncService.syncFirebaseIdsToSheet(ssSecondarie);

    Logger.log(`Completato. Trovati ${countS} ID da popolare.`);
    ui.alert(`Operazione completata per le Secondarie.\n\n${countS} ID popolati.`);
  } catch (e) {
    Logger.log(`ERRORE in uiPopulateFirebaseIds_Secondarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}


/**
 * NUOVA FUNZIONE: Legge dal file sorgente PRIMARIE e aggiorna le proposte.
 */
function uiUpdateAcceptedProposals_Primarie() {
  const ui = Utilities_.getUi();
  const nodeName = CONFIG.ASSEGNAZIONI_PRIMARIE_NODE;
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_PRIMARY; // Usa lo stesso file sorgente

  if (ui.alert(`Stai per leggere dal file sorgente "${fileName}" e aggiornare le assegnazioni "SI/NO" sul nodo "${nodeName}".\n\nContinuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    const ssPrimarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssPrimarie) {
      throw new Error(`File sorgente Primarie non trovato: ${fileName}`);
    }

    Logger.log('--- Aggiornamento proposte per Primarie ---');
    const statsP = SheetSyncService.syncProposalsFromFile(ssPrimarie, nodeName);

    Logger.log(`Completato. Righe inviate: ${statsP.successCount}, Errori: ${statsP.errorCount}`);
    ui.alert(`Operazione completata per le Primarie.\n\nRighe inviate: ${statsP.successCount}\nErrori: ${statsP.errorCount}`);
  } catch (e) {
    Logger.log(`ERRORE in uiUpdateAcceptedProposals_Primarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}

/**
 * NUOVA FUNZIONE: Legge dal file sorgente SECONDARIE e aggiorna le proposte.
 */
function uiUpdateAcceptedProposals_Secondarie() {
  const ui = Utilities_.getUi();
  const nodeName = CONFIG.ASSEGNAZIONI_SECONDARIE_NODE;
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_SECONDARY; // Usa lo stesso file sorgente

  if (ui.alert(`Stai per leggere dal file sorgente "${fileName}" e aggiornare le assegnazioni "SI/NO" sul nodo "${nodeName}".\n\nContinuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    const ssSecondarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssSecondarie) {
      throw new Error(`File sorgente Secondarie non trovato: ${fileName}`);
    }
    
    Logger.log('--- Aggiornamento proposte per Secondarie ---');
    const statsS = SheetSyncService.syncProposalsFromFile(ssSecondarie, nodeName);

    Logger.log(`Completato. Righe inviate: ${statsS.successCount}, Errori: ${statsS.errorCount}`);
    ui.alert(`Operazione completata per le Secondarie.\n\nRighe inviate: ${statsS.successCount}\nErrori: ${statsS.errorCount}`);
  } catch (e) {
    Logger.log(`ERRORE in uiUpdateAcceptedProposals_Secondarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}


/**
 * Wrapper UI per la funzione PERICOLOSA di ripopolamento
 * SOLO PER LE PRIMARIE.
 */
function uiRepopulatePrimarie() {
  const ui = Utilities_.getUi();
  const nodeName = CONFIG.REQUESTS_NODE_PRIMARY;
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_PRIMARY;
  
  const response = ui.alert('ATTENZIONE! PERICOLO!', `Questa operazione CANCELLERÀ IRREVERSIBILMENTE il nodo "${nodeName}" su Firebase e lo sostituirà con i dati letti dal file "${fileName}".\n\nContinuare?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('Operazione annullata.');
    return;
  }

  try {
    Logger.log(`--- AVVIO POPOLAMENTO (SOLO PRIMARIE) da ${fileName} ---`);
    
    const ssPrimarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssPrimarie) {
      throw new Error(`File sorgente Primarie non trovato: ${fileName}`);
    }

    const countPrimarie = SheetSyncService.repopulateNodeFromSpreadsheet(
      ssPrimarie,
      nodeName
    );
    
    ui.alert(`Popolamento Primarie completato! Inviati ${countPrimarie} record a "${nodeName}".`);

  } catch (e) {
    Logger.log('ERRORE CRITICO durante il popolamento Primarie: ' + e.toString());
    ui.alert('Si è verificato un errore critico. Controllare i log: ' + e.message);
  }
}

/**
 * Wrapper UI per la funzione PERICOLOSA di ripopolamento
 * SOLO PER LE SECONDARIE.
 */
function uiRepopulateSecondarie() {
  const ui = Utilities_.getUi();
  const nodeName = CONFIG.REQUESTS_NODE_SECONDARY;
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_SECONDARY;

  const response = ui.alert('ATTENZIONE! PERICOLO!', `Questa operazione CANCELLERÀ IRREVERSIBILMENTE il nodo "${nodeName}" su Firebase e lo sostituirà con i dati letti dal file "${fileName}".\n\nContinuare?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('Operazione annullata.');
    return;
  }

  try {
    Logger.log(`--- AVVIO POPOLAMENTO (SOLO SECONDARIE) da ${fileName} ---`);
    
    const ssSecondarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssSecondarie) {
      throw new Error(`File sorgente Secondarie non trovato: ${fileName}`);
    }

    const countSecondarie = SheetSyncService.repopulateNodeFromSpreadsheet(
      ssSecondarie,
      nodeName
    );
    
    ui.alert(`Popolamento Secondarie completato! Inviati ${countSecondarie} record a "${nodeName}".`);

  } catch (e) {
    Logger.log('ERRORE CRITICO durante il popolamento Secondarie: ' + e.toString());
    ui.alert('Si è verificato un errore critico. Controllare i log: ' + e.message);
  }
}


/**
 * Wrapper UI per eseguire il processo completo (Matching + Sincronizzazione Adesioni)
 * SOLO PER LE SECONDARIE.
 */
function uiRunMatchingSecondarie() {
  const ui = Utilities_.getUi();
  if (ui.alert('Stai per eseguire l\'algoritmo di assegnazione e aggiornare il foglio "Adesioni" SOLO per le SCUOLE SECONDARIE. Il processo può richiedere diversi minuti. Continuare?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    Logger.log("--- AVVIO FASE 1: PROCESSO DI MATCHING (SOLO SECONDARIE) ---");
    MatchingEngine.runMatchingSecondarie();
    Logger.log("--- ✅ PROCESSO DI MATCHING (SOLO SECONDARIE) COMPLETATO ---");
    
    Logger.log("--- AVVIO FASE 2: AGGIORNAMENTO FOGLI ADESIONE (SOLO SECONDARIE) ---");
    const count = ResultsTransferService.syncResultsSecondarie();
    Logger.log(`--- ✅ AGGIORNAMENTO FOGLI (SOLO SECONDARIE) COMPLETATO (${count} righe) ---`);

    ui.alert(`Processo completo per le Secondarie terminato. ${count} righe sono state aggiornate nel foglio Adesioni.`);
  } catch (e) {
    Logger.log(`🛑 ERRORE CRITICO nel processo completo (SECONDARIE): ${e.toString()}\nStack: ${e.stack}`);
    ui.alert(`Si è verificato un errore critico: ${e.message}. Controllare i log.`);
  }
}


/**
 * Wrapper UI per eseguire il processo completo (Matching + Sincronizzazione Adesioni)
 * SOLO PER LE PRIMARIE.
 */
function uiRunMatchingPrimarie() {
  const ui = Utilities_.getUi();
  if (ui.alert('Stai per eseguire l\'algoritmo di assegnazione e aggiornare il foglio "Adesioni" SOLO per le SCUOLE PRIMARIE. Il processo può richiedere diversi minuti. Continuare?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  try {
    Logger.log("--- AVVIO FASE 1: PROCESSO DI MATCHING (SOLO PRIMARIE) ---");
    MatchingEngine.runMatchingPrimarie();
    Logger.log("--- ✅ PROCESSO DI MATCHING (SOLO PRIMARIE) COMPLETATO ---");
    
    Logger.log("--- AVVIO FASE 2: AGGIORNAMENTO FOGLI ADESIONE (SOLO PRIMARIE) ---");
    const count = ResultsTransferService.syncResultsPrimarie();
    Logger.log(`--- ✅ AGGIORNAMENTO FOGLI (SOLO PRIMARIE) COMPLETATO (${count} righe) ---`);

    ui.alert(`Processo completo per le Primarie terminato. ${count} righe sono state aggiornate nel foglio Adesioni.`);
  } catch (e) {
    Logger.log(`🛑 ERRORE CRITICO nel processo completo (PRIMARIE): ${e.toString()}\nStack: ${e.stack}`);
    ui.alert(`Si è verificato un errore critico: ${e.message}. Controllare i log.`);
  }
}

/**
 * Wrapper UI per sincronizzare SOLO le nuove righe dal file PRIMARIE.
 */
function uiSyncNewRequests_Primarie() {
  const ui = Utilities_.getUi();
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_PRIMARY;

  if (ui.alert(`SYNC INCREMENTALE PRIMARIE\n\nStai per leggere il file "${fileName}".\n\nVerranno cercate righe che NON esistono su Firebase e aggiunte (generando il firebase_id).\nLe righe esistenti verranno ignorate.\n\nContinuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  try {
    const ssPrimarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssPrimarie) {
      throw new Error(`File sorgente Primarie non trovato: ${fileName}`);
    }
    
    Logger.log('--- Avvio Sync Incrementale Primarie ---');
    const addedCount = SheetSyncService.syncNewRequestsToFirebase(ssPrimarie);
    
    Logger.log(`Completato. Aggiunte ${addedCount} nuove righe.`);
    ui.alert(`Operazione completata per le Primarie.\n\nNuove richieste aggiunte: ${addedCount}`);

  } catch (e) {
    Logger.log(`ERRORE in uiSyncNewRequests_Primarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}

/**
 * Wrapper UI per sincronizzare SOLO le nuove righe dal file SECONDARIE.
 */
function uiSyncNewRequests_Secondarie() {
  const ui = Utilities_.getUi();
  const fileName = CONFIG.REPOPULATE_SOURCE_FILE_SECONDARY;

  if (ui.alert(`SYNC INCREMENTALE SECONDARIE\n\nStai per leggere il file "${fileName}".\n\nVerranno cercate righe che NON esistono su Firebase e aggiunte (generando il firebase_id).\nLe righe esistenti verranno ignorate.\n\nContinuare?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  try {
    const ssSecondarie = Utilities_.findSpreadsheetByName(fileName);
    if (!ssSecondarie) {
      throw new Error(`File sorgente Secondarie non trovato: ${fileName}`);
    }
    
    Logger.log('--- Avvio Sync Incrementale Secondarie ---');
    const addedCount = SheetSyncService.syncNewRequestsToFirebase(ssSecondarie);
    
    Logger.log(`Completato. Aggiunte ${addedCount} nuove righe.`);
    ui.alert(`Operazione completata per le Secondarie.\n\nNuove richieste aggiunte: ${addedCount}`);

  } catch (e) {
    Logger.log(`ERRORE in uiSyncNewRequests_Secondarie: ${e.toString()}\n${e.stack}`);
    ui.alert(`Errore: ${e.message}`);
  }
}