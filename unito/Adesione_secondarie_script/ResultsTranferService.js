/**
 * @fileoverview Servizio per trasferire i risultati del matching
 * da Firebase ai fogli di calcolo "Adesioni" finali.
 */

var ResultsTransferService = (function() {

  /**
   * Esegue la sincronizzazione dei risultati SOLO per le Primarie.
   * @returns {number} Numero di righe aggiornate.
   */
  function syncResultsPrimarie() {
    Logger.log(`Recupero dati da: ${CONFIG.MATCHING_RESULTS_PRIMARY_NODE}...`);
    const ssPrimarie = _findSpreadsheet(CONFIG.ADESIONI_PRIMARIE_FILENAME);
    if (!ssPrimarie) {
      Logger.log('❌ File Adesioni Primarie non trovato');
      return 0;
    }
    
    const dataParent = FirebaseService.firebaseGet(CONFIG.MATCHING_RESULTS_PRIMARY_NODE);
    return _syncLatestResults(dataParent, ssPrimarie, "PRIMARIE");
  }

  /**
   * Esegue la sincronizzazione dei risultati SOLO per le Secondarie.
   * @returns {number} Numero di righe aggiornate.
   */
  function syncResultsSecondarie() {
    Logger.log(`Recupero dati da: ${CONFIG.MATCHING_RESULTS_SECONDARY_NODE}...`);
    const ssSecondarie = _findSpreadsheet(CONFIG.ADESIONI_SECONDARIE_FILENAME);
    if (!ssSecondarie) {
      Logger.log('❌ File Adesioni Secondarie non trovato');
      return 0;
    }
    
    const dataParent = FirebaseService.firebaseGet(CONFIG.MATCHING_RESULTS_SECONDARY_NODE);
    return _syncLatestResults(dataParent, ssSecondarie, "SECONDARIE");
  }

  /**
   * Esegue la sincronizzazione dei risultati per entrambe.
   */
  function syncAllResults() {
    const countP = syncResultsPrimarie();
    const countS = syncResultsSecondarie();
    Logger.log(`✅ Aggiornamento completato. Primarie: ${countP}, Secondarie: ${countS}.`);
  }

  /**
   * Trova il foglio di calcolo "Adesioni" nella cartella specificata.
   * @private
   * @param {string} fileName Il nome (o parte del nome) del file.
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null} Lo spreadsheet.
   */
  function _findSpreadsheet(fileName) {
    try {
      const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName().includes(fileName)) {
          return SpreadsheetApp.open(file);
        }
      }
    } catch (e) {
      Logger.log(`Errore durante la ricerca del file ${fileName}: ${e.toString()}`);
    }
    return null;
  }

  /**
   * Trova i dati del timestamp più recente e avvia la sincronizzazione.
   * @private
   */
  function _syncLatestResults(dataParent, ssDestinazione, livello) {
    if (!dataParent) {
      Logger.log(`❌ Dati non trovati per ${livello}. Assicurati di eseguire prima il matching.`);
      return 0;
    }
    
    const allTimestamps = Object.keys(dataParent);
    if (allTimestamps.length === 0) {
      Logger.log(`❌ Il nodo ${livello} esiste ma è vuoto. Nessun risultato da sincronizzare.`);
      return 0;
    }
    
    allTimestamps.sort();
    const latestTimestampKey = allTimestamps[allTimestamps.length - 1];
    Logger.log(`   -> Trovato sotto-nodo ${livello} più recente: ${latestTimestampKey}`);
    
    const dataToSync = dataParent[latestTimestampKey];
    return _sincronizzaLivello(dataToSync, ssDestinazione, livello);
  }



  /**
   * Logica di sincronizzazione effettiva su un foglio di destinazione.
   * MODIFICATO: Scrive SOLO nelle colonne target per preservare i link nelle altre colonne.
   * @private
   */
  function _sincronizzaLivello(datiAssegnazioni, ssDestinazione, livello) {
    if (!ssDestinazione || !datiAssegnazioni) {
      Logger.log(`❌ Dati o file Adesioni per ${livello} non validi, sincronizzazione saltata.`);
      return 0;
    }

    const shAdesioni = ssDestinazione.getSheets()[0];
    if (!shAdesioni) throw new Error(`❌ Nessun foglio trovato nel file Adesioni ${livello}.`);

    // === 1. MAPPATURA DATI ASSEGNAZIONI (Da Firebase) ===
    const dataA = Object.values(datiAssegnazioni);
    const mappa = {};
    dataA.forEach(row => {
      const id = row.id;
      if (id) {
        let dataText = String(row.dataAssegnata || '');
        if (dataText.startsWith("'")) {
          dataText = dataText.substring(1);
        }
        let labFinale = row.labAssegnato || row.labRichiesto || '';

        mappa[id] = {
          dataAssegnata: dataText,
          labAssegnato: labFinale,
          tipoLab: row.faseAssegnazione,
          sedeLab: row.sede,
          durataLab: row.durata_incontro
        };
      }
    });

    // === 2. LETTURA INTESTAZIONI ===
    const lastRow = shAdesioni.getLastRow();
    if (lastRow < 2) return 0; // Solo intestazione o vuoto

    const lastCol = shAdesioni.getLastColumn();
    const headers = shAdesioni.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const firebaseIdIndex = headers.indexOf('firebase_id');
    if (firebaseIdIndex === -1) {
      throw new Error(`❌ Colonna obbligatoria 'firebase_id' mancante nel file Adesioni ${livello}.`);
    }

    // Recupera tutti gli ID presenti nel foglio (per sapere quale riga aggiornare)
    // Nota: +1 perché getRange è 1-based, e saltiamo la riga intestazione
    const sheetIds = shAdesioni.getRange(2, firebaseIdIndex + 1, lastRow - 1, 1).getValues().map(r => r[0]);

    let rowsUpdatedTotal = 0;

    // === 3. AGGIORNAMENTO COLONNA PER COLONNA (Non distruttivo) ===
    // Invece di riscrivere tutto il foglio, iteriamo solo sulle colonne che dobbiamo toccare.
    
    CONFIG.ADESIONI_COLUMNS_TO_UPDATE.forEach(colName => {
      const colIndex = headers.indexOf(colName);
      if (colIndex === -1) {
        Logger.log(`⚠️ Colonna "${colName}" non trovata nel foglio. Salto.`);
        return;
      }

      // Leggiamo i valori attuali DI QUESTA SOLA COLONNA
      const colRange = shAdesioni.getRange(2, colIndex + 1, lastRow - 1, 1);
      
      // Imposta formato testo se necessario per date/ore
      if (['Data e ora proposta/accettata', 'Sede incontro', 'Durata incontro'].includes(colName)) {
        colRange.setNumberFormat('@');
      }

      const colValues = colRange.getValues();
      let colModified = false;

      // Iteriamo su ogni riga del foglio
      for (let i = 0; i < sheetIds.length; i++) {
        const fid = sheetIds[i];
        
        // Se abbiamo dati nuovi per questo ID
        if (mappa[fid]) {
          const mapData = mappa[fid];
          let newValue = undefined;

          // Determina il valore da scrivere in base alla colonna
          switch (colName) {
            case 'consiglio AI':
              newValue = mapData.tipoLab;
              break;
            case 'Proposta accettata':
              // FIX: Se non c'è una data (es. NESSUNA DISPONIBILITÀ), il campo deve essere vuoto.
              if (!mapData.dataAssegnata) {
                  newValue = "";
              } else {
                  newValue = "proposta da elaborare";
              }
              break;
            case 'Nome laboratorio proposto/accettato':
              newValue = mapData.labAssegnato;
              break;
            case 'Data e ora proposta/accettata':
              newValue = mapData.dataAssegnata;
              break;
            case 'Sede incontro':
              newValue = mapData.sedeLab;
              break;
            case 'Durata incontro':
              newValue = mapData.durataLab;
              break;
          }

          // Se c'è un valore valido, aggiorniamo l'array in memoria
          if (newValue !== undefined && newValue !== null) {
            colValues[i][0] = newValue;
            colModified = true;
            // Contiamo l'aggiornamento solo una volta per ID (usiamo un set o contatore approssimativo)
          }
        }
      }

      // Se la colonna ha subito modifiche, scriviamo SOLO quella colonna
      if (colModified) {
        colRange.setValues(colValues);
      }
    });

    // Calcolo approssimativo delle righe toccate (basato sulla mappa)
    rowsUpdatedTotal = Object.keys(mappa).filter(id => sheetIds.includes(id)).length;

    Logger.log(`✅ ${livello}: Aggiornate le colonne target per circa ${rowsUpdatedTotal} righe.`);
    return rowsUpdatedTotal;
  }

  // Interfaccia pubblica
  return {
    syncResultsPrimarie: syncResultsPrimarie,
    syncResultsSecondarie: syncResultsSecondarie,
    syncAllResults: syncAllResults
  };

})();