/**
 * @fileoverview Funzioni di utilità generiche.
 */

var Utilities_ = (function() {

  return {
    /**
     * Ottiene l'oggetto UI di SpreadsheetApp.
     * @returns {GoogleAppsScript.Base.Ui} L'oggetto UI.
     */
    getUi: function() {
      return SpreadsheetApp.getUi();
    },

    /**
     * Ottiene un foglio tramite il nome DAL FOGLIO ATTIVO. Lancia un errore se non trovato.
     * @param {string} sheetName Il nome del foglio.
     * @returns {GoogleAppsScript.Spreadsheet.Sheet} L'oggetto Foglio.
     */
    getSheet: function(sheetName) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Foglio "${sheetName}" non trovato nel foglio di calcolo attivo.`);
      }
      return sheet;
    },
    
    /**
     * NUOVA FUNZIONE: Trova uno Spreadsheet in Drive tramite nome.
     * @param {string} fileName Il nome del file (o parte di esso).
     * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null} Lo spreadsheet.
     */
    findSpreadsheetByName: function(fileName) {
      try {
        const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
        const files = folder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          if (file.getName().includes(fileName)) {
            Logger.log(`Trovato file: ${file.getName()}`);
            return SpreadsheetApp.open(file);
          }
        }
      } catch (e) {
        Logger.log(`Errore durante la ricerca del file ${fileName}: ${e.toString()}`);
      }
      Logger.log(`File non trovato: ${fileName}`);
      return null;
    },

    /**
     * Converte un oggetto Firebase (con chiavi) in un array di oggetti.
     * @param {Object} firebaseObject L'oggetto da Firebase.
     * @returns {Array<Object>} Un array di oggetti, ognuno con la proprietà '__firebaseKey'.
     */
    toArray: function(firebaseObject) {
      if (!firebaseObject) return [];
      return Object.keys(firebaseObject).map(key => ({
        ...firebaseObject[key],
        __firebaseKey: key
      }));
    },

    /**
     * Ottiene un valore da un oggetto, provando più chiavi in ordine.
     * Restituisce una stringa trimmata o null.
     * @param {Object} req L'oggetto da cui leggere.
     * @param {...string} keys Le chiavi da provare.
     * @returns {string|null} Il valore trovato o null.
     */
    getField: function(req, ...keys) {
      for (const key of keys) {
        if (req && req[key] !== undefined && req[key] !== null && req[key] !== "") {
          return String(req[key]).trim();
        }
      }
      return null;
    },

    /**
     * Tenta di pulire e normalizzare una stringa di data.
     * Gestisce solo la manipolazione di stringhe per evitare errori di parsing.
     * @param {string|Date} dateString La stringa o oggetto data.
     * @returns {string|null} La data normalizzata (dd/MM/yyyy HH:mm:ss) o null/originale.
     */
    cleanDateString: function(dateString) {
      if (!dateString) return null;

      if (dateString instanceof Date && !isNaN(dateString.getTime())) {
        return Utilities.formatDate(dateString, CONFIG.SCRIPT_TIMEZONE, "dd/MM/yyyy HH:mm:ss");
      }

      if (typeof dateString !== 'string') return null;

      const trimmed = dateString.trim();
      const match = trimmed.match(/^(\d{1,2})[\/\\](\d{1,2})[\/\\](\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
      
      if (match) {
        const day = match[1].padStart(2, '0');
        const month = match[2].padStart(2, '0');
        const year = match[3];
        const hours = match[4].padStart(2, '0');
        const minutes = match[5].padStart(2, '0');
        const seconds = (match[6] || '00').padStart(2, '0');
        return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
      }

      Logger.log(`ATTENZIONE: Formato data non standardizzato: '${trimmed}'. Verrà usato il valore grezzo.`);
      return trimmed;
    },

    /**
     * Formatta una data (potenzialmente da un formato grezzo) per l'output finale.
     * Esegue solo la normalizzazione come stringa (alias di cleanDateString).
     * @param {string|Date} rawDate La data grezza.
     * @returns {string} La data normalizzata come stringa.
     */
    formatRawDateForOutput: function(rawDate) {
      const cleaned = Utilities_.cleanDateString(rawDate);
      return cleaned || ''; // Restituisce stringa pulita o stringa vuota
    },

    /**
     * Genera un ID univoco in stile Firebase (Push ID).
     * @returns {string} Un ID univoco.
     */
    generateUniqueId: function() {
      const PUSH_CHARS = '-0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz';
      let lastPushTime = 0,
        lastRandChars = [];
      let now = new Date().getTime(),
        duplicateTime = (now === lastPushTime);
      lastPushTime = now;
      let timeStampChars = new Array(8);
      for (let i = 7; i >= 0; i--) {
        timeStampChars[i] = PUSH_CHARS.charAt(now % 64);
        now = Math.floor(now / 64);
      }
      let id = timeStampChars.join('');
      if (!duplicateTime) {
        for (let i = 0; i < 12; i++) {
          lastRandChars[i] = Math.floor(Math.random() * 64);
        }
      } else {
        for (let i = 11; i >= 0 && lastRandChars[i] === 63; i--) {
          lastRandChars[i] = 0;
        }
        lastRandChars[i]++;
      }
      for (let i = 0; i < 12; i++) {
        id += PUSH_CHARS.charAt(lastRandChars[i]);
      }
      return id;
    }
  };
})();