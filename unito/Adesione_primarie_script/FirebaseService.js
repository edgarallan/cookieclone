/**
 * @fileoverview Servizio centralizzato per la comunicazione con Firebase.
 * Utilizza un segreto (da Script Properties) per autenticare tutte le richieste.
 */

var FirebaseService = (function() {

  let _secret = null;

  /**
   * Ottiene il segreto di Firebase dalle proprietà dello script.
   * Se non trovato, usa il fallback (insicuro) da Config.gs.
   * @private
   * @returns {string} Il segreto di Firebase.
   */
  function _getFirebaseSecret() {
    if (_secret) {
      return _secret;
    }
    
    const secretFromProps = PropertiesService.getScriptProperties().getProperty(CONFIG.FIREBASE_SECRET_PROPERTY);
    
    if (secretFromProps) {
      _secret = secretFromProps;
      return _secret;
    }
    
    Logger.log('ATTENZIONE: Chiave segreta di Firebase non trovata nelle Proprietà Script. Utilizzo del fallback hardcodato da Config.gs. Questo è INSICURO.');
    _secret = CONFIG.FIREBASE_SECRET_FALLBACK;
    return _secret;
  }

  /**
   * Costruisce l'URL completo per una richiesta Firebase, includendo l'autenticazione.
   * @param {string} path Il percorso del nodo (es. "richiesteprimarie").
   * @returns {string} L'URL completo con autenticazione.
   */
  function _buildUrl(path) {
    const secret = _getFirebaseSecret();
    const cleanPath = path.replace(/^\/+/, '').replace(/\/+$/, '');
    const pathWithExtension = cleanPath ? `${cleanPath}.json` : '.json';
    return `${CONFIG.FIREBASE_URL}/${pathWithExtension}?auth=${secret}`;
  }

  /**
   * Esegue una richiesta HTTP generica a Firebase.
   * @param {string} path Il percorso del nodo.
   * @param {'get' | 'patch' | 'post' | 'delete'} method Il metodo HTTP.
   * @param {Object} [payload] Il payload JSON per PATCH o POST.
   * @returns {Object|null} I dati JSON di risposta (per GET) o null.
   * @throws {Error} Se la richiesta fallisce.
   */
  function _fetch(path, method, payload = null) {
    const url = _buildUrl(path);
    const options = {
      method: method,
      muteHttpExceptions: true,
      contentType: 'application/json',
    };

    if (payload) {
      options.payload = JSON.stringify(payload);
    }

    Logger.log(`FirebaseService: Esecuzione ${method.toUpperCase()} su ${path}`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      if (responseText && responseText !== 'null') {
        return JSON.parse(responseText);
      }
      return null; // Operazione riuscita ma senza corpo (es. DELETE) o 'null'
    }

    // Gestione specifica 404 per le letture (nodo non esistente)
    if (method === 'get' && responseCode === 404) {
      Logger.log(`FirebaseService: Nodo non trovato (404) su ${path}. Restituisco null.`);
      return null;
    }

    Logger.log(`ERRORE FirebaseService: Metodo ${method.toUpperCase()} su ${path} fallito. Codice: ${responseCode}. Risposta: ${responseText}`);
    throw new Error(`Errore Firebase ${method.toUpperCase()} (${responseCode}): ${responseText}`);
  }

  // Interfaccia pubblica
  return {
    /**
     * Legge i dati da un nodo Firebase.
     * @param {string} path Il percorso del nodo.
     * @returns {Object|null} I dati JSON.
     */
    firebaseGet: function(path) {
      return _fetch(path, 'get');
    },

    /**
     * Aggiorna parzialmente i dati in un nodo Firebase (PATCH).
     * @param {string} path Il percorso del nodo.
     * @param {Object} data L'oggetto con i dati da aggiornare.
     * @returns {Object|null} I dati aggiornati.
     */
    firebasePatch: function(path, data) {
      return _fetch(path, 'patch', data);
    },

    /**
     * Inserisce nuovi dati in un nodo Firebase (POST), generando un ID univoco.
     * @param {string} path Il percorso del nodo.
     * @param {Object} data L'oggetto con i dati da inserire.
     * @returns {Object|null} La risposta di POST (spesso { name: "ID_GENERATO" }).
     */
    firebasePost: function(path, data) {
      return _fetch(path, 'post', data);
    },

    /**
     * Cancella tutti i dati da un nodo Firebase (DELETE).
     * @param {string} path Il percorso del nodo.
     * @returns {null}
     */
    firebaseDelete: function(path) {
      return _fetch(path, 'delete');
    }
  };

})();