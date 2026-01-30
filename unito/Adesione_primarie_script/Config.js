/**
 * @fileoverview Configurazione centralizzata per l'applicazione.
 * Contiene tutte le costanti globali per Firebase, Fogli Google e l'algoritmo di matching.
 */

// Oggetto di configurazione globale
const CONFIG = {

  // --- Configurazione Firebase ---
  FIREBASE_URL: 'https://bimbiunito-test2.europe-west1.firebasedatabase.app',
  FIREBASE_SECRET_PROPERTY: 'FIREBASE_SECRET',


  // Nodi principali
  LABORATORI_NODE: 'laboratori',
  REQUESTS_NODE_PRIMARY: 'richiesteprimarie',
  REQUESTS_NODE_SECONDARY: 'richiestesecondarie',
  ASSEGNAZIONI_PRIMARIE_NODE: 'assegnazioni_primarie',
  ASSEGNAZIONI_SECONDARIE_NODE: 'assegnazioni_secondarie',
  
  // Nodi dei risultati del matching
  MATCHING_RESULTS_PRIMARY_NODE: 'risultati_assegnazione_primarie',
  MATCHING_RESULTS_SECONDARY_NODE: 'risultati_assegnazione_secondarie',


  // --- Configurazione Fogli Google ---
  
  /**
   * Nome del foglio di calcolo *originale* (dove arrivano le risposte
   * del form e dove si lanciano 'populateIds' e 'updateAccepted').
   */
  ORIGINAL_FORM_SHEET_NAME: 'Risposte del modulo 2',
  
  /**
   * ID della cartella Drive contenente i fogli "Adesioni" E i fogli "Sorgente".
   */
  DESTINATION_FOLDER_ID: '1kvbOiM3PFno91jfshnqhkKfC7z0dTG9P',
  
  // Nomi dei file "Adesioni" (Destinazione)
  ADESIONI_PRIMARIE_FILENAME: 'Adesione a _Un giorno all_università_ 25-26  - PRIMARIE (Risposte)',
  ADESIONI_SECONDARIE_FILENAME: 'Adesione a _Un giorno all_università_ 25-26  - SECONDO GRADO (Risposte)',
  ADESIONI_COLUMNS_TO_UPDATE: [
    'Proposta accettata',
    'Nome laboratorio proposto/accettato',
    'Data e ora proposta/accettata',
    'consiglio AI',
    'Sede incontro',
    'Durata incontro'
  ],

  // --- NUOVO: Configurazione per Ripopolamento da File Specifici ---
  /**
   * Nome del file sorgente per il ripopolamento delle PRIMARIE.
   */
  REPOPULATE_SOURCE_FILE_PRIMARY: 'Adesione a _Un giorno all_università_ 25-26  - PRIMARIE (Risposte)',
  
  /**
   * Nome del file sorgente per il ripopolamento delle SECONDARIE.
   */
  REPOPULATE_SOURCE_FILE_SECONDARY: 'Adesione a _Un giorno all_università_ 25-26  - SECONDO GRADO (Risposte)',
  
  /**
   * Nome del foglio da cui leggere i dati all'interno dei file sorgente.
   */
  REPOPULATE_SOURCE_SHEET_NAME: 'Risposte del modulo 2',


  // --- Configurazione Generale & Matching ---
  SCRIPT_TIMEZONE: 'Europe/Rome',
  BATCH_SIZE: 250,
  PRIMARY_DESTINATION: 'Scuola Primaria',
  SECONDARY_DESTINATION: 'Scuola Secondaria di I grado',

  // Mappa priorità circoscrizioni
  CIRCOSCRIZIONE_PRIORITY_MAP: {
    '2': 4, '6': 4, '0': 4,
    '8': 3, '7': 3,
    '3': 2, '4': 2, '5': 2,
    '1': 1
  },
  
  // Pesi dell'algoritmo di punteggio
  SCORE_WEIGHTS: {
    TERRITORIAL_EQUITY_BASE_WEIGHT: 1000,
    INSTITUTE_EQUITY_BASE_WEIGHT: 1500,
    INSTITUTE_EQUITY_MULTIPLIER: 10.0,
    EMAIL_EQUITY_MULTIPLIER: 1.0,
    EMAIL_MAX_ASSIGNMENTS_PENALTY: 50000,
    LAB_POPULARITY_PENALTY_WEIGHT: 500,
    REQUEST_PRIORITY_WEIGHT: 100,
    CIRC_VARIETY_PENALTY_WEIGHT: 40000,
    BONUS_MULTIPLIER: 1.0,
    PENALTY_MULTIPLIER: 0.5,
  }
};