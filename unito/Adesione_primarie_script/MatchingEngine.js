/**
 * @fileoverview Motore di matching e assegnazione.
 * VERSIONE OPTIMIZED & SEPARATED:
 * 1. Fetch dei dati separato: Scarica SOLO Primarie o SOLO Secondarie in base al run.
 * 2. Calcolo posti prenotati e storico separato per tipo.
 * 3. Logica destinazione robusta, filtro primavera ed etichette "terzo rifiuto" mantenute.
 * 4. FIX ASSEGNAZIONE: Il nome del laboratorio assegnato viene passato esplicitamente alla finalizzazione.
 */
var MatchingEngine = (function() {

  // Normalizza stringhe per confronti sicuri
  function _normalizeAndClean(str) {
    if (!str) return '';
    return String(str).trim().toLowerCase().replace(/\s+/g, ' ');
  }

  function _normalizeLabName(name) {
    if (typeof name !== 'string') return '';
    return name.trim().replace(/\s+/g, ' ');
  }

  // Helper data parsing
  function _parseStringDateToObj(dateStr) {
    if (!dateStr) return null;
    const clean = Utilities_.cleanDateString(dateStr); 
    if (!clean) return null;
    
    const parts = clean.split(' ');
    const dateParts = parts[0].split('/');
    const timeParts = parts[1] ? parts[1].split(':') : ['00', '00', '00'];
    
    return new Date(
      parseInt(dateParts[2], 10),
      parseInt(dateParts[1], 10) - 1,
      parseInt(dateParts[0], 10),
      parseInt(timeParts[0], 10),
      parseInt(timeParts[1], 10),
      parseInt(timeParts[2] || '00', 10)
    );
  }

  // ========================================================
  // ORCHESTRATORI
  // ========================================================
  function runMatchingPrimarie() {
    _runMatchingProcessGeneric({
      TYPE: 'PRIMARIE',
      REQUESTS_NODE: CONFIG.REQUESTS_NODE_PRIMARY,
      ASSIGNMENTS_NODE: CONFIG.ASSEGNAZIONI_PRIMARIE_NODE, // Nodo specifico assegnazioni
      DESTINATION: CONFIG.PRIMARY_DESTINATION,
      RESULTS_NODE: CONFIG.MATCHING_RESULTS_PRIMARY_NODE
    });
  }

  function runMatchingSecondarie() {
    _runMatchingProcessGeneric({
      TYPE: 'SECONDARIE',
      REQUESTS_NODE: CONFIG.REQUESTS_NODE_SECONDARY,
      ASSIGNMENTS_NODE: CONFIG.ASSEGNAZIONI_SECONDARIE_NODE, // Nodo specifico assegnazioni
      DESTINATION: CONFIG.SECONDARY_DESTINATION,
      RESULTS_NODE: CONFIG.MATCHING_RESULTS_SECONDARY_NODE
    });
  }

  function runFullMatchingProcess() {
    runMatchingPrimarie();
    runMatchingSecondarie();
  }

  function _runMatchingProcessGeneric(config) {
    Logger.log(`\n=== AVVIO PROCESSO DI MATCHING: ${config.TYPE} ===`);
    Logger.log(`Target Destinazione: "${config.DESTINATION}"`);
    
    // 1. Fetch Dati (SEPARATO: Scarica solo ciò che serve per questo TYPE)
    const allData = _fetchDataForMatching(config);
    if (!allData) return;

    // 2. Pre-processo (Analizza solo lo storico specifico)
    const { filteredRequests, bookedSlots, rejectedSlotsByRequest, requestAssignmentStats } = _preProcessAndFilterData(allData, config);

    // 3. Costruzione lista laboratori (Usa bookedSlots specifici)
    const masterLaboratori = _buildMasterLaboratoriList(allData.laboratori, bookedSlots);

    // 4. Esecuzione algoritmo
    const results = _performMatchingForType(config, filteredRequests, masterLaboratori, rejectedSlotsByRequest, requestAssignmentStats);
    
    _saveResultsToFirebase(config.RESULTS_NODE, results.assignments);
    Logger.log(`=== FINE PROCESSO ${config.TYPE} ===\n`);
  }

  // ========================================================
  // FASE 1: CARICAMENTO DATI (SPECIFICO PER TIPO)
  // ========================================================
  function _fetchDataForMatching(config) {
    try {
      Logger.log(`[DATA] Recupero dati da Firebase SOLO per ${config.TYPE}...`);
      
      // Scarica sempre i laboratori (nodo comune)
      const laboratori = FirebaseService.firebaseGet(CONFIG.LABORATORI_NODE) || {};
      
      // Scarica SOLO le richieste e le assegnazioni pertinenti
      const requests = FirebaseService.firebaseGet(config.REQUESTS_NODE) || {};
      const assignments = FirebaseService.firebaseGet(config.ASSIGNMENTS_NODE) || {};

      Logger.log(`[DATA] Scaricati: ${Object.keys(requests).length} richieste, ${Object.keys(assignments).length} assegnazioni storiche.`);

      return {
        laboratori: laboratori,
        requests: requests,
        assignments: assignments
      };

    } catch (e) {
      Logger.log(`[ERROR] Errore recupero dati: ${e.toString()}`);
      return null;
    }
  }

  // ========================================================
  // FASE 2: PRE-ELABORAZIONE (SU DATI SPECIFICI)
  // ========================================================
  function _preProcessAndFilterData(allData, config) {
    Logger.log(`[PRE-PROC] Analisi assegnazioni e calcolo slot occupati (${config.TYPE})...`);
    
    const rejectedSlotsByRequest = {}; 
    const bookedSlots = {}; 
    const requestAssignmentStats = {};

    // Usa solo le assegnazioni scaricate per questo tipo
    const specificAssignmentsRaw = Utilities_.toArray(allData.assignments);

    const assignmentsByRequest = {};
    specificAssignmentsRaw.forEach(ass => {
      const rid = Utilities_.getField(ass, 'id_firebase', 'firebase_id');
      if (!rid) return;
      if (!assignmentsByRequest[rid]) assignmentsByRequest[rid] = [];
      assignmentsByRequest[rid].push(ass);
    });

    const normLab = (name) => _normalizeLabName((name || '').toString()).replace(/\s\(\d+\)$/, '');
    const normState = (raw) => {
      const v = (raw || '').toString().trim().toUpperCase();
      if (v === 'SI' || v === 'SÌ') return 'SI';
      if (v.includes('PROPOSTA') && v.includes('ELABORARE')) return 'PROPOSTA DA ELABORARE';
      if (v === 'NO') return 'NO';
      return v;
    };

    for (const reqId in assignmentsByRequest) {
      const records = assignmentsByRequest[reqId];

      // Cerca SI storico
      const siRecord = records.find(r => normState(Utilities_.getField(r, 'proposta_accettata')) === 'SI');

      if (siRecord) {
        const labName = normLab(Utilities_.getField(siRecord, 'nome_lab'));
        const rawDate = Utilities_.getField(siRecord, 'data_lab');
        const dateNorm = Utilities_.cleanDateString(rawDate);

        requestAssignmentStats[reqId] = { 
          rejectedCounter: 0, 
          assignedDates: new Set(), 
          isFullyAssigned: true 
        };

        if (labName && dateNorm) {
          if (!bookedSlots[labName]) bookedSlots[labName] = {};
          bookedSlots[labName][dateNorm] = (bookedSlots[labName][dateNorm] || 0) + 1;
          requestAssignmentStats[reqId].assignedDates.add(dateNorm);
        }
      } else {
        // Max Contatore NO
        records.sort((a, b) => {
          const ca = parseInt(a.contatore_no || '0', 10);
          const cb = parseInt(b.contatore_no || '0', 10);
          return cb - ca; 
        });

        const activeRecord = records[0]; 
        const currentStatus = normState(Utilities_.getField(activeRecord, 'proposta_accettata'));
        const maxCounter = parseInt(activeRecord.contatore_no || '0', 10);
        const labName = normLab(Utilities_.getField(activeRecord, 'nome_lab'));
        const rawDate = Utilities_.getField(activeRecord, 'data_lab');
        const dateNorm = Utilities_.cleanDateString(rawDate);

        requestAssignmentStats[reqId] = { 
          rejectedCounter: maxCounter, 
          assignedDates: new Set(), 
          isFullyAssigned: false 
        };

        if (currentStatus === 'PROPOSTA DA ELABORARE') {
          if (labName && dateNorm) {
            if (!bookedSlots[labName]) bookedSlots[labName] = {};
            bookedSlots[labName][dateNorm] = (bookedSlots[labName][dateNorm] || 0) + 1;
            requestAssignmentStats[reqId].isFullyAssigned = true;
            requestAssignmentStats[reqId].assignedDates.add(dateNorm);
          }
        } 
        
        // Costruisce Blacklist
        records.forEach(rec => {
           const s = normState(Utilities_.getField(rec, 'proposta_accettata'));
           if (s === 'NO') {
             const l = normLab(Utilities_.getField(rec, 'nome_lab'));
             const d = Utilities_.cleanDateString(Utilities_.getField(rec, 'data_lab'));
             if (l && d) {
               if (!rejectedSlotsByRequest[reqId]) rejectedSlotsByRequest[reqId] = {};
               if (!rejectedSlotsByRequest[reqId][l]) rejectedSlotsByRequest[reqId][l] = new Set();
               rejectedSlotsByRequest[reqId][l].add(d);
             }
           }
        });
      }
    }

    // --- FILTRAGGIO RICHIESTE (Usa allData.requests specifico) ---
    const filteredRequests = {};
    const originalRequests = allData.requests;
    
    const now = new Date();
    const currentMonth = now.getMonth(); 
    const isReassignmentPeriod = (currentMonth >= 1 && currentMonth <= 6); // Feb-Lug

    const filterFunction = (req, firebaseKey) => {
      const stats = requestAssignmentStats[firebaseKey] || { isFullyAssigned: false, rejectedCounter: 0 };
      if (stats.isFullyAssigned) return false;

      const rejectionLimit = isReassignmentPeriod ? 999 : 3;
      
      if (stats.rejectedCounter >= rejectionLimit) {
        return false;
      }
      return true;
    };

    for (const key in originalRequests) {
      if (Object.prototype.hasOwnProperty.call(originalRequests, key) && filterFunction(originalRequests[key], key)) {
        filteredRequests[key] = originalRequests[key];
      }
    }

    return { filteredRequests, bookedSlots, rejectedSlotsByRequest, requestAssignmentStats };
  }

  // ========================================================
  // MASTER LABS
  // ========================================================
  function _buildMasterLaboratoriList(laboratoriData, bookedSlots) {
    Logger.log("[MASTER] Costruzione lista laboratori...");
    const masterLaboratori = {};
    const now = new Date(); 

    Utilities_.toArray(laboratoriData)
      .filter(labData => labData.date_disponibili)
      .forEach(labData => {
        const labName = _normalizeLabName(labData.titolo);
        if (!labName) return;
        
        // Logica occupazione (Usa bookedSlots passati, specifici per il tipo)
        const bookedCountsForThisLab = {};
        for(const bookedLabName in bookedSlots) {
            const baseBookedName = _normalizeLabName(bookedLabName.replace(/\s\(\d+\)$/, ''));
            if (baseBookedName === labName) {
                const slots = bookedSlots[bookedLabName];
                for (const date in slots) {
                    bookedCountsForThisLab[date] = (bookedCountsForThisLab[date] || 0) + slots[date];
                }
            }
        }

        const availableDateMap = {};
        let totalAvailableSlots = 0;
        const allDates = Utilities_.toArray(labData.date_disponibili).map(d => d.datetime).filter(Boolean);
        
        for (const dateStr of allDates) {
          const dateObj = _parseStringDateToObj(dateStr);
          if (dateObj && dateObj < now) continue; 
          
          const normalizedDate = Utilities_.cleanDateString(dateStr);
          if (!normalizedDate) continue;

          if (bookedCountsForThisLab[normalizedDate] && bookedCountsForThisLab[normalizedDate] > 0) {
            bookedCountsForThisLab[normalizedDate]--;
          } else {
            availableDateMap[dateStr] = (availableDateMap[dateStr] || 0) + 1;
            totalAvailableSlots++;
          }
        }

        if (totalAvailableSlots > 0) {
          masterLaboratori[labName] = {
            destinatari: labData.destinatari,
            slotDisponibili: totalAvailableSlots,
            dateDisponibiliMap: availableDateMap,
            durata_incontro: labData.durata_incontro || null,
            sede: labData.punto_incontro || 'Sede non specificata',
            area_tematica: labData.area_tematica || 'Non definita'
          };
        }
      });

    return masterLaboratori;
  }

  // ========================================================
  // PREPARAZIONE DATI
  // ========================================================
  function _prepareDataForType(config, filteredRequests, masterLaboratori) {
    const laboratori = {};
    const targetDest = _normalizeAndClean(config.DESTINATION);
    
    let labsFound = 0;
    let labsRejected = 0;

    Logger.log(`[FILTER LABS] Cerco laboratori per: "${targetDest}"`);

    for (const labName in masterLaboratori) {
      const labDestRaw = masterLaboratori[labName].destinatari;
      const labDest = _normalizeAndClean(labDestRaw);

      if (labDest === targetDest || labDest.includes(targetDest) || targetDest.includes(labDest)) {
        laboratori[labName] = masterLaboratori[labName];
        labsFound++;
      } else {
        if (labsRejected < 5) {
            Logger.log(`   [SKIP] Lab "${labName}" scartato. Destinazione Lab: "${labDestRaw}"`);
        }
        labsRejected++;
      }
    }
    Logger.log(`[FILTER LABS] Labs Accettati: ${labsFound}. Scartati: ${labsRejected}.`);

    const rawRequestsArray = Utilities_.toArray(filteredRequests);
    Logger.log(`[FILTER REQS] Processo ${rawRequestsArray.length} richieste filtrate.`);
    
    const seen = new Set();
    const requests = [];

    rawRequestsArray.forEach(req => {
      const id = Utilities_.getField(req, 'id') || req.__firebaseKey;
      if (!id || seen.has(id)) return;
      seen.add(id);

      let labChoicesRaw = req['laboratori_richiesti'];
      let labNames = [];
      if (Array.isArray(labChoicesRaw)) {
        labNames = labChoicesRaw.map(n => _normalizeLabName(String(n))).filter(Boolean);
      } else if (typeof labChoicesRaw === 'string') {
        labNames = labChoicesRaw.split(',').map(n => _normalizeLabName(n)).filter(Boolean);
      }

      let hasValidChoice = false;
      labNames.forEach((labName, idx) => {
        if (laboratori[labName]) {
           hasValidChoice = true;
           requests.push({
            id: id,
            labRichiesto: labName,
            Priorita: idx + 1,
            email: Utilities_.getField(req, 'email'),
            istituto: Utilities_.getField(req, 'istituto_codice_mecc'),
            circoscrizione: Utilities_.getField(req, 'istituto_circoscrizione'),
            classe: Utilities_.getField(req, 'classe_livello'),
            sezione: Utilities_.getField(req, 'classe_sezione'),
            insegnante_cellulare: Utilities_.getField(req, 'insegnante_cellulare'),
           });
        }
      });
    });
    
    return { laboratori, requests };
  }

  // ========================================================
  // CONSUMO SLOT
  // ========================================================
  function _consumeNextAvailableSlot(labName, labData, excludedDatesSet, reqId) {
    if (!labData || labData.slotDisponibili <= 0) return null;

    const availableDatesRaw = Object.keys(labData.dateDisponibiliMap).sort();
    let dateToAssignRaw = null;

    for (const rawDate of availableDatesRaw) {
      const normalizedDate = Utilities_.cleanDateString(rawDate);
      if (normalizedDate && !excludedDatesSet.has(normalizedDate)) {
        dateToAssignRaw = rawDate;
        break;
      }
    }

    if (!dateToAssignRaw) return null;

    labData.dateDisponibiliMap[dateToAssignRaw]--;
    labData.slotDisponibili--;
    
    if (labData.dateDisponibiliMap[dateToAssignRaw] <= 0) {
      delete labData.dateDisponibiliMap[dateToAssignRaw];
    }
    return dateToAssignRaw;
  }

  function _getPerLabRejectedDates(rejectedByReq, reqId, labNameRaw) {
    const labKey = _normalizeLabName((labNameRaw || '').toString()).replace(/\s\(\d+\)$/, '');
    const byReq = rejectedByReq[reqId] || {};
    return byReq[labKey] || new Set();
  }

  // ========================================================
  // ASSEGNAZIONE (MAIN LOOP)
  // ========================================================
  function _performMatchingForType(config, filteredRequests, masterLaboratori, rejectedSlotsByRequest, requestAssignmentStats) {
    const deepCopiedMaster = JSON.parse(JSON.stringify(masterLaboratori));
    const { laboratori, requests } = _prepareDataForType(config, filteredRequests, deepCopiedMaster);
    
    if (Object.keys(laboratori).length === 0) {
        Logger.log("⚠️ ATTENZIONE: Nessun laboratorio trovato. Controlla Config Destinazione e DB.");
    }
    if (requests.length === 0) {
        Logger.log("⚠️ ATTENZIONE: Nessuna richiesta valida trovata.");
    }

    return _processAssignments(laboratori, requests, config, rejectedSlotsByRequest, requestAssignmentStats);
  }

  function _processAssignments(laboratori, requests, config, rejectedSlotsByRequest, requestAssignmentStats) {
    const assignments = [];
    const assignedClasses = new Set();
    const equityCounters = _initEquityCounters();

    requests.forEach(r => _updateEquityCounters(r, equityCounters, true));

    // --- FASE 0: REPAIR ---
    const repairRequests = requests.filter(r => {
        const stats = requestAssignmentStats[r.id];
        return stats && stats.rejectedCounter > 0 && !assignedClasses.has(r.id);
    });

    repairRequests.sort((a, b) => {
        const statsA = requestAssignmentStats[a.id].rejectedCounter;
        const statsB = requestAssignmentStats[b.id].rejectedCounter;
        if (statsA !== statsB) return statsB - statsA; 
        return a.Priorita - b.Priorita; 
    });

    Logger.log(`[FASE 0] Repair per ${repairRequests.length} richieste.`);
    
    repairRequests.forEach(req => {
        if (assignedClasses.has(req.id)) return; 
        
        const labData = laboratori[req.labRichiesto];
        if (labData && labData.slotDisponibili > 0) {
            const perLabRejected = _getPerLabRejectedDates(rejectedSlotsByRequest, req.id, req.labRichiesto);
            const assignedToMe = (requestAssignmentStats[req.id] && requestAssignmentStats[req.id].assignedDates) || new Set();
            const excludedDates = new Set([...perLabRejected, ...assignedToMe]);
            
            const slot = _consumeNextAvailableSlot(req.labRichiesto, labData, excludedDates, req.id);
            
            if (slot) {
                const currentCounter = requestAssignmentStats[req.id] ? requestAssignmentStats[req.id].rejectedCounter : 0;
                let dynamicPhaseLabel = `Priorità Rifiuto (Fase ${currentCounter})`;
                
                if (currentCounter >= 3) {
                   dynamicPhaseLabel = "terzo rifiuto - da rivedere in primavera";
                }
                // FIX: Passo esplicitamente il laboratorio richiesto
                _finalizeAssignment(req, labData, slot, dynamicPhaseLabel, assignments, assignedClasses, equityCounters, req.labRichiesto);
            }
        }
    });

    // --- STANDARD FASES ---
    const residualRequests = requests.filter(r => !assignedClasses.has(r.id));
    const residualDemand = {};
    residualRequests.forEach(r => residualDemand[r.labRichiesto] = (residualDemand[r.labRichiesto] || 0) + 1);

    const lowDemandLabs = new Set();
    const highDemandLabs = new Set();
    for (const labName in laboratori) {
        if (laboratori[labName].slotDisponibili > 0) {
            if ((residualDemand[labName] || 0) <= 3) lowDemandLabs.add(labName);
            else highDemandLabs.add(labName);
        }
    }

    // FASE A
    const lowDemandReqs = residualRequests
        .filter(r => lowDemandLabs.has(r.labRichiesto))
        .map(r => ({ ...r, score: _calculateScore(r, equityCounters, false) }))
        .sort((a, b) => b.score - a.score);

    lowDemandReqs.forEach(req => {
        if (assignedClasses.has(req.id)) return;
        const labData = laboratori[req.labRichiesto];
        if (labData && labData.slotDisponibili > 0) {
             const perLabRejected = _getPerLabRejectedDates(rejectedSlotsByRequest, req.id, req.labRichiesto);
             const assignedToMe = (requestAssignmentStats[req.id] && requestAssignmentStats[req.id].assignedDates) || new Set();
             const excluded = new Set([...perLabRejected, ...assignedToMe]);
             
             const slot = _consumeNextAvailableSlot(req.labRichiesto, labData, excluded, req.id);
             if (slot) {
                 // FIX: Passo esplicitamente il laboratorio richiesto
                 _finalizeAssignment(req, labData, slot, 'Bassa Domanda', assignments, assignedClasses, equityCounters, req.labRichiesto);
             }
        }
    });

    // FASE B
    highDemandLabs.forEach(labName => {
        const labData = laboratori[labName];
        while (labData.slotDisponibili > 0) {
            const candidates = requests
                .filter(r => r.labRichiesto === labName && !assignedClasses.has(r.id))
                .map(r => ({ ...r, score: _calculateScore(r, equityCounters, false) }))
                .sort((a, b) => b.score - a.score);
            
            if (candidates.length === 0) break;
            
            let assigned = false;
            for (const candidate of candidates) {
                const perLabRejected = _getPerLabRejectedDates(rejectedSlotsByRequest, candidate.id, candidate.labRichiesto);
                const assignedToMe = (requestAssignmentStats[candidate.id] && requestAssignmentStats[candidate.id].assignedDates) || new Set();
                const excluded = new Set([...perLabRejected, ...assignedToMe]);
                
                const slot = _consumeNextAvailableSlot(labName, labData, excluded, candidate.id);
                if (slot) {
                    // FIX: Passo il labName attuale del ciclo
                    _finalizeAssignment(candidate, labData, slot, 'Alta Domanda', assignments, assignedClasses, equityCounters, labName);
                    assigned = true;
                    break; 
                }
            }
            if (!assigned) break;
        }
    });

    // FASE EXTRA
    const unassignedIds = [...new Set(requests.filter(r => !assignedClasses.has(r.id)).map(r => r.id))];
    const bestUnassigned = unassignedIds.map(id => {
        return requests.find(r => r.id === id);
    }).filter(Boolean).map(r => ({...r, score: _calculateScore(r, equityCounters, true)})).sort((a,b) => b.score - a.score);

    bestUnassigned.forEach(req => {
        if (assignedClasses.has(req.id)) return;
        
        const allLabs = Object.keys(laboratori);
        for (const labName of allLabs) {
            const labData = laboratori[labName];
            if (labData.slotDisponibili > 0) {
                 const perLabRejected = _getPerLabRejectedDates(rejectedSlotsByRequest, req.id, labName);
                 const assignedToMe = (requestAssignmentStats[req.id] && requestAssignmentStats[req.id].assignedDates) || new Set();
                 const excluded = new Set([...perLabRejected, ...assignedToMe]);
                 
                 const slot = _consumeNextAvailableSlot(labName, labData, excluded, req.id);
                 if (slot) {
                     // FIX: Passo il labName attuale (diverso dalla richiesta originale)
                     _finalizeAssignment(req, labData, slot, 'Extra (Ripiego)', assignments, assignedClasses, equityCounters, labName);
                     return;
                 }
            }
        }
    });

    // === FIX MANCATA CANCELLAZIONE ===
    // Controlliamo se ci sono richieste di "Repair" (chi aveva detto NO) 
    // che NON sono state soddisfatte in nessuna fase.
    // Dobbiamo forzare un aggiornamento vuoto per cancellare la data vecchia sul foglio.
    
    repairRequests.forEach(req => {
        if (!assignedClasses.has(req.id)) {
            // È una richiesta che doveva essere riparata ma non ha trovato posto.
            // Aggiungiamo un oggetto che "pulisce" i campi.
            assignments.push({
                id: req.id,
                labRichiesto: req.labRichiesto, // Manteniamo il lab richiesto originale
                labAssegnato: "",     // VUOTO -> Cancella sul foglio
                dataAssegnata: "",    // VUOTO -> Cancella sul foglio
                punteggioEquita: 0,
                faseAssegnazione: "NESSUNA DISPONIBILITÀ (Dopo Rifiuto)",
                assegnato: false,
                durata_incontro: "",
                sede: ""
            });
            // Non serve aggiungerlo ad assignedClasses perché non consuma risorse
            Logger.log(`>>> [CLEANUP] RIMOZIONE VECCHIA DATA: ${req.id} -> Nessun nuovo posto trovato.`);
        }
    });

    return { assignments };
  }

  function _finalizeAssignment(req, labData, rawDate, phase, assignments, assignedClasses, equityCounters, actualLabName) {
      assignments.push(Object.assign({}, req, {
          dataAssegnata: Utilities_.formatRawDateForOutput(rawDate),
          punteggioEquita: req.score || 0,
          faseAssegnazione: phase,
          assegnato: true,
          durata_incontro: labData.durata_incontro,
          sede: labData.sede,
          // FIX: Usa il nome reale se passato, altrimenti fallback (safety)
          labAssegnato: actualLabName || req.labRichiesto 
      }));
      assignedClasses.add(req.id);
      _updateEquityCounters(req, equityCounters);
      Logger.log(`>>> [SUCCESS] ASSEGNATO: ${req.id} -> ${actualLabName || req.labRichiesto} (${rawDate}) [${phase}]`);
  }

  // ========================================================
  // EQUITY & HELPERS
  // ========================================================
  function _initEquityCounters() {
      return {
          laboratoriRequestsCount: {},
          circoscrizioniAssignmentCount: {},
          instituteAssignmentCount: {},
          emailAssignmentCount: {},
          circoscrizioneInstituteAssignments: {}
      };
  }

  function _updateEquityCounters(req, counters, onlyDemand = false) {
      if (onlyDemand) {
          counters.laboratoriRequestsCount[req.labRichiesto] = (counters.laboratoriRequestsCount[req.labRichiesto] || 0) + 1;
          return;
      }
      counters.circoscrizioniAssignmentCount[req.circoscrizione] = (counters.circoscrizioniAssignmentCount[req.circoscrizione] || 0) + 1;
      counters.instituteAssignmentCount[req.istituto] = (counters.instituteAssignmentCount[req.istituto] || 0) + 1;
      counters.emailAssignmentCount[req.email] = (counters.emailAssignmentCount[req.email] || 0) + 1;
      const key = `${req.istituto}|${req.circoscrizione}`;
      counters.circoscrizioneInstituteAssignments[key] = (counters.circoscrizioneInstituteAssignments[key] || 0) + 1;
  }

  function _calculateScore(req, counters, isExtra) {
    const W = CONFIG.SCORE_WEIGHTS;
    let score = 0;
    const circ = req.circoscrizione || 'NA';
    score += W.TERRITORIAL_EQUITY_BASE_WEIGHT * (CONFIG.CIRCOSCRIZIONE_PRIORITY_MAP[circ] || 0);
    score -= (counters.laboratoriRequestsCount[req.labRichiesto] || 0) * W.LAB_POPULARITY_PENALTY_WEIGHT;
    score -= (counters.circoscrizioneInstituteAssignments[`${req.istituto}|${circ}`] || 0) * W.CIRC_VARIETY_PENALTY_WEIGHT;
    
    const instCount = counters.instituteAssignmentCount[req.istituto] || 0;
    score += Math.max(0, W.INSTITUTE_EQUITY_BASE_WEIGHT - (instCount * W.INSTITUTE_EQUITY_MULTIPLIER));
    
    const emailCount = counters.emailAssignmentCount[req.email] || 0;
    score += Math.max(0, W.EMAIL_EQUITY_MULTIPLIER - emailCount) * 100;
    
    if (isExtra) {
        score *= W.BONUS_MULTIPLIER;
    }
    return score;
  }

  function _saveResultsToFirebase(node, assignments) {
      const timestamp = new Date().getTime();
      const isoCreatedAt = Utilities.formatDate(new Date(), CONFIG.SCRIPT_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss'Z'");
      Logger.log(`[SAVE] Salvataggio ${assignments.length} assegnazioni su ${node}/${timestamp}`);
      
      const dataToSave = {};
      assignments.forEach(a => {
          if (a.id) dataToSave[a.id] = a;
      });

      const payload = Object.assign({ _meta: { createdAt: isoCreatedAt, count: assignments.length } }, dataToSave);
      FirebaseService.firebasePatch(`${node}/${timestamp}`, payload);
  }

  return {
    runMatchingPrimarie: runMatchingPrimarie,
    runMatchingSecondarie: runMatchingSecondarie,
    runFullMatchingProcess: runFullMatchingProcess
  };

})();