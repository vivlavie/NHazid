/*
  HAZID Workshop Table App
  - Data model with hazards, causes, consequences, measures, recommendations
  - Dynamic rendering with rowspans to align uneven cause/consequence counts
  - CRUD operations with cascading deletes
  - Copy/paste for items and hazards
  - Local storage persistence and JSON import/export
*/

;(function () {
  const STORAGE_KEY = 'hazid_v1';

  /**
   * Data Types (JS Doc for readability)
   * Hazard = { id, title, description, causes: Cause[], consequences: Consequence[], recommendations: Recommendation[] }
   * Cause = { id, text, preventionMeasures: Measure[] }
   * Consequence = { id, text, mitigationMeasures: Measure[], risk: Risk }
   * Recommendation = { id, action, responsible }
   * Measure = { id, text }
   * Risk = { severityCategory, severityLevel, likelihoodLevel, riskScore }
   */

  // Utilities
  const generateId = () => Math.random().toString(36).slice(2, 10);

  const deepClone = (obj) => JSON.parse(JSON.stringify(obj));

  const byId = (id) => document.getElementById(id);

  const qs = (sel, el = document) => el.querySelector(sel);

  const qsa = (sel, el = document) => Array.from(el.querySelectorAll(sel));

  const createEl = (tag, attrs = {}, children = []) => {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') el.className = v;
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else if (k === 'text') el.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v);
    }
    for (const child of children) el.append(child);
    return el;
  };

  // Helper function to get button text based on compact mode
  function getButtonText(normalText, compactChar) {
    return state.compactMode ? compactChar : normalText;
  }

  // Helper function to create button with tooltip
  function createButtonWithTooltip(attrs, normalText, compactChar) {
    const button = createEl('button', attrs);
    button.textContent = getButtonText(normalText, compactChar);
    if (state.compactMode) {
      button.title = normalText; // Show full text as tooltip in compact mode
    }
    return button;
  }

  // Persistence
  const persist = {
    load() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const data = JSON.parse(raw);
        return Array.isArray(data) ? data : [];
      } catch (e) {
        console.error('Failed to load data', e);
        return [];
      }
    },
    save(hazards) {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(hazards));
      } catch (e) {
        console.error('Failed to save data', e);
      }
    }
  };

  // App State
  const state = {
    hazards: [],
    autosave: true,
    compactMode: false,
    clipboard: null, // holds a copied item or hazard
    bowtieColors: null, // holds custom bow-tie diagram colors
    riskMatrix: {
      likelihood: [
        { id: 'A', label: 'A', description: 'Very unlikely' },
        { id: 'B', label: 'B', description: 'Unlikely' },
        { id: 'C', label: 'C', description: 'Possible' },
        { id: 'D', label: 'D', description: 'Likely' },
        { id: 'E', label: 'E', description: 'Very likely' }
      ],
      severity: [
        { id: '1', label: '1', description: 'Negligible effect', category: 'personnel' },
        { id: '2', label: '2', description: 'Minor effect', category: 'personnel' },
        { id: '3', label: '3', description: 'Moderate effect', category: 'personnel' },
        { id: '4', label: '4', description: 'Major effect', category: 'personnel' },
        { id: '5', label: '5', description: 'Severest effect', category: 'personnel' }
      ],
      severityDescriptions: {
        '1': {
          personnel: 'Minor injury, no lost time',
          asset: 'Minor damage, easily repairable',
          environmental: 'Minimal environmental impact',
          reputation: 'No reputation impact',
          operation: 'No operational impact'
        },
        '2': {
          personnel: 'Minor injury, some lost time',
          asset: 'Moderate damage, repairable',
          environmental: 'Minor environmental impact',
          reputation: 'Minor local reputation impact',
          operation: 'Minor operational disruption'
        },
        '3': {
          personnel: 'Serious injury, significant lost time',
          asset: 'Major damage, expensive repair',
          environmental: 'Moderate environmental impact',
          reputation: 'Moderate reputation impact',
          operation: 'Moderate operational disruption'
        },
        '4': {
          personnel: 'Major injury, permanent disability',
          asset: 'Severe damage, major repair cost',
          environmental: 'Major environmental impact',
          reputation: 'Major reputation impact',
          operation: 'Major operational disruption'
        },
        '5': {
          personnel: 'Multiple fatalities',
          asset: 'Total loss of facility',
          environmental: 'Permanent environmental damage',
          reputation: 'Major international effect',
          operation: 'Loss of operation up to a year'
        }
      },
      riskLevels: [
        { id: 'low', label: 'Low', color: '#28a745' },
        { id: 'medium', label: 'Medium', color: '#ffc107' },
        { id: 'high', label: 'High', color: '#dc3545' }
      ],
      matrix: {} // will be computed from likelihood × severity
    }
  };

  // Risk helpers
  const computeRiskScore = (severityLevel, likelihoodLevel) => {
    if (!severityLevel || !likelihoodLevel) return '';
    const key = `${likelihoodLevel}-${severityLevel}`;
    const riskLevel = state.riskMatrix.matrix[key];
    return riskLevel ? riskLevel.label : '';
  };

  const getRiskLevelColor = (severityLevel, likelihoodLevel) => {
    if (!severityLevel || !likelihoodLevel) return '';
    const key = `${likelihoodLevel}-${severityLevel}`;
    const riskLevel = state.riskMatrix.matrix[key];
    return riskLevel ? riskLevel.color : '';
  };

  const getRiskLevel = (severityLevel, likelihoodLevel) => {
    if (!severityLevel || !likelihoodLevel) return '';
    const key = `${likelihoodLevel}-${severityLevel}`;
    const riskLevel = state.riskMatrix.matrix[key];
    return riskLevel ? riskLevel.label : '';
  };

  const updateRiskMatrix = () => {
    state.riskMatrix.matrix = {};
    // Default risk assignment: higher likelihood + higher severity = higher risk
    state.riskMatrix.likelihood.forEach((lik, li) => {
      state.riskMatrix.severity.forEach((sev, si) => {
        const key = `${lik.id}-${sev.id}`;
        let riskLevel = state.riskMatrix.riskLevels[0]; // default to low
        
        // Simple risk calculation: sum of indices
        const riskSum = li + si;
        const maxSum = (state.riskMatrix.likelihood.length - 1) + (state.riskMatrix.severity.length - 1);
        const riskRatio = riskSum / maxSum;
        
        if (riskRatio >= 0.7) riskLevel = state.riskMatrix.riskLevels[2]; // high
        else if (riskRatio >= 0.4) riskLevel = state.riskMatrix.riskLevels[1]; // medium
        
        state.riskMatrix.matrix[key] = riskLevel;
      });
    });
  };

  // Model factories
  const createHazard = () => ({
    id: generateId(),
    title: '',
    description: '',
    causes: [],
    consequences: [],
    recommendations: []
  });

  const createCause = () => ({ id: generateId(), text: '', preventionMeasures: [] });
  const createConsequence = () => ({ id: generateId(), text: '', mitigationMeasures: [], risk: { severityCategory: '', severityLevel: '', likelihoodLevel: '', riskScore: '' } });
  const createMeasure = () => ({ id: generateId(), text: '' });
  const createRecommendation = () => ({ id: generateId(), action: '', responsible: '' });

  // Initialization
  function init() {
    const saved = persist.load();
    if (saved && saved.hazards) {
      state.hazards = saved.hazards;
      if (saved.riskMatrix) {
        state.riskMatrix = { ...state.riskMatrix, ...saved.riskMatrix };
      }
      if (saved.compactMode !== undefined) {
        state.compactMode = saved.compactMode;
      }
      if (saved.bowtieColors) {
        state.bowtieColors = saved.bowtieColors;
      }
    } else if (Array.isArray(saved)) {
      // Legacy format
      state.hazards = saved;
    }
    
    if (state.hazards.length === 0) {
      // Create a starter hazard for convenience
      const hz = createHazard();
      hz.title = 'New hazard';
      hz.causes.push(createCause());
      hz.consequences.push(createConsequence());
      hz.recommendations.push(createRecommendation());
      state.hazards.push(hz);
    }

    updateRiskMatrix();
    
    const container = byId('table-container');
    container.innerHTML = '';
    renderTable(container);
    renderRiskMatrixConfig();
    wireGlobalActions();
    
    // Apply compact mode class if enabled
    if (state.compactMode) {
      document.body.classList.add('compact-mode');
      document.getElementById('compact-mode').checked = true;
    }
  }

  // Rendering
  function renderTable(container) {
    const tmpl = byId('hazid-table-template');
    const tableEl = tmpl.content.firstElementChild.cloneNode(true);
    const tbody = qs('tbody', tableEl);

    state.hazards.forEach((hazard, hazardIndex) => {
      const tr = document.createElement('tr');

      // Hazard cell
      const tdHazard = createEl('td');
      tdHazard.append(renderHazardCell(hazard, hazardIndex));
      tr.append(tdHazard);

      // Causes group
      const tdCause = createEl('td');
      tdCause.append(renderCauseSegments(hazard, hazardIndex));
      tr.append(tdCause);

      const tdCauseMeasures = createEl('td');
      tdCauseMeasures.append(renderCauseMeasuresSegments(hazard, hazardIndex));
      tr.append(tdCauseMeasures);

      // Consequences group
      const tdConseq = createEl('td');
      tdConseq.append(renderConsequenceSegments(hazard, hazardIndex));
      tr.append(tdConseq);

      const tdConseqMeasures = createEl('td');
      tdConseqMeasures.append(renderConsequenceMeasuresSegments(hazard, hazardIndex));
      tr.append(tdConseqMeasures);

      // Risk columns – segmented aligned to consequences only
      const tdSevCat = createEl('td');
      tdSevCat.append(renderRiskSegments(hazard, hazardIndex, 'severityCategory'));
      tr.append(tdSevCat);

      const tdSevLvl = createEl('td');
      tdSevLvl.append(renderRiskSegments(hazard, hazardIndex, 'severityLevel'));
      tr.append(tdSevLvl);

      const tdLikeLvl = createEl('td');
      tdLikeLvl.append(renderRiskSegments(hazard, hazardIndex, 'likelihoodLevel'));
      tr.append(tdLikeLvl);

      const tdRisk = createEl('td');
      tdRisk.append(renderRiskSegments(hazard, hazardIndex, 'riskScore', true));
      tr.append(tdRisk);

      // Recommendations and actions (single, full cell each)
      const tdReco = createEl('td');
      tdReco.append(renderRecommendationsCell(hazard, hazardIndex));
      tr.append(tdReco);

      const tdActions = createEl('td');
      tdActions.append(renderHazardActions(hazardIndex));
      tr.append(tdActions);

      tbody.append(tr);
    });

    container.append(tableEl);
    // After the table is in the DOM, align paired segments
    requestAnimationFrame(() => alignSegments(container));
  }

  // Cell renderers
  function renderHazardCell(hazard, hazardIndex) {
    const wrap = createEl('div', { class: 'stack' });
    const title = createEl('input', {
      type: 'text', value: hazard.title, placeholder: 'Hazard title',
      oninput: (e) => { hazard.title = e.target.value; scheduleSave(); },
      onblur: () => { scheduleSaveAndRerender(); }
    });
    const desc = createEl('textarea', {
      placeholder: 'Hazard description (optional)'
    });
    desc.value = hazard.description || '';
    desc.addEventListener('input', (e) => { hazard.description = e.target.value; scheduleSave(); });

    const causeBtnRow = createEl('div', { class: 'inline-controls' }, [
      createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.causes.push(createCause()); scheduleSaveAndRerender(); } }, '+ Cause', '+'),
      createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.consequences.push(createConsequence()); scheduleSaveAndRerender(); } }, '+ Consequence', '+'),
      createButtonWithTooltip({ class: 'icon edit-button', onclick: () => showBowtieDiagram(hazardIndex) }, 'Bow-tie', 'B'),
      createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { state.hazards.splice(hazardIndex, 1); scheduleSaveAndRerender(); } }, 'Remove hazard', '×')
    ]);

    wrap.append(title, desc, causeBtnRow);
    return wrap;
  }

  function renderCauseCell(hazard, hazardIndex, cause, rowIndex) {
    const wrap = createEl('div', { class: 'stack' });
    if (!cause) {
      const btn = createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.causes.push(createCause()); scheduleSaveAndRerender(); } }, '+ Add cause', '+');
      wrap.append(btn);
      return wrap;
    }
    const input = createEl('input', { type: 'text', value: cause.text, placeholder: 'Cause', oninput: (e) => { cause.text = e.target.value; scheduleSave(); } });
    const actions = createEl('div', { class: 'inline-controls' }, [
      createButtonWithTooltip({ class: 'icon copy-button', onclick: () => copyItem({ type: 'cause', hazardIndex, rowIndex }) }, 'Copy', 'C'),
      createButtonWithTooltip({ class: 'icon muted paste-button', onclick: () => pasteItem({ type: 'cause', hazardIndex, rowIndex }) }, 'Paste', 'P'),
      createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { hazard.causes.splice(rowIndex, 1); scheduleSaveAndRerender(); } }, 'Remove', '×')
    ]);
    wrap.append(input, actions);
    return wrap;
  }

  function renderConsequenceCell(hazard, hazardIndex, consequence, rowIndex) {
    const wrap = createEl('div', { class: 'stack' });
    if (!consequence) {
      const btn = createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.consequences.push(createConsequence()); scheduleSaveAndRerender(); } }, '+ Add consequence', '+');
      wrap.append(btn);
      return wrap;
    }
    const input = createEl('input', { type: 'text', value: consequence.text, placeholder: 'Consequence', oninput: (e) => { consequence.text = e.target.value; scheduleSave(); } });
    const actions = createEl('div', { class: 'inline-controls' }, [
      createButtonWithTooltip({ class: 'icon copy-button', onclick: () => copyItem({ type: 'consequence', hazardIndex, rowIndex }) }, 'Copy', 'C'),
      createButtonWithTooltip({ class: 'icon muted paste-button', onclick: () => pasteItem({ type: 'consequence', hazardIndex, rowIndex }) }, 'Paste', 'P'),
      createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { hazard.consequences.splice(rowIndex, 1); scheduleSaveAndRerender(); } }, 'Remove', '×')
    ]);
    wrap.append(input, actions);
    return wrap;
  }

  function renderMeasureStack(hazard, hazardIndex, ownerType, rowIndex) {
    const isCause = ownerType === 'cause';
    const owner = isCause ? hazard.causes[rowIndex] : hazard.consequences[rowIndex];
    const label = isCause ? 'Prevention measure' : 'Mitigation measure';
    const key = isCause ? 'preventionMeasures' : 'mitigationMeasures';
    const empty = createEl('div');
    if (!owner) return empty;
    owner[key] = owner[key] || [];

    const count = Math.max(owner[key].length, 1);
    const container = segmentedContainer(count);

    if (owner[key].length === 0) {
      const seg = createEl('div', { class: 'segment' });
      seg.append(createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { owner[key].push(createMeasure()); scheduleSaveAndRerender(); } }, `+ ${label}`, '+'));
      container.append(seg);
      return container;
    }

    owner[key].forEach((m, mi) => {
      const seg = createEl('div', { class: 'segment' });
      const input = createEl('input', { type: 'text', value: m.text, placeholder: label, oninput: (e) => { m.text = e.target.value; scheduleSave(); } });
      const actions = createEl('div', { class: 'inline-controls' }, [
        createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { owner[key].splice(mi + 1, 0, createMeasure()); scheduleSaveAndRerender(); } }, '+ Insert below', '+'),
        createButtonWithTooltip({ class: 'icon copy-button', onclick: () => copyItem({ type: 'measure', ownerType, hazardIndex, rowIndex, measureIndex: mi }) }, 'Copy', 'C'),
        createButtonWithTooltip({ class: 'icon muted paste-button', onclick: () => pasteItem({ type: 'measure', ownerType, hazardIndex, rowIndex, measureIndex: mi }) }, 'Paste', 'P'),
        createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { owner[key].splice(mi, 1); scheduleSaveAndRerender(); } }, 'Remove', '×')
      ]);
      seg.append(input, actions);
      container.append(seg);
    });

    return container;
  }

  function renderRiskField(hazard, hazardIndex, rowIndex, field, isComputed = false) {
    const consequence = hazard.consequences[rowIndex];
    const risk = consequence ? (consequence.risk || (consequence.risk = { severityCategory: '', severityLevel: '', likelihoodLevel: '', riskScore: '' })) : null;
    const wrap = createEl('div');
    if (!risk) return wrap;

    if (isComputed) {
      risk.riskScore = computeRiskScore(risk.severityLevel, risk.likelihoodLevel);
      const color = getRiskLevelColor(risk.severityLevel, risk.likelihoodLevel);
      wrap.textContent = risk.riskScore || '';
      if (color) wrap.style.backgroundColor = color;
      if (color) wrap.style.color = 'white';
      return wrap;
    }

    if (field === 'severityCategory') {
      const select = createEl('select', {
        value: risk[field] || '',
        onchange: (e) => { risk[field] = e.target.value; scheduleSave(); }
      });
      
      const emptyOption = createEl('option', { value: '', text: 'Select category' });
      select.append(emptyOption);
      
      ['personnel', 'asset', 'environmental', 'reputation', 'operation'].forEach(cat => {
        const option = createEl('option', { value: cat, text: cat.charAt(0).toUpperCase() + cat.slice(1) });
        if (cat === risk[field]) option.selected = true;
        select.append(option);
      });
      
      wrap.append(select);
    } else if (field === 'severityLevel') {
      const select = createEl('select', {
        value: risk[field] || '',
        onchange: (e) => { risk[field] = e.target.value; scheduleSaveAndRerender(); }
      });
      
      const emptyOption = createEl('option', { value: '', text: 'Select severity' });
      select.append(emptyOption);
      
      state.riskMatrix.severity.forEach(sev => {
        const option = createEl('option', { value: sev.id, text: `${sev.label} - ${sev.description}` });
        if (sev.id === risk[field]) option.selected = true;
        select.append(option);
      });
      
      wrap.append(select);
    } else if (field === 'likelihoodLevel') {
      const select = createEl('select', {
        value: risk[field] || '',
        onchange: (e) => { risk[field] = e.target.value; scheduleSaveAndRerender(); }
      });
      
      const emptyOption = createEl('option', { value: '', text: 'Select likelihood' });
      select.append(emptyOption);
      
      state.riskMatrix.likelihood.forEach(lik => {
        const option = createEl('option', { value: lik.id, text: `${lik.label} - ${lik.description}` });
        if (lik.id === risk[field]) option.selected = true;
        select.append(option);
      });
      
      wrap.append(select);
    }

    return wrap;
  }

  // Segmented containers
  function segmentedContainer(segmentCount) {
    const container = createEl('div', { class: 'segments' });
    // style set at runtime for equal division
    container.style.display = 'grid';
    container.style.gridTemplateRows = `repeat(${Math.max(segmentCount, 1)}, 1fr)`;
    container.style.height = '100%';
    return container;
  }

  function renderCauseSegments(hazard, hazardIndex) {
    const count = Math.max(hazard.causes.length, 1);
    const container = segmentedContainer(count);
    if (hazard.causes.length === 0) {
      const seg = createEl('div', { class: 'segment' });
      seg.append(createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.causes.push(createCause()); scheduleSaveAndRerender(); } }, '+ Add cause', '+'));
      container.append(seg);
      return container;
    }
    hazard.causes.forEach((cause, i) => {
      const seg = createEl('div', { class: 'segment', dataset: { hazardIndex: String(hazardIndex), kind: 'cause', segIndex: String(i) } });
      seg.append(renderCauseCell(hazard, hazardIndex, cause, i));
      container.append(seg);
    });
    return container;
  }

  function renderCauseMeasuresSegments(hazard, hazardIndex) {
    const count = Math.max(hazard.causes.length, 1);
    const container = segmentedContainer(count);
    if (hazard.causes.length === 0) {
      container.append(createEl('div', { class: 'segment' }));
      return container;
    }
    hazard.causes.forEach((_, i) => {
      const seg = createEl('div', { class: 'segment', dataset: { hazardIndex: String(hazardIndex), kind: 'cause-measures', segIndex: String(i) } });
      seg.append(renderMeasureStack(hazard, hazardIndex, 'cause', i));
      container.append(seg);
    });
    return container;
  }

  function renderConsequenceSegments(hazard, hazardIndex) {
    const count = Math.max(hazard.consequences.length, 1);
    const container = segmentedContainer(count);
    if (hazard.consequences.length === 0) {
      const seg = createEl('div', { class: 'segment' });
      seg.append(createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazard.consequences.push(createConsequence()); scheduleSaveAndRerender(); } }, '+ Add consequence', '+'));
      container.append(seg);
      return container;
    }
    hazard.consequences.forEach((consequence, i) => {
      const seg = createEl('div', { class: 'segment', dataset: { hazardIndex: String(hazardIndex), kind: 'consequence', segIndex: String(i) } });
      seg.append(renderConsequenceCell(hazard, hazardIndex, consequence, i));
      container.append(seg);
    });
    return container;
  }

  function renderConsequenceMeasuresSegments(hazard, hazardIndex) {
    const count = Math.max(hazard.consequences.length, 1);
    const container = segmentedContainer(count);
    if (hazard.consequences.length === 0) {
      container.append(createEl('div', { class: 'segment' }));
      return container;
    }
    hazard.consequences.forEach((_, i) => {
      const seg = createEl('div', { class: 'segment', dataset: { hazardIndex: String(hazardIndex), kind: 'consequence-measures', segIndex: String(i) } });
      seg.append(renderMeasureStack(hazard, hazardIndex, 'consequence', i));
      container.append(seg);
    });
    return container;
  }

  // Align heights of paired segments between causes and their measures, and consequences and their measures
  function alignSegments(root) {
    // Reset any previous inline heights
    qsa('.segments .segment', root).forEach((el) => { el.style.minHeight = ''; });

    // For each hazard index present in DOM
    const hazardIndices = new Set(qsa('[data-hazard-index]', root).map((el) => el.dataset.hazardIndex));
    hazardIndices.forEach((hIdx) => {
      // Align causes with cause-measures
      const causeSegs = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="cause"]`, root);
      const causeMeaSegs = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="cause-measures"]`, root);
      const n = Math.max(causeSegs.length, causeMeaSegs.length);
      for (let i = 0; i < n; i += 1) {
        const a = causeSegs[i];
        const b = causeMeaSegs[i];
        if (!a || !b) continue;
        const target = Math.max(a.scrollHeight, b.scrollHeight);
        a.style.minHeight = `${target}px`;
        b.style.minHeight = `${target}px`;
      }

      // Align consequences with consequence-measures and risk columns
      const consSegs = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="consequence"]`, root);
      const consMeaSegs = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="consequence-measures"]`, root);
      const riskSevCat = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="risk-sev-cat"]`, root);
      const riskSev = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="risk-sev"]`, root);
      const riskLike = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="risk-like"]`, root);
      const riskScore = qsa(`.segment[data-hazard-index="${hIdx}"][data-kind="risk-score"]`, root);
      const m = consSegs.length;
      for (let i = 0; i < m; i += 1) {
        const elems = [consSegs[i], consMeaSegs[i], riskSevCat[i], riskSev[i], riskLike[i], riskScore[i]].filter(Boolean);
        if (elems.length === 0) continue;
        const target = Math.max(...elems.map(el => el.scrollHeight));
        elems.forEach(el => { el.style.minHeight = `${target}px`; });
      }
    });

    // Re-run after images/fonts load or window resize
  }

  window.addEventListener('resize', () => {
    const container = byId('table-container');
    alignSegments(container);
  });

  function renderRiskSegments(hazard, hazardIndex, field, isComputed = false) {
    const count = Math.max(hazard.consequences.length, 1);
    const container = segmentedContainer(count);
    if (hazard.consequences.length === 0) {
      container.append(createEl('div', { class: 'segment' }));
      return container;
    }
    hazard.consequences.forEach((_, i) => {
      let kind = 'risk-score';
      if (field === 'severityCategory') kind = 'risk-sev-cat';
      else if (field === 'severityLevel') kind = 'risk-sev';
      else if (field === 'likelihoodLevel') kind = 'risk-like';
      const seg = createEl('div', { class: 'segment', dataset: { hazardIndex: String(hazardIndex), kind, segIndex: String(i) } });
      seg.append(renderRiskField(hazard, hazardIndex, i, field, isComputed));
      container.append(seg);
    });
    return container;
  }

  function renderRecommendationsCell(hazard, hazardIndex) {
    const wrap = createEl('div', { class: 'stack' });
    hazard.recommendations.forEach((r, ri) => {
      const seg = createEl('div', { class: 'row-segment' });
      const action = createEl('input', { type: 'text', value: r.action, placeholder: 'Action', oninput: (e) => { r.action = e.target.value; scheduleSave(); } });
      const resp = createEl('input', { type: 'text', value: r.responsible, placeholder: 'Responsible', oninput: (e) => { r.responsible = e.target.value; scheduleSave(); } });
      const actions = createEl('div', { class: 'inline-controls' }, [
        createButtonWithTooltip({ class: 'icon copy-button', onclick: () => copyItem({ type: 'recommendation', hazardIndex, recoIndex: ri }) }, 'Copy', 'C'),
        createButtonWithTooltip({ class: 'icon muted paste-button', onclick: () => pasteItem({ type: 'recommendation', hazardIndex, recoIndex: ri }) }, 'Paste', 'P'),
        createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { hazards(hazardIndex).recommendations.splice(ri, 1); scheduleSaveAndRerender(); } }, 'Remove', '×')
      ]);
      seg.append(action, resp, actions);
      wrap.append(seg);
    });

    const addBtn = createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { hazards(hazardIndex).recommendations.push(createRecommendation()); scheduleSaveAndRerender(); } }, '+ Recommendation', '+');
    wrap.append(addBtn);
    return wrap;
  }

  function renderHazardActions(hazardIndex) {
    const wrap = createEl('div', { class: 'cell-actions' });
    const addAbove = createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { state.hazards.splice(hazardIndex, 0, createHazard()); scheduleSaveAndRerender(); } }, 'Add above', '+');
    const addBelow = createButtonWithTooltip({ class: 'icon primary add-button', onclick: () => { state.hazards.splice(hazardIndex + 1, 0, createHazard()); scheduleSaveAndRerender(); } }, 'Add below', '+');
    const duplicate = createButtonWithTooltip({ class: 'icon duplicate-button', onclick: () => { const clone = deepClone(state.hazards[hazardIndex]); clone.id = generateId(); state.hazards.splice(hazardIndex + 1, 0, clone); scheduleSaveAndRerender(); } }, 'Duplicate', 'D');
    const copyBtn = createButtonWithTooltip({ class: 'icon copy-button', onclick: () => { state.clipboard = { type: 'hazard', data: deepClone(state.hazards[hazardIndex]) }; } }, 'Copy', 'C');
    const pasteBtn = createButtonWithTooltip({ class: 'icon muted paste-button', onclick: () => {
      if (!state.clipboard || state.clipboard.type !== 'hazard') return;
      const clone = deepClone(state.clipboard.data);
      clone.id = generateId();
      state.hazards.splice(hazardIndex + 1, 0, clone);
      scheduleSaveAndRerender();
    } }, 'Paste', 'P');
    const remove = createButtonWithTooltip({ class: 'icon danger remove-button', onclick: () => { state.hazards.splice(hazardIndex, 1); scheduleSaveAndRerender(); } }, 'Remove', '×');
    wrap.append(addAbove, addBelow, duplicate, copyBtn, pasteBtn, remove);
    return wrap;
  }

  // Copy/paste logic for granular items
  function copyItem(ref) {
    const { type } = ref;
    if (type === 'cause') {
      const hazard = state.hazards[ref.hazardIndex];
      const item = hazard.causes[ref.rowIndex];
      state.clipboard = { type: 'cause', data: deepClone(item) };
    } else if (type === 'consequence') {
      const hazard = state.hazards[ref.hazardIndex];
      const item = hazard.consequences[ref.rowIndex];
      state.clipboard = { type: 'consequence', data: deepClone(item) };
    } else if (type === 'measure') {
      const hazard = state.hazards[ref.hazardIndex];
      const owner = ref.ownerType === 'cause' ? hazard.causes[ref.rowIndex] : hazard.consequences[ref.rowIndex];
      const key = ref.ownerType === 'cause' ? 'preventionMeasures' : 'mitigationMeasures';
      const item = owner?.[key]?.[ref.measureIndex];
      if (item) state.clipboard = { type: 'measure', ownerType: ref.ownerType, data: deepClone(item) };
    } else if (type === 'recommendation') {
      const hazard = state.hazards[ref.hazardIndex];
      const item = hazard.recommendations[ref.recoIndex];
      state.clipboard = { type: 'recommendation', data: deepClone(item) };
    }
  }

  function pasteItem(ref) {
    if (!state.clipboard) return;
    const clip = state.clipboard;
    if (ref.type === 'cause' && clip.type === 'cause') {
      const hazard = state.hazards[ref.hazardIndex];
      const clone = deepClone(clip.data); clone.id = generateId();
      hazard.causes.splice(ref.rowIndex + 1, 0, clone);
      scheduleSaveAndRerender();
    } else if (ref.type === 'consequence' && clip.type === 'consequence') {
      const hazard = state.hazards[ref.hazardIndex];
      const clone = deepClone(clip.data); clone.id = generateId();
      hazard.consequences.splice(ref.rowIndex + 1, 0, clone);
      scheduleSaveAndRerender();
    } else if (ref.type === 'measure' && clip.type === 'measure' && ref.ownerType === clip.ownerType) {
      const hazard = state.hazards[ref.hazardIndex];
      const owner = ref.ownerType === 'cause' ? hazard.causes[ref.rowIndex] : hazard.consequences[ref.rowIndex];
      const key = ref.ownerType === 'cause' ? 'preventionMeasures' : 'mitigationMeasures';
      const clone = deepClone(clip.data); clone.id = generateId();
      owner[key].splice((ref.measureIndex ?? owner[key].length) + 1, 0, clone);
      scheduleSaveAndRerender();
    } else if (ref.type === 'recommendation' && clip.type === 'recommendation') {
      const hazard = state.hazards[ref.hazardIndex];
      const clone = deepClone(clip.data); clone.id = generateId();
      hazard.recommendations.splice((ref.recoIndex ?? hazard.recommendations.length) + 1, 0, clone);
      scheduleSaveAndRerender();
    }
  }

  // Helpers: access and save/rerender
  const hazards = (i) => state.hazards[i];

  let rerenderTimer = null;
  function scheduleSave() {
    if (state.autosave) persist.save({ hazards: state.hazards, riskMatrix: state.riskMatrix, compactMode: state.compactMode, bowtieColors: state.bowtieColors });
  }
  function scheduleSaveAndRerender() {
    scheduleSave();
    clearTimeout(rerenderTimer);
    rerenderTimer = setTimeout(() => {
      const container = byId('table-container');
      container.innerHTML = '';
      renderTable(container);
    }, 0);
  }

  // Global UI actions
  function wireGlobalActions() {
    // Tab switching
    byId('tab-hazards').addEventListener('click', () => switchTab('hazards'));
    byId('tab-risk-matrix').addEventListener('click', () => switchTab('risk-matrix'));

    byId('add-hazard').addEventListener('click', () => {
      state.hazards.push(createHazard());
      scheduleSaveAndRerender();
    });

    byId('clear-all').addEventListener('click', () => {
      if (!confirm('Clear all data?')) return;
      state.hazards = [];
      persist.save({ hazards: state.hazards, riskMatrix: state.riskMatrix });
      scheduleSaveAndRerender();
    });

    byId('autosave-toggle').addEventListener('change', (e) => {
      state.autosave = !!e.target.checked;
      if (state.autosave) persist.save({ hazards: state.hazards, riskMatrix: state.riskMatrix });
    });

    byId('compact-mode').addEventListener('change', (e) => {
      state.compactMode = !!e.target.checked;
      document.body.classList.toggle('compact-mode', state.compactMode);
      scheduleSaveAndRerender(); // Re-render to update button styles
    });

    byId('export-json').addEventListener('click', () => {
      const blob = new Blob([JSON.stringify({ hazards: state.hazards, riskMatrix: state.riskMatrix }, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'hazid.json'; a.click();
      URL.revokeObjectURL(url);
    });

    byId('export-excel').addEventListener('click', exportToExcel);

    // Bow-tie colors management
    byId('import-bowtie-colors').addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = 'application/json';
      input.onchange = () => {
        const file = input.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const data = JSON.parse(String(reader.result || '{}'));
            loadBowtieColorsFromJSON(data);
          } catch (e) {
            alert('Failed to import bow-tie colors JSON: ' + e.message);
          }
        };
        reader.readAsText(file);
      };
      input.click();
    });

    byId('export-bowtie-colors').addEventListener('click', exportBowtieColors);
    byId('load-default-bowtie-colors').addEventListener('click', loadDefaultBowtieColors);

    byId('import-json').addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file'; input.accept = 'application/json';
      input.onchange = () => {
        const file = input.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const data = JSON.parse(String(reader.result || '[]'));
            if (Array.isArray(data)) {
              // Legacy format
              state.hazards = data;
            } else if (data.hazards) {
              // New format
              state.hazards = data.hazards;
              if (data.riskMatrix) {
                state.riskMatrix = { ...state.riskMatrix, ...data.riskMatrix };
                updateRiskMatrix();
              }
            } else {
              throw new Error('Invalid file format');
            }
            persist.save({ hazards: state.hazards, riskMatrix: state.riskMatrix });
            scheduleSaveAndRerender();
            renderRiskMatrixConfig();
          } catch (e) {
            alert('Failed to import JSON: ' + e.message);
          }
        };
        reader.readAsText(file);
      };
      input.click();
    });

    // Risk matrix import/export buttons
    byId('import-risk-matrix').addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file'; input.accept = 'application/json';
      input.onchange = () => {
        const file = input.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const data = JSON.parse(String(reader.result || '{}'));
            loadRiskMatrixFromJSON(data);
            scheduleSaveAndRerender();
            renderRiskMatrixConfig();
            alert('Risk matrix imported successfully!');
          } catch (e) {
            alert('Failed to import risk matrix JSON: ' + e.message);
          }
        };
        reader.readAsText(file);
      };
      input.click();
    });

    byId('export-risk-matrix').addEventListener('click', () => {
      const config = {
        likelihoodLevels: state.riskMatrix.likelihood.length,
        likelihoodDescriptions: state.riskMatrix.likelihood.map(l => ({ id: l.id, label: l.label, description: l.description })),
        severityCategories: ['personnel', 'asset', 'environmental', 'reputation', 'operation'],
        severityDescriptions: state.riskMatrix.severityDescriptions,
        riskLevels: state.riskMatrix.riskLevels.length,
        riskLevelDescriptions: state.riskMatrix.riskLevels.map(r => ({ id: r.id, label: r.label, color: r.color })),
        matrix: Object.fromEntries(
          Object.entries(state.riskMatrix.matrix).map(([key, riskLevel]) => [key, riskLevel.id])
        )
      };
      
      const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'risk-matrix.json'; a.click();
      URL.revokeObjectURL(url);
    });

    byId('load-default-matrix').addEventListener('click', () => {
      if (confirm('Load default risk matrix? This will replace your current configuration.')) {
        loadDefaultRiskMatrix();
        scheduleSaveAndRerender();
        renderRiskMatrixConfig();
        alert('Default risk matrix loaded!');
      }
    });
  }

  // Excel export using ExcelJS - rewritten from scratch
  async function exportToExcel() {
    try {
      if (!window.ExcelJS) { alert('Excel export library not loaded'); return; }
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('HAZID');
      const hazards = state.hazards;

      // Header row
      const headers = ['Hazard','Causes','Prevention measures','Consequences','Mitigation measures','Severity category','Severity level','Likelihood level','Risk','Recommendations'];
      sheet.addRow(headers);
      
      // Style header
      const headerRow = sheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      headerRow.eachCell((cell) => { 
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF024F75' } }; 
        cell.border = headerBorders(); 
      });

      let currentRow = 2;
      const borderColor = '#024F75';

      state.hazards.forEach((hazard) => {
        const causeCount = Math.max(hazard.causes.length, 1);
        const consCount = Math.max(hazard.consequences.length, 1);
        
        // Calculate total prevention measures (causes with no measures count as 1)
        const totalPreventionMeasures = hazard.causes.reduce((sum, cause) => {
          const measures = cause ? (cause.preventionMeasures || []).length : 0;
          return sum + Math.max(measures, 1);
        }, 0);
        
        // Calculate total mitigation measures (consequences with no measures count as 1)
        const totalMitigationMeasures = hazard.consequences.reduce((sum, cons) => {
          const measures = cons ? (cons.mitigationMeasures || []).length : 0;
          return sum + Math.max(measures, 1);
        }, 0);
        
        // Total rows needed = max of prevention vs mitigation measures
        const blockRows = Math.max(totalPreventionMeasures, totalMitigationMeasures, 1);
        const startRow = currentRow;

        // Pre-create all rows for this hazard
        for (let i = 0; i < blockRows; i += 1) {
          sheet.addRow(['','','','','','','','','','']);
        }

        // Merge Hazard column across all rows
        mergeAndSet(sheet, startRow, 1, blockRows, hazard.title + (hazard.description ? `\n${hazard.description}` : ''));

        // Process Causes and Prevention Measures
        let causeRowOffset = 0;
        hazard.causes.forEach((cause, causeIndex) => {
          const actualMeasures = cause ? (cause.preventionMeasures || []).length : 0;
          let causeRows = Math.max(actualMeasures, 1);
          
          // Last cause expands to fill remaining rows
          if (causeIndex === hazard.causes.length - 1) {
            causeRows = blockRows - causeRowOffset;
          }
          
          const causeStartRow = startRow + causeRowOffset;
          
          // Merge cause text across its rows
          mergeAndSet(sheet, causeStartRow, 2, causeRows, cause ? (cause.text || '') : '');
          
          // Add prevention measures
          if (actualMeasures > 0) {
            const measures = cause.preventionMeasures;
            measures.forEach((measure, measureIndex) => {
              const measureRow = causeStartRow + measureIndex;
              // Last measure expands to fill remaining rows
              const measureRows = (measureIndex === measures.length - 1) ? 
                (causeRows - measureIndex) : 1;
              mergeAndSet(sheet, measureRow, 3, measureRows, measure.text || '');
            });
          }
          
          causeRowOffset += causeRows;
        });

        // Process Consequences and Mitigation Measures
        let consRowOffset = 0;
        hazard.consequences.forEach((cons, consIndex) => {
          const actualMeasures = cons ? (cons.mitigationMeasures || []).length : 0;
          let consRows = Math.max(actualMeasures, 1);
          
          // Last consequence expands to fill remaining rows
          if (consIndex === hazard.consequences.length - 1) {
            consRows = blockRows - consRowOffset;
          }
          
          const consStartRow = startRow + consRowOffset;
          
          // Merge consequence text across its rows
          mergeAndSet(sheet, consStartRow, 4, consRows, cons ? (cons.text || '') : '');
          
          // Add mitigation measures
          if (actualMeasures > 0) {
            const measures = cons.mitigationMeasures;
            measures.forEach((measure, measureIndex) => {
              const measureRow = consStartRow + measureIndex;
              // Last measure expands to fill remaining rows
              const measureRows = (measureIndex === measures.length - 1) ? 
                (consRows - measureIndex) : 1;
              mergeAndSet(sheet, measureRow, 5, measureRows, measure.text || '');
            });
          }
          
          // Risk columns - merge across consequence rows
          const risk = cons ? cons.risk || {} : {};
          mergeAndSet(sheet, consStartRow, 6, consRows, risk.severityCategory || '');
          mergeAndSet(sheet, consStartRow, 7, consRows, risk.severityLevel || '');
          mergeAndSet(sheet, consStartRow, 8, consRows, risk.likelihoodLevel || '');
          mergeAndSet(sheet, consStartRow, 9, consRows, risk.riskScore || '');
          
          consRowOffset += consRows;
        });

        // Merge Recommendations across all rows
        mergeAndSet(sheet, startRow, 10, blockRows, hazard.recommendations.map(r => `${r.action} — ${r.responsible}`).join('\n'));

        // Apply borders and styling
        for (let r = startRow; r < startRow + blockRows; r += 1) {
          for (let c = 1; c <= 10; c += 1) {
            const cell = sheet.getCell(r, c);
            cell.border = allBorders(borderColor);
            cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
          }
        }

        // Apply risk colors after all merging and styling is complete
        consRowOffset = 0;
        hazard.consequences.forEach((cons, consIndex) => {
          const actualMeasures = cons ? (cons.mitigationMeasures || []).length : 0;
          let consRows = Math.max(actualMeasures, 1);
          
          // Last consequence expands to fill remaining rows
          if (consIndex === hazard.consequences.length - 1) {
            consRows = blockRows - consRowOffset;
          }
          
          const consStartRow = startRow + consRowOffset;
          const risk = cons ? cons.risk || {} : {};
          const riskColor = getRiskLevelColor(risk.severityLevel, risk.likelihoodLevel);
          
          if (riskColor && risk.riskScore) {
            // Apply color to all rows in the merged range
            for (let r = consStartRow; r < consStartRow + consRows; r++) {
              const riskCell = sheet.getCell(r, 9);
              riskCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cssHexToARGB(riskColor) } };
              riskCell.font = { color: { argb: 'FFFFFFFF' } };
            }
          }
          
          consRowOffset += consRows;
        });

        currentRow = startRow + blockRows;
      });

      // Column widths for readability
      const widths = [30,24,28,24,28,18,14,18,12,30];
      sheet.columns = widths.map(w => ({ width: w }));

      // Create Summary Worksheet
      try {
        const summaryWorksheet = workbook.addWorksheet('Risk Summary');
        
        // Summary headers
        const summaryHeaders = ['Hazard', 'Consequence', 'Severity Category', 'Severity Index', 'Likelihood Index', 'Risk Ranking'];
        summaryHeaders.forEach((header, index) => {
          const cell = summaryWorksheet.getCell(1, index + 1);
          cell.value = header;
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF024F75' }
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Add summary data with hazard spanning
        let summaryRow = 2;
        hazards.forEach((hazard) => {
          // Count risk-ranked consequences for this hazard
          const riskRankedConsequences = hazard.consequences.filter(consequence => 
            consequence.risk && consequence.risk.severityLevel && consequence.risk.likelihoodLevel
          );
          
          if (riskRankedConsequences.length > 0) {
            const hazardStartRow = summaryRow;
            let consequenceCount = 0;
            
            riskRankedConsequences.forEach((consequence) => {
              // Consequence
              summaryWorksheet.getCell(summaryRow, 2).value = consequence.text || 'Untitled Consequence';
              
              // Severity Category
              summaryWorksheet.getCell(summaryRow, 3).value = consequence.risk.severityCategory || '';
              
              // Severity Index
              summaryWorksheet.getCell(summaryRow, 4).value = consequence.risk.severityLevel || '';
              
              // Likelihood Index
              summaryWorksheet.getCell(summaryRow, 5).value = consequence.risk.likelihoodLevel || '';
              
              // Risk Ranking
              const riskLevel = getRiskLevel(consequence.risk.severityLevel, consequence.risk.likelihoodLevel);
              summaryWorksheet.getCell(summaryRow, 6).value = riskLevel || '';
              
              // Apply risk color to risk ranking cell
              const riskColor = getRiskLevelColor(consequence.risk.severityLevel, consequence.risk.likelihoodLevel);
              if (riskColor) {
                const riskCell = summaryWorksheet.getCell(summaryRow, 6);
                riskCell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: cssHexToARGB(riskColor) }
                };
                riskCell.font = { color: { argb: 'FFFFFFFF' } };
              }
              
              summaryRow++;
              consequenceCount++;
            });
            
            // Add hazard name to first row and merge if multiple consequences
            summaryWorksheet.getCell(hazardStartRow, 1).value = hazard.title || 'Untitled Hazard';
            if (consequenceCount > 1) {
              summaryWorksheet.mergeCells(hazardStartRow, 1, hazardStartRow + consequenceCount - 1, 1);
            }
          }
        });

        // Style summary worksheet
        // Set column widths
        summaryWorksheet.getColumn(1).width = 30; // Hazard
        summaryWorksheet.getColumn(2).width = 40; // Consequence
        summaryWorksheet.getColumn(3).width = 20; // Severity Category
        summaryWorksheet.getColumn(4).width = 15; // Severity Index
        summaryWorksheet.getColumn(5).width = 15; // Likelihood Index
        summaryWorksheet.getColumn(6).width = 15; // Risk Ranking

        // Add borders to all cells
        if (summaryRow > 2) {
          // Apply borders to all data cells
          for (let row = 1; row < summaryRow; row++) {
            for (let col = 1; col <= 6; col++) {
              const cell = summaryWorksheet.getCell(row, col);
              if (row === 1) {
                // Header borders (white)
                cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
                };
              } else {
                // Data cell borders (blue)
                cell.border = {
                  top: { style: 'thin', color: { argb: 'FF024F75' } },
                  left: { style: 'thin', color: { argb: 'FF024F75' } },
                  bottom: { style: 'thin', color: { argb: 'FF024F75' } },
                  right: { style: 'thin', color: { argb: 'FF024F75' } }
                };
              }
            }
          }
        }
      } catch (summaryError) {
        console.error('Summary worksheet creation failed:', summaryError);
        // Continue with main worksheet even if summary fails
      }

      // Create Recommendations Worksheet
      try {
        const recommendationsWorksheet = workbook.addWorksheet('Recommendations');
        
        // Recommendations headers
        const recommendationsHeaders = ['Hazard', 'Action', 'Responsible Party'];
        recommendationsHeaders.forEach((header, index) => {
          const cell = recommendationsWorksheet.getCell(1, index + 1);
          cell.value = header;
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF024F75' }
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Add recommendations data with hazard spanning
        let recommendationsRow = 2;
        hazards.forEach((hazard) => {
          if (hazard.recommendations && hazard.recommendations.length > 0) {
            const hazardStartRow = recommendationsRow;
            let recommendationCount = 0;
            
            hazard.recommendations.forEach((recommendation) => {
              // Action
              recommendationsWorksheet.getCell(recommendationsRow, 2).value = recommendation.action || '';
              
              // Responsible Party
              recommendationsWorksheet.getCell(recommendationsRow, 3).value = recommendation.responsible || '';
              
              recommendationsRow++;
              recommendationCount++;
            });
            
            // Add hazard name to first row and merge if multiple recommendations
            recommendationsWorksheet.getCell(hazardStartRow, 1).value = hazard.title || 'Untitled Hazard';
            if (recommendationCount > 1) {
              recommendationsWorksheet.mergeCells(hazardStartRow, 1, hazardStartRow + recommendationCount - 1, 1);
            }
          }
        });

        // Style recommendations worksheet
        // Set column widths
        recommendationsWorksheet.getColumn(1).width = 30; // Hazard
        recommendationsWorksheet.getColumn(2).width = 50; // Action
        recommendationsWorksheet.getColumn(3).width = 25; // Responsible Party

        // Add borders to all cells
        if (recommendationsRow > 2) {
          // Apply borders to all data cells
          for (let row = 1; row < recommendationsRow; row++) {
            for (let col = 1; col <= 3; col++) {
              const cell = recommendationsWorksheet.getCell(row, col);
              if (row === 1) {
                // Header borders (white)
                cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
                };
              } else {
                // Data cell borders (blue)
                cell.border = {
                  top: { style: 'thin', color: { argb: 'FF024F75' } },
                  left: { style: 'thin', color: { argb: 'FF024F75' } },
                  bottom: { style: 'thin', color: { argb: 'FF024F75' } },
                  right: { style: 'thin', color: { argb: 'FF024F75' } }
                };
              }
            }
          }
        }
      } catch (recommendationsError) {
        console.error('Recommendations worksheet creation failed:', recommendationsError);
        // Continue even if recommendations worksheet fails
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'hazid.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Excel export failed', err);
      alert('Excel export failed: ' + err.message + '. Please check the console for details.');
    }
  }

  function allBorders(hex) {
    const argb = cssHexToARGB(hex);
    return {
      top: { style: 'thin', color: { argb } },
      left: { style: 'thin', color: { argb } },
      bottom: { style: 'thin', color: { argb } },
      right: { style: 'thin', color: { argb } }
    };
  }

  function headerBorders() {
    const whiteArgb = 'FFFFFFFF';
    const blueArgb = cssHexToARGB('#024F75');
    return {
      top: { style: 'thin', color: { argb: blueArgb } },
      left: { style: 'thin', color: { argb: whiteArgb } },
      bottom: { style: 'thin', color: { argb: blueArgb } },
      right: { style: 'thin', color: { argb: whiteArgb } }
    };
  }

  function cssHexToARGB(hex) {
    const clean = hex.replace('#','');
    if (clean.length === 6) return 'FF' + clean.toUpperCase();
    if (clean.length === 3) return 'FF' + clean.split('').map(x => x + x).join('').toUpperCase();
    return 'FF000000';
  }

  // Merge helper for Excel export
  function mergeAndSet(sheet, startRow, col, rowSpan, value) {
    const endRow = startRow + Math.max(rowSpan, 1) - 1;
    if (endRow > startRow) {
      sheet.mergeCells(startRow, col, endRow, col);
    }
    const cell = sheet.getCell(startRow, col);
    cell.value = value == null ? '' : String(value);
    cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
  }

  // Risk Matrix UI
  function renderRiskMatrixConfig() {
    renderSeverityDescriptionsTable();
    renderRiskMatrixTable();
  }

  function renderRiskMatrixTable() {
    const container = byId('risk-matrix-table');
    container.innerHTML = '';
    
    if (state.riskMatrix.likelihood.length === 0 || state.riskMatrix.severity.length === 0) {
      container.textContent = 'Add likelihood and severity levels to see the matrix';
      return;
    }
    
    // Create a proper HTML table
    const table = createEl('table', { class: 'risk-matrix-table' });
    const thead = createEl('thead');
    const tbody = createEl('tbody');
    
    // Header row: empty cell + likelihood headers
    const headerRow = createEl('tr');
    const emptyHeader = createEl('th', { class: 'risk-header', text: 'L/S' });
    headerRow.append(emptyHeader);
    
    state.riskMatrix.likelihood.forEach(lik => {
      const th = createEl('th', { class: 'risk-header', text: lik.label });
      headerRow.append(th);
    });
    thead.append(headerRow);
    
    // Data rows: severity label + risk cells
    state.riskMatrix.severity.forEach(sev => {
      const row = createEl('tr');
      
      // Severity label in first column
      const sevCell = createEl('th', { class: 'risk-header', text: sev.label });
      row.append(sevCell);
      
      // Risk cells for this severity level across all likelihood levels
      state.riskMatrix.likelihood.forEach(lik => {
        const key = `${lik.id}-${sev.id}`;
        const riskLevel = state.riskMatrix.matrix[key];
        const cell = createEl('td', { 
          class: 'risk-cell',
          style: `background: ${riskLevel ? riskLevel.color : '#ccc'};`,
          text: riskLevel ? riskLevel.label : '?'
        });
        row.append(cell);
      });
      
      tbody.append(row);
    });
    
    table.append(thead, tbody);
    container.append(table);
    
    console.log('Risk matrix table created:', {
      likelihoodCount: state.riskMatrix.likelihood.length,
      severityCount: state.riskMatrix.severity.length,
      totalCells: state.riskMatrix.likelihood.length * state.riskMatrix.severity.length
    });
  }

  // Risk matrix import/export functions
  function loadDefaultRiskMatrix() {
    state.riskMatrix = {
      likelihood: [
        { id: 'A', label: 'A', description: 'Very unlikely' },
        { id: 'B', label: 'B', description: 'Unlikely' },
        { id: 'C', label: 'C', description: 'Possible' },
        { id: 'D', label: 'D', description: 'Likely' },
        { id: 'E', label: 'E', description: 'Very likely' }
      ],
      severity: [
        { id: '1', label: '1', description: 'Negligible effect', category: 'personnel' },
        { id: '2', label: '2', description: 'Minor effect', category: 'personnel' },
        { id: '3', label: '3', description: 'Moderate effect', category: 'personnel' },
        { id: '4', label: '4', description: 'Major effect', category: 'personnel' },
        { id: '5', label: '5', description: 'Severest effect', category: 'personnel' }
      ],
      severityDescriptions: {
        '1': {
          personnel: 'Minor injury, no lost time',
          asset: 'Minor damage, easily repairable',
          environmental: 'Minimal environmental impact',
          reputation: 'No reputation impact',
          operation: 'No operational impact'
        },
        '2': {
          personnel: 'Minor injury, some lost time',
          asset: 'Moderate damage, repairable',
          environmental: 'Minor environmental impact',
          reputation: 'Minor local reputation impact',
          operation: 'Minor operational disruption'
        },
        '3': {
          personnel: 'Serious injury, significant lost time',
          asset: 'Major damage, expensive repair',
          environmental: 'Moderate environmental impact',
          reputation: 'Moderate reputation impact',
          operation: 'Moderate operational disruption'
        },
        '4': {
          personnel: 'Major injury, permanent disability',
          asset: 'Severe damage, major repair cost',
          environmental: 'Major environmental impact',
          reputation: 'Major reputation impact',
          operation: 'Major operational disruption'
        },
        '5': {
          personnel: 'Multiple fatalities',
          asset: 'Total loss of facility',
          environmental: 'Permanent environmental damage',
          reputation: 'Major international effect',
          operation: 'Loss of operation up to a year'
        }
      },
      riskLevels: [
        { id: 'low', label: 'Low', color: '#28a745' },
        { id: 'medium', label: 'Medium', color: '#ffc107' },
        { id: 'high', label: 'High', color: '#dc3545' }
      ],
      matrix: {}
    };
    updateRiskMatrix();
  }

  function loadRiskMatrixFromJSON(data) {
    // Load likelihood levels
    if (data.likelihoodDescriptions && Array.isArray(data.likelihoodDescriptions)) {
      state.riskMatrix.likelihood = data.likelihoodDescriptions.map(l => ({
        id: l.id,
        label: l.label,
        description: l.description
      }));
    }

    // Load severity descriptions
    if (data.severityDescriptions) {
      state.riskMatrix.severityDescriptions = data.severityDescriptions;
    }

    // Load risk levels
    if (data.riskLevelDescriptions && Array.isArray(data.riskLevelDescriptions)) {
      state.riskMatrix.riskLevels = data.riskLevelDescriptions.map(r => ({
        id: r.id,
        label: r.label,
        color: r.color
      }));
    }

    // Load matrix assignments (convert IDs back to full objects)
    if (data.matrix) {
      state.riskMatrix.matrix = {};
      Object.entries(data.matrix).forEach(([key, riskLevelId]) => {
        const riskLevel = state.riskMatrix.riskLevels.find(r => r.id === riskLevelId);
        if (riskLevel) {
          state.riskMatrix.matrix[key] = riskLevel;
        }
      });
    } else {
      updateRiskMatrix();
    }
  }

  function renderSeverityDescriptionsTable() {
    const container = byId('severity-descriptions-table');
    container.innerHTML = '';
    
    const categories = ['personnel', 'asset', 'environmental', 'reputation', 'operation'];
    const severityLevels = state.riskMatrix.severity.map(s => s.id);
    
    // Create table
    const table = createEl('table', { class: 'severity-descriptions-table' });
    const thead = createEl('thead');
    const tbody = createEl('tbody');
    
    // Header row: empty cell + category headers
    const headerRow = createEl('tr');
    const emptyHeader = createEl('th', { text: 'Severity/Category' });
    headerRow.append(emptyHeader);
    
    categories.forEach(cat => {
      const th = createEl('th', { text: cat.charAt(0).toUpperCase() + cat.slice(1) });
      headerRow.append(th);
    });
    thead.append(headerRow);
    
    // Data rows: severity level + description inputs
    severityLevels.forEach(severityId => {
      const row = createEl('tr');
      
      // Severity level in first column
      const sevCell = createEl('th', { text: severityId });
      row.append(sevCell);
      
      // Description inputs for each category
      categories.forEach(category => {
        const cell = createEl('td');
        const input = createEl('textarea', {
          class: 'severity-description-input',
          placeholder: `Enter description for severity ${severityId} - ${category}`,
          oninput: (e) => {
            if (!state.riskMatrix.severityDescriptions[severityId]) {
              state.riskMatrix.severityDescriptions[severityId] = {};
            }
            state.riskMatrix.severityDescriptions[severityId][category] = e.target.value;
            scheduleSave();
          }
        });
        
        // Set the current value
        const currentValue = state.riskMatrix.severityDescriptions[severityId]?.[category] || '';
        input.value = currentValue;
        cell.append(input);
        row.append(cell);
      });
      
      tbody.append(row);
    });
    
    table.append(thead, tbody);
    container.append(table);
  }

  // Tab switching
  function switchTab(tabName) {
    // Update tab buttons
    qsa('.tab-button').forEach(btn => btn.classList.remove('active'));
    byId(`tab-${tabName}`).classList.add('active');
    
    // Update panels
    qsa('.panel').forEach(panel => panel.classList.remove('active'));
    byId(`${tabName}-panel`).classList.add('active');
  }

  // Bow-tie Diagram functionality
  function showBowtieDiagram(hazardIndex) {
    const hazard = state.hazards[hazardIndex];
    const modal = document.getElementById('bowtie-modal');
    const title = document.getElementById('bowtie-title');
    const svg = document.getElementById('bowtie-svg');

    title.textContent = `Bow-tie Diagram: ${hazard.title || 'Untitled Hazard'}`;

    // Clear previous content
    svg.innerHTML = '';

    // Set up SVG dimensions
    const width = 1000;
    const height = 600;
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);

    // Load bow-tie colors from state or use defaults
    const bowtieColors = state.bowtieColors || {
      hazard: { fill: '#ff6b6b', stroke: '#d63031' },
      cause: { fill: '#74b9ff', stroke: '#0984e3' },
      consequence: { fill: '#fd79a8', stroke: '#e84393' },
      measure: { fill: '#55a3ff', stroke: '#2d3436' },
      line: '#636e72',
      measureLine: '#55a3ff'
    };
    
    // Center coordinates
    const centerX = width / 2;
    const centerY = height / 2;
    
    // Draw hazard in center
    const hazardRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    hazardRect.setAttribute('x', centerX - 60);
    hazardRect.setAttribute('y', centerY - 30);
    hazardRect.setAttribute('width', 120);
    hazardRect.setAttribute('height', 60);
    hazardRect.setAttribute('rx', 10);
    hazardRect.setAttribute('fill', bowtieColors.hazard.fill);
    hazardRect.setAttribute('stroke', bowtieColors.hazard.stroke);
    hazardRect.setAttribute('stroke-width', '2');
    svg.appendChild(hazardRect);
    
    const hazardText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    hazardText.setAttribute('x', centerX);
    hazardText.setAttribute('y', centerY);
    hazardText.setAttribute('class', 'bowtie-text');
    hazardText.setAttribute('text-anchor', 'middle');
    hazardText.setAttribute('dominant-baseline', 'middle');
    hazardText.textContent = hazard.title || 'Hazard';
    svg.appendChild(hazardText);
    
    // Draw causes on the left - center around hazard
    const causes = hazard.causes || [];
    const leftX = 150;
    const maxCauseConsequence = Math.max(causes.length, hazard.consequences.length, 1);
    const totalSpacing = height - 200; // Leave 100px margin top and bottom
    const itemSpacing = totalSpacing / Math.max(maxCauseConsequence - 1, 1);
    
    // Calculate starting position to center causes around hazard
    const totalCauseHeight = (causes.length - 1) * itemSpacing;
    const causeStartY = centerY - totalCauseHeight / 2;
    
    causes.forEach((cause, i) => {
      // Position causes centered around the hazard
      const causeY = causeStartY + i * itemSpacing;
      
      // Cause rectangle
      const causeRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      causeRect.setAttribute('x', leftX - 50);
      causeRect.setAttribute('y', causeY - 20);
      causeRect.setAttribute('width', 100);
      causeRect.setAttribute('height', 40);
      causeRect.setAttribute('rx', 5);
      causeRect.setAttribute('fill', bowtieColors.cause.fill);
      causeRect.setAttribute('stroke', bowtieColors.cause.stroke);
      causeRect.setAttribute('stroke-width', '2');
      svg.appendChild(causeRect);
      
      // Cause text with proper centering
      const causeText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      causeText.setAttribute('x', leftX);
      causeText.setAttribute('y', causeY);
      causeText.setAttribute('class', 'bowtie-text');
      causeText.setAttribute('text-anchor', 'middle');
      causeText.setAttribute('dominant-baseline', 'middle');
      causeText.textContent = cause.text || `Cause ${i + 1}`;
      svg.appendChild(causeText);
      
      // Line from cause to hazard
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', leftX + 50);
      line.setAttribute('y1', causeY);
      line.setAttribute('x2', centerX - 60);
      line.setAttribute('y2', centerY);
      line.setAttribute('stroke', bowtieColors.line);
      line.setAttribute('stroke-width', '2');
      line.setAttribute('fill', 'none');
      svg.appendChild(line);
      
      // Prevention measures on the line
      const measures = cause.preventionMeasures || [];
      measures.forEach((measure, mIndex) => {
        const measureX = leftX + 50 + (mIndex + 1) * (centerX - 60 - leftX - 50) / (measures.length + 1);
        const measureY = causeY + (centerY - causeY) * (mIndex + 1) / (measures.length + 1);
        
        // Measure circle - make it larger to accommodate text
        const measureCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        measureCircle.setAttribute('cx', measureX);
        measureCircle.setAttribute('cy', measureY);
        measureCircle.setAttribute('r', 25);
        measureCircle.setAttribute('fill', bowtieColors.measure.fill);
        measureCircle.setAttribute('stroke', bowtieColors.measure.stroke);
        measureCircle.setAttribute('stroke-width', '1');
        svg.appendChild(measureCircle);
        
        // Measure text
        const measureText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        measureText.setAttribute('x', measureX);
        measureText.setAttribute('y', measureY);
        measureText.setAttribute('class', 'bowtie-measure-text');
        measureText.setAttribute('text-anchor', 'middle');
        measureText.setAttribute('dominant-baseline', 'middle');
        measureText.textContent = measure.text || 'Prevention measure';
        svg.appendChild(measureText);
        
        // Tooltip for measure
        const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
        title.textContent = measure.text || 'Prevention measure';
        measureCircle.appendChild(title);
        
        // Add mouse events for better tooltip support
        measureCircle.addEventListener('mouseenter', (e) => {
          const tooltip = document.createElement('div');
          tooltip.className = 'bowtie-tooltip';
          tooltip.textContent = measure.text || 'Prevention measure';
          tooltip.style.position = 'absolute';
          tooltip.style.background = '#333';
          tooltip.style.color = 'white';
          tooltip.style.padding = '5px 8px';
          tooltip.style.borderRadius = '4px';
          tooltip.style.fontSize = '12px';
          tooltip.style.zIndex = '10000';
          tooltip.style.pointerEvents = 'none';
          tooltip.style.maxWidth = '200px';
          tooltip.style.wordWrap = 'break-word';
          document.body.appendChild(tooltip);
          
          const updateTooltip = (e) => {
            tooltip.style.left = (e.clientX + 10) + 'px';
            tooltip.style.top = (e.clientY - 30) + 'px';
          };
          
          updateTooltip(e);
          measureCircle.addEventListener('mousemove', updateTooltip);
          measureCircle._tooltip = tooltip;
        });
        
        measureCircle.addEventListener('mouseleave', () => {
          if (measureCircle._tooltip) {
            document.body.removeChild(measureCircle._tooltip);
            measureCircle._tooltip = null;
          }
        });
      });
    });
    
    // Draw consequences on the right - center around hazard
    const consequences = hazard.consequences || [];
    const rightX = width - 150;
    
    // Calculate starting position to center consequences around hazard
    const totalConsequenceHeight = (consequences.length - 1) * itemSpacing;
    const consequenceStartY = centerY - totalConsequenceHeight / 2;
    
    consequences.forEach((consequence, i) => {
      // Position consequences centered around the hazard
      const consY = consequenceStartY + i * itemSpacing;
      
      // Consequence rectangle with risk-based coloring
      // Increase height if risk information is present
      const hasRiskInfo = consequence.risk && (consequence.risk.severityLevel || consequence.risk.likelihoodLevel);
      const rectHeight = hasRiskInfo ? 50 : 40;
      const rectY = hasRiskInfo ? consY - 25 : consY - 20;
      
      const consRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      consRect.setAttribute('x', rightX - 50);
      consRect.setAttribute('y', rectY);
      consRect.setAttribute('width', 100);
      consRect.setAttribute('height', rectHeight);
      consRect.setAttribute('rx', 5);
      
      // Apply risk color if available
      const riskColor = getRiskLevelColor(consequence.risk?.severityLevel, consequence.risk?.likelihoodLevel);
      if (riskColor) {
        consRect.setAttribute('fill', riskColor);
        consRect.setAttribute('stroke', '#333');
      } else {
        consRect.setAttribute('fill', bowtieColors.consequence.fill);
        consRect.setAttribute('stroke', bowtieColors.consequence.stroke);
      }
      consRect.setAttribute('stroke-width', '2');
      svg.appendChild(consRect);
      
      // Consequence text with proper centering
      const consText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      consText.setAttribute('x', rightX);
      consText.setAttribute('y', consY - 8); // Adjust vertical position for main text
      consText.setAttribute('class', 'bowtie-text');
      consText.setAttribute('text-anchor', 'middle');
      consText.setAttribute('dominant-baseline', 'middle');
      consText.textContent = consequence.text || `Consequence ${i + 1}`;
      svg.appendChild(consText);
      
      // Compact risk description below consequence
      if (consequence.risk && (consequence.risk.severityLevel || consequence.risk.likelihoodLevel)) {
        const riskDesc = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        riskDesc.setAttribute('x', rightX);
        riskDesc.setAttribute('y', consY + 15);
        riskDesc.setAttribute('class', 'bowtie-risk-text');
        riskDesc.setAttribute('text-anchor', 'middle');
        riskDesc.setAttribute('dominant-baseline', 'middle');
        
        // Format: Severity(Category)-Likelihood (e.g., "2(Asset)-B")
        if (consequence.risk.severityLevel && consequence.risk.likelihoodLevel) {
          const severityCategory = consequence.risk.severityCategory || 'Unknown';
          const categoryCapitalized = severityCategory.charAt(0).toUpperCase() + severityCategory.slice(1);
          const likelihoodDesc = state.riskMatrix.likelihood.find(l => l.id === consequence.risk.likelihoodLevel);
          riskDesc.textContent = `${consequence.risk.severityLevel}(${categoryCapitalized})-${likelihoodDesc?.label || consequence.risk.likelihoodLevel}`;
        } else if (consequence.risk.severityLevel) {
          const severityCategory = consequence.risk.severityCategory || 'Unknown';
          const categoryCapitalized = severityCategory.charAt(0).toUpperCase() + severityCategory.slice(1);
          riskDesc.textContent = `${consequence.risk.severityLevel}(${categoryCapitalized})`;
        } else if (consequence.risk.likelihoodLevel) {
          const likelihoodDesc = state.riskMatrix.likelihood.find(l => l.id === consequence.risk.likelihoodLevel);
          riskDesc.textContent = `${likelihoodDesc?.label || consequence.risk.likelihoodLevel}`;
        } else {
          riskDesc.textContent = 'Unrated';
        }
        
        svg.appendChild(riskDesc);
      }
      
      // Line from hazard to consequence
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', centerX + 60);
      line.setAttribute('y1', centerY);
      line.setAttribute('x2', rightX - 50);
      line.setAttribute('y2', consY);
      line.setAttribute('stroke', bowtieColors.line);
      line.setAttribute('stroke-width', '2');
      line.setAttribute('fill', 'none');
      svg.appendChild(line);
      
      // Mitigation measures on the line
      const measures = consequence.mitigationMeasures || [];
      measures.forEach((measure, mIndex) => {
        const measureX = centerX + 60 + (mIndex + 1) * (rightX - 50 - centerX - 60) / (measures.length + 1);
        const measureY = centerY + (consY - centerY) * (mIndex + 1) / (measures.length + 1);
        
        // Measure circle - make it larger to accommodate text
        const measureCircle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        measureCircle.setAttribute('cx', measureX);
        measureCircle.setAttribute('cy', measureY);
        measureCircle.setAttribute('r', 25);
        measureCircle.setAttribute('fill', bowtieColors.measure.fill);
        measureCircle.setAttribute('stroke', bowtieColors.measure.stroke);
        measureCircle.setAttribute('stroke-width', '1');
        svg.appendChild(measureCircle);
        
        // Measure text
        const measureText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        measureText.setAttribute('x', measureX);
        measureText.setAttribute('y', measureY);
        measureText.setAttribute('class', 'bowtie-measure-text');
        measureText.setAttribute('text-anchor', 'middle');
        measureText.setAttribute('dominant-baseline', 'middle');
        measureText.textContent = measure.text || 'Mitigation measure';
        svg.appendChild(measureText);
        
        // Tooltip for measure
        const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
        title.textContent = measure.text || 'Mitigation measure';
        measureCircle.appendChild(title);
        
        // Add mouse events for better tooltip support
        measureCircle.addEventListener('mouseenter', (e) => {
          const tooltip = document.createElement('div');
          tooltip.className = 'bowtie-tooltip';
          tooltip.textContent = measure.text || 'Mitigation measure';
          tooltip.style.position = 'absolute';
          tooltip.style.background = '#333';
          tooltip.style.color = 'white';
          tooltip.style.padding = '5px 8px';
          tooltip.style.borderRadius = '4px';
          tooltip.style.fontSize = '12px';
          tooltip.style.zIndex = '10000';
          tooltip.style.pointerEvents = 'none';
          tooltip.style.maxWidth = '200px';
          tooltip.style.wordWrap = 'break-word';
          document.body.appendChild(tooltip);
          
          const updateTooltip = (e) => {
            tooltip.style.left = (e.clientX + 10) + 'px';
            tooltip.style.top = (e.clientY - 30) + 'px';
          };
          
          updateTooltip(e);
          measureCircle.addEventListener('mousemove', updateTooltip);
          measureCircle._tooltip = tooltip;
        });
        
        measureCircle.addEventListener('mouseleave', () => {
          if (measureCircle._tooltip) {
            document.body.removeChild(measureCircle._tooltip);
            measureCircle._tooltip = null;
          }
        });
      });
    });
    
    // Show modal
    modal.style.display = 'block';
  }
  
  function closeBowtieModal() {
    const modal = document.getElementById('bowtie-modal');
    modal.style.display = 'none';
  }
  
  // Wire modal events
  function wireBowtieModal() {
    const modal = document.getElementById('bowtie-modal');
    const closeBtn = modal.querySelector('.close');
    const exportPngBtn = document.getElementById('export-bowtie-png');
    const exportSvgBtn = document.getElementById('export-bowtie-svg');

    closeBtn.onclick = closeBowtieModal;
    exportPngBtn.onclick = exportBowtieToPNG;
    exportSvgBtn.onclick = exportBowtieToSVG;

    window.onclick = (event) => {
      if (event.target === modal) {
        closeBowtieModal();
      }
    };
  }

  // Bow-tie colors management
  function loadBowtieColorsFromJSON(jsonData) {
    try {
      state.bowtieColors = jsonData;
      scheduleSave();
      alert('Bow-tie colors loaded successfully!');
    } catch (e) {
      alert('Failed to load bow-tie colors: ' + e.message);
    }
  }

  function exportBowtieColors() {
    const defaultColors = {
      hazard: { fill: '#ff6b6b', stroke: '#d63031' },
      cause: { fill: '#74b9ff', stroke: '#0984e3' },
      consequence: { fill: '#fd79a8', stroke: '#e84393' },
      measure: { fill: '#55a3ff', stroke: '#2d3436' },
      line: '#636e72',
      measureLine: '#55a3ff'
    };
    
    const colors = state.bowtieColors || defaultColors;
    const blob = new Blob([JSON.stringify(colors, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bowtie-colors.json';
    a.click();
    URL.revokeObjectURL(url);
  }

  function loadDefaultBowtieColors() {
    if (confirm('Load default bow-tie colors? This will replace your current color scheme.')) {
      state.bowtieColors = null; // Reset to defaults
      scheduleSave();
      alert('Default bow-tie colors loaded!');
    }
  }

  // Export bowtie diagram to PNG
  function exportBowtieToPNG() {
    const svg = document.getElementById('bowtie-svg');
    if (!svg) {
      alert('No bowtie diagram to export');
      return;
    }

    try {
      // Get the SVG as a string
      const svgData = new XMLSerializer().serializeToString(svg);
      
      // Create a canvas element
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      
      // Set canvas dimensions to match SVG viewBox
      const viewBox = svg.getAttribute('viewBox');
      if (viewBox) {
        const [, , width, height] = viewBox.split(' ').map(Number);
        canvas.width = width;
        canvas.height = height;
      } else {
        canvas.width = 1000;
        canvas.height = 600;
      }
      
      // Create an image from the SVG data
      const img = new Image();
      const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(svgBlob);
      
      img.onload = () => {
        // Fill with white background
        ctx.fillStyle = 'white';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // Draw the SVG image
        ctx.drawImage(img, 0, 0);
        
        // Convert canvas to PNG and download
        canvas.toBlob((blob) => {
          const downloadUrl = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = downloadUrl;
          a.download = `bowtie-diagram-${new Date().toISOString().slice(0, 10)}.png`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          
          // Clean up URLs
          URL.revokeObjectURL(url);
          URL.revokeObjectURL(downloadUrl);
        }, 'image/png');
      };
      
      img.onerror = () => {
        alert('Failed to export bowtie diagram. Please try again.');
        URL.revokeObjectURL(url);
      };
      
      img.src = url;
      
    } catch (error) {
      console.error('Export error:', error);
      alert('Failed to export bowtie diagram: ' + error.message);
    }
  }

  // Export bowtie diagram to SVG
  function exportBowtieToSVG() {
    const svg = document.getElementById('bowtie-svg');
    if (!svg) {
      alert('No bowtie diagram to export');
      return;
    }

    try {
      // Clone the SVG to avoid modifying the original
      const svgClone = svg.cloneNode(true);
      
      // Add namespace declaration if not present
      if (!svgClone.getAttribute('xmlns')) {
        svgClone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      }
      if (!svgClone.getAttribute('xmlns:xlink')) {
        svgClone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
      }
      
      // Get the SVG as a string
      const svgData = new XMLSerializer().serializeToString(svgClone);
      
      // Create and download the SVG file
      const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(svgBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `bowtie-diagram-${new Date().toISOString().slice(0, 10)}.svg`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
    } catch (error) {
      console.error('SVG export error:', error);
      alert('Failed to export bowtie diagram as SVG: ' + error.message);
    }
  }

  // Start
  document.addEventListener('DOMContentLoaded', () => {
    init();
    wireBowtieModal();
  });
})();


