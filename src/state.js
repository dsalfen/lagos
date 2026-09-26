// Shared application state, history and change notification.
import { normalizeDoc, clone } from './model.js';

const STORAGE_KEY = 'lagos.flowchart.v1';
const MAX_HISTORY = 200;

export const app = {
  doc: null,
  sel: { nodes: new Set(), edges: new Set(), lane: null, phase: null },
  view: { x: 40, y: 40, k: 1 },
  mode: 'edit',
  undoStack: [],
  redoStack: [],
  listeners: new Set(),
  lastSaved: null,
  connectDefaults: { kind: null, routing: 'orthogonal' },
  lastShape: 'process',
};

export function on(fn) { app.listeners.add(fn); return () => app.listeners.delete(fn); }

// what: {doc, sel, view, inspector}
export function emit(what = {}) {
  for (const fn of app.listeners) fn(what);
}

export function snapshot() { return JSON.stringify(app.doc); }

export function checkpoint() {
  app.undoStack.push(snapshot());
  if (app.undoStack.length > MAX_HISTORY) app.undoStack.shift();
  app.redoStack.length = 0;
}

// Run a mutation as one undoable step.
export function change(fn, { inspector = true } = {}) {
  checkpoint();
  fn(app.doc);
  pruneSelection();
  emit({ doc: true, inspector });
}

// Record a step whose mutation already happened (e.g. after a drag) given the prior snapshot.
export function commitFrom(before, { inspector = true } = {}) {
  if (before === snapshot()) { emit({ doc: true, inspector }); return false; }
  app.undoStack.push(before);
  if (app.undoStack.length > MAX_HISTORY) app.undoStack.shift();
  app.redoStack.length = 0;
  emit({ doc: true, inspector });
  return true;
}

export function undo() {
  if (!app.undoStack.length) return;
  app.redoStack.push(snapshot());
  app.doc = normalizeDoc(JSON.parse(app.undoStack.pop()));
  pruneSelection();
  emit({ doc: true, inspector: true, history: true });
}

export function redo() {
  if (!app.redoStack.length) return;
  app.undoStack.push(snapshot());
  app.doc = normalizeDoc(JSON.parse(app.redoStack.pop()));
  pruneSelection();
  emit({ doc: true, inspector: true, history: true });
}

export function setDoc(doc, { keepHistory = false } = {}) {
  if (keepHistory && app.doc) checkpoint();
  else { app.undoStack.length = 0; app.redoStack.length = 0; }
  app.doc = normalizeDoc(doc);
  clearSelection(false);
  emit({ doc: true, inspector: true, fit: true, history: true });
}

export function pruneSelection() {
  const nodeIds = new Set(app.doc.nodes.map((n) => n.id));
  const edgeIds = new Set(app.doc.edges.map((e) => e.id));
  for (const id of [...app.sel.nodes]) if (!nodeIds.has(id)) app.sel.nodes.delete(id);
  for (const id of [...app.sel.edges]) if (!edgeIds.has(id)) app.sel.edges.delete(id);
  if (app.sel.lane && !app.doc.lanes.items.some((l) => l.id === app.sel.lane)) app.sel.lane = null;
  if (app.sel.phase && !app.doc.phases.items.some((l) => l.id === app.sel.phase)) app.sel.phase = null;
}

export function clearSelection(notify = true) {
  app.sel.nodes.clear(); app.sel.edges.clear(); app.sel.lane = null; app.sel.phase = null;
  if (notify) emit({ sel: true, inspector: true });
}

export function select({ nodes = [], edges = [], lane = null, phase = null } = {}, additive = false) {
  if (!additive) { app.sel.nodes.clear(); app.sel.edges.clear(); app.sel.lane = null; app.sel.phase = null; }
  if (phase) app.sel.phase = phase;
  nodes.forEach((id) => app.sel.nodes.add(id));
  edges.forEach((id) => app.sel.edges.add(id));
  if (lane) { app.sel.lane = lane; }
  emit({ sel: true, inspector: true });
}

export function toggleSelect(kind, id) {
  const set = kind === 'node' ? app.sel.nodes : app.sel.edges;
  if (set.has(id)) set.delete(id); else set.add(id);
  app.sel.lane = null; app.sel.phase = null;
  emit({ sel: true, inspector: true });
}

export const nodeById = (id) => app.doc.nodes.find((n) => n.id === id);
export const edgeById = (id) => app.doc.edges.find((e) => e.id === id);
export const selectedNodes = () => app.doc.nodes.filter((n) => app.sel.nodes.has(n.id));
export const selectedEdges = () => app.doc.edges.filter((e) => app.sel.edges.has(e.id));

// ---- persistence ----
let saveTimer = null;
export function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, 300);
}
export function saveNow() {
  try {
    localStorage.setItem(STORAGE_KEY, snapshot());
    app.lastSaved = new Date();
  } catch (_) { /* storage may be unavailable */ }
  emit({ saved: true });
}
export function loadSaved() {
  try {
    const s = localStorage.getItem(STORAGE_KEY);
    if (s) return normalizeDoc(JSON.parse(s));
  } catch (_) { /* ignore */ }
  return null;
}

export { clone };
