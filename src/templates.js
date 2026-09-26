import { emptyDoc, normalizeDoc, defaultFields } from './model.js';
import { parseDelimited, guessRoles, buildFromRows } from './spreadsheet.js';

// A fictional audit walkthrough, built from spreadsheet-style rows (the same path as File → Import spreadsheet).
const WALKTHROUGH_ROWS = `Step,Activity,Owner,Phase,Type,Next,What happens,System / data,Output,Control,Control #,Risk point to probe
START,Customer places order,Customer,Order,start,OTC-01,,,,,,
OTC-01,Receive purchase order,Sales,Order,manual input,OTC-02,Orders arrive by email or through the web shop. Emailed orders are keyed into the order system by the sales desk.,Web shop; order system,Sales order,,,Orders keyed in by hand could be entered with the wrong quantity or price.
OTC-02,Within credit limit?,Finance,Order,decision,Yes: OTC-04; No: OTC-03,Finance compares the order value and overdue balance with the customer's credit limit. Orders over the limit are held.,Order system credit screen,Released or held order,Orders over the credit limit are blocked automatically until Finance approves the release.,C1,Orders could be released to customers who are over their credit limit.
OTC-03,Hold order and contact customer,Sales,Order,process,Cleared: OTC-02,"Sales contacts the customer to collect the overdue balance or agree a prepayment, then asks Finance to re-check.",Email; CRM notes,Cleared or cancelled order,,,
OTC-04,Pick and pack,Warehouse,Fulfil,process,OTC-05,The warehouse picks the released order and checks the items against the packing list.,Warehouse system,Packed order,,,
OTC-05,Ship and record delivery,Warehouse,Fulfil,data,OTC-06; ~delivery data: OTC-07,"The carrier collects the parcel; the shipment is confirmed in the warehouse system, which passes delivery data to Finance.",Warehouse system → accounting system,Proof of delivery,,,Goods could be shipped without being recorded.
OTC-06,Customer receives goods,Customer,Fulfil,process,,The customer signs for the delivery.,Carrier portal,Signed delivery note,,,
OTC-07,Raise invoice,Finance,Bill and collect,document,OTC-08,Invoices are generated from confirmed shipments at the agreed price.,Accounting system,Customer invoice,Weekly report of shipped but not invoiced orders is reviewed by the billing lead.,C2,Shipped orders might not be invoiced (completeness).
OTC-08,Send invoice,Finance,Bill and collect,process,OTC-09,Invoices are emailed to the customer's accounts payable address.,Accounting system; email,Sent invoice,,,
OTC-09,Customer pays,Customer,Bill and collect,process,OTC-10,The customer pays by bank transfer quoting the invoice number.,Bank,Payment,,,
OTC-10,Apply cash to invoices,Finance,Bill and collect,database,OTC-11,Receipts from the daily bank file are matched to open invoices.,Bank file; accounting system,Updated receivables,Unapplied cash is reviewed and cleared daily; exceptions go to the Finance manager.,C3,Cash could be applied to the wrong customer or invoice.
OTC-11,Paid in full?,Finance,Bill and collect,decision,Yes: OTC-13; No: OTC-12,Short payments and overdue invoices appear on the aged receivables report.,Aged receivables report,Closed or open balance,,,
OTC-12,Chase overdue balance,Sales,Bill and collect,process,Reminder: OTC-09,Sales follows up overdue balances with the customer.,Email; CRM notes,Payment promise,,,
OTC-13,Reconcile receivables to ledger,Finance,Bill and collect,process,END,Monthly reconciliation of the receivables subledger to the general ledger. Who reviews it has not been confirmed yet.,Accounting system; spreadsheet,Reconciliation (to be confirmed),,,The receivables subledger might not agree to the ledger.
END,Order closed,Finance,Bill and collect,end,,,,,,,`;

function walkthrough() {
  const rows = parseDelimited(WALKTHROUGH_ROWS);
  const { doc } = buildFromRows(rows, guessRoles(rows[0], rows.slice(1)), { title: 'Order-to-cash walkthrough (example)' });
  // stable ids: use the step references
  const ids = new Map(doc.nodes.map((n) => [n.id, n.tag || n.text.toUpperCase().replace(/\W+/g, '-')]));
  doc.nodes.forEach((n) => { n.id = ids.get(n.id); });
  doc.edges.forEach((e) => { e.from.node = ids.get(e.from.node); e.to.node = ids.get(e.to.node); });
  const tbc = doc.nodes.find((n) => n.id === 'OTC-13');
  tbc.style = { dash: '5 4' };
  doc.subtitle = 'Fictional example · Steps OTC-01 to OTC-13 · shows lanes, phases, controls, risk points and step details';
  doc.notes = 'Everything here is invented to show what the editor can do.\nOTC-13 has a dashed outline because its review has not been confirmed yet.';
  doc.settings.detailHint = 'Select any shape to see what happens at that step.';
  doc.settings.initialSelection = 'OTC-02';
  doc.settings.noTagLabel = 'Connector';
  doc.shapeLabels = { document: 'Report', data: 'Data transfer', terminator: 'Start / end' };
  doc.dashedNodeLabel = 'TBC';
  return normalizeDoc(doc);
}

function basicFlow() {
  const x = 200;
  return normalizeDoc({
    title: 'Basic flowchart',
    subtitle: 'Double-click any shape to edit its text. Drag from the blue dots to connect.',
    settings: { showTags: true },
    fields: defaultFields(),
    nodes: [
      { id: 'start', shape: 'terminator', x: x + 14, y: 40, w: 156, h: 40, text: 'Start' },
      { id: 'a', shape: 'manualInput', x, y: 120, w: 184, h: 60, text: 'Receive request', tag: 'S-01', fields: { what: 'A request arrives by email or the web form.' } },
      { id: 'b', shape: 'process', x, y: 220, w: 184, h: 60, text: 'Review request', tag: 'S-02', fields: { what: 'Check the request for completeness.' } },
      { id: 'c', shape: 'decision', x, y: 320, w: 184, h: 80, text: 'Complete?', tag: 'S-03' },
      { id: 'd', shape: 'process', x, y: 440, w: 184, h: 60, text: 'Fulfil request', tag: 'S-04', badges: ['risk'], fields: { risk: 'Requests could be fulfilled without approval.' } },
      { id: 'f', shape: 'document', x: x + 260, y: 330, w: 184, h: 60, text: 'Ask for missing info', tag: 'S-05' },
      { id: 'g', shape: 'database', x: x + 260, y: 438, w: 184, h: 64, text: 'Request log' },
      { id: 'end', shape: 'terminator', x: x + 14, y: 540, w: 156, h: 40, text: 'Done' },
    ],
    edges: [
      { from: { node: 'start', side: 'auto' }, to: { node: 'a', side: 'auto' } },
      { from: { node: 'a', side: 'auto' }, to: { node: 'b', side: 'auto' } },
      { from: { node: 'b', side: 'auto' }, to: { node: 'c', side: 'auto' } },
      { from: { node: 'c', side: 'b' }, to: { node: 'd', side: 't' }, label: 'Yes' },
      { from: { node: 'c', side: 'r' }, to: { node: 'f', side: 'l' }, label: 'No' },
      { from: { node: 'f', side: 't' }, to: { node: 'b', side: 'r' }, label: 'resubmitted' },
      { from: { node: 'd', side: 'r' }, to: { node: 'g', side: 'l' }, kind: 'feed', label: 'logged' },
      { from: { node: 'd', side: 'auto' }, to: { node: 'end', side: 'auto' } },
    ],
  });
}

function swimlane(horizontal = false) {
  const lanes = [
    { id: 'l1', title: 'Customer', size: horizontal ? 150 : 240 },
    { id: 'l2', title: 'Sales', size: horizontal ? 150 : 240 },
    { id: 'l3', title: 'Operations', size: horizontal ? 150 : 240 },
  ];
  const P = (lane, step) => (horizontal
    ? { x: 80 + step * 230, y: lane * 150 + 45 }
    : { x: lane * 240 + 28, y: 70 + step * 110 });
  const n = (id, shape, lane, step, text, extra = {}) => {
    const h = shape === 'decision' ? 70 : 60;
    const p = P(lane, step);
    if (horizontal) p.y = lane * 150 + 75 - h / 2;
    return { id, shape, ...p, w: 184, h, text, ...extra };
  };
  return normalizeDoc({
    title: horizontal ? 'Order fulfilment (horizontal lanes)' : 'Order fulfilment',
    subtitle: 'Swimlane template — rename lanes in the right-hand panel',
    lanes: { orientation: horizontal ? 'horizontal' : 'vertical', header: 40, items: lanes },
    phases: horizontal ? { items: [] } : { header: 34, items: [{ id: 'p1', title: 'Order', size: 330 }, { id: 'p2', title: 'Fulfil', size: 330 }] },
    nodes: [
      n('o1', 'manualInput', 0, 0, 'Place order', { tag: 'OF-01' }),
      n('o2', 'process', 1, 1, 'Confirm pricing', { tag: 'OF-02' }),
      n('o3', 'decision', 1, 2, 'Credit OK?', { tag: 'OF-03' }),
      n('o4', 'process', 2, 4, 'Pick and ship', { tag: 'OF-05' }),
      n('o5', 'document', 2, 5, 'Invoice', { tag: 'OF-06' }),
      n('o6', 'process', 0, 3, 'Pay upfront', { tag: 'OF-04' }),
    ],
    edges: [
      { from: { node: 'o1', side: 'auto' }, to: { node: 'o2', side: 'auto' } },
      { from: { node: 'o2', side: 'auto' }, to: { node: 'o3', side: 'auto' } },
      { from: { node: 'o3', side: 'auto' }, to: { node: 'o4', side: 'auto' }, label: 'Yes' },
      { from: { node: 'o3', side: 'auto' }, to: { node: 'o6', side: 'auto' }, label: 'No' },
      { from: { node: 'o6', side: 'auto' }, to: { node: 'o4', side: 'auto' } },
      { from: { node: 'o4', side: 'auto' }, to: { node: 'o5', side: 'auto' } },
    ],
  });
}

export const TEMPLATES = [
  { id: 'blank', name: 'Blank canvas', desc: 'Start from nothing.', make: () => normalizeDoc(emptyDoc()) },
  { id: 'basic', name: 'Basic flowchart', desc: 'Start, steps, a decision and a loop.', make: basicFlow },
  { id: 'swim', name: 'Swimlanes (columns)', desc: 'Three vertical lanes for hand-offs between teams.', make: () => swimlane(false) },
  { id: 'swimh', name: 'Swimlanes (rows)', desc: 'Three horizontal lanes, flow left to right.', make: () => swimlane(true) },
  { id: 'walkthrough', name: 'Audit walkthrough (example)', desc: 'Order-to-cash example with lanes, phases, controls, risk points and step details.', make: walkthrough },
];
