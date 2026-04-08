// Worker thread for parsing contingency MDB files off the main thread.
// Keeps the main process event loop free so the UI never freezes during load.
const { parentPort } = require('worker_threads');
const fs = require('fs');

const TABLES = [
  'OrderInfo', 'ConstructionData', 'InspectionData', 'TestingData',
  'CleaningData', 'PaintingData', 'NamePlateData', 'PackagingData',
  'PurchasingData', 'QCdata', 'RevNotes', 'Attachments',
];

parentPort.on('message', async (msg) => {
  try {
    const MDBReader = (await import('mdb-reader')).default;
    const orderMap = new Map();

    for (const dbPath of msg.dbPaths) {
      try {
        const buf = fs.readFileSync(dbPath);
        const db = new MDBReader(Buffer.from(buf));
        for (const tn of TABLES) {
          try {
            const rows = db.getTable(tn).getData();
            for (const row of rows) {
              const ol = row.OrderLine;
              if (!ol) continue;
              if (!orderMap.has(ol)) orderMap.set(ol, {});
              const entry = orderMap.get(ol);
              if (tn === 'RevNotes') {
                if (!entry.RevNotes) entry.RevNotes = [];
                entry.RevNotes.push(row);
              } else {
                entry[tn] = row;
              }
            }
          } catch {}
        }
      } catch {}
    }

    // Convert Map to array of entries for structured clone transfer
    parentPort.postMessage({ entries: [...orderMap.entries()] });
  } catch (e) {
    parentPort.postMessage({ error: e.message });
  }
});
