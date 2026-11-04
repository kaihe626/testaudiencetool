/* Background reader: parses the file and runs the import gate off the main
   page so the dashboard stays responsive. Runs in this browser only. */
'use strict';
importScripts('app.js');
onmessage = e => {
  const { buf, name, profile } = e.data;
  try {
    state.profile = profile;
    postMessage({ stage: 'Reading the file…', w: '35%' });
    let rows;
    if (/\.xlsx?$/i.test(name)) {
      importScripts('vendor/xlsx.full.min.js');
      const wb = XLSX.read(buf, { type: 'array', cellDates: false, cellFormula: false, cellHTML: false });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '', blankrows: false });
    } else {
      let text = new TextDecoder('utf-8').decode(buf);
      if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
      rows = parseCSV(text);
    }
    postMessage({ stage: 'Running the import gate…', w: '75%' });
    const data = runGate(rows, name);
    rows = null;
    postMessage({ data });
  } catch (err) {
    postMessage({ error: (err && err.message) || String(err) });
  }
};
