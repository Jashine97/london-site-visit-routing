// ===== london-site-visit-routing / Code.gs =====

// OFFICE ADDRESS
const OFFICE_ADDRESS = '5 Clarendon Road, London N22 6XJ';

// SHEET NAMES
const TRACKER_SHEET_NAME = 'Site Visit Tracker';
const TO_VISIT_SHEET_NAME = 'To Visit';

// Build custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Site visits')
    .addItem('Refresh "To Visit" data', 'refreshToVisit')
    .addToUi();
}

// Refresh "To Visit" sheet
function refreshToVisit() {
  const ss = SpreadsheetApp.getActive();
  const tracker = ss.getSheetByName(TRACKER_SHEET_NAME);
  const toVisit = ss.getSheetByName(TO_VISIT_SHEET_NAME);

  const data = tracker.getDataRange().getValues();
  const header = data[0];

  const idxSite = header.indexOf('Site');
  const idxStatus = header.indexOf('Status');
  const idxDeal = header.indexOf('Deal no');

  const filtered = data
    .slice(1)
    .filter(r => r[idxStatus] && r[idxStatus].toString().toLowerCase() === 'visit')
    .map(r => ({
      site: r[idxSite],
      deal: r[idxDeal]
    }));

  const out = [['Cluster','Site','Deal no','Distance','Duration','Duration (mins)','Directions link']];
  filtered.forEach(row => {
    const cluster = extractClusterFromAddress(row.site);
    const api = Maps.newDirectionFinder()
      .setOrigin(OFFICE_ADDRESS)
      .setDestination(row.site)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    const leg = api.routes[0].legs[0];
    const distance = leg.distance.text;
    const duration = leg.duration.text;
    const durationMin = Math.round(leg.duration.value / 60);

    const link =
      'https://www.google.com/maps/dir/?api=1' +
      '&origin=' + encodeURIComponent(OFFICE_ADDRESS) +
      '&destination=' + encodeURIComponent(row.site) +
      '&travelmode=driving';

    out.push([cluster, row.site, row.deal, distance, duration, durationMin, link]);
  });

  toVisit.clear();
  toVisit.getRange(1,1,out.length,out[0].length).setValues(out);
}

// Extract outward UK postcode
function extractClusterFromAddress(address) {
  const match = address.toUpperCase().match(/\b[A-Z]{1,2}\d[A-Z0-9]?\b/g);
  if (match) return match[match.length-1];
  return '';
}

// Provide data to web app
function getVisitData() {
  const ss = SpreadsheetApp.getActive();
  const vs = ss.getSheetByName(TO_VISIT_SHEET_NAME);
  const data = vs.getDataRange().getValues();
  const header = data[0];

  const idxCluster = header.indexOf('Cluster');
  const idxSite = header.indexOf('Site');
  const idxDeal = header.indexOf('Deal no');
  const idxDist = header.indexOf('Distance');
  const idxDur = header.indexOf('Duration');
  const idxDurMin = header.indexOf('Duration (mins)');
  const idxLink = header.indexOf('Directions link');

  const rows = data.slice(1).map(r => ({
    cluster: r[idxCluster],
    site: r[idxSite],
    deal: r[idxDeal],
    distance: r[idxDist],
    duration: r[idxDur],
    minutes: r[idxDurMin],
    link: r[idxLink]
  }));

  return {
    office: OFFICE_ADDRESS,
    rows: rows
  };
}

// Build multi-site Google Maps URL
function buildMultiRouteUrl(office, rows) {
  if (!rows.length) return '';

  const dest = rows[rows.length-1].site;
  const waypoints = rows.slice(0,-1).map(r => r.site);

  let url = 'https://www.google.com/maps/dir/?api=1' +
    '&origin=' + encodeURIComponent(office) +
    '&destination=' + encodeURIComponent(dest) +
    '&travelmode=driving';

  if (waypoints.length) {
    url += '&waypoints=' + encodeURIComponent(waypoints.join('|'));
  }
  return url;
}

// Serve the HTML web app
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Site Visit Planner')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
