const RETIRED_API = Object.freeze({
  status: 'error',
  code: 'api_moved',
  message: 'このAPIはSupabase Edge Functionへ移行しました。最新版のKEIKOを開いてください。',
});

function doGet() {
  return jsonOutput_(RETIRED_API);
}

function doPost() {
  return jsonOutput_(RETIRED_API);
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
