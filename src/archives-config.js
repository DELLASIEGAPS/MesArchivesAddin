window.MES_ARCHIVES_CONFIG = Object.freeze({
  pstUrls: Object.freeze({
    "2024": "https://dscoie091-my.sharepoint.com/personal/primosylvestredellasiega-nkoume_dscoie091_onmicrosoft_com/Documents/Personnel/ADR.pst"
  }),
  uploadFolderUrl: "https://www.office.com/launch/onedrive",
  // Configure an authorized HTTPS service: GET ?year=YYYY returns fileExists/mounted;
  // POST accepts annee, Monte, Couleur, and sharePointPath. Never put Azure keys here.
  webStatusEndpoint: "https://defaultdb63b3bbeebc464bbc108dbb46365f.c3.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/08/workflows/e727039e40ab4c50b115ff20b075a9b4/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=abTai-tFuAykHxs09QxnLl8QMGzTivvZFF_FDa454nQ"
});
