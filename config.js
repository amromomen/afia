// App settings. Paste your Application (client) ID from Microsoft Entra between the quotes.
window.APP_CONFIG = {
  clientId: 'PASTE-YOUR-CLIENT-ID-HERE',
  folder: '/Finance',                 // OneDrive folder that holds the two files below
  workbook: 'Finance.xlsx',   // the Excel database
  config: 'finance.json',          // accounts, cards, plan and categories
  app: 'app.html',                 // the app itself (kept private in OneDrive)
  authority: 'consumers'           // personal Microsoft account
};
