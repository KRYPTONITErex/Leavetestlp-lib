

// function checkSlackConfig() {
//     const props = PropertiesService.getScriptProperties().getProperties();
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Logs");
    
//     sheet.appendRow(["Timestamp", "Config Check"]);
    
//     for (const key in props) {
//       if (key.includes('SLACK') || key.includes('URL')) {
//         // Only log the first few characters of tokens for security
//         const value = props[key].startsWith('xoxb-') ? 
//                       props[key].substring(0, 10) + '...' : 
//                       props[key];
        
//         sheet.appendRow([new Date(), `${key}: ${value}`]);
//       }
//     }
//   }


// function fixScriptProperties() {
//     const scriptProperties = PropertiesService.getScriptProperties();
//     const currentProps = scriptProperties.getProperties();
    
//     // Log current values
//     logToSheet("Current script properties:");
//     for (const [key, value] of Object.entries(currentProps)) {
//       if (key.includes('URL') || key.includes('TOKEN')) {
//         // Mask tokens for security
//         const maskedValue = value.startsWith('xoxb-') ? value.substring(0, 10) + '...' : value;
//         logToSheet(`${key}: ${maskedValue}`);
//       }
//     }
    
//     // Fix LOOK_UP_BY_EMAIL_URL if it has quotes
//     if (currentProps.LOOK_UP_BY_EMAIL_URL && 
//         currentProps.LOOK_UP_BY_EMAIL_URL.includes('"')) {
//       const fixedValue = currentProps.LOOK_UP_BY_EMAIL_URL.replace(/"/g, '');
//       scriptProperties.setProperty('LOOK_UP_BY_EMAIL_URL', fixedValue);
//       logToSheet(`Fixed LOOK_UP_BY_EMAIL_URL: ${fixedValue}`);
//     }
    
//     // Fix other URL properties if needed
//     if (currentProps.USER_LIST_URL && currentProps.USER_LIST_URL.includes('"')) {
//       const fixedValue = currentProps.USER_LIST_URL.replace(/"/g, '');
//       scriptProperties.setProperty('USER_LIST_URL', fixedValue);
//       logToSheet(`Fixed USER_LIST_URL: ${fixedValue}`);
//     }
    
//     if (currentProps.POST_MSG_URL && currentProps.POST_MSG_URL.includes('"')) {
//       const fixedValue = currentProps.POST_MSG_URL.replace(/"/g, '');
//       scriptProperties.setProperty('POST_MSG_URL', fixedValue);
//       logToSheet(`Fixed POST_MSG_URL: ${fixedValue}`);
//     }
    
//     // Verify SLACK_BOT_TOKEN format
//     if (currentProps.SLACK_BOT_TOKEN && !currentProps.SLACK_BOT_TOKEN.startsWith('xoxb-')) {
//       logToSheet(`Warning: SLACK_BOT_TOKEN may have an incorrect format. Tokens usually start with 'xoxb-'`);
//     }
    
//     logToSheet("Script properties check/fix completed");
//   }
  