/*
GLOBAL VARIABLES
*/
//Weebhook endpoint
var WEB_APP_URL = "";

/*
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
1- IMPORTING DATA FROM THE APPENGINE: The JSON should have the same number of categories in the sheet and in the JSON. 

------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Validation")
      .addItem('Valider et envoyer les données', 'collect_data')
      .addItem("Import N1 data", 'import_n1_data')
      .addToUi();
}

function collect_data (){
  
  var sheet_name = SpreadsheetApp.getActive().getActiveSheet().getName();
  //Logger.log("Sheet name: " + sheet_name);
  
  if (sheet_name.indexOf("CEX commercial") >= 0) {
    
    try {
      collect_data_CEX_commercial();
      
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
  }  
}







