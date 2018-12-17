function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Prévisions")
      .addItem('Valider et envoyer les données', 'collect_data')
      .addItem("Importer les données N-1", 'import_data')
      .addToUi();
}


function collect_data (){
  
  loadSpinner();
  var sheet_name = SpreadsheetApp.getActive().getActiveSheet().getName();
  
  if (sheet_name.indexOf("CEX commercial") >= 0) {
    
    try {
      collect_data_CEX_commercial();
      SpreadsheetApp.getUi().alert("Les prévisions ont été envoyées.");
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
    
  } else if (sheet_name.indexOf("Autres Eléments du CEX") >= 0){
       
    try {
      collect_data_autres_elements();
      SpreadsheetApp.getUi().alert("Les prévisions ont été envoyées.");
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
  }  
}


function import_data (){
  
  loadSpinner();
  var sheet_name = SpreadsheetApp.getActive().getActiveSheet().getName();
  
  if (sheet_name.indexOf("CEX commercial") >= 0) {
    
    try {
      import_n1_data();
      SpreadsheetApp.getUi().alert("Les données ont été importées.");
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
    
  } else if (sheet_name.indexOf("Autres Eléments du CEX") >= 0){
       
    try {
      import_n1_data_autres_elements();
      SpreadsheetApp.getUi().alert("Les données ont été importées");
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
  }  
}


function loadSpinner(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var htmlApp = HtmlService
  .createHtmlOutput('spinner')
     .setTitle("Chargement ...")
     .setWidth(300)
     .setHeight(1);

 SpreadsheetApp.getActiveSpreadsheet().show(htmlApp);
}

