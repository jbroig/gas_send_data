/*
GLOBAL VARIABLES
*/
//
var WEB_APP_URL = "https://webhook.site/...";

/*
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
  
  if (sheet_name.indexOf("CEX commercial") >= 0) {
    
    try {
      collect_data_CEX_commercial();
      
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }
    
  } else if (sheet_name.indexOf("Autres Eléments du CEX") >= 0){
       
    try {
      collect_data_autres_elements();
      
    } catch (e){
      SpreadsheetApp.getUi().alert("Error: " + e);
    }

  }  
}

//Works for CEX commercial - Rxx sheets
function collect_data_CEX_commercial (){
  var ss = SpreadsheetApp.getActive();
  var data = [];
      
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
     
  var categories = [];
  
  var ray = {};
  ray.year = 2018;

  var start = 49;
  
  var forecast = [];
  var month = 0;
  
  //Range 49 - 60
  for (var i = start; i < (start + 12); i++){
    
    month += 1;
    var turnover = sheet.getRange("G"+[i]+":G"+[i]).getValue();
    if (turnover == "" || turnover == undefined ){turnover = null;}
    var turnoverEvolution = sheet.getRange("H"+[i]+":H"+[i]).getValue();
    if (turnoverEvolution == "" || turnoverEvolution == undefined ){turnoverEvolution = null;}
    
    var forecast_obj = {}
    forecast_obj.month = month;
    forecast_obj.turnover = turnover;
    forecast_obj.turnoverEvolution = turnoverEvolution;
    forecast_obj.profitability = null;
    forecast_obj.profitabilityRate = null;
        
    forecast.push(forecast_obj);
    
    var obj = {};
    obj.id = "test ID";
    obj.forecasts = forecast;
      
  }
  categories.push(obj);
  
  var start = 67;
  
  var forecast = [];
  var month = 0;
  
  //Range 67 - 78
  for (var i = start; i < (start + 12); i++){
    
    month += 1;
    var profitability = sheet.getRange("G"+[i]+":G"+[i]).getValue();
    if (profitability == "" || profitability == undefined ){profitability = null;}
    var profitabilityRate = sheet.getRange("H"+[i]+":H"+[i]).getValue();
    if (profitabilityRate == "" || profitabilityRate == undefined ){profitabilityRate = null;}
    
    var forecast_obj = {}
    forecast_obj.month = month;
    forecast_obj.turnover = null;
    forecast_obj.turnoverEvolution = null;
    forecast_obj.profitability = profitability;
    forecast_obj.profitabilityRate = profitabilityRate;
        
    forecast.push(forecast_obj);
    
    var obj = {};
    obj.id = "test ID";
    obj.forecasts = forecast;
       
  }
  categories.push(obj);
  
  ray.categories = categories; 
  data = JSON.stringify(ray);
    
  var payload = {
    "magasin": sheet.getRange("B2").getValue(),
    "departement": sheet.getRange("B3").getValue(),
    "data": data
  };
  var options = {
    "method": "POST",
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    "payload": payload
  };
  
  var response = UrlFetchApp.fetch(WEB_APP_URL, options);
  
  if (response.getResponseCode() == 200){
    SpreadsheetApp.getUi().alert("The data has been validated.");
  }
  
}

//Works for Autres Eléments du CEX sheets
function collect_data_autres_elements (){
  
  var ss = SpreadsheetApp.getActive();
  var data = [];
      
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();

  var indexes = 6;
  var categories = [];      
      
  //RANGE AR7:BC19
  var categName = sheet.getRange("AR5").getValue();                  
  var categValues = sheet.getRange("AR" +(indexes+1)+ ":BC" +(indexes+13)).getValues();
  var obj = {};

  obj[categName] = categValues;
      
  //RANGE CE7:CQ19
  // Where can I found the category name? 
  var categ2Name = "Category2";                 
  var categ2Values = sheet.getRange("CE" +(indexes+1)+ ":CQ" +(indexes+13)).getValues();

  obj[categ2Name] = categ2Values;      
      
  categories.push(obj);
  var ray = {};
  ray[rayon] = categories;
    
  data.push(ray);   
    
  var payload = {
    //Where can i found the magasin value?
    "magasin": sheet.getRange("B2").getValue(),
    //Where can i found the magasin value?
    "departement": sheet.getRange("B3").getValue(),
    "data": JSON.stringify(data)
  };
  var options = {
    "method": "POST",
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    "payload": payload
  };
  
  var response = UrlFetchApp.fetch(WEB_APP_URL, options);
   
}







