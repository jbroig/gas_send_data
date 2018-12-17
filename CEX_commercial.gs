function getURL(storeID, sectionID){

  var WEB_APP_URL = "";
  Logger.log(WEB_APP_URL);
  
  return WEB_APP_URL;
}

//DONE
//Works for CEX commercial - Rxx sheets
function collect_data_CEX_commercial (){
  var ss = SpreadsheetApp.getActive();
      
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
     
  var storeID = sheet.getRange("C2").getValue();
  var sectionID = sheet.getRange("B4").getValue();
  var WEB_APP_URL = getURL(storeID, sectionID);
  
  var categories = [];
  
  var ray = {};
  ray.year = 2019; 

  var indexes = getAllCategoriesIndexes(sheet);
  
  var forecast;
  var month;
  var offset = 18;
  
  var turnover, turnoverEvolution, profitability, profitabilityRate, estimatedGrossTurnover, estimatedTurnover, estimatedProfitability, turnoverProjection, profitabilityProjection;
  //Recorre indices
  for (var j in indexes){
    month = 0;
    forecast = [];
    var index = indexes[j];     
    
    var id = sheet.getRange("C"+(index-5)+":C"+(index-5)).getValue();
        
    // Recorre por meses
    for (var i = index; i < (index + 12); i++){
    
      month += 1;
          
      //TurnoverProjection, turnover, turoverEvolution, EstimatedTurnover
      var turnoverValues = sheet.getRange(i, 5, 1, 13).getValues();
      
      //profitability, profitabilityRate, estimatedProfitability, profitabilityProjection
      var profitabilityValues = sheet.getRange((i+offset), 5, 1, 4).getValues();      
      
      //If the value is not a number, then it will be converted to null
      turnover = turnoverValues[0][2];
      if (typeof turnover != "number"){turnover = null}
      turnoverEvolution = turnoverValues[0][3];
      if (typeof turnoverEvolution != "number"){turnoverEvolution = null}
      profitability = profitabilityValues[0][2];
      if (typeof profitability != "number"){profitability = null}
      profitabilityRate = profitabilityValues[0][3];
      if (typeof profitabilityRate != "number"){profitabilityRate = null}
      estimatedGrossTurnover = turnoverValues[0][12];
      if (typeof estimatedGrossTurnover != "number"){estimatedGrossTurnover = null}
      estimatedTurnover = turnoverValues[0][4];
      if (typeof estimatedTurnover != "number"){estimatedTurnover = null}
      estimatedProfitability = profitabilityValues[0][4];
      if (typeof estimatedProfitability != "number"){estimatedProfitability = null}
      turnoverProjection = turnoverValues[0][0];
      if (typeof turnoverProjection != "number"){turnoverProjection = null}
      profitabilityProjection = profitabilityValues[0][0];
      if (typeof profitabilityProjection != "number"){profitabilityProjection = null}
      
      var forecast_obj = {}
      
      forecast_obj.month = month;
      forecast_obj.turnover = turnover;
      forecast_obj.turnoverEvolution = turnoverEvolution;
      forecast_obj.profitability = profitability;
      forecast_obj.profitabilityRate = profitabilityRate;
      forecast_obj.estimatedGrossTurnover = estimatedGrossTurnover;
      forecast_obj.estimatedTurnover = estimatedTurnover;
      forecast_obj.estimatedProfitability = estimatedProfitability;
      forecast_obj.turnoverProjection = turnoverProjection;
      forecast_obj.profitabilityProjection = profitabilityProjection;
      
      forecast.push(forecast_obj);
      
      //categories_obj
      var obj = {};
      obj.id = id;
      obj.forecasts = forecast;
      
    }
    categories.push(obj);   
  }
  ray.categories = categories; 
  var data = JSON.stringify(ray);
       
  var options = {
    "method": "POST",
    'contentType': 'application/json',
    "headers": {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
    },
    "payload": data
  };
    

  var response = UrlFetchApp.fetch(WEB_APP_URL, options);
  
}  


function getAllCategoriesIndexes(sheet){
    
  categorieCol = sheet.getRange("A1:A").getValues();
  var end = 0;
  const categKey = "CATEGORIE";
  var indexes = [];
  
  //Index of the "RAYON" range
  indexes.push(12);
  
  for(var i = 0; i < categorieCol.length; i++) {
    var value = categorieCol[i][0];
    //Logger.log(value);
    if(value == '') {
      end++;
    }
    else {
      end = 0;
      if(typeof(value) == "string" && value.indexOf(categKey) !== -1) {
        //+6: because we have 5 rows between the index and the real data, and we start the for in 0;
        indexes.push(i+6);
      }
    }
    if(end > 5) {
      
      return indexes;
    }
  }
}

function collect_n1_data (storeID, sectionID){
      
  var payload = {
    "year": 2018, 
  };
  
  var json = JSON.stringify(payload);
  
  
  var options = {
    "method": "POST",
    'contentType': 'application/json',
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    "payload":  json
    
  };
  
  var url = "";
  Logger.log(url);
    
  var response = UrlFetchApp.fetch(url, options);
  console.log(response.getContentText());
  
  return response.getContentText(); 
}

function import_n1_data (){

  var ss = SpreadsheetApp.getActive();
  var data = [];    
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
  
  var storeID = sheet.getRange("C2").getValue();
  var sectionID = sheet.getRange("B4").getValue();
  
  var json_data_str = collect_n1_data(storeID, sectionID); 
  var json_data = JSON.parse(json_data_str);
  //Logger.log(JSON.stringify(json_data))
  
  var num_categories = json_data.categories;
  var index_list = getAllCategoriesIndexes(sheet);  
  
  
  var ids_list_ss = [];
  
  var list_obj = [];
  for (var m=0; m < index_list.length; m++){
    var id_ss = sheet.getRange("C"+(index_list[m]-5)+":C"+(index_list[m]-5)).getValue(); 
    ids_list_ss.push(id_ss);
    
    var obj = {};
    obj.row = parseInt(index_list[m])-5;
    obj.id = id_ss;
    list_obj.push(obj);
    Logger.log("list_obj: " + obj.id + " -------" + obj.row)
  }
  
  
  var row_separator = 18;
    
  if (num_categories.length > index_list.length){
    SpreadsheetApp.getUi().alert("Oops! The number of categories is greater in the JSON than in the Spreadsheet.");
    return;
  }  
  

  for (var i in num_categories){
        
    var data = json_data.categories[i].data; 
    var categoryID_JSON = parseInt(json_data.categories[i].id); //JSON ID

    Logger.log("categoryID_JSON: " + categoryID_JSON);
    
    var index;
    for(var j in list_obj){
      //Logger.log(list_obj[j]);  
      if(list_obj[j].id == categoryID_JSON) {  
        index = list_obj[j].row + 4;
        break;
      }
      
    }
    Logger.log("INDEX: " + index ); 
    var categoryID = sheet.getRange("C"+(index-4)+":C"+(index-4)).getValue(); //ID SS   
    var counter = 0;

    //months || j < 12 
    for (var j in data){
      counter ++;
      var range_B_D = sheet.getRange("B"+(index+counter)+":D"+(index+counter));
      var range_P = sheet.getRange("P"+(index+counter)+":P"+(index+counter));   
      var range_B_D_profitability = sheet.getRange("B"+(index+counter+row_separator)+":D"+(index+counter+row_separator));
      var range_N_O = sheet.getRange("N"+(index+counter+row_separator)+":O"+(index+counter+row_separator));
           
      var month = data[j].month;
      var turnover = data[j].turnover;
      var calendarEffect = data[j].calendarEffect;
      var turnoverMarketTrends = data[j].turnoverMarketTrends;
      var grossTurnover = data[j].grossTurnover;
      var profitability = data[j].profitability;
      var profitabilityRate = data[j].profitabilityRate;
      var profitabilityMarketTrends = data[j].profitabilityMarketTrends;
      var markdownProfit = data[j].markdownProfit;
      var antiWastageProfit = data[j].antiWastageProfit;
      
      var data_B_D = [turnover,calendarEffect,turnoverMarketTrends]      
      var data_B_D_profitability = [profitability, profitabilityRate, profitabilityMarketTrends]
      var data_N_O = [markdownProfit, antiWastageProfit];
      
      
      range_B_D.setValues([data_B_D]);
      range_P.setValue([grossTurnover])
      range_B_D_profitability.setValues([data_B_D_profitability]);
      range_N_O.setValues([data_N_O]);
      
      data_B_D = [];
      data_B_D_profitability = [];
      data_N_O = [];
    }
  }   
}


