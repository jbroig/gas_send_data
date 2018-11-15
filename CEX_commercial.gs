
function collect_data_CEX_commercial (){
  var ss = SpreadsheetApp.getActive();
  var data = [];
      
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
     
  var test = getAllCategoriesIndexes(sheet);
  
  var categories = [];
  
  var ray = {};
  ray.year = 2018;

  var indexes = getAllCategoriesIndexes(sheet);
  
  var forecast;
  var month;
  var row_separator = 18;
  
  //Recorre indices
  for (var j in indexes){
    month = 0;
    forecast = [];
    var index = indexes[j];     
    var id = sheet.getRange("B"+(index-5)+":B"+(index-5)).getValue();
    
    // Recorre por meses
    for (var i = index; i < (index + 12); i++){
    
      month += 1;
      var turnover = sheet.getRange("G"+i+":G"+i).getValue(); 
      if (turnover == "" || turnover == undefined ){turnover = null;}
      var turnoverEvolution = sheet.getRange("H"+i+":H"+i).getValue();
      if (turnoverEvolution == "" || turnoverEvolution == undefined ){turnoverEvolution = null;}
      var profitability = sheet.getRange("G"+(i + row_separator)+":G"+(i + row_separator)).getValue();
      if (profitability == "" || profitability == undefined ){profitability = null;}
      var profitabilityRate = sheet.getRange("H"+(i + row_separator)+":H"+(i + row_separator)).getValue();
      if (profitabilityRate == "" || profitabilityRate == undefined ){profitabilityRate = null;}
      
      var forecast_obj = {}
      forecast_obj.month = month;
      forecast_obj.turnover = turnover;
      forecast_obj.turnoverEvolution = turnoverEvolution;
      forecast_obj.profitability = profitability;
      forecast_obj.profitabilityRate = profitabilityRate;
      
      forecast.push(forecast_obj);
      
      //categories_obj
      var obj = {};
      obj.id = id;
      obj.forecasts = forecast;
      
    }
        
    categories.push(obj);
        
  }
  ray.categories = categories; 
  data = JSON.stringify(ray);

  var payload = {
    "magasin": sheet.getRange("B2").getValue(),
    "departement": sheet.getRange("B3").getValue(),
    "data": data
  };
  var options = {
    "method": "POST",
    'contentType': 'application/json',
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    "payload": payload
  };
  
  var response = UrlFetchApp.fetch(WEB_APP_URL, options);
  
  if (response.getResponseCode() == 200){
    SpreadsheetApp.getUi().alert("The data has been validated");
  }
}  


function getAllCategoriesIndexes(sheet){
    
  categorieCol = sheet.getRange("A1:A").getValues();
  var end = 0;
  const categKey = "CATEGORIE";
  var indexes = [];
  for(var i = 0; i < categorieCol.length; i++) {
    var value = categorieCol[i][0];
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

function collect_n1_data (){
    
  var payload = {
    "year": 2018, 
  };
  
  var json = JSON.stringify(payload);  
  var options = {
    "method": "POST",
    'contentType': 'application/json',
    "payload":  json
    
  };
  
  //URL endpoint
  var response = UrlFetchApp.fetch("", options);
  return response.getContentText(); 
}

function import_n1_data (){

  var ss = SpreadsheetApp.getActive();
  var data = [];    
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
  
  var json_data_str = collect_n1_data(); 
  Logger.log(json_data_str);
  var json_data = JSON.parse(json_data_str);
 
  var num_categories = json_data.categories;
      
  var index_list = getAllCategoriesIndexes(sheet);  
  var row_separator = 18;
  
  if (num_categories.length != index_list.length){
    SpreadsheetApp.getUi().alert("Oops! The number of categories should be the same in the JSON and in the Spreadsheet.");
    return;
  }  
  
  for (var i in num_categories){
        
    var data = json_data.categories[i].data; 
    var index = index_list[i]-1;
    var counter = 0;

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
