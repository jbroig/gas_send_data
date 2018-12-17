function getURL_autres(magasinID, departmentID){
    
  var WEB_APP_URL = "";
  
  Logger.log(WEB_APP_URL);
  return WEB_APP_URL;
}

function collect_data_autres_elements (){
 
  var ss = SpreadsheetApp.getActive();
      
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
     
  var magasinID = sheet.getRange("C1").getValue();
  var departmentID = sheet.getRange("G1").getValue();
  var web_app_url = getURL_autres(magasinID, departmentID);

  var sections = [];
  
  var ray = {};
  ray.year = 2019;

  var indexes = getAllCategoriesIndexes_AutresElements(sheet);
  Logger.log("Indice 1: " + indexes[0] + " Indice 2: " + indexes[1]);
  var index = indexes[0];
  
  var data_list = [];
  var month = 0;
    
  var deferredCostsInput, consumerBenefitsInput, accountingGap, deferredCosts, consumerBenefits, grossBenefits, grossBenefitsRate, previousYearGrossBenefitsRate, deferredCostsPreTaxTurnoverRatio, consumerBenefitsPreTaxTurnoverRatio; 
  
  //Will iterate through all the elements until the next index. (index -1).
  for (var i =0; i < (indexes[1]-indexes[0]); i++){
   
    
    month = 0;
    data_list = [];
    var section_ID = sheet.getRange("B" + (index+i) + ":B" + (index+i)).getValue();
    
    if (section_ID.toString() == "Département PGC" || section_ID.toString() == "Pft Yc Restauration" || section_ID.toString() == "EQUIPEMENT ET SERVICES"   ){
      Logger.log("¿¿¿???: " + section_ID)
      break;
    }
       
    //We can escape the empty lines
    if (typeof section_ID == "string" || (section_ID == "" || section_ID === undefined || section_ID === null) ){ //parseInt(section_ID.toString()[0]) != departmentID
      continue;
    }
    
    Logger.log("Section_ID: " + section_ID);
    
    var obj = {};
    obj.id = section_ID;
    obj.data = data_list;
    
    var positions = check_ids (sheet, section_ID); 
          
    //recorremos meses
    for (var j = 0; j < 12; j++){
      
      month ++;
      var data_obj = {}
      data_obj.month = month;
      
      var row_data = sheet.getRange((positions[0]), (17 + j), 1, 129).getValues();
      
      //17: Column Q is number 17
      deferredCostsInput = row_data[0][0];
      //44: Column AR is number 44
      consumerBenefitsInput = row_data[0][27];
      //83: Column CE is number 71
      accountingGap = row_data[0][54];
      //109: Column DE is number 109
      grossBenefits = row_data[0][92];
      //135: Column EE is number 135 
      grossBenefitsRate = row_data[0][118];
      //122: Column DR is number 122
      previousYearGrossBenefitsRate = row_data[0][105];
      //31: column AE is number 31
      deferredCostsPreTaxTurnoverRatio = sheet.getRange((positions[0]), 30, 1, 1).getValue();
      //58: column BE is number 58
      consumerBenefitsPreTaxTurnoverRatio = sheet.getRange((positions[0]), 57, 1, 1).getValue();
      
      //var row_data_2 = sheet.getRange((positions[1]), (17 + j), 1, 39).getValues();      
      var row_data_3 = sheet.getRange((positions[2]), (17 + j), 1, 39).getValues();
      
      //17: Column Q is number 17
      deferredCosts = row_data_3[0][0];
      //44: Column AR is number 44
      consumerBenefits = row_data_3[0][27];
      
      if (typeof deferredCostsInput == "string" || (deferredCostsInput == "" || deferredCostsInput == undefined)){deferredCostsInput = null;}
      if (typeof consumerBenefitsInput == "string" || (consumerBenefitsInput == "" || consumerBenefitsInput == undefined)){consumerBenefitsInput = null;}
      if (typeof accountingGap == "string" || (accountingGap == "" || accountingGap == undefined)){accountingGap = null;}
      if (typeof deferredCosts == "string" || (deferredCosts == "" || deferredCosts == undefined)){deferredCosts = null;}
      if (typeof consumerBenefits == "string" || (consumerBenefits == "" || consumerBenefits == undefined)){consumerBenefits = null;}
      if (typeof grossBenefits == "string" || (grossBenefits == "" || grossBenefits == undefined)){grossBenefits = null;}
      if (typeof grossBenefitsRate == "string" || (grossBenefitsRate == "" || grossBenefitsRate == undefined)){grossBenefitsRate = null;}
      if (typeof previousYearGrossBenefitsRate == "string" || (previousYearGrossBenefitsRate == "" || previousYearGrossBenefitsRate == undefined)){previousYearGrossBenefitsRate = null;}
      if (typeof deferredCostsPreTaxTurnoverRatio == "string" || (deferredCostsPreTaxTurnoverRatio == "" || deferredCostsPreTaxTurnoverRatio == undefined)){deferredCostsPreTaxTurnoverRatio = null;}
      if (typeof consumerBenefitsPreTaxTurnoverRatio == "string" || (consumerBenefitsPreTaxTurnoverRatio == "" || consumerBenefitsPreTaxTurnoverRatio == undefined)){consumerBenefitsPreTaxTurnoverRatio = null;}
      
      data_obj.deferredCostsInput = deferredCostsInput;
      data_obj.consumerBenefitsInput = consumerBenefitsInput;
      data_obj.accountingGap = accountingGap;      
      data_obj.deferredCosts = deferredCosts;      
      data_obj.consumerBenefits = consumerBenefits;     
      data_obj.grossBenefits = grossBenefits;
      data_obj.grossBenefitsRate = grossBenefitsRate;
      data_obj.previousYearGrossBenefitsRate = previousYearGrossBenefitsRate;
      data_obj.deferredCostsPreTaxTurnoverRatio = deferredCostsPreTaxTurnoverRatio;
      data_obj.consumerBenefitsPreTaxTurnoverRatio = consumerBenefitsPreTaxTurnoverRatio;
      
      data_list.push(data_obj);
           
    }
    
    sections.push(obj);
    ray.sections = sections;     
   
  }
    
  var data = JSON.stringify(ray);
  
  var options = {
    "method": "POST",
    'contentType': 'application/json',
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    //"muteHttpExceptions": true,
    "payload": data
  };
  
  
  //We need the App Engine Endpoint 
  var response = UrlFetchApp.fetch(web_app_url, options);
  Logger.log("getResponseCode: " + response.getResponseCode())
  
}


function getAllCategoriesIndexes_AutresElements(sheet){
     
  //var sheet = SpreadsheetApp.getActive().getActiveSheet();
  //Logger.log("Sheet name: " + sheet.getName())
  categorieCol = sheet.getRange("B1:B").getValues();
  var end = 0;
  const categKey = "RAYONS";
  var indexes = [];
  for(var i = 0; i < categorieCol.length; i++) {
    var value = categorieCol[i][0];
    if(value == '') {
      end++;
    }
    else {
      end = 0;
      if(typeof(value) == "string" && value.indexOf(categKey) !== -1) {
        //+2: because we have 1 rows between the index and the real data, and we start the for in 0;
        //Logger.log(i+i);
        indexes.push(i+2);
      }
    }
  
  }
  Logger.log(indexes)
  return indexes;
}


function collect_n1_data_autres_elements (storeID, departmentID){
    
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  
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
    
  //TO-DO: Update the URL endpoint ->
  var response = UrlFetchApp.fetch("", options);
  Logger.log("RESPONSE CODE: " + response.getResponseCode());
  Logger.log(response.getContentText());
  return response.getContentText(); 
}


function import_n1_data_autres_elements (){

  var ss = SpreadsheetApp.getActive();
  var data = [];    
  var sheet = ss.getActiveSheet();
  var rayon = ss.getActiveSheet().getName();
    
  var storeID = sheet.getRange("C1").getValue();
  var departmentID = sheet.getRange("G1").getValue();
    
  var json_data_str = collect_n1_data_autres_elements (storeID, departmentID);
  var json_data = JSON.parse(json_data_str);
  
  var index_list = getAllCategoriesIndexes_AutresElements(sheet);
  Logger.log("INDEX LIST: " + index_list);
  
  
  var num_sections = json_data.sections.length;
  Logger.log("num_sections: " + num_sections);
  
  var number_of_sections = extractIDs(sheet, index_list[1]);
  Logger.log("LLEGAN LOS ID'S: " + number_of_sections);
  
  
  for (var i=0; i < num_sections; i++){
    var data = json_data.sections[i].data;
    var JSON_SECTION_ID = json_data.sections[i].id;
    //Logger.log("Data length: " + data.length);
            
    for(var j in number_of_sections){
      if(number_of_sections[j].id == JSON_SECTION_ID) {  
        var index = number_of_sections[j].row;
        var id = number_of_sections[j].id;
        
        Logger.log("A VER ESE ID: " + id);
        Logger.log("LA  ROW: "  + index)
        
        break;
      }      
    }
    
  
    var positions = check_ids(sheet, id);
    Logger.log("positions: " + positions);
    
    //Iterate months
    //j < 12 
    for (var j=0; j<data.length; j++) {
           
      var month = data[j].month;
      
      var separator = 0;
      
      switch (month){
        case 1:
          break;
        case 2:
          separator +=1;
          break;
        case 3:
          separator +=2;
          break;
        case 4:
          separator +=3;
          break;
        case 5:
          separator +=4;
          break;
        case 6:
          separator +=5;
          break;
        case 7:
          separator +=6;
          break;    
        case 8:
          separator +=7;
          break;
        case 9:
          separator +=8;
          break;
        case 10:
          separator +=9;
          break;
        case 11:
          separator +=10;
          break;
        case 12:
          separator +=11;
          break;
        default:
          break;
      }
        
      var accountingGap = data[j].accountingGap;
      var grossBenefits = data[j].grossBenefits;
      var preTaxTurnover = data[j].preTaxTurnover;
      var deferredCosts = data[j].deferredCosts;
      var consumerBenefits = data[j].consumerBenefits;
      
      
      //COLUMN START = 70 --> BR --> BF --> 58
      var range_accountingGap = sheet.getRange((index), (58+separator), 1, 1).setValue(accountingGap);
      //COLUMN START = 96 --> CR --> CF --> 84
      var range_grossBenefits = sheet.getRange((index), (84+separator), 1, 1).setValue(grossBenefits);
      //COLUMN START = 4 --> D || ROW: INDEX + 19
      var range_preTaxTurnover = sheet.getRange((positions[1]), (4+separator), 1, 1).setValue(preTaxTurnover);    
      //COLUMN START = 4 --> D || ROW: INDEX + 37
      var range_deferredCosts = sheet.getRange((positions[2]), (4+separator), 1, 1).setValue(deferredCosts);  
      //COLUMN START = 31 --> AE || ROW: INDEX + 37
      var range_consumerBenefits = sheet.getRange((positions[2]), (31+separator), 1, 1).setValue(consumerBenefits); 
      
    }
  }  
}

function check_ids (sheet, id){
  
  Logger.log("TEST || " + " INDEX: " +id);
  var values_column_b = sheet.getRange("B1:B").getValues();
  
  var positions = [];
  for (var i=0; i < values_column_b.length; i++){
    if (values_column_b[i] == id){positions.push(i+1);}
  }
  return positions;
}

function extractIDs (sheet, end){
    
  var obj_list = []; 
  for (var i=0; i < end; i++){
    var obj = {};
    var id = sheet.getRange("B"+(i+1)+":B"+(i+1)).getValue();
    
    obj.row = i+1;
    obj.id = id;
    
    obj_list.push(obj);
  }
  
  Logger.log("obj_list: " + obj_list);
  return obj_list;
}


