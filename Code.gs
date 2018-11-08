var WEBHOOK = "https://webhook.site/.........." //Use your webhook endpoint

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Validation")
      .addItem('Validate data', 'main')
      .addToUi();
}

function main (){
  //SpreadsheetApp.getUi().alert("In progress");
  collect_data();
  
}

function collect_data (){
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var data = [];
  
  sheets.forEach(function(sheet) {
    
    var rayon = sheet.getName();
    
    //Works for CEX commercial - Rxx sheets
    if (rayon.indexOf("CEX commercial") >= 0) {
    
      var indexes = [48,66]; 
      
      var categories = [];
      for(var i = 0; i < indexes.length; i++) {
        
        var categName = sheet.getRange("B44").getValue();    
        var categValues = sheet.getRange("G" +(indexes[i]+1)+ ":H" +(indexes[i]+12)).getValues();
        var obj = {};

        obj[categName] = categValues;
        categories.push(obj);
      }
      var ray = {};
      ray[rayon] = categories;
      data.push(ray);
    
    //Works for Autres Eléments du CEX sheets
    } else if (rayon.indexOf("Autres Eléments du CEX") >= 0){
      
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
            
    }
  });
    
  var json_final = [];
  
  for (var line in data){
    json_final.push(JSON.stringify(data[line]));
  }
   
  as = ss.getActiveSheet();
  var payload = {
    "magasin": as.getRange("B2").getValue(),
    "departement": as.getRange("B3").getValue(),
    "data": JSON.stringify(data)
  };
  var options = {
    "method": "POST",
    "headers": {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    "payload": payload
  };
  
  var response = UrlFetchApp.fetch("WEBHOOK", options);
  
}













