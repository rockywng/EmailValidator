// retrieve sheet object
var sheetObj = SpreadsheetApp.getActiveSpreadsheet();

// tab with api key
var apiTab =  sheetObj.getSheetByName("Settings")

// tab with user input and output 
var validatorTab = sheetObj.getSheetByName("Validator")

// find the last row of the validator tab to obtain row count
var lastRowValidator = validatorTab.getLastRow();

// retrieve the api key from the api tab
var apiKey = apiTab.getRange(1, 2).getValue();

// use to build ui 
var ui = SpreadsheetApp.getUi();

// make api call to obtain info on email
function checkEmailValidator() {
  // for loop to loop through column of email inputs
  for(var i=2;i<=lastRowValidator;i++){
    // retrieve email from column
    var email = validatorTab.getRange(i, 1).getValue();

    // url for api call
    var url = "http://apilayer.net/api/check?access_key=" + apiKey + "&email=" + email + "&smtp=1&format=1"

    // use url to fetch
    var fetch = UrlFetchApp.fetch(url)

    // when api key has not been added or api key is not valid
    if(!apiKey){
    
      // return error message
      var htmlOutput = HtmlService
        .createHtmlOutput('<p style="text-align:center">Remember to enter an API key in the  <b>Settings</b> tab in cell <b>B1</b>.<br><br></p>')
        .setWidth(300)
        .setHeight(300);

      // dialogue title
      return ui.showModalDialog(htmlOutput, 'API Key Required');
    }

    // if api responds 
    if(fetch.getResponseCode() == 200) {
        // format response into sheets cells
        var response = JSON.parse(fetch.getContentText());
        validatorTab.getRange(i, 2).setValue(response.domain)
        validatorTab.getRange(i, 3).setValue(response.catch_all)
        validatorTab.getRange(i, 4).setValue(response.format_valid)
        validatorTab.getRange(i, 5).setValue(response.mx_found)
        validatorTab.getRange(i, 6).setValue(response.smtp_check)
        validatorTab.getRange(i, 7).setValue(response.free)
        validatorTab.getRange(i, 8).setValue(response.score)
    } 
    
    // if api key has run out of credits
    else if(fetch.getResponseCode()  == 104){
     var htmlOutput = HtmlService
    .createHtmlOutput('<p style="text-align:center">You have run out of credits with this API key. Please set up a  new API key.<br><br</p>')
    .setWidth(250)
    .setHeight(300);
    return ui.showModalDialog(htmlOutput, 'No Remaining API Credits!');
    }
  
  // if error arises
  else{
     validatorTab.getRange(i, 9).setValue("Exception Occured: Response code "+ fetch.getResponseCode() 
     +"Checkout https://bit.ly/EmailValidatorAPIDocs for more info")
  }
  }
  
}

// called when validator function is first called
function onOpen(){
  // create ui
  ui.createMenu("Email Validator")
    .addItem("Run Email Validator",'checkEmailValidator')
    .addToUi()

}


