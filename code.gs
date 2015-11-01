//push new events to calendar
function alertar() {Browser.msgBox("Calendar Update Complete")}

function createCalEvent(flag,startDate,endDate,campaignName,formatName){

var baseFlagUrl = "https://raw.githubusercontent.com/stevenrskelton/flag-icon/master/gif/"

  switch(flag){
  
    case "Spain":
    var icono = baseFlagUrl + "es.gif"
    break;
    case "Germany":
    var icono = baseFlagUrl + "de.gif"
    break;
    case "UK":
    var icono = baseFlagUrl + "gb.gif"
    break;
    case "Italy":
    var icono = baseFlagUrl + "it.gif"
    break;
    default:
    var icono = baseFlagUrl + "europeanunion.gif"
    break;
  }  
  
var event = {
  'summary': flag + " - " + campaignName + " - " + formatName,
  'start': {
    'date': startDate
  },
  'end': {
    'date': endDate
  },
 "gadget": {
 "display":"icon",
 "type": "application/x-google-gadgets+xml",
 "title": "Word of the Day",
 "link": "https://raw.githubusercontent.com/zardoya/mariano/master/gadget_test.xml",
 "iconLink": icono
}
};

Calendar.Events.insert(event, 'gh97ai8n0k7jgcieeglgvtg7fg@group.calendar.google.com');
}

function pushToCalendar(person) {
   
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow(); 
  var range = sheet.getRange(2,1,lastRow,8);
  var values = range.getValues();   
  //var updateRange = sheet.getRange('G1'); 
   
  //calendar variables
  var calendar = CalendarApp.getCalendarById('gh97ai8n0k7jgcieeglgvtg7fg@group.calendar.google.com')
   
  //show updating message
  //updateRange.setFontColor('red');
   
  var numValues = 0;
  for (var i = 0; i < values.length; i++) {   
    //check to see if name and type are filled out - date is left off because length is "undefined"
    if ((values[i][0].length > 0) && (values[i][2].length > 0)) {
      //check if it's been entered before          
      if (values[i][6] != 'Uploaded') {                       
         
        //create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        var country = values[i][0];
        var advertiser = values[i][2];
        var format = values[i][4];
        var date = new Date(values[i][1]).toJSON();
        var end = new Date(date);
        end.setDate(end.getDate()+1);
        date = date.slice(0,date.indexOf("T"))
        //Browser.msgBox(date);
        end = end.toJSON();
        end = end.slice(0,end.indexOf("T"));
        //Browser.msgBox(end);
        createCalEvent(country,date,end,advertiser,format);
        //var newEventTitle = values[i][0] + ' - ' + values[i][2] + ' - ' + values[i][4];
        //var newEvent = calendar.createAllDayEvent(newEventTitle, values[i][1], {location:newplace.getContent()});
         
        //get ID
        //var newEventId = newEvent.getId();
         
        //mark as entered, enter ID
        sheet.getRange(i+2,7).setValue('Uploaded');
        sheet.getRange(i+2,8).setValue(person);
         
      } //could edit here with an else statement
    }
    numValues++;
  }
   
  //hide updating message
 // updateRange.setFontColor('white');
 // Browser.msgBox("Calendar Update Complete")
 
}

function showSidebar() {
 var html = HtmlService.createHtmlOutputFromFile('siderBar').setSandboxMode(HtmlService.SandboxMode.NATIVE).setTitle('ACT ID Calendar Uploader').setWidth(300);
 //html.append('<div id="dvImportSegments" class="fileupload "><fieldset><legend>Upload your CSV File</legend><input type="file" name="File Upload" id="txtFileUpload" accept=".csv" /></fieldset></div>')
  SpreadsheetApp.getUi().showSidebar(html);
   /*var app = UiApp.createApplication().setTitle("Upload CSV to Sheet");
  var formContent = app.createVerticalPanel();
  var form = app.createFormPanel();
  formContent.add(app.createHTML('<p style="margin:10px;"><b>Ad Creative Tech uploader</b></p>'));
  var lb = app.createListBox(false).setId('myId').setName('myLbName');
   // add items to ListBox
   lb.setVisibleItemCount(1);
   lb.addItem('Select Uploader Name');
   lb.addItem('Ehab Haddad');
   lb.addItem('Graziano Muscas');
   lb.addItem('Michael Parry');
   lb.addItem('Ulisses Alves Machado');
  formContent.add(lb);
  formContent.add(app.createHTML('<p style="margin:10px;"><b>Salesforce report to upload</b></p>'));
  formContent.add(app.createFileUpload().setName('thefile'));
  formContent.add(app.createHTML('<p style="margin:10px;"><b>Hola</b></p>'));
  formContent.add(app.createSubmitButton('Start Upload'));
   form.add(formContent);
   app.add(form);
  SpreadsheetApp.getUi().showSidebar(app);*/
}

function uploadReport() {
//SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     showSidebar();
}

function createChart() {
   var sampleData = Charts.newDataTable()
       .addColumn(Charts.ColumnType.STRING, "Year")
       .addColumn(Charts.ColumnType.NUMBER, "Sales")
       .addColumn(Charts.ColumnType.NUMBER, "Expenses")
       .addRow(["2004", 1000, 400])
       .addRow(["2005", 1170, 460])
       .addRow(["2006", 660, 1120])
       .addRow(["2007", 1030, 540])
       .addRow(["2008", 800, 600])
       .addRow(["2009", 943, 678])
       .addRow(["2010", 1020, 550])
       .addRow(["2011", 910, 700])
       .addRow(["2012", 1230, 840])
       .build();
   
   var chart = Charts.newColumnChart()
       .setTitle('Sales vs. Expenses')
       .setXAxisTitle('Year')
       .setYAxisTitle('Amount (USD)')
       .setDimensions(600, 500)
       .setDataTable(sampleData)
       .build();
   
   return UiApp.createApplication().add(chart);
 }

//add a menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Update Calendar", functionName: "pushToCalendar"},{name: "Show Sidebar", functionName: "uploadReport"},{name: "Create Chart", functionName: "createChart"}); 
  sheet.addMenu("Event Calendar", menuEntries);  
}
 
/* 
Notes: 
         To trigger an alert use Browser.msgBox(); 
         onEdit triggers when spreadsheet is edited, however was unreliable
         To support edits, we need to grab calendar event by id (currently is not supported), then compare the date of the event with the date in the spreadsheet and adjust if they don't match
*/