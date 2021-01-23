function createAndUpdateCalendarEvents() {
  const TESTING = false; //CAREFUL set testing to false, when on production

  var spreadsheet = SpreadsheetApp.getActiveSheet();
  if (TESTING){
    spreadsheet = SpreadsheetApp.getActive().getSheetByName("SheetTest");
  }
  Logger.log("Executing script on sheet named: "+spreadsheet.getName());
  var lr = spreadsheet.getLastRow();
  var calendarId = spreadsheet.getRange("A3").getValue(); 
  var eventCal = CalendarApp.getCalendarById(calendarId);
  
  //select data
  Logger.log("Last Row = " +lr);
  var range = spreadsheet.getRange("A7:"+lr);
  var dates = range.getValues();

  const today = new Date()
  const yesterday = new Date(today)
  yesterday.setDate(yesterday.getDate() - 1);
  
  for (i=0 ; i<dates.length; i++){
    var row = dates[i];
   
    var dateStart = row[map.get(cDataInici)];
    var dateFinal = row[map.get(cDataFinal)];
    var concept = row[map.get(cConcept)];
    
    if (concept == "" || dateStart == "" || dateFinal == ""){
      //Logger.log("Either one of the following is empty. Cannot create Event (concept, dateStart, dateFinal): ", concept, dateStart, dateFinal);
      continue;
    }
    
    if (dateFinal < yesterday){ //examen passat
      continue;
    }
    
    var preu = row[map.get(cPreu)];
    var assigned = row[map.get(cFreelancer)];
    var calendarEventId = row[map.get(cCalendarEventID)];
    
    if (assigned){
      concept += (' assigned: ' + assigned); 
    }
    
    var event;

    if (calendarEventId){
      event = eventCal.getEventById(calendarEventId);

      if (!event){
        Logger.log("Not found event in this calendar, but calendarEventId row is set: "+ calendarEventId + "\nTitle is: "+concept);
        
      }

       if (row[map.get(cDelete)].toString().toUpperCase() == "DELETE"){
         Logger.log("About to delete: " + concept);
        
        dates[i].forEach( (value, j) => {
                         dates[i][j] = "";
                         });
        event.deleteEvent();
         continue;
      } 
      
      //update event
      event.setTitle(concept);
    } else {
      //Logger.log(concept, dateStart, dateFinal);
      
      try{
        event = eventCal.createEvent( concept, dateStart, dateFinal);
        Logger.log("Created event: " + concept);
        //Logger.log(event.getId());
        dates[i][map.get(cCalendarEventID)] = event.getId();
      }catch (error){
        throw("No se puede crear evento en Calendar. Check if Concept, DataInici y DataFinal están correctos.\nFunction: CalendarApp.Calendar.createEvent\nconcept: "+concept+" DataInici: "+dateStart + " DataFinal: "+dateFinal);
       
      }

      //write now all the rows to the Google Sheet so that if there is an error further down the Event ID is saved
      range.setValues(dates);
    }
    
    if (preu){
     event.setColor("10");  //green = paid
    } else {
      if (assigned){
        event.setColor("7"); //cyan = assigned
      } else {
        event.setColor("11"); //red
      }
    }
    
  }
  
  range.setValues(dates);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Sync Now', 'createAndUpdateCalendarEvents')
      .addToUi();
}
