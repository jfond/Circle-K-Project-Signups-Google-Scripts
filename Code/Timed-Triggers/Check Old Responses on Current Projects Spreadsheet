function removeOldResponses(){
  //This removes the old responses from the spreadsheet.
  //That stops the projects from appearing on the "delete a project" form, and the "sign up for a project" form
    
  var currentProjectsSpreadsheet = SpreadsheetApp.openById("1T6QaKngOqfaxf6kyLDtSdCMeATiFGF5jxlsDwE5ABNw"); 
  var currentProjectsSheet = currentProjectsSpreadsheet.getSheets()[0];
  var dataRange = currentProjectsSheet.getDataRange();
  var numProjects = dataRange.getHeight()-1;
  var numParameters = dataRange.getWidth();
  var signupEndDateIndex;
  var projectNameIndex;
  var projectStartTimeIndex;
  var projectEndTimeIndex;
  var projectChairIndex;
  var projectChairEmailIndex;
  var projectTypeIndex;
  
  var unsubmittedHoursSpreadsheet = SpreadsheetApp.openById("1tcEum8LBGMEmhxt9ku7-KSSnggTxvWbJwNlTNT3zEok");
  var unsubmittedHoursSheet = unsubmittedHoursSpreadsheet.getSheets()[0];
  var ushHeight = unsubmittedHoursSheet.getDataRange().getHeight();
  
  var now = new Date();
  
  for (var i=0;i<numParameters;i++){
    if (currentProjectsSheet.getRange(1, i+1).getValue() == "Signup End Date"){
      signupEndDateIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Event/Project Name"){
      projectNameIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Project Start Date/Time"){
      projectStartTimeIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Project End Date/Time"){
      projectEndTimeIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Project Chairperson"){
      projectChairIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Project Chairperson Email Address"){
      projectChairEmailIndex = i+1;
    }
    else if (currentProjectsSheet.getRange(1, i+1).getValue() == "Type of Event"){
      projectTypeIndex = i+1; 
    }
    
  }  
     
  //Do a for loop through the current projects spreadsheet and delete entries that have passed signup date
  for (var j=0;j<numProjects;j++){
    if (currentProjectsSheet.getRange(j+2,signupEndDateIndex).getValue().getTime() <= now.getTime()){
      //First access the spreadsheet which holds the names of the people who sign up for the project and delete it.
      var startTime = currentProjectsSheet.getRange(j+2, projectStartTimeIndex).getValue();
      var endTime = currentProjectsSheet.getRange(j+2, projectEndTimeIndex).getValue();
      var projectName = currentProjectsSheet.getRange(j+2, projectNameIndex).getValue();
      var projectChair = currentProjectsSheet.getRange(j+2, projectChairIndex).getValue();
      var projectChairEmail = currentProjectsSheet.getRange(j+2, projectChairEmailIndex).getValue();
      var projectType = currentProjectsSheet.getRange(j+2, projectTypeIndex).getValue();
      var signupsSheetName = projectName + " - " + DatetoWords(startTime, endTime);
      var sheetFile = DriveApp.getFilesByName(signupsSheetName).next();
      var ushHeight = unsubmittedHoursSheet.getDataRange().getHeight()+1; 
      unsubmittedHoursSheet.getRange(ushHeight, 1, 1, numParameters).setValues(currentProjectsSheet.getRange(j+2, 1,1,numParameters).getValues()); 
      if (projectType == "Service Project"){
        sendProjectSignupEndEmailandForm(sheetFile, projectName, projectChair, projectChairEmail, ushHeight);
      }
      else {
        sendEventRosterEmail(sheetFile, projectName, projectChair, projectChairEmail); 
      }
      currentProjectsSheet.getRange(j+2, 1, 1, numParameters).clearContent();
      for (var n=j;n<numProjects+1;n++){
        currentProjectsSheet.getRange(n+2, 1, 1, numParameters).setValues(currentProjectsSheet.getRange(n+3, 1,1,numParameters).getValues())      
      }
      j-=1;
      numProjects -=1;
    }
  }

  //Here I put the refreshForms call outside ofthe loop and if statement because I will also be checking for outdated projects.
  //I want to always refreshForms every 30 minutes, not ONLY when there is one to remove.
  refreshForms();
  
  
}

function sendProjectSignupEndEmailandForm(signupSpreadsheetFile, projectName, projectChair, projectChairEmail,ushIndex){
  //The signup spreadsheet is a spreadsheet, and every individual sheet of the spreadsheet corresponds to a different shift
  //Each sheet has information regarding the names, email, timestamps, and need ride information of the people who signed up
  
  var nameList = [];
  var emailList = [];
  //Now, we need to create the form:  
  var signupSpreadsheet = SpreadsheetApp.open(signupSpreadsheetFile);
  var signupSheets = signupSpreadsheet.getSheets();
  
  //First iterate through the number of shifts
  for (var i=0;i<signupSheets.length;i++){
    var numNames = signupSheets[i].getDataRange().getHeight()-1;
    var numParameters = signupSheets[i].getDataRange().getWidth();
    var nameIndex;
    var emailIndex;
    var name;
    var email;
    //Find where the signup names are located in this sheet.
    for (var n=0;n<numParameters;n++){
      if (signupSheets[i].getRange(1, n+1).getValue() == "First and Last Name"){
        nameIndex = n+1;
      }
      else if (signupSheets[i].getRange(1, n+1).getValue() == "Email Address"){
        emailIndex = n+1;
      }
    }
    //Iterate through every entry and add the name to the list if it is new.
    for (var j=0;j<numNames;j++){
      name = signupSheets[i].getRange(j+2, nameIndex).getValue();
      email = signupSheets[i].getRange(j+2, emailIndex).getValue();
      //indexOf returns -1 if the list doesnt have that entry.
      //So here, if either the name or the email is unique, append that entry to the lists.
      //Thus you can have duplicate names if the emails are different (because there could possibly be two people with the same name)
      //But they would have different emails
      
      //Each entry is a combination of a name and email. To be considered a duplicate, the entry must match both of those categories.
      if ((nameList.indexOf(name) == -1) || (emailList.indexOf(email) == -1)){
        nameList.push(name);
        emailList.push(email);        
      } 
    }
  }
  
  //Now that we have a list of all of the names of the people who signed up
  //We just have to populate a form with that
  var documentHoursFormTemplateFile = DriveApp.getFilesByName("Project Record Hours Form Template").next();
  var chairDocumentHoursFormFile = documentHoursFormTemplateFile.makeCopy();
  
  //In a google form there are two identifying strings: the form's title and the form's name.
  //The name is the name of the file in the drive, and the title is the bolded letters at the top of the form. 
  //In this context, I make both of them the same, so later when I want to reference the drive file, I can just call
  //Form.getTitle, and use that as an argument in DriveApp.getFilesByName(name). This is an easy way to corcumvent the problem that 
  //Form.getName does not exist.
  chairDocumentHoursFormFile.setName(projectName + "  - Document Service Hours");
  
  var chairDocumentHoursFormURL = "";
  while (chairDocumentHoursFormURL == ""){
    try{
      chairDocumentHoursFormURL = chairDocumentHoursFormFile.getUrl();
       }
    catch(err){
    }
  }
  var chairDocumentHoursForm = FormApp.openByUrl(chairDocumentHoursFormURL);
  var chairDocumentHoursFormID = chairDocumentHoursForm.getId();
  
  //Now we create a new field in the "Unsubmitted Hours Spreadsheet" with the URL of the form so we can reference it later.
  //We had to do this in this function, rather than in the removeOldResponses() function because we just created the form in this function  
  var unsubmittedHoursSpreadsheet = SpreadsheetApp.openById("1tcEum8LBGMEmhxt9ku7-KSSnggTxvWbJwNlTNT3zEok");
  var unsubmittedHoursSheet = unsubmittedHoursSpreadsheet.getSheets()[0];
  var ushNumParameters = unsubmittedHoursSheet.getDataRange().getWidth();
  //This for loop loops through the parameters and finds the index of the one we are looking for and puts the newly added project's URL in the right place.
  for (var p=0;p<ushNumParameters;p++){
    if (unsubmittedHoursSheet.getRange(1, p+1).getValue() == "Document Hours Form ID"){
      unsubmittedHoursSheet.getRange(ushIndex, p+1).setValue(chairDocumentHoursFormID);
    }
  }  

  chairDocumentHoursForm.setTitle(projectName + "  - Document Service Hours");
  
  //Do a for loop for each name set the title of a new item as the person's name
  //Then, duplicate it. At the end of the loop there will be one extra item, so delete it.
  var myTextItem = chairDocumentHoursForm.getItems(FormApp.ItemType.TEXT)[0].asTextItem();
  for (var k=0;k<nameList.length;k++){
    myTextItem.setTitle(nameList[k] + " (" + emailList[k] + ")");
    myTextItem.setRequired(true);
    myTextItem = myTextItem.duplicate(); 
    chairDocumentHoursForm.moveItem(myTextItem.getIndex(), k+1);
  }
  chairDocumentHoursForm.deleteItem(nameList.length);
   
  var MessageHTML = "Dear " + projectChair +  ",<br><br>";
  MessageHTML += "The signup period for your project: " + projectName + " has just finished.";
  MessageHTML += "<br><br>";
  MessageHTML +="Below is a link to a spreadsheet with all of the people who have signed up for the project ";
  MessageHTML +="and a form which you can use to document how many hours each person volunteered."
  MessageHTML +="<br><br>";
  MessageHTML +="Please submit the form in a timely manner and note that you may only submit it once."
  MessageHTML +="<br><br>";
  MessageHTML +="If you accidentally enter in the incorrect number of hours for someone"
  MessageHTML +=" then please contact the secretary."
  MessageHTML +="<br><br><br>";
  MessageHTML +="Thank you,";
  MessageHTML +="<br><br>";
  MessageHTML +="Circle K Project Signup Creator Bot";
  MessageHTML +="<br><br><br><br><br><br>";
  MessageHTML +='<a href="'+signupSpreadsheet.getUrl() + '">Signup Spreadsheet Link</a><br>';
  MessageHTML +='<a href="'+chairDocumentHoursForm.getPublishedUrl() + '">Document Hours Form</a>';
  
  
  GmailApp.sendEmail(projectChairEmail, "Project: " +projectName + " signup list and document hours form", "", {htmlBody: MessageHTML});
  
}

function sendEventRosterEmail(signupSpreadsheetFile, projectName, projectChair, projectChairEmail){
  //The signup spreadsheet is a spreadsheet, and every individual sheet of the spreadsheet corresponds to a different shift
  //Each sheet has information regarding the names, email, timestamps, and need ride information of the people who signed up
  
  var nameList = [];
  var emailList = [];
  //Now, we need to create the form:  
  var signupSpreadsheet = SpreadsheetApp.open(signupSpreadsheetFile);
  var signupSheets = signupSpreadsheet.getSheets();
  
  //First iterate through the number of shifts
  for (var i=0;i<signupSheets.length;i++){
    var numNames = signupSheets[i].getDataRange().getHeight()-1;
    var numParameters = signupSheets[i].getDataRange().getWidth();
    var nameIndex;
    var emailIndex;
    var name;
    var email;
    //Find where the signup names are located in this sheet.
    for (var n=0;n<numParameters;n++){
      if (signupSheets[i].getRange(1, n+1).getValue() == "First and Last Name"){
        nameIndex = n+1;
      }
      else if (signupSheets[i].getRange(1, n+1).getValue() == "Email Address"){
        emailIndex = n+1;
      }
    }
    //Iterate through every entry and add the name to the list if it is new.
    for (var j=0;j<numNames;j++){
      name = signupSheets[i].getRange(j+2, nameIndex).getValue();
      email = signupSheets[i].getRange(j+2, emailIndex).getValue();
      //indexOf returns -1 if the list doesnt have that entry.
      //So here, if either the name or the email is unique, append that entry to the lists.
      //Thus you can have duplicate names if the emails are different (because there could possibly be two people with the same name)
      //But they would have different emails
      
      //Each entry is a combination of a name and email. To be considered a duplicate, the entry must match both of those categories.
      if ((nameList.indexOf(name) == -1) || (emailList.indexOf(email) == -1)){
        nameList.push(name);
        emailList.push(email);        
      } 
    }
  }

    
  var MessageHTML = "Dear " + projectChair +  ",<br><br>";
  MessageHTML += "The signup period for your social event: " + projectName + " has just finished.";
  MessageHTML += "<br><br>";
  //MessageHTML +="Below is a table of the names and emails of everyone who signed up.";
  MessageHTML +="If you are interested in copy and pasting emails to send a message to everyone, we ";
  MessageHTML +="find it is easiest to do so from the spreadsheet, included in this message.";
  MessageHTML +="<br><br><br>";
  MessageHTML +="Thank you,";
  MessageHTML +="<br><br>";
  MessageHTML +="Circle K Project Signup Creator Bot";
  
  MessageHTML +="<br><br><br><br><br><br>";
  
  MessageHTML +='<a href="'+signupSpreadsheet.getUrl() + '">Signup Spreadsheet Link</a><br>';  
  //MessageHTML +='<table border="1" style="width:100%">';
  //MessageHTML +="<tr>";
  //MessageHTML +='<td>Name</td>';
  //MessageHTML +='<td>Email</td>';
  //MessageHTML +='</tr>';
  //for (var k=0;k<nameList.length;k++){
  //  MessageHTML +='<tr>';
  //  MessageHTML +='<td>'+ nameList[k] +'</td>';
  //  MessageHTML +='<td>'+ emailList[k] +'</td>';
  //  MessageHTML +='</tr>';
  //}
  //MessageHTML +='</table>';
  
  
  GmailApp.sendEmail(projectChairEmail, "Social: " +projectName + " signup list", "", {htmlBody: MessageHTML});
  
}

function getDateObjectFromString(DateString){
  //DateString comes in the format "YYYY-MM-DD HH:MM" 
  //When the function parseInt sees the "0" in front of the string it interprets it as an octal integer (base 8), which causes problems for the values 08 and 09.
  //To solve this, we pass a second argument to it specifying the standard base 10 system. 
  Logger.log("getDateObjectFromString: " + DateString);
  var YYYY = parseInt(DateString.substr(0,4), 10);
  var MM = parseInt(DateString.substr(5,2), 10)-1; //0-11;
  var DD = parseInt(DateString.substr(8,2), 10);
  var HH = parseInt(DateString.substr(11,2), 10);
  var MINS = parseInt(DateString.substr(14,2), 10);
  
  Logger.log("Day Substring: " + DateString.substr(8,2));
  Logger.log("Year: " + YYYY);
    Logger.log("Month: " + MM);
    Logger.log("Day: " + DD);
    Logger.log("Hour: " + HH);
    Logger.log("Minutes: " + MINS);
  var myDate = new Date(YYYY, MM, DD, HH, MINS, 0, 0);
  Logger.log("DateTime: " + myDate.getTime());
  return myDate;
}

function ConvertDateToString(date){
    //Format for the date is: "YYYY-MM-DD-YYYY HH:MM" when working in google forms
    //Unfortunately formatting the date properly takes several lines of code
  Logger.log("Convert Date to String: " + date);
  var DD = date.getDate();
  var MM = date.getMonth()+1; //January is 0!
  var YYYY = date.getFullYear();
  var HH = date.getHours(); //0-23
  var MINS = date.getMinutes(); //0-59
  var SECS = date.getSeconds(); //0-59
  
  if(DD<10) {
    //Convert months 1-9 to 01-09.
    DD='0'+DD;
  } 
  
  if(MM<10) {
    //Convert days 1-9 to 01-09.
    MM='0'+MM; 
  } 
  
  if (MINS<10){
    //Convert Mins 1-9 to 01-09
    MINS = '0'+MINS;
  }

  var nowDate = YYYY+"-"+MM+"-"+DD+" "+HH+":"+MINS;
  return nowDate;
}

function DatetoWords(startDate, endDate){
 //This function expects Javascript Date objects as its arguments. 
  
  var MM = startDate.getMonth(); //0-11, January is 0!
  var DD = startDate.getDate();
  var YYYY = startDate.getFullYear();
  var HH = startDate.getHours();//0-23
  var MINS = startDate.getMinutes();//0-59
  var hourSet = "AM";
  var startDate = new Date(YYYY,MM,DD,HH,MINS,0,0);
  
  //Find integers corresponding to the end time
  var MM2 = endDate.getMonth(); //0-11, January is 0!
  var DD2 = endDate.getDate();
  var YYYY2 = endDate.getFullYear();
  var HH2 = endDate.getHours();//0-23
  var MINS2 = endDate.getMinutes();//0-59
  var hourSet2 = "AM";
  var endDate = new Date(YYYY2,MM2,DD2,HH2,MINS2,0,0);
  
  //Determine the words corresponding to the integers given to us for the month, day
  var monthWord = getMonthFromInt(parseInt(MM), 10);
  var dayWord = getDayFromInt(startDate.getDay());
  var monthWord2 = getMonthFromInt(parseInt(MM2), 10);
  var dayWord2 = getDayFromInt(endDate.getDay());
  
  //Check for am vs pm
  if (HH > 11){
    hourSet = "PM";
    HH -= 12; 
  }
  if (HH == 0){
    HH = 12; 
  }
  HH = HH.toString()
  if (HH2 > 11){
    hourSet2 = "PM";
    HH2 -= 12; 
  }
  if (HH2 == 0){
    HH2 = 12; 
  }
  HH2 = HH2.toString();
  
  //Recast Minutes to have leading '0' if necessary
  if (MINS < 10){
    MINS = "0" + MINS.toString();
  }
  else{
    MINS = MINS.toString();
  }
  if (MINS2 < 10){
    MINS2 = "0"+MINS2.toString();
  }
  else {
    MINS2.toString();
  }

  //Now that everything has been examined, cast it all to string type
  YYYY = YYYY.toString();
  YYYY2= YYYY2.toString();
  MM = MM.toString();
  MM2 = MM2.toString();
  DD = DD.toString();
  DD2 = DD2.toString();

  var Words = "Error: Date not Parsed Correctly";
  if ((YYYY==YYYY2) && (DD==DD2) && (MM==MM2)){
    Words = dayWord + ", " + monthWord + " " + DD + ", " + YYYY + " ";
    Words += HH + ":" + MINS + " " + hourSet + " - " + HH2 + ":" + MINS2 + " " + hourSet2;  
  }
  else{
    Words = dayWord + ", " + monthWord + " " + DD + ", " + YYYY + " " + HH+":"+MINS + " " + hourSet + " - ";
    Words += dayWord2 + ", " + monthWord2 + " " + DD2 + ", " + YYYY2 + " " + HH2+":"+MINS2 + " " + hourSet2;
  }
      
  return Words;  
  
}
 
function getMonthFromInt(monthInt){
 var monthWord = "Error Parsing Month";
  
  switch (monthInt){
    case 0:
      monthWord = "January";
      break;
    case 1:
      monthWord = "February";
      break;
    case 2:
      monthWord = "March";
      break;
    case 3:
      monthWord = "April";
      break;
    case 4:
      monthWord = "May";
      break;
    case 5:
      monthWord = "June";
      break;
    case 6:
      monthWord = "July";
      break;
    case 7:
      monthWord = "August";
      break;
    case 8:
      monthWord = "September";
      break;
    case 9:
      monthWord = "October";
      break;
    case 10:
      monthWord = "November";
      break;
    case 11:
      monthWord = "December";
      break;
  }
  return monthWord;
  
}

function getDayFromInt(dayInt){
   
  var dayWord = "Error Parsing Day of the Week";
  
  switch (dayInt){
    case 0:
      dayWord = "Sunday";
      break;
    case 1:
      dayWord = "Monday";
      break;
    case 2:
      dayWord = "Tuesday";
      break;
    case 3:
      dayWord = "Wednesday";
      break;
    case 4:
      dayWord = "Thursday";
      break;
    case 5:
      dayWord = "Friday";
      break;
    case 6:
      dayWord = "Saturday";
      break;
  }
  return dayWord;
}

function getRosterSpreadSheetFile(sheetName){
  //Sheetname includes not only the event name, but also the date, as shown in all of the event roster spreadsheets.
  
  var potentialFiles = DriveApp.getFilesByName(sheetName);
  var myFile;
  var mySpreadsheet;
  var mySheets;
  
  //If there is no spreadsheet with that name found:
  if (!(potentialFiles.hasNext())){
     Logger.log("ERROR: Problem retrieving spreadsheet for " + sheetName);
     return null;
  }
  else {
    myFile = potentialFiles.next();
    if (potentialFiles.hasNext()){
      Logger.log("ERROR: Multiple Instances of spreadsheet: " + sheetName + " have been found. returning the first instance found");
    }
    return myFile;
  }
  
  return null;
  
}

//This function was adapted from the google script service here: https://developers.google.com/apps-script/advanced/url-shortener
function getShortUrl(inputUrl) {
  var shortUrl = UrlShortener.Url.insert({
    longUrl: inputUrl
  });
  return shortUrl.id;
}

