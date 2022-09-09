/*
3:59:15 PM	Info	#0, First Name
3:59:15 PM	Info	#1, Last Name
3:59:15 PM	Info	#2, Move in Date
3:59:15 PM	Info	#3, Floor Plan
3:59:15 PM	Info	#4, If your floor plan choice is not available, what would your second choice be? 
3:59:15 PM	Info	#5, Do you have a location preference?
3:59:15 PM	Info	#6, Do you have specific roommates you want to live with?
3:59:15 PM	Info	#7, What is your gender?
3:59:15 PM	Info	#8, What gender of roommate(s) do you want to live with?
3:59:15 PM	Info	#9, Do you have any allergies?
3:59:15 PM	Info	#10, What are your hobbies?
3:59:16 PM	Info	#11, Special considerations or preferences?
3:59:16 PM	Info	#12, What is your major/area of study?
3:59:16 PM	Info	#13, What year of school are you going into/in?
3:59:16 PM	Info	#14, What years do you prefer to live with?
3:59:16 PM	Info	#15, My Study habits are...:
3:59:16 PM	Info	#16, Do you consider yourself quiet?
3:59:16 PM	Info	#17, Do you consider yourself neat?
3:59:16 PM	Info	#18, When do you go to sleep?
3:59:16 PM	Info	#19, How often do you entertain at your apartment?
3:59:16 PM	Info	#20, Are you comfortable living with an animal?
3:59:17 PM	Info	#21, Do you drink?
3:59:17 PM	Info	#22, Do you smoke?
*/
var questionName = {
  0 : "firstName",
  1 : "lastName" ,
  2 : "moveIn",
  3 : "floorPlan",
  4 : "secFloor",
  5 : "locPref",
  6 : "roommates",
  7 : "gender",
  8 : "unitGender",
  9 : "allergies",
  10: "hobbies",
  11: "specConsid",
  12: "major",
  13: "schoolYear",
  14: "unitYear",
  15: "studyHabits",
  16: "noiseLevel",
  17: "clean",
  18: "bedTime",
  19: "enterFreq",
  20: "animalPref",
  21: "drink",
  22: "smoke"
};
//REPLACE THIS ID WITH THE FORM FOR THE NEW YEAR
function onSubmit(){
  main(-1);
}

function selectedReplace(){
  main(51);
}

function main(num){
  var form = FormApp.openById(ScriptProperties.getProperty("formID"));
  var folder = DriveApp.getFolderById(ScriptProperties.getProperty("folderID"));
  var template = DriveApp.getFileById(ScriptProperties.getProperty("templateID"));
  var [responseData,done] = generate_Preferences(form.getResponses(), num);
  create_pdf(responseData, folder, template);
  if(!done){
    num++;
    //main(num);
  }
}

function generate_Preferences(responses, num) {
  var done = false
  var retData = {};
  if(num == -1){
    response = responses[responses.length -1];
    done = true;
  }else{
    if(num == responses.length-1){
      done = true;
    }
    response = responses[num];
  }
  var itemResponses = response.getItemResponses();
  
  retData["email"] = response.getRespondentEmail().toLowerCase();
  retData["timeStamp"] = response.getTimestamp();
  for(var i =0; i < itemResponses.length; i++){
    var response = itemResponses[i].getResponse();
    var repltxt = questionName[i];
    if(repltxt == "firstName" || repltxt == "lastName"){
      response = response.trim();
      response = response.toLowerCase();
      response = response.charAt(0).toUpperCase() + response.slice(1);
    }
    retData[repltxt] = response;
  }
  return [retData, done];
}

function create_pdf(retData, folder, template){
  var fileName = retData["firstName"]+ "_"+retData["lastName"] + "_PreferenceForm";
  var dupes = DriveApp.getFilesByName(fileName);
  while(dupes.hasNext()){
    var dupe = dupes.next();
    console.log("Deleted file %s", fileName);
    dupe.setTrashed(true);
  }
  var tempFolder = DriveApp.getFolderById(ScriptProperties.getProperty("tempFolderID"));
  var tempFile = template.makeCopy(tempFolder);
  var OpenDoc = DocumentApp.openById(tempFile.getId());
  var body = OpenDoc.getBody();
  var footer = OpenDoc.getFooter();
  for(const [key,value] of Object.entries(retData)){
    body.replaceText("{{"+key+"}}", value);
    footer.replaceText("{{"+key+"}}", value);
  }
  OpenDoc.saveAndClose();
  
  const blopPDF = tempFile.getAs(MimeType.PDF);
  var pdfFile = folder.createFile(blopPDF).setName(fileName);
  tempFile.setTrashed(true);
  console.log("Created file %s", fileName);
}
