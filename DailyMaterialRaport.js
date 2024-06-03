function CheckMaterialsBelowMinimumLevel() {
  //Get actual date
  var data = new Date();
  var miesiace = ["styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień"];
  var aktualnyMiesiac = miesiace[data.getMonth()];
  //Initiate Sheets and Ranges
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(aktualnyMiesiac + " 2024");
  var rangeData = dataSheet.getRange(1, 1,500,10);
  var limitSheet = ss.getSheetByName("MaterialLimits");
  var rangeLimit = limitSheet.getRange(1,1,16,2);

  //Do not execute in weekends
  if(data.getDay() == 0 || data.getDay() == 6)
  {
    return;
  }
  //Validate if choosed date exist in sheet
  var dataWeek = Utilities.formatDate(data, "GMT", "w"); 
  var emptyText = true;
  var mailText = "Proper day not found"

  //Set MaterialsLimit values
  var PA12Limit = rangeLimit.getCell(2,2).getValue();
  var PA11Limit = rangeLimit.getCell(3,2).getValue();
  var FlexaSoftLimit = rangeLimit.getCell(4,2).getValue();
  var FlexaGreyLimit = rangeLimit.getCell(5,2).getValue();
  var FlexaBrightLimit = rangeLimit.getCell(6,2).getValue();
  var FlexaBlackLimit = rangeLimit.getCell(7,2).getValue();
  var TPELimit = rangeLimit.getCell(8,2).getValue();
  var SealerLimit = rangeLimit.getCell(9,2).getValue();
  var PA11ESDLimit = rangeLimit.getCell(10,2).getValue();
  var PPLimit = rangeLimit.getCell(11,2).getValue();
  var PA11CFLimit = rangeLimit.getCell(12,2).getValue();
  var ScierniwoLimit = rangeLimit.getCell(13,2).getValue();
  var PA12IndustrialLimit = rangeLimit.getCell(14,2).getValue();
  var FlexaPerformanceLimit = rangeLimit.getCell(15,2).getValue();
  var PBTLimit = rangeLimit.getCell(16,2).getValue();

  //Set material values cells offsets
  var PA12Offset = 4;
  var PA11Offset = 11;
  var FlexaSoftOffset = 16;
  var FlexaGreyOffset = 18;
  var FlexaBrightOffset = 21;
  var FlexaBlackOffset = 23;
  var TPEOffset = 26;
  var SealerOffset = 28;
  var PA11ESDOffset = 29;
  var PPOffset = 31;
  var PA11CFOffset = 33;
  var ScierniwoOffset = 35;
  var PA12IndustrialOffset = 37;
  var FlexaPerformanceOffset = 39;
  var PBTOffset = 41;

  //Finde today's raport
  for(var i = 1; i < 500; i = i + 45)
  {
    if(rangeData.getCell(i,2).getValue() == dataWeek)
    {
      mailText = "Material level raport: <br />\n<br />\n"
      for(var j = 0; j < 5; j++)
      {
        var today = new Date(rangeData.getCell(i + 2,2 + j).getValue());
        //Compare limit values with proper date
        if(today.getUTCDate() + 1 == data.getUTCDate())
        {
          //Check Pa12
          var PowderOnStock = rangeData.getCell(i + PA12Offset, 2 + j).getValue();
          if(PowderOnStock < PA12Limit)
          {
            
            mailText += "PA12 below Limit: " + PowderOnStock + ", limits is " + PA12Limit +"<br />\n"
            emptyText = false;
          }
          //Check Pa11
          PowderOnStock = rangeData.getCell(i + PA11Offset, 2 + j).getValue();
          if(PowderOnStock < PA11Limit)
          {
            mailText += "PA11 below Limit: " + PowderOnStock + ", limits is " + PA11Limit +"<br />\n"
            emptyText = false;
          }
          //Check FlexaSoft
          PowderOnStock = rangeData.getCell(i + FlexaSoftOffset, 2 + j).getValue();
          if(PowderOnStock < FlexaSoftLimit)
          {
            mailText += "FlexaSoft below Limit: " + PowderOnStock + ", limits is " + FlexaSoftLimit +"<br />\n"
            emptyText = false;
          }
          //Check FlexaGrey
          PowderOnStock = rangeData.getCell(i + FlexaGreyOffset, 2 + j).getValue();
          if(PowderOnStock < FlexaGreyLimit)
          {
            mailText += "FlexaGrey below Limit: " + PowderOnStock + ", limits is " + FlexaGreyLimit +"<br />\n"
            emptyText = false;
          }
          //Check FlexaBright
          PowderOnStock = rangeData.getCell(i + FlexaBrightOffset, 2 + j).getValue();
          if(PowderOnStock < FlexaBrightLimit)
          {
            mailText += "FlexaBright below Limit: " + PowderOnStock + ", limits is " + FlexaBrightLimit +"<br />\n"
            emptyText = false;
          }
          //Check FlexaBlack
          PowderOnStock = rangeData.getCell(i + FlexaBlackOffset, 2 + j).getValue();
          if(PowderOnStock < FlexaBlackLimit)
          {
            mailText += "FlexaBlack below Limit: " + PowderOnStock + ", limits is " + FlexaBlackLimit +"<br />\n"
            emptyText = false;
          }
          //Check TPE
          PowderOnStock = rangeData.getCell(i + TPEOffset, 2 + j).getValue();
          if(PowderOnStock < TPELimit)
          {
            mailText += "TPE below Limit: " + PowderOnStock + ", limits is " + TPELimit +"<br />\n"
            emptyText = false;
          }
          //Check Sealer
          PowderOnStock = rangeData.getCell(i + SealerOffset, 2 + j).getValue();
          if(PowderOnStock < SealerLimit)
          {
            mailText += "Sealer below Limit: " + PowderOnStock + ", limits is " + SealerLimit +"<br />\n"
            emptyText = false;
          }
          //Check PA11ESD
          PowderOnStock = rangeData.getCell(i + PA11ESDOffset, 2 + j).getValue();
          if(PowderOnStock < PA11ESDLimit)
          {
            mailText += "PA11ESD below Limit: " + PowderOnStock + ", limits is " + PA11ESDLimit +"<br />\n"
            emptyText = false;
          }
          //Check PP
          PowderOnStock = rangeData.getCell(i + PPOffset, 2 + j).getValue();
          if(PowderOnStock < PPLimit)
          {
            mailText += "PP below Limit: " + PowderOnStock + ", limits is " + PPLimit +"<br />\n"
            emptyText = false;
          }
          //Check PA11CF
          PowderOnStock = rangeData.getCell(i + PA11CFOffset, 2 + j).getValue();
          if(PowderOnStock < PA11CFLimit)
          {
            mailText += "PA11CF below Limit: " + PowderOnStock + ", limits is " + PA11CFLimit +"<br />\n"
            emptyText = false;
          }
          //Check Scierniwo
          PowderOnStock = rangeData.getCell(i + ScierniwoOffset, 2 + j).getValue();
          if(PowderOnStock < ScierniwoLimit)
          {
            mailText += "Scierniwo below Limit: " + PowderOnStock + ", limits is " + ScierniwoLimit +"<br />\n"
            emptyText = false;
          }
          //Check PA12Industrial
          PowderOnStock = rangeData.getCell(i + PA12IndustrialOffset, 2 + j).getValue();
          if(PowderOnStock < PA12IndustrialLimit)
          {
            mailText += "PA12Industrial below Limit: " + PowderOnStock + ", limits is " + PA12IndustrialLimit +"<br />\n"
            emptyText = false;
          }
          //Check FlexaPerformance
          PowderOnStock = rangeData.getCell(i + FlexaPerformanceOffset, 2 + j).getValue();
          if(PowderOnStock < FlexaPerformanceLimit)
          {
            mailText += "FlexaPerformance below Limit: " + PowderOnStock + ", limits is " + FlexaPerformanceLimit +"<br />\n"
            emptyText = false;
          }
          //Check PBT
          PowderOnStock = rangeData.getCell(i + PBTOffset, 2 + j).getValue();
          if(PowderOnStock < PBTLimit)
          {
            mailText += "PBT below Limit: " + PowderOnStock + ", limits is " + PBTLimit +"<br />\n"
            emptyText = false;
          }
        }
      }
    }
  }
  //Set empty mail text
  if(emptyText)
  {
    mailText += "All materials above minimum level."
  }


  //Send email with raport to proper addressees
  MailApp.sendEmail({
    to: "warhaouseCompany@gmail.com, johnDoeCompany@gmail.com",
    subject: "Material level raport - " + data.toLocaleDateString("en-US").toString(),
    htmlBody: mailText
  });
}
