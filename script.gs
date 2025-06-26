function onOpen()
{
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Functions')
    .addItem('Filtered Sum','filter_sum')
    .addItem('monthly sum','monthly_spending_total')
    .addToUi();
}

function filter_sum() 
{
  var sheet=SpreadsheetApp.getActiveSpreadsheet()
  var data=sheet.getRange("E4:G").getValues()
  var totalSumYes=0
  var taxYessum=0
  var totalSumNo=0
  var taxNosum=0

   for(var i=0;i<data.length;i++)
   {
    var number=data[i][1]
    var tax=data[i][0]
    var yes_or_no=data[i][2]

    if(data[i][2]=="YES")
    {totalSumYes+=number
    taxYessum+=tax}
    if(data[i][2]=="NO")
    {totalSumNo+=number
    taxNosum+=tax}
   }
  sheet.getRange("K4").setValue(totalSumYes)
  sheet.getRange("K3").setValue(taxYessum)
  sheet.getRange("M4").setValue(totalSumNo)
  sheet.getRange("M3").setValue(taxNosum) 
}

function monthly_spending_total()
{
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var raw_data=sheet.getRange("E4:G").getValues()                     //raw data (array of arrays)
  Logger.log(raw_data)
  var monthly_data=[]                                                //initialize 2D array
  var current_row=4                                                 //start row                                 
  var column=5                                                      //always check tax column (can be changed)
  var month_row=6                                                   //start at J2
  var month_column=11                                            //always on column J (10)
  var cell=sheet.getRange(current_row,column)                     //current cell variable
  var raw_data_row=0                                         //individual array from raw data
  var totals_sum=0                                          //total money sum


    while(!cell.isBlank())                                      //loops through the data untill blank cell detected

    {
      totals_sum=0                                                       //everytime its back to main loop, new total
      monthly_data=[]                                                     //clears the data array

      for(var i=0;i<raw_data.length;i++)                    //loops through rows to detect merge cells & this loop runs monthly
      {
        var new_row=raw_data[raw_data_row]                 //the row from raw data is now the new row for monthly data
        cell=sheet.getRange(current_row,column)            //current cell variable used for checking merge

        if (!cell.isPartOfMerge() && !cell.isBlank())       // Only run code if cell is not merged nor blank
        {
          monthly_data.push([new_row[0],new_row[1],new_row[2]])    //add first data row values to monthly data. numbers 0,1,2
          //                                                  correspond to each column entry from the row (tax,total,YES/NO) 
          
          raw_data_row+=1                                            //makes sure the array from raw data is the next one
          totals_sum+=monthly_data[i][1]                                 //add each "total" entry to the total sum
          current_row++                                                 //current row changes once code is run
        }

        else
        {
          raw_data_row+=1                                     //counter for current row from data array
          current_row++                                       //current row changes once code is run
          break                                              //once blank or merged cell is detected it returns it to main loop
        }
      }
      
      cell=sheet.getRange(current_row,column)                      //before checking main loop conditional, uptade current cell 
      sheet.getRange(month_row,month_column).setValue(totals_sum)    //place the total
      month_row+=1                                                   //place total on the next row
      
    }
}
