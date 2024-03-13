function onOpen() 
{
  var ui = null;
  ui = SpreadsheetApp.getUi();
  ui.createMenu('Financial Handler')
    .addItem("TAccount", "TAccount")
    .addToUi();
}

function TAccount() 
{
  var selectedRanges = null;
  var newRanges = null;
  var confirmToReplace = false;
  var ok = false ;
  var rowshould = 0 ;
  var colshould = 0 ;

  rowshould = 2;
  colshould = 5;

  selectedRanges = SpreadsheetApp.getActiveSheet().getSelection().getActiveRange();

  ok = (selectedRanges.getNumColumns() == colshould)
   && (selectedRanges.getNumRows() > rowshould)

  if (ok == true) 
  {
    newRanges = selectedRanges.offset(0,selectedRanges.getNumColumns()+1,1000);
    confirmToReplace = ConfirmToReplace(newRanges);

    if (confirmToReplace == true) 
    {
      newRanges.clear();
      SpreadsheetApp.flush();
    }
    TAccountUtil(selectedRanges,[newRanges.getRow(),newRanges.getColumn(),newRanges.getNumRows(),newRanges.getNumColumns()]);   
  }
}

function TAccountUtil(ranges) 
{
  var retVal1 = null;
  var retVal2 = null;
  retVal1 = SpiltRangesByCols(ranges,[1,2,6]);
  retVal2 = SplitEqualRanges(retVal1[1],[ranges.getNumRows(),2]);
  TAccountUtil2(retVal2[0][0],retVal2[0][1],retVal1[0]);
}

function TAccountUtil2(ranges1,ranges2)
{
  var rowlen = 0 ;
  var collen = 0;
  var isEmptyItem = false;
  var currentItem = null;
  var currentValue = 0;
  var currentDate = new Date();
  var ranges1_rowidx,ranges1_colidx,ranges3_colPos,ranges3_rowPos;
  var ranges3Info = {};
  var ranges3Infos = [] ;
  var itemsinT = [];
  var prevDate;
  var rangesDate;
  var ranges1Values;
  var ranges2Values;
  var rangesDateValues;

  ranges1Values = ranges1.getValues();
  ranges2Values = ranges2.getValues();
  
  rangesDate = ranges1.offset(0,-1,ranges1.getNumRows(),1);
  rangesDateValues = rangesDate.getValues();
  
  rowlen = ranges1.getNumRows();
  collen = ranges1.getNumColumns();

  Object.assign(itemsinT,{});
  Object.assign(ranges3Info,{});
  
  ranges3Infos = [];

  ranges3_colPos = ranges2.getColumn() + 1 + 1;
  ranges3_rowPos = ranges2.getRow();

  for(ranges1_rowidx = 0 ; ranges1_rowidx<=rowlen-1;ranges1_rowidx++)
  {
    for(isEmptyItem = true,ranges1_colidx = 0 ; ranges1_colidx <= collen-1; ranges1_colidx ++)
    {
      currentItem = ranges1Values[ranges1_rowidx][ranges1_colidx];
      currentValue = ranges2Values[ranges1_rowidx][ranges1_colidx];
      prevDate = rangesDateValues[ranges1_rowidx][0];
      prevDate = Date.parse(prevDate);
      if(!(prevDate == null || prevDate == undefined || isNaN(prevDate)))
      {
        currentDate = new Date(prevDate);
      }
      if(!(currentItem == null || currentItem == undefined || currentItem == ""))
      {
        isEmptyItem = false;
      }
      if(isEmptyItem != true)
      {
        break;
      }
    }

    if(isEmptyItem == true)
    {
      throw Error("ERROR!!! The values of one or more row among accounting items are both null or undefined. At ranges1_rowidx:"+ranges1_rowidx);
    }
    if(typeof(currentItem)!="string")
    {
      throw Error("ERROR!!! The value among accounting items is not string type. At ranges1_rowidx:"+ranges1_rowidx);
    }

    if((itemsinT[currentItem] == null || itemsinT[currentItem] == undefined))
    {
      ranges3Info = Object.create({});
      Object.defineProperty(ranges3Info,'科目',{value:currentItem,writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'借方金額筆數',{value:0,writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'借方日期陣列',{value:[],writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'借方金額陣列',{value:[],writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'貸方金額筆數',{value:0,writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'貸方日期陣列',{value:[],writable:true,enumerable:true});
      Object.defineProperty(ranges3Info,'貸方金額陣列',{value:[],writable:true,enumerable:true});
      ranges3Infos.push(ranges3Info);
      itemsinT.push(currentItem);
    }

    if(ranges1_colidx == 0)
    {
      ranges3Info['借方金額筆數'] += 1;
      ranges3Info['借方日期陣列'].push(currentDate);
      ranges3Info['借方金額陣列'].push(currentValue);
    }

    if(ranges1_colidx == 1)
    {
      ranges3Info['貸方金額筆數'] += 1;
      ranges3Info['貸方日期陣列'].push(currentDate);
      ranges3Info['貸方金額陣列'].push(currentValue);
    }
  }
  WriteDataToSheet([ranges3Infos,ranges3_rowPos,ranges3_colPos + 1]);
}

function AskToConfirm(title,text)
{
  var ui = null;
  var result = null;
  ui = SpreadsheetApp.getUi();
  result = ui.alert(title,text,ui.ButtonSet.YES_NO);
  return (result=="YES") ;
}

function ConfirmToReplace(newRanges) 
{
  var hasValues = false;
  var result = false;

  hasValues = ( newRanges.isBlank() == true ? false : true ) ; 
  result = false;
  if(hasValues == true)
  {
    result = AskToConfirm('Replace Confirm','Has value in the specified grid of cells.'+'Do you want to replace these values in the grid of cells?'+'ranges:('+newRanges.getRow()+','+newRanges.getColumn()+','+newRanges.getEndRow()+','+newRanges.getEndColumn()+')');
  }
  return result;
}

function WriteDataToSheet([ranges3Infos,ranges3_rowPos,ranges3_colPos])
{
  var ui = null;
  var sheet = null;
  var rowStartPos = 0;
  var colStartPos = 0;
  var currentRowPos = 0 ;
  var currentColPos = 0;
  var infolen = 0 ;
  var ranges3Info = {};
  var item = "";
  var num_record1 = 0 ;
  var num_record2 = 0;
  var record1 = [] ;
  var record2 = [];
  var max_num_record = 0 ;
  var record1_date;
  var record2_date;
  var array2D=[];

  ui = SpreadsheetApp.getUi();
  sheet = SpreadsheetApp.getActiveSheet();
  infolen = ranges3Infos.length;

  rowStartPos = ranges3_rowPos;
  colStartPos = ranges3_colPos;

  currentRowPos = rowStartPos ; 
  currentColPos = colStartPos ;

  try
  {
    for(var infoidx = 0;infoidx<=infolen-1;infoidx++)
    {
      ranges3Info = Object.create(ranges3Infos[infoidx]);    
      item = ranges3Info['科目'];
      num_record1 = ranges3Info['借方金額筆數'];
      record1 = ranges3Info['借方金額陣列']
      record1_date = ranges3Info['借方日期陣列'];

      num_record2 = ranges3Info['貸方金額筆數'];
      record2 = ranges3Info['貸方金額陣列']
      record2_date = ranges3Info['貸方日期陣列'];

      max_num_record = Math.max(num_record1,num_record2);

      currentColPos = colStartPos;
      sheet.getRange(currentRowPos,currentColPos,1,2).setValues([['科目',item]]);

      currentRowPos += 1;

      if(num_record1 >= 1)
      { 
        currentColPos = colStartPos;
        array2D = [record1_date.concat(record1)];
        sheet.getRange(currentRowPos,currentColPos,num_record1,2).setValues(array2D);
      }

      if(num_record2 >= 1)
      {
        currentColPos =  colStartPos  + 2 ; 
        array2D = [record2_date.concat(record2)];
        sheet.getRange(currentRowPos,currentColPos,num_record2,2).setValues(array2D);
      }
      currentRowPos += (max_num_record + 1);

      //flush the app about Google SpreadSheet.
      SpreadsheetApp.flush();
    }
  }catch(err)
  {
    throw Error("Error occurs at infoidx:"+infoidx+",detailed message:\n"+err.message);
  }
}

function SpiltRangesByCols(ranges,parts)
{
  var ui = null;
  var sheet = null;
  var ranges_array1D = [];

  var nextColIdx = 0;

  ui = SpreadsheetApp.getUi();
  sheet = SpreadsheetApp.getActiveSheet();

  try 
  {
    ranges_array1D = [];
    for(var colidx = 0 ;colidx<= parts.length - 2 ;colidx+=1)
    {
      nextColIdx = colidx + 1 ; 
      ranges_temp = sheet.getRange(ranges.getRow(),parts[colidx],ranges.getNumRows(),parts[nextColIdx]-parts[colidx] );
      ranges_array1D.push(ranges_temp);
    }
    return ranges_array1D;  
  }
  catch(err)
  {
    throw Error("There are an error at colidx:"+colidx+",detailed message:"+"\n"+err.message);
  }
}

function SplitEqualRanges(ranges,parts)
{
  var sheet = null;
  var ranges_array1D = [];
  var ranges_array2D = [];

  sheet = SpreadsheetApp.getActiveSheet();

  ranges_array2D = [];
  for (var rowidx= ranges.getRow();rowidx<= ranges.getEndRow() ; rowidx+= parts[0])
  {
    ranges_array1D = [];
    for(var colidx = ranges.getColumn() ;colidx<= ranges.getEndColumn() ;colidx+=parts[1])
    {
      ranges_temp = sheet.getRange(rowidx,colidx,parts[0],parts[1]);
      ranges_array1D.push(ranges_temp);
    }
    ranges_array2D.push(ranges_array1D);
  }
  return ranges_array2D;
}
