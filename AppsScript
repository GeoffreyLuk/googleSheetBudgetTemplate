const ui = SpreadsheetApp.getUi();
function onOpen() { //only one onOpen exists
  // createSampleMenu() //use as sample
  createMenuWithSubMenu() //use as sample
}

// function createSampleMenu(){
// ui.createMenu('Custom Menu')
//       .addItem('First item', 'menuItem1')
//       .addSeparator()
//       .addSubMenu(ui.createMenu('Sub-menu')
//           .addItem('Second item', 'menuItem2'))
//       .addToUi();
// }

function createMenuWithSubMenu(){
  var subMenu = ui.createMenu("Templates")
    .addItem("Current Month", "budgetTemplateCurrentMonth")
    .addItem("Custom Month", "budgetTemplateMonth");
  
  ui.createMenu("Budget Settings")
    .addSubMenu(subMenu)
    .addSeparator()
    .addItem('Categories','listCategories')
    .addToUi();
}


// everything below is for setting logic

let app = SpreadsheetApp;
let spreadSheet = app.getActiveSpreadsheet();
let currentSheet = spreadSheet.getActiveSheet();
let targetMonth = 3;
let targetYear = 2023
const date = new Date();
const currentYear = date.getFullYear();
const currentMonth = date.getMonth() + 1; // 👈️ months are 0-based
function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}


let daysInTargetMonth = parseInt(getDaysInMonth(targetYear, targetMonth));

const monthConverter = {
  1: "Jan",
  2: "Feb",
  3: "Mar",
  4: "Apr",
  5: "May",
  6: "Jun",
  7: "Jul",
  8: "Aug",
  9: "Sep",
  10: "Oct",
  11: "Nov",
  12: "Dec"
}

const categories = {
  transportation: ["Transporation"],
  food: ["Food"],
  entertainment: ["Date", "Hangouts"],
  miscellaneous: ["Body care", "Household", "Groceries", "Haircut", "Electronics", "Clothes"],
  work: ["Coding"],
  adhoc: ["Health"],
  utilities: ["Grammarly", "Hong Kong Free Press", "Impact HK", "Google", "Spotify"]
}

const startingPosition = {
  row: 2,
  column: 2
}

const saturdays = []; //need to -1 for current month in the saturdays.push
function saturdaysfinder(){
for (let i = 1;i<=daysInTargetMonth;i++){
  if(new Date(targetYear,targetMonth-1,i).getDay() == 6){
    saturdays.push(i)
  }
}
}

function convertToLetters(number,string = ""){
  remainingNum = 0
  let numberDic = {
    1:'A',
    2:'B',
    3:'C',
    4:'D',
    5:'E',
    6:'F',
    7:'G',
    8:'H',
    9:'I',
    10:'J',
    11:'K',
    12:'L',
    13:'M',
    14:'N',
    15:'O',
    16:'P',
    17:'Q',
    18:'R',
    19:'S',
    20:'T',
    21:'U',
    22:'V',
    23:'W',
    24:'X',
    25:'Y',
    26:'Z',
  }
  if (number > 26){
    remainingNum = number-26
    string += 'A'
    return convertToLetters(remainingNum,string)
  }else{
    string += numberDic[number]
    return string
  }
}

const lengthOfCategories = lengthCalculator(categories)
function lengthCalculator(categories){
  let catLength = 0
  for(keys in categories){
    catLength += categories[keys].length
  }
  return catLength
}

function settingRange(row, startingP, numberofdays, month) {
  let input;
  let firstRow = startingPosition.row +1;
  let lastRow = startingPosition.row + lengthOfCategories
  let startLetter = startingPosition.column+1;
  for (let i = 0; i < numberofdays + 2; i++) {
    if (i < numberofdays) {
      currentSheet.getRange(row+lengthOfCategories+1, startingP + i).setValue(`=SUM(${convertToLetters(startLetter+i)}${firstRow}:${convertToLetters(startLetter+i)}${lastRow})`)
      input = String(i + 1 + " " + monthConverter[month])
    } else { input = "" }
    currentSheet.getRange(row, startingP + i).setValue(input)
  }
  currentSheet.getRange(row+lengthOfCategories+1, startingPosition.column+daysInTargetMonth+1).setValue(`=SUM(${convertToLetters(startingPosition.column+daysInTargetMonth+1)}${firstRow}:${convertToLetters(startingPosition.column+daysInTargetMonth+1)}${lastRow})`)
}

function settingCategories() {
  let i = 1
  let ii = 1
  let iii;
  for (keys in categories) {
    iii = categories[keys].length
    for (elem of categories[keys]) {
      let startLetter = convertToLetters(startingPosition.column+1);
      let currentRow = startingPosition.row +i;
      let endLetter = convertToLetters(startingPosition.column+daysInTargetMonth)
      currentSheet.getRange(startingPosition.row + i, startingPosition.column).setValue(elem)
      currentSheet.getRange(startingPosition.row+i,startingPosition.column+daysInTargetMonth+1).setValue(`=SUM(${startLetter}${currentRow}:${endLetter}${currentRow})`)
      i++
    }
    function titleCase(str) {
      return str.toLowerCase().split(' ').map(function (word) {
        return (word.charAt(0).toUpperCase() + word.slice(1));
      }).join(' ');
    }

    let properKeys = titleCase(String(keys));
    currentSheet.getRange(startingPosition.row + ii, startingPosition.column - 1).setValue(properKeys)
    currentSheet.getRange(startingPosition.row + ii, startingPosition.column - 1, iii, 1).mergeVertically()
    ii += iii
  }
  currentSheet.getRange(startingPosition.row + i, startingPosition.column).setValue('Total')
  currentSheet.getRange(startingPosition.row + i+1, startingPosition.column).setValue('Subtotal')
}

function bordersForWeek(length=1){
  let i = 0
  let firstDay = startingPosition.column+1
  let weekend = saturdays[0];
  let subtotalRow = startingPosition.row+lengthOfCategories+2
  let lastSaturday = saturdays[saturdays.length-1]
  while(i<saturdays.length){
    let endOfSubtotal = firstDay+weekend-1
    currentSheet.getRange(subtotalRow,firstDay).setValue(`=SUM(${convertToLetters(firstDay)}${subtotalRow-1}:${convertToLetters(endOfSubtotal)}${subtotalRow-1})`)
    currentSheet.getRange(startingPosition.row,firstDay,length+3,weekend).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
    currentSheet.getRange(startingPosition.row+1,firstDay,length+1,weekend).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
    firstDay += weekend
    i++;
    weekend = saturdays[i] - saturdays[i-1]
  }
  currentSheet.getRange(startingPosition.row,firstDay,length+3,daysInTargetMonth-lastSaturday).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
  currentSheet.getRange(startingPosition.row+1,firstDay,length+1,daysInTargetMonth-lastSaturday).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
  currentSheet.getRange(subtotalRow,firstDay).setValue(`=SUM(${convertToLetters(firstDay)}${subtotalRow-1}:${convertToLetters(daysInTargetMonth+2)}${subtotalRow-1})`)
}

function bordersForCategories(){
  catRow = startingPosition.row+1
  for (key in categories){
    let catLength = categories[key].length
    currentSheet.getRange(catRow,startingPosition.column-1,catLength,daysInTargetMonth+3).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
    currentSheet.getRange(catRow,startingPosition.column-1,catLength,1).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.null);
    catRow += catLength
  }
}

function shoveToString(){
  let message = ''
  for (keys in categories){
    message += `\n\nTopic: ${keys}: \nSubtopic: `
    for(elem of categories[keys]){
      message += `${elem}, `
    }
  }
  return message
}

//ordering menu logic

function budgetTemplateCurrentMonth(){
targetMonth = parseInt(currentMonth)
targetYear = parseInt(currentYear)

daysInTargetMonth = parseInt(getDaysInMonth(targetYear, targetMonth))

saturdaysfinder()
settingRange(startingPosition.row, startingPosition.column + 1, daysInTargetMonth, targetMonth)
settingCategories()
bordersForWeek(lengthOfCategories)
bordersForCategories()
currentSheet.autoResizeColumns(1,daysInTargetMonth+2)
}

function budgetTemplateMonth(){
targetMonth = parseInt(ui.prompt('Which Month (number only): MM').getResponseText());
targetYear = parseInt(ui.prompt('Which Year (number only): YYYY').getResponseText());

daysInTargetMonth = parseInt(getDaysInMonth(targetYear, targetMonth))
saturdaysfinder()
settingRange(startingPosition.row, startingPosition.column + 1, daysInTargetMonth, targetMonth)
settingCategories()
bordersForWeek(lengthOfCategories)
bordersForCategories()
currentSheet.autoResizeColumns(1,daysInTargetMonth+2)
}

function listCategories(){
  ui.prompt(shoveToString())
}
