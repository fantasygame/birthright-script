var goldBars = 0;
var regencyPoints = 0;
var maintenanceCost = 0;

var maxRegencyPoints;
var courtCost;
var embassies;
var spyNetwork;
var maxReserve;
var currentReserve;
var currentTreasury;

var militaryCost;

var otherHoldingsGoldBars;
var otherHoldingsRegentPoints;

var totalMaintCost;

function calculate() {
  regentCardData();
  militaryCardData();
  otherHoldingsData();
  sumProvinces();
  updateRegencyPoints();
  updateGoldBars();
  getSheets()[0].getRange("A2").setValue(regencyPoints);
  promptMessage();
}

function promptMessage() {
  s = goldBars.toFixed(2);
  t = totalMaintCost.toFixed(2);
  u = (goldBars - totalMaintCost).toFixed(2);
  smsg = "You have earned " + regencyPoints + " Regency Points and "
  smsg = smsg + s + " Gold Bars. Maintenance Costs equals " + t
  smsg = smsg + " Gold Bars for a net of " + u + " Gold Bars."
  SpreadsheetApp.getUi().alert(smsg);
}

function regentCardData() {
  var regentSheet = getRegentSheet();
  maxRegencyPoints = regentSheet.getRange("C4").getValue();
  courtCost = regentSheet.getRange("D14").getValue();
  embassies = regentSheet.getRange("D17").getValue();
  spyNetwork = regentSheet.getRange("D30").getValue();
  maxReserve = regentSheet.getRange("D8").getValue();
  currentReserve = regentSheet.getRange("D7").getValue();
  currentTreasury = regentSheet.getRange("D11").getValue();
}

function militaryCardData() {
  var militarySheet = getMilitarySheet();
  militaryCost = militarySheet.getRange("L2").getValue();
}

function otherHoldingsData() {
  var holdingsSheet = getOtherHoldingsSheet();
  otherHoldingsGoldBars = holdingsSheet.getRange("G48").getValue();
  otherHoldingsRegentPoints = holdingsSheet.getRange("E48").getValue();
}

function sumProvinces() {
  var sheets = getSheetsOfType("Province");
  for (i = 0; i < sheets.length; i++) {
    regencyPoints += sheets[i].getRange("F5").getValue();
    goldBars += sheets[i].getRange("G5").getValue();
    maintenanceCost += sheets[i].getRange("H37").getValue();
  }
}

function updateRegencyPoints() {
  var regentSheet = getRegentSheet();

  if (regencyPoints > maxRegencyPoints) {
    regencyPoints = maxRegencyPoints;
  }
  updateRegencyEarned(regentSheet, regencyPoints);

  if ((currentReserve + regencyPoints) > maxReserve) {
    currentReserve = maxReserve;
  } else {
    currentReserve = currentReserve + regencyPoints + otherHoldingsRegentPoints;
  }
  updateCurrentReserve(regentSheet, currentReserve);
}

function updateGoldBars() {
  var regentSheet = getRegentSheet();
  updateGoldBarsEarned(regentSheet, goldBars);
  totalMaintCost = maintenanceCost + militaryCost + courtCost + embassies + spyNetwork;
  updateTreasury(regentSheet, (goldBars + currentTreasury + otherHoldingsGoldBars) - totalMaintCost);
  updatetotalMaintCost(regentSheet, totalMaintCost);
}

function updateRegencyEarned(regentSheet, regencyPoints) {
  regentSheet.getRange("D6").setValue(regencyPoints);
}

function updateCurrentReserve(regentSheet, currentReserve) {
  regentSheet.getRange("D7").setValue(currentReserve);
}

function updateGoldBarsEarned(regentSheet, goldBars) {
  regentSheet.getRange("D10").setValue(goldBars);
}

function updateTreasury(regentSheet, treasury) {
  regentSheet.getRange("D11").setValue(treasury);
}

function updatetotalMaintCost(regentSheet, totalMaintCost) {
  regentSheet.getRange("D12").setValue(totalMaintCost);
}

function getSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets();
}

function getSheetsOfType(type) {
  var sheetsOfType = [];
  var sheets = getSheets();
  for (i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheetType(sheet) === type) {
      sheetsOfType.push(sheet);
    }
  }
  return sheetsOfType;
}

function sheetType(sheet) {
  return firstCellValue = sheet.getRange("A1").getValue();
}

function getRegentSheet() {
  return getSheetsOfType("Regent")[0];
}

function getMilitarySheet() {
  return getSheetsOfType("Military")[0];
}

function getOtherHoldingsSheet() {
  return getSheetsOfType("OtherHoldings")[0];
}

