function main(workbook: ExcelScript.Workbook) {
  
  let selectedSheet = workbook.getActiveWorksheet();

  console.log(selectedSheet.getName());

  if (selectedSheet.getName().substr(-12) == "_karta_pracy") {
    console.log("uruchamiaj ten skrypt /będąc/ na FCP a nie na karcie pracy.");
  } else {

    var row_start = find_first_row(workbook);
    var last_day_string = selectedSheet.getRange("A" + row_start).getText();
    var day_end_string = last_day_string.substr(0, 2);
    var month_end_string = last_day_string.substr(3, 2);
    var year_end_string = last_day_string.substr(6, 4);

    var day_end_int: number = +day_end_string;
    var month_end_int: number = +month_end_string;
    var year_day_int: number = +year_end_string;

    let range = selectedSheet.getRange("A" + row_start + ":D" + (row_start + 113)).getValues(); // 300

    var days_hours = [];
    for (var i = 0; i < range.length; i++) {
      days_hours.push(range[i]);
    }

    generate_KP(workbook);
    put_titles(workbook);
    put_data(workbook, days_hours);

  }

}



function put_data(workbook: ExcelScript.Workbook, days_hours) {

// console.log(days_hours);
// 44166 to jest wartosc z pliku
let date4 = new Date( (days_hours[0][0]-25567-2) * 1000 * 60 * 60 * 24 );

//console.log(date4);

var last_day_of_month = date4.getDate();

var ile_odjac = days_hours[0][0] - last_day_of_month;
// console.log("ile_odjac", ile_odjac);

var current = days_hours[0][0];
var date = [];
var day = [];
var start = [];
var end = [];
var sum = [];

for (var i = 0; i < days_hours.length; i++) {
  days_hours[i][4] = days_hours[i][0] - ile_odjac;
}


for (var i = 0; i <= last_day_of_month; i++) {
  day[i] = days_hours[0][0] - last_day_of_month + i +1;
	start[i] = 1;
	end[i] = 0;
	sum[i] = 0;
}


for (var i = 0; i < days_hours.length; i++) {
  
  if (days_hours[i][1] <= start[days_hours[i][4]]) {
    start[days_hours[i][4]] = days_hours[i][1];
  }  

  if (days_hours[i][2] >= end[days_hours[i][4]]) {
    end[days_hours[i][4]] = days_hours[i][2];
  }  

  if (days_hours[i][3] > 0) {
    sum[days_hours[i][4]] = sum[days_hours[i][4]] + days_hours[i][3];
  }  
  
}

  let selectedSheet = workbook.getActiveWorksheet();

  selectedSheet.getRange("A12:A41").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  selectedSheet.getRange("A12:A41").getFormat().setIndentLevel(0);
  selectedSheet.getRange("C11:F44").setNumberFormatLocal("g:mm;@");
  selectedSheet.getRange("A11:A44").setNumberFormatLocal("rrrr-mm-dd;@");
  selectedSheet.getRange("A1").getFormat().setColumnWidth(71.25);
  selectedSheet.getRange("B1").getFormat().setColumnWidth(90);

  for (i=1; i < start.length; i++) {
    date4 = new Date((day[i - 1] - 25567 - 2) * 1000 * 60 * 60 * 24);
    if ((date4.getDay() == 0) || (date4.getDay() == 6)) {
    selectedSheet.getRange("A" + (i + 11).toString() + ":F" + (i + 11).toString())
      .getFormat()
      .getFill()
      .setColor("FFF2CC");
  }
  
  selectedSheet.getRange("B" + (i + 11).toString()).setValue(return_name_of_day(date4));
  selectedSheet.getRange("A" + (i + 11).toString()).setValue(day[i-1]);
  selectedSheet.getRange("C"+(i+11).toString()).setValue(start[i]);
  selectedSheet.getRange("D"+(i+11).toString()).setValue(end[i]);
  selectedSheet.getRange("E"+(i+11).toString()).setValue(sum[i]);
  if ((end[i] - start[i] - sum[i]) < 0.00001) {
    selectedSheet.getRange("F" + (i + 11).toString()).setValue(0);
  }
  else {
    selectedSheet.getRange("F" + (i + 11).toString()).setValue(end[i] - start[i] - sum[i]);
  }
}

}



function generate_KP(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  var current_sheet_name = selectedSheet.getName()
  
  var delete_worksheet_id = -1;
  
  for (let worksheet of workbook.getWorksheets()) {
  
    if (current_sheet_name + '_karta_pracy' == worksheet.getName()) {
      delete_worksheet_id = 1;
      worksheet.activate();
      worksheet.delete();
    }
  }
  
  var a = workbook.addWorksheet(current_sheet_name + '_karta_pracy');
  a.activate();
}

function put_titles(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  selectedSheet.getRange("A1").setValue("Karta czasu pracy wygenerowana (Nie edytować bo i tak zostanie nadpisana)");
  selectedSheet.getRange("A3").setValue("Karta czasu pracy");
  selectedSheet.getRange("A4").setValue("Miesiąc");
  selectedSheet.getRange("A5").setValue("Imię i nazwisko");
  selectedSheet.getRange("A6").setValue("Process");
  selectedSheet.getRange("A7").setValue("Stanowisko");

  
  selectedSheet.getRange("A10").setValue("Dzień");
  selectedSheet.getRange("B10").setValue("Dzień tyg.");
  selectedSheet.getRange("C10").setValue("Start");
  selectedSheet.getRange("D10").setValue("Koniec");
  selectedSheet.getRange("E10").setValue("Suma");
  selectedSheet.getRange("F10").setValue("Przerwa");
}



function find_first_row(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  var row_start = -1;

  for (let i = 1; i < 10; i++) {
    var tmp = selectedSheet.getRange("A" + i).getText();
    if ((tmp.toString().length == 10) && (tmp.toString()[2] == ".") && (tmp.toString()[5] == ".")) {
      row_start = i;
      break;
    }
  }
  console.log("dane zaczynają się w wierszu numer = " + row_start);
  return row_start;
}



function return_name_of_day(mydate) {
  var day;
  switch (new Date(mydate).getDay()) {
    case 0:
      day = "Niedziela";
      break;
    case 1:
      day = "Poniedziałek";
      break;
    case 2:
      day = "Wtorek";
      break;
    case 3:
      day = "Środa";
      break;
    case 4:
      day = "Czwartek";
      break;
    case 5:
      day = "Piątek";
      break;
    case 6:
      day = "Sobota";
  }
  return day;
}
