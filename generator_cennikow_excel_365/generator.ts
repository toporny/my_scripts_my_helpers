
function main(workbook: ExcelScript.Workbook) {

  if (does_generator_sheet_exist(workbook)) {


    let selectedSheet = workbook.getActiveWorksheet();
    // Clear ExcelScript.ClearApplyTo.contents from range generator danych do cennika PDF!A11:N30
    selectedSheet.getRange("A11:N99")
      .clear(ExcelScript.ClearApplyTo.contents);
    // Clear ExcelScript.ClearApplyTo.contents from range generator danych do cennika PDF!A6
    selectedSheet.getRange("A6")
      .clear(ExcelScript.ClearApplyTo.contents);


    let heater_name = selectedSheet.getRange("A2").getValue().toString();
    let heater_kind = selectedSheet.getRange("A3").getValue().toString();

    if ((heater_name == "") || (heater_kind == "")) {
      console.log("podaj heater_name oraz heater_kind w polu A2 i A3");
    } else {
      console.log("heater_name: " + heater_name);
      console.log("heater_kind: " + heater_kind);
      switch (heater_kind) {
        case "zwykły":
          let array_with_prices = get_array_with_prices(heater_name, workbook);
          //console.log(array_with_prices)
          //console.log('-----------------')
          //console.log(array_with_powers)

          populate_prices_to_normal_heater(heater_name, array_with_prices, array_with_prices, workbook);
          populate_powers(heater_name, workbook);
          //populate_powers_to_normal_heater(heater_name, array_with_powers, workbook);
          break;
        case "elektryczny":
          console.log("ten skrypt nie potrafi jeszcze generować tabeli dla  grzejników elektrycznych");
          break;
        default:
          console.log("w polu A3 podaj rodzaj grzejnika jako 'zwykły' lub 'elektryczny'");
          break;
      }
    }
  } else {
    create_sheet_with_empty_template_normal_heater(workbook);
  }
}


function populate_powers(heater_mame: string, workbook: ExcelScript.Workbook) {
  console.log('populate_powers START');
  var start_row = 11;

  let metryki = workbook.getWorksheet("METRYKI WG");
  let usedRange = metryki.getUsedRange();
  let found1 = usedRange.find(heater_mame, {
    completeMatch: false,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.forward
  });

  console.log("metryki wg getRowIndex1 = " + found1.getRowIndex());

  let found2 = usedRange.find(heater_mame, {
    completeMatch: false,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.backwards
  });

  console.log("metryki wg getRowIndex2 = " + found2.getRowIndex());

  let selected = metryki.getRangeByIndexes(
    found1.getRowIndex(),
    found1.getColumnIndex(),
    found2.getRowIndex() + 1 - found1.getRowIndex() ,
    usedRange.getLastColumn().getColumnIndex() - found2.getColumnIndex() + 1);

  let array_text = selected.getTexts();
  console.log("array_text.length = " + array_text.length);

  // zrob tablice indeksow w formacie 
  // [0] 540300farba
  // [1] WGALE054030chrom
  let array_text2 = Array();
  for (let i = 0; i < array_text.length; i++) {
    let index = array_text[i][2] + '' + array_text[i][3] + '0';
    // console.log(index)
    array_text2[(1 + i + found1.getRowIndex())] =  array_text[i][2] + '-' + array_text[i][3] + '-' + array_text[i][7] ;
  }
  console.log("array_text2 = ", array_text2);

  let selectedSheet = workbook.getActiveWorksheet();

  console.log("procedura wypełniająca");

  for (let i = 0; i < 100; i++) {
    let tmp = selectedSheet.getRange("A" + (start_row + i)).getValue();
  
    if (tmp.toString().length > 0) {
      let index_tmp = selectedSheet.getRange("A" + (start_row + i)).getValue().toString();
      index_tmp += "-";
      index_tmp += selectedSheet.getRange("B" + (start_row + i)).getValue().toString();

      let index_farba = index_tmp + '-'+ 'farba';
      let index_chrom = index_tmp + '-' + 'chrom'; 

      //  if (array_text2)
      let tmp1 = array_text2.indexOf(index_farba);
      let tmp2 = array_text2.indexOf(index_chrom);

      if (tmp1 != -1) { // farba
        selectedSheet.getRange("C" + (start_row + i)).setValue("='METRYKI WG'!E" + tmp1);
        selectedSheet.getRange("D" + (start_row + i)).setValue("='METRYKI WG'!F"+tmp1);
        selectedSheet.getRange("E" + (start_row + i)).setValue("='METRYKI WG'!G"+tmp1);
      }
      if (tmp2 != -1) { //chrom
        selectedSheet.getRange("F" + (start_row + i)).setValue("='METRYKI WG'!E" + tmp1);
        selectedSheet.getRange("G" + (start_row + i)).setValue("='METRYKI WG'!F" + tmp1);
        selectedSheet.getRange("H" + (start_row + i)).setValue("='METRYKI WG'!G" + tmp1);
      }

    } else {
        break;
    }
  }

//  return array_text;
 // ???? 

} // function populate_powers - koniec



function does_generator_sheet_exist(workbook: ExcelScript.Workbook) {
  var worksheets = workbook.getWorksheets();
  var worksheet_exist = false;
  for (var i = 0; i < worksheets.length; i++) {
    if (worksheets[i].getName() == "generator danych do cennika PDF") {
      worksheets[i].activate();

      // console.log("does_generator_sheet_exist: worksheet_exist");
      worksheet_exist = true;
    }
  }
  return worksheet_exist;
}



function populate_prices_to_normal_heater(
  heater_name: string,
  array_with_prices: String[],
  array_with_powers: String[],
  workbook: ExcelScript.Workbook) {

  let selectedSheet = workbook.getActiveWorksheet();
  const start_row = 11;
  let typoszeregi_tmp = Array();
  for (let i = 0; i < array_with_prices.length; i++ ) {
    typoszeregi_tmp.push(array_with_prices[i][2]);
  }

  //console.log(typoszeregi_tmp);
  var typoszeregi = toUniqueArray(typoszeregi_tmp); // ["red","green"]
  // console.log(typoszeregi);
  let index = 0;
  for (var i = 0; i < array_with_prices.length; i++) {
    // zabezpieczenie jezeli przez przypadek w zbiorze danych wystepuje inny string niz heater name 
    if (array_with_prices[i][2].indexOf(heater_name) == -1) {
      continue;
    }
          
    // console.log(array_with_prices[2]);
    index = typoszeregi.indexOf(array_with_prices[i][2])
    // console.log('index = ' + index);

    let dimension_a = array_with_prices[i][2].substr(5, 4);
    let dimension_b = array_with_prices[i][2].substr(9, 2) + '0';
    let price = array_with_prices[i][6];

    switch (array_with_prices[i][4].substr(0,5)) {

      case "Biały":
        selectedSheet.getRange("A" + Math.trunc((i / 3) + start_row)).setValue(dimension_a);
        selectedSheet.getRange("B" + Math.trunc((i / 3) + start_row)).setValue(dimension_b);
        selectedSheet.getRange("I" + Math.trunc((i / 3) + start_row)).setValue(price);
        break;

      case "Farba":
        selectedSheet.getRange("A" + Math.trunc((i / 3) + start_row)).setValue(dimension_a);
        selectedSheet.getRange("B" + Math.trunc((i / 3) + start_row)).setValue(dimension_b);
        selectedSheet.getRange("J" + Math.trunc((i / 3) + start_row)).setValue(price);
        break;

      case "Galwa":
        selectedSheet.getRange("A" + Math.trunc((i / 3) + start_row)).setValue(dimension_a);
        selectedSheet.getRange("B" + Math.trunc((i / 3) + start_row)).setValue(dimension_b);
        selectedSheet.getRange("K" + Math.trunc((i / 3) + start_row)).setValue(price);
        break;
    }
  }
}

function toUniqueArray(a:string[]) {
  var newArr = [];
  for (var i = 0; i < a.length; i++) {
    if (newArr.indexOf(a[i]) === -1) {
      newArr.push(a[i]);
    }
  }
  return newArr;
}



// funkcja generuje pusty szablon formularza do wspomagania generowania tabelek 
function create_sheet_with_empty_template_normal_heater(workbook: ExcelScript.Workbook) {

  var a = workbook.addWorksheet('generator danych do cennika PDF');
  a.activate();

  let selectedSheet = workbook.getActiveWorksheet();
  selectedSheet.getRange("A1").setValue("Dane konfiguracyjne do generatora tabel z cenami (do cennika PDF)");

  selectedSheet.getRange("C2").setValue("W polu A2 podaj nazwę grzejnika (np WGALE albo inny)");
  selectedSheet.getRange("C3").setValue("W polu A3 podaj rodzaj grzejnika (zwykły, elektryczny))");
  selectedSheet.getRange("A8").setValue("A");
  selectedSheet.getRange("B8").setValue("B");
  selectedSheet.getRange("C8").setValue("farba");
  selectedSheet.getRange("D8").setValue("farba");
  selectedSheet.getRange("E8").setValue("farba");
  selectedSheet.getRange("F8").setValue("chrom");
  selectedSheet.getRange("G8").setValue("chrom");
  selectedSheet.getRange("H8").setValue("chrom");
  selectedSheet.getRange("I8").setValue("Kolor");
  selectedSheet.getRange("J8").setValue("standard");
  selectedSheet.getRange("K8").setValue("RAL i kolory specjalne");
  selectedSheet.getRange("L8").setValue("Kod produktu");

  selectedSheet.getRange("C9").setValue("75/65/20°C");
  selectedSheet.getRange("D9").setValue("55/45/20°C");
  selectedSheet.getRange("E9").setValue("[-.-]");

  selectedSheet.getRange("F9").setValue("75/65/20°C");
  selectedSheet.getRange("G9").setValue("55/45/20°C");
  selectedSheet.getRange("H9").setValue("[-.-]");

}



// funkcja pobiera dane z pliku cennika dla typu okreslonego grzejnika
function get_array_with_prices(heater_mame: string, workbook: ExcelScript.Workbook) {
  let sheet = workbook.getWorksheet("CENNIK 2021");
  let usedRange = sheet.getUsedRange();
  let found1 = usedRange.find(heater_mame, {
    completeMatch: false,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.forward
  });

  console.log("getRowIndex1 = " + found1.getRowIndex());

  let found2 = usedRange.find(heater_mame, {
    completeMatch: false,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.backwards
  });

  console.log("getRowIndex2 = " + found2.getRowIndex());

  let selected = sheet.getRangeByIndexes(
    found1.getRowIndex(),
    found1.getColumnIndex(),
    found2.getRowIndex() + 1 - found1.getRowIndex() ,
    usedRange.getLastColumn().getColumnIndex() - found2.getColumnIndex() + 1);
  let array_text = selected.getTexts();
  console.log("array_text.length = " + array_text.length);
  // ponizej wyrzucenie z tablicy wszystkich wierszy
  // ktore mają w kolumnie E wyraz "ZAMÓWIENIE"
  let index_zamowienie = Array();
  let array_text2 = Array();
  var data_inconsistency = "";
  for (var i=0; i < array_text.length; i++ ) {
    if (array_text[i][4].indexOf("ZAMÓWIENIE") == -1 ) {
      array_text2.push(array_text[i]);
    } 
    if (array_text[i][2].indexOf(heater_mame) == -1) {
      data_inconsistency = "brak konsystencji danych w poblizu wiersza " + (found1.getRowIndex() + i + 1 + " (" + array_text[i][2]+"). To oznacza że dane grzejnika w wielu miejscach pliku cennika.");
    }
  }
  // to jest zabezpieczenie jesli tposzeregi nie beda po sobie
  // czyli np troche grzejnika alex bedzie na koncu pliku a trocha na poczatku
  let selectedSheet = workbook.getActiveWorksheet();
  if (data_inconsistency != "") {
    selectedSheet.getRange("A5").getFormat()
      .getFont()
      .setColor("FF0000");
    selectedSheet.getRange("A5").setValue(data_inconsistency);
  } else {
    selectedSheet.getRange("A5").setValue("");
  }
  return array_text2;
}

