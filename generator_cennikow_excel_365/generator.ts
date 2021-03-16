// makro pomocne przy tworzeniu tabel do cenników PDF
// Przemysław Rzeźnik Marzec 2021

function main(workbook: ExcelScript.Workbook) {
  console.log('main start');

  if (does_generator_sheet_exist(workbook)) {

    console.log('main does_generator_sheet_exist');

    let selectedSheet = workbook.getActiveWorksheet();
    selectedSheet.getRange("A10:N99").clear(ExcelScript.ClearApplyTo.contents);

    let validation_results = validate_all_parameters(workbook);

    if (validation_results) {
      // create_empty_template_normal_heater(workbook);
      create_header_of_table(workbook);
      let heater_name = selectedSheet.getRange("A4").getValue().toString();
      let heater_kind = selectedSheet.getRange("A5").getValue().toString();

      switch (heater_kind) {
        case "zwykły":
          do_normal_heater(workbook);
          break;
        case "elektryczny":
          console.log("ten skrypt nie potrafi jeszcze generować tabeli dla  grzejników elektrycznych");
          break;
        default:
          console.log("w polu A3 podaj rodzaj grzejnika jako 'zwykły' lub 'elektryczny'");
          break;
      }
    } else {
      //console.log("xxxxxxxxxxxx x x xxxxxr");
      //create_empty_template_normal_heater(workbook);
    }
  } else {
    create_sheet(workbook);
    create_empty_template_normal_heater(workbook);

  }
}


function do_normal_heater(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  let heater_name = selectedSheet.getRange("A4").getValue().toString();

  let array_with_prices = get_array_with_prices(heater_name, workbook);
  console.log('array_with_prices', array_with_prices);
  populate_prices_to_normal_heater(heater_name, array_with_prices, workbook);
  populate_powers(heater_name, workbook);

  //populate_powers_to_normal_heater(heater_name, array_with_powers, workbook);

}


function populate_powers(heater_mame: string, workbook: ExcelScript.Workbook) {
  console.log('populate_powers START');
  var start_row = 21;

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
    found2.getRowIndex() + 1 - found1.getRowIndex(),
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
    array_text2[(1 + i + found1.getRowIndex())] = array_text[i][2] + '-' + array_text[i][3] + '-' + array_text[i][7];
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

      let index_farba = index_tmp + '-' + 'farba';
      let index_chrom = index_tmp + '-' + 'chrom';

      //  if (array_text2)
      let tmp1 = array_text2.indexOf(index_farba);
      let tmp2 = array_text2.indexOf(index_chrom);

      if (tmp1 != -1) { // farba
        selectedSheet.getRange("C" + (start_row + i)).setValue("='METRYKI WG'!E" + tmp1);
        selectedSheet.getRange("D" + (start_row + i)).setValue("='METRYKI WG'!F" + tmp1);
        selectedSheet.getRange("E" + (start_row + i)).setValue("='METRYKI WG'!G" + tmp1);
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
  workbook: ExcelScript.Workbook) {

  console.log('populate_prices_to_normal_heater');

  let selectedSheet = workbook.getActiveWorksheet();
  const start_row = 21;
  let typoszeregi_tmp = Array();
  for (let i = 0; i < array_with_prices.length; i++) {
    typoszeregi_tmp.push(array_with_prices[i][2]);
  }

  //console.log(typoszeregi_tmp);
  var typoszeregi = toUniqueArray(typoszeregi_tmp); // ["red","green"]
  console.log("typoszeregi", typoszeregi);
  let index = 0;

  // kolor_standard, kolor_ral, kolor_chrom
  let painting = selectedSheet.getRange("A7").getValue().toString();

  let kolor_standard = 0;
  let kolor_ral = 0;
  let kolor_chrom = 0;

  if (painting.indexOf('kolor_standard') > -1) {
     kolor_standard = 1;
  }

  if (painting.indexOf('kolor_ral') > -1) {
    kolor_ral = 1;
  }

  if (painting.indexOf('kolor_chrom') > -1) {
    kolor_chrom = 1;
  }

  var dimension_old_a = "";
  var dimension_old_b = "";

  let row_counter_destination = -1;

  for (var i = 0; i < array_with_prices.length; i++) {
    // zabezpieczenie jezeli przez przypadek w zbiorze danych wystepuje inny string niz heater name 
    if (array_with_prices[i][2].indexOf(heater_name) == -1) {
      continue;
    }

    index = typoszeregi.indexOf(array_with_prices[i][2])
    let dimension_a = array_with_prices[i][2].substr(5, 4);
    let dimension_b = array_with_prices[i][2].substr(9, 2) + '0';
    let price = array_with_prices[i][9];

    if ((dimension_old_a == dimension_a) && (dimension_old_b == dimension_b))
    {

    } else {
      dimension_old_a = dimension_a;
      dimension_old_b = dimension_b;
      row_counter_destination++;

      selectedSheet.getRange("A" + (row_counter_destination + start_row)).setValue(dimension_a);
      selectedSheet.getRange("B" + (row_counter_destination + start_row)).setValue(dimension_b);

      selectedSheet.getRange("L" + (row_counter_destination + start_row)).setValue(array_with_prices[i][2]+'...');


    }



    switch (array_with_prices[i][4].substr(0, 5)) {
      case "Biały":
        if (kolor_standard) {

          selectedSheet.getRange("I" + (row_counter_destination + start_row)).setValue(price);
        }
        break;

      case "Farba":
        if (kolor_ral) {
/*          selectedSheet.getRange("A" + (row_counter_destination + start_row)).setValue(dimension_a);
          selectedSheet.getRange("B" + (row_counter_destination + start_row)).setValue(dimension_b);
          */
          selectedSheet.getRange("J" + (row_counter_destination + start_row)).setValue(price);
        }
        break;

      case "Galwa":
        if (kolor_chrom) {
          /*selectedSheet.getRange("A" + (row_counter_destination + start_row)).setValue(dimension_a);
          selectedSheet.getRange("B" + (row_counter_destination + start_row)).setValue(dimension_b);
          */
          selectedSheet.getRange("K" + (row_counter_destination + start_row)).setValue(price);
        }
        break;
    }
  }
}

function toUniqueArray(a: string[]) {
  var newArr = [];
  for (var i = 0; i < a.length; i++) {
    if (newArr.indexOf(a[i]) === -1) {
      newArr.push(a[i]);
    }
  }
  return newArr;
}



function create_sheet(workbook: ExcelScript.Workbook) {
  var a = workbook.addWorksheet('generator danych do cennika PDF');
  a.activate();
}


function create_header_of_table(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  selectedSheet.getRange("A19").setValue("A");
  selectedSheet.getRange("B19").setValue("B");
  selectedSheet.getRange("C19").setValue("farba");
  selectedSheet.getRange("D19").setValue("farba");
  selectedSheet.getRange("E19").setValue("farba");
  selectedSheet.getRange("F19").setValue("chrom");
  selectedSheet.getRange("G19").setValue("chrom");
  selectedSheet.getRange("H19").setValue("chrom");
  selectedSheet.getRange("I19").setValue("Kolor standard");
  selectedSheet.getRange("J19").setValue("RAL i kolory specjalne");
  selectedSheet.getRange("K19").setValue("Chrom");
  selectedSheet.getRange("L19").setValue("Kod produktu");
  selectedSheet.getRange("C20").setValue("75/65/20°C");
  selectedSheet.getRange("D20").setValue("55/45/20°C");
  selectedSheet.getRange("E20").setValue("[-.-]");
  selectedSheet.getRange("F20").setValue("75/65/20°C");
  selectedSheet.getRange("G20").setValue("55/45/20°C");
  selectedSheet.getRange("H20").setValue("[-.-]");
}

// funkcja generuje pusty szablon formularza do wspomagania generowania tabelek 
function create_empty_template_normal_heater(workbook: ExcelScript.Workbook) {

  let selectedSheet = workbook.getActiveWorksheet();
  selectedSheet.getRange("A1").setValue("Dane konfiguracyjne do generatora tabel z cenami (do cennika PDF)");

  selectedSheet.getRange("C2").setValue("W polu A2 podaj w której zakładce są ceny (np. CENNIK 2021)");

  selectedSheet.getRange("A2").setValue("CENNIK 2021");

  selectedSheet.getRange("C3").setValue("W polu A3 podaj w której zakładce są metryki (np. METRYKI WG)");

  selectedSheet.getRange("A3").setValue("METRYKI WG");

  selectedSheet.getRange("C4").setValue("W polu A4 podaj nazwę grzejnika (np WGALE albo inny)");
  selectedSheet.getRange("A4").setValue("WGALE");

  selectedSheet.getRange("C5").setValue("W polu A5 podaj rodzaj grzejnika (zwykły, elektryczny)");

  selectedSheet.getRange("A5").setValue("zwykły");

  selectedSheet.getRange("C6").setValue("W polu A6 podaj liczbę w której kolumnie cennika znajduje się cena (licząc od zera. np 9)");
  selectedSheet.getRange("A6").setValue("9");

  selectedSheet.getRange("C7").setValue("W polu A7 podaj ile jakie kolumny chcesz mieć wygenerowane (kolor_standard, kolor_ral, kolor_chrom) lub (kolor_standard, kolor_chrom)");

  selectedSheet.getRange("A7").setValue("kolor_standard, kolor_ral, kolor_chrom");

 

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
    found2.getRowIndex() + 1 - found1.getRowIndex(),
    usedRange.getLastColumn().getColumnIndex() - found2.getColumnIndex() + 1);
  let array_text = selected.getTexts();
  console.log("array_text.length = " + array_text.length);
  // ponizej wyrzucenie z tablicy wszystkich wierszy
  // ktore mają w kolumnie E wyraz "ZAMÓWIENIE"
  let index_zamowienie = Array();
  let array_text2 = Array();
  var data_inconsistency = "";
  for (var i = 0; i < array_text.length; i++) {
    if (array_text[i][4].indexOf("ZAMÓWIENIE") == -1) {
      array_text2.push(array_text[i]);
    }
    if (array_text[i][2].indexOf(heater_mame) == -1) {
      data_inconsistency = "brak konsystencji danych w pobliżu wiersza " + (found1.getRowIndex() + i + 1 + " (" + array_text[i][2] + "). To oznacza że dane grzejnika w wielu miejscach pliku cennika. A powinny być po sobie (jeden po drugim)");
    }
  }

  console.log('array_text2', array_text2);

  // to jest zabezpieczenie jeśli typoszeregi nie beda po sobie
  // czyli np troche grzejnika alex bedzie na koncu pliku a trocha na poczatku
  let selectedSheet = workbook.getActiveWorksheet();
  if (data_inconsistency != "") {
    show_error(workbook, data_inconsistency);
  }
  return array_text2;
}


function show_error(workbook: ExcelScript.Workbook, error: String) {
  let start_error_row = 10;
  console.log("show_error");

  let selectedSheet = workbook.getActiveWorksheet();
  for (let i = 0; i < 10; i++) {
    let numer_komorki = "A" + (10 + i);
    console.log(numer_komorki);
    let value = selectedSheet.getRange(numer_komorki).getValue();
    if (value == "") {
      selectedSheet.getRange(numer_komorki).setValue(error);
      selectedSheet.getRange(numer_komorki).getFormat()
        .getFont()
        .setColor("FF0000");
      break;
    }
  }

}

function validate_all_parameters(workbook: ExcelScript.Workbook) {
  var result = true;

  let selectedSheet = workbook.getActiveWorksheet();

  let price_sheet = selectedSheet.getRange("A2").getValue().toString();
  let metrics_sheet = selectedSheet.getRange("A3").getValue().toString();
  let heater_name = selectedSheet.getRange("A4").getValue().toString();
  let heater_kind = selectedSheet.getRange("A5").getValue().toString();
  let price_column = selectedSheet.getRange("A6").getValue().toString();
  let painting = selectedSheet.getRange("A7").getValue().toString();

  var tmp = workbook.getWorksheets();
  var all_names = Array();
  for (let i = 0; i < tmp.length; i++) {
    all_names.push(tmp[i].getName());
  }

  let price_sheet_exist = all_names.indexOf(price_sheet);
  if (price_sheet_exist == -1) {
    selectedSheet.getRange("I2").setValue("Nie ma takiego cennika w zakładkach");
    selectedSheet.getRange("I2").getFormat().getFont().setColor("FF0000");
    result = false;
  } else {
    selectedSheet.getRange("I2").setValue("");
    selectedSheet.getRange("I2").getFormat().getFont().setColor("000000");
  }

  let metrics_sheet_exist = all_names.indexOf(metrics_sheet);
  if (metrics_sheet_exist == -1) {
    selectedSheet.getRange("I3").setValue("Nie ma takiego pliku z metrykami w zakładkach");
    selectedSheet.getRange("I3").getFormat().getFont().setColor("FF0000");
    result = false;
  } else {
    selectedSheet.getRange("I3").setValue("");
    selectedSheet.getRange("I3").getFormat().getFont().setColor("000000");
  }

  let sheet = workbook.getWorksheet(price_sheet);
  let usedRange = sheet.getUsedRange();
  let found1 = usedRange.find(heater_name, {
    completeMatch: false,
    matchCase: false,
    searchDirection: ExcelScript.SearchDirection.forward
  });


  // nazwa grzejnika
  if (found1 === undefined) {
    selectedSheet.getRange("I4").setValue("Nie ma takiego grzejnika w pliku z cenami (" + price_sheet + ")");
    selectedSheet.getRange("I4").getFormat().getFont().setColor("FF0000");
    result = false;
  } else {
    selectedSheet.getRange("I4").setValue("");
    selectedSheet.getRange("I4").getFormat().getFont().setColor("000000");
  }


  // rodzaj grzejnika
  if (['zwykły', 'elektryczny'].indexOf(heater_kind) == -1) {
    selectedSheet.getRange("I5").setValue("'zwykły' lub 'elektryczny' to jedynie dozwolone wartości");
    selectedSheet.getRange("I5").getFormat().getFont().setColor("FF0000");
    result = false;
  } else {
    selectedSheet.getRange("I5").setValue("");
    selectedSheet.getRange("I5").getFormat().getFont().setColor("000000");
  }

  // kolumna cennika
  if ((parseInt(price_column, 10)) && (parseInt(price_column, 10)>0)) {
    selectedSheet.getRange("I6").setValue("");
    selectedSheet.getRange("I6").getFormat().getFont().setColor("000000");
  } else {
    selectedSheet.getRange("I6").setValue("tutaj musi być liczba całkowita większa o zera");
    selectedSheet.getRange("I6").getFormat().getFont().setColor("FF0000");
    result = false;
  }

  // painting
  // console.log(painting.indexOf('kolor_standard') !== -1);
  // 
  let tmp_bool = false;

  if (painting.indexOf('kolor_standard') > -1) {
    tmp_bool = true;
  }
  if (painting.indexOf('kolor_ral') > -1) {
    tmp_bool = true;
  }
  if (painting.indexOf('kolor_chrom') > -1) {
    tmp_bool = true;
  }

  if (tmp_bool == false) {
    selectedSheet.getRange("I7").setValue("'kolor_standard', 'kolor_ral', 'kolor_chrom' to jedynie dozwolone wartości (mogą być oddzielone przecinkiem)");
    selectedSheet.getRange("I7").getFormat().getFont().setColor("FF0000");
    result = false;
  } else {
    selectedSheet.getRange("I7").setValue("");
    selectedSheet.getRange("I7").getFormat().getFont().setColor("000000");
  }





  return result;
}

