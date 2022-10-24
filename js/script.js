// <script src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
//
let base = []; //

var ExcelToJSON = function () {
  this.parseExcel = function (file) {
    var reader = new FileReader();

    reader.onload = function (e) {
      var data = e.target.result;
      var workbook = XLSX.read(data, {
        type: "binary",
      });

      exelToObj(workbook);
    };

    reader.onerror = function (ex) {
      console.log(ex);
    };

    reader.readAsBinaryString(file);
  };
};

function handleFileSelect(evt) {
  var files = evt.target.files; // FileList object
  var xl2json = new ExcelToJSON();
  try {
    xl2json.parseExcel(files[0]);
  } catch (error) {
    xl2json.parseExcel(evt.dataTransfer.files[0]);
  }
}

function exelToObj(exselObj) {
  let contentMap = new Map();
  let activeSheet = [0];
  let stopParseSpec = false;
  if (Object.values(exselObj)[5]["Для цеху"]) {
    activeSheet = "Для цеху";
  }
  //Нoмер специфікації та замовник
  let specification = new Specification();
  specification.Num = specification.Num(exselObj, activeSheet) || "БЕЗ НОМЕРУ";
  if (specification.Num == "БЕЗ НОМЕРУ") {
    stopParseSpec = true;
    alert("УВАГА!\n ВІДСУТНІ ЗАГОЛОВКИ: \n Номер Специфікації.");
    return;
  }
  specification.Client =
    specification.Client(exselObj, activeSheet) || "Відсутній клієнт";
  if (specification.Client == "Відсутній клієнт") {
    stopParseSpec = true;
    alert("УВАГА!\n ВІДСУТНІ ЗАГОЛОВКИ: \n Замовник.");
    return;
  }
  //Адреси заголовків
  specification.titlesAdr.Marking =
    specification.titlesAdr(exselObj, activeSheet, /Маркування/i) || "-";
  specification.titlesAdr.Metall =
    specification.titlesAdr(exselObj, activeSheet, /Вибір металу:/i) || "-";
  specification.titlesAdr.Material =
    specification.titlesAdr(exselObj, activeSheet, /Матеріали/i) || "-";
  specification.titlesAdr.Mat =
    specification.titlesAdr(exselObj, activeSheet, /Використати мат:/i) || "-";

  specification.titlesAdr.MarkingSTOP = () => {
    if (specification.titlesAdr.Marking == "-") {
      alert("УВАГА!\n ВІДСУТНІ ЗАГОЛОВКИ: \n Маркування.");
    }
    if (specification.titlesAdr.Metall != "-") {
      return specification.titlesAdr.Metall[1].valueOf();
    } else if (specification.titlesAdr.Material != "-") {
      return specification.titlesAdr.Material[1].valueOf();
    } else if (specification.titlesAdr.Mat != "-") {
      return specification.titlesAdr.Mat[1].valueOf();
    } else {
      stopParseSpec = true;
      alert(
        "УВАГА!\n ВІДСУТНІ ЗАГОЛОВКИ: \n Вибір металу, Матеріали, Використати мат."
      );
      return 0;
    }
  };

  specification.titlesAdr.Name = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Назва/i
  ) || ["z", "1"];
  specification.titlesAdr.Length = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Довжина/i
  ) || ["z", "1"];
  specification.titlesAdr.Width = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Ширина/i
  ) || ["z", "1"];
  specification.titlesAdr.Quantity = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Кількість/i
  ) || ["z", "1"];
  specification.titlesAdr.Notes = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Примітки/i
  ) || ["z", "1"];
  specification.titlesAdr.Notch = specification.titlesAdr(
    exselObj,
    activeSheet,
    /Вирізи/i
  ) || ["z", "1"];
  //["z","1"] -у випадку, якщо не знайдено заголовок

  let nameCell =
    "S" +
    specification.Num.replace(/([ЗМІНА]{5}\d*)/g, "_$1_") +
    "___" +
    specification.Client;
  nameCell =
    nameCell.replace(/[\s]/g, "").replace(/[/]/g, "_").replace(/["]/g, "__") +
    [];
  console.log(nameCell, base);

  base.push(nameCell);
  base.push(specification.Num);
  base.push(specification.Client);
  base.push(detailsMarking()); //запис Маркування + характеристики деталей
  base.push([]); // для подальших записів ВТК

  // відправка на сервер
  function onSuccess(baseReturn) {
    baseReturn = JSON.parse(baseReturn);
    console.log(baseReturn);
    var div = document.getElementById("output");
    div.innerHTML = baseReturn;
  }

  ///////////////////////////////////////////////////////////////////////////////////

  console.log(base);

  // google.script.run.withSuccessHandler(onSuccess).getBase(base);
  base = [];
  //////////////////////////////////////////////////////////////////////////////////////////

  function detailsMarking() {
    if (!stopParseSpec) {
      let details = [];
      let markingSTOP =
        specification.titlesAdr.MarkingSTOP() -
        specification.titlesAdr.Marking[1] * 1;
      for (let n = 1; n < markingSTOP; n++) {
        let detail = []; // Маркування + характеристики деталі
        let markingAdr =
          specification.titlesAdr.Marking[0] +
          (specification.titlesAdr.Marking[1] * 1 + n);
        let nameAdr =
          specification.titlesAdr.Name[0] +
          (specification.titlesAdr.Name[1] * 1 + n);
        let lengthAdr =
          specification.titlesAdr.Length[0] +
          (specification.titlesAdr.Length[1] * 1 + n);
        let widthAdr =
          specification.titlesAdr.Width[0] +
          (specification.titlesAdr.Width[1] * 1 + n);
        let quantityAdr =
          specification.titlesAdr.Quantity[0] +
          (specification.titlesAdr.Quantity[1] * 1 + n);
        let notesAdr =
          specification.titlesAdr.Notes[0] +
          (specification.titlesAdr.Notes[1] * 1 + n);
        let notchAdr =
          specification.titlesAdr.Notch[0] +
          (specification.titlesAdr.Notch[1] * 1 + n);

        let tempBuffer;
        //маркування.(Назва, Довжина, Ширина, Кількість, Примітки, Вирізи)  ( маркування [])
        try {
          tempBuffer = specification.tableContent(
            exselObj,
            activeSheet,
            markingAdr
          );
          if (tempBuffer) {
            detail.push(tempBuffer);
            [
              nameAdr,
              lengthAdr,
              widthAdr,
              quantityAdr,
              notesAdr,
              notchAdr,
            ].forEach((item) => {
              tempBuffer = specification.tableContent(
                exselObj,
                activeSheet,
                item
              );
              tempBuffer ? detail.push(tempBuffer) : detail.push("");
            });
            details.push(detail);
          }
        } catch (e) {}
      }
      return details;
    }
  }
}

let Specification = function () {
  //Отримання № специфікаціі
  this.Num = function (exselObj, activeSheet) {
    let maxRow = 10;

    let Spec =
      /([\u0410-\u042F\u0456i]{12}\s*[\u2116]\s*)(\d{1,})(\s*-?\d*)\s*(\u002F\s*\u0447?[\u043e\u0426]{0,2}\s*\u002F\u0423?[\u0415|e]?)/i;
    let SpecTest = /[СПЕЦИФІКАЦІЯi]{12}\s/i;
    let SpecAnswer = "$2$4"; //«АФЕТ-БУД»

    for (let row = 1; row < maxRow; row++) {
      for (let column = 65; column < 90; column++) {
        let cell = String.fromCharCode(column) + (row + "");
        try {
          if (
            SpecTest.test(Object.values(exselObj)[5][activeSheet][cell]["w"])
          ) {
            let val = Object.values(exselObj)[5][activeSheet][cell]["w"];
            val = val
              .replace(Spec, SpecAnswer)
              .replace(/\s{2,}/, "s")
              .replace(/[«||»]/g, '"')
              .toUpperCase();
            return val;
          }
        } catch (error) {}
      }
    }
  };
  this.Client = function (exselObj, activeSheet) {
    let maxRow = 10;

    let zamovnyk = /\s*([Замовник]{8}\s*:*\s*)([\s\u0410-\u042F\u0456A-Z]*)/i;
    let zamovnykTest = /[Замовник]{8}\s*:*\s*/i;
    let zamovnykAnswer = "$2";

    for (let row = 1; row < maxRow; row++) {
      for (let column = 65; column < 90; column++) {
        let cell = String.fromCharCode(column) + (row + "");
        try {
          if (
            zamovnykTest.test(
              Object.values(exselObj)[5][activeSheet][cell]["w"]
            )
          ) {
            let val = Object.values(exselObj)[5][activeSheet][cell]["w"];
            val = val
              .replace(zamovnyk, zamovnykAnswer)
              .replace(/\s{2,}/, "s")
              .replace(/[«||»]/g, '"')
              .toUpperCase();
            return val;
          }
        } catch (error) {}
      }
    }
  };

  this.titlesAdr = function (exselObj, activeSheet, titleRegEx) {
    //отримання адреси ячейки заголовку
    let maxRow = 10;
    if (titleRegEx.test("/Вибір металу:/gi/Матеріали/ig/Використати мат:/gi")) {
      maxRow = 1024;
    }

    for (let row = 1; row < maxRow; row++) {
      for (let column = 65; column < 90; column++) {
        let cell = String.fromCharCode(column) + (row + "");
        try {
          if (
            titleRegEx.test(Object.values(exselObj)[5][activeSheet][cell]["w"])
          ) {
            return [String.fromCharCode(column), row];
          }
        } catch (error) {}
      }
    }
  };
  this.tableContent = function (exselObj, activeSheet, adressCell = "G1") {
    //отримання контенту з ячейки
    try {
      return Object.values(exselObj)[5][activeSheet][adressCell]["w"];
    } catch (error) {
      //console.error(adressCell, ": error")
      return false;
    }
  };
};
