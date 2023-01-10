import { CSVToArray } from "./CSV.js";

const opzioniImport = {
    types: [
      {
        description: 'File CSV',
        accept: {
          'excel/*': ['.csv']
        }
      },
    ],
    excludeAcceptAllOption: true,
    multiple: false
};
  
const passSets = {
    "lettere": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "numeri": "0123456789"
}

const regexRangeRighe = /[^:0-9]/g;
const regexRangeColonne = /[^:A-Za-z]/g;

var loadinPlace, git_values, fileHandle;
var extImportData = { settings_rowstart: null, totalrow: null };
var csvData = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
  
      document.getElementById("msg").style.display = "none";
      document.getElementById("body").style.display = "block";
      document.getElementById("produttore").onchange = updateDataset;   //FATTO
      document.getElementById("btn_genInterni").onclick = genInterni;   //DA FARE
      document.getElementById("btn_genContatti").onclick = genContatti; //FATTO
      document.getElementById("btn_addTelefono").onclick = addTelefono; //FATTO
      document.getElementById("btn_genPagine").onclick = genPagine;     //FATTO
      document.getElementById("btn_importCSV").onclick = importCSV;     //FATTO
      document.getElementById("btn_importEXT").onclick = importEXT;     //FATTO
      document.getElementById("export_EXT").onclick = export_csv;       //FATTO
      document.getElementById("export_CONTACT").onclick = export_csv;   //FATTO
      if (window!=window.top){
        loadinPlace = "web";
        document.getElementById("webimport").style.display = "block";
      }
      gitVal();
    }
});
  
//FUNZIONI SYNC
function combineTemplate(sourceA, sourceB, offset) {
    let tmpArrayY = new Array;  
    for (let i = 0; i < sourceA.length; i++) {
      let tmpArrayX = [...sourceB];
      for (let j = 0; j < sourceA[i].length; j++) {
        tmpArrayX[offset[j] - 1] = sourceA[i][j];
      }
      tmpArrayY.push(tmpArrayX);
    }
    return tmpArrayY;
}

//RITORNA IL VALORE MINIMO E IL MASSIMO DI UN ARRAY
function getMinMax(arr) {
  let tmpArr = new Array();
  for (let i = 0; i < arr.length; i++) {
    if(arr[i] != ""){
      tmpArr.push(arr[i]);
    }
  }
  let max = Math.max(...tmpArr);
  let min = Math.min(...tmpArr);
  if(tmpArr.length == 0){
    return null;
  }
  return {min: min, max: max};
}

//AGGIORNA IL TAG SELECT CON IL DATASET DEL OPTION SELEZIONATO
function updateDataset(event) {
    let numRiga = event.target.options[event.target.selectedIndex].dataset.numRiga;
    event.target.dataset.numRiga = numRiga;
}

//GENERA CREDENZIALI - 0 PER PIN, 1 PER PASSWORD
function genPass(type, length, pattern) {
  let tmpSet = passSets.numeri;
  let psw = "";
  for (let i = 0, n = tmpSet.length; i < length; ++i) {
    psw += tmpSet.charAt(Math.floor(Math.random() * n));
  }
  if(type){
    psw = psw.split("");
    pattern = pattern.split("");
    pattern.sort(function(a, b){return 0.5 - Math.random()});
    for (let i = 0; i < length; i++) {
      if(pattern[i] == "x"){
        psw[i] = (passSets.lettere.charAt(Math.floor(Math.random()* passSets.lettere.length))).toLocaleLowerCase();
      }
      if(pattern[i] == "X"){
        psw[i] = (passSets.lettere.charAt(Math.floor(Math.random()* passSets.lettere.length)));
      }
      
      if(pattern[i] == "1"){
        psw[i] = (passSets.numeri.charAt(Math.floor(Math.random()* passSets.numeri.length)));
      }
    }
    psw = psw.join("");
  }
  return psw;
}

//FUNZIONI ASYNC
// GENERAZIONE CONTATTI
export async function genContatti() {
    try {
      await Excel.run(async (context) => {
 
        //RECUPERO SELEZIONE
        let srcRange = await selezioneADV(context, { start: "A", end: "I", rowend: null });
        let destRange = await selezioneADV(context, { start: "A", end: "N", rowend: null });
  
        //INIZIO PRELIEVO DATI
        const Contatti = context.workbook.worksheets.getItem("Contatti");
        let rangeContatti = Contatti.getRange(srcRange);
        rangeContatti.load("values");
        await context.sync();
  
        //UNISCO I DATI PRELEVATI AL TEMPLATE DEI CONTATTI
        let sourceA = rangeContatti.values;
        let sourceB = combineTemplate(sourceA, git_values.conf.contatti.template, git_values.conf.contatti.offsets);

        //CARICO DATI
        const UscitaContatti = context.workbook.worksheets.getItem("Uscita contatti");
        let rangeUscitaContatti = UscitaContatti.getRange(destRange);
        rangeUscitaContatti.load("values");
        await context.sync();
        rangeUscitaContatti.values = sourceB;
  
      });
    } catch (error) {
      genPagine();
    }
}

//GENRAZIONE PAGINE 
export async function genPagine() {
    try {
      await Excel.run(async (context) => {
        //CREATE PAGE 
        const names = context.workbook.worksheets.load("items/name");
        let paginePresenti = new Array();
        await context.sync();
        names.items.forEach((nomePagina) => {
          paginePresenti.push(nomePagina.name);
        })
        git_values.pagine.forEach((pagina, i) => {
          if (paginePresenti.includes(pagina.nome) == false) {
            context.workbook.worksheets.add(pagina.nome);
          }
        })
  
      });

      //LOAD DATA
      await Excel.run(async (context) => {
        git_values.pagine.forEach(pagina => {
            const selPagina = context.workbook.worksheets.getItem(pagina.nome);
            const range = selPagina.getRange(pagina.range);
            range.values = pagina.celle;
            const range2 = selPagina.getRange(pagina.range.replace(/\d+/g, ""));
            range2.numberFormat = "@";
        });
      });
  
    } catch (error) {
        console.log(error);
    }
}

//AGGIUNGE MUNU A TENDINA PER SELEZIONARE MODELLI DEI TELEFONI
export async function addTelefono() {
  let selectVal = await document.getElementById("produttore");
  if (selectVal.value != 0) {
    try {
      await Excel.run(async (context) => {

        // IMPOSTO MODELLI DA SELEZIONARE
        const Interni = context.workbook.worksheets.getItem("Interni");
        let modello = await selezioneADV(context, { start: "G", end: "G", rowend: null });
        let rangeModello = Interni.getRange(modello);
        rangeModello.values = "Seleziona";
        rangeModello.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: `${(git_values.produttori[selectVal.value].modelli).join(",")}`
          }
        }; 

        //METTO RIGA INTERNO
        let riga = await selezioneADV(context, { start: "I", end: "I", rowend: null });
        let rangeRiga = Interni.getRange(riga);
        rangeRiga.values = selectVal.dataset.numRiga;
        await context.sync();

      });
    } catch (error) {
      genPagine();
    }
  }
}

//ESPORTA INTERNI IN FORMATO CSV SEPARATO DA VIRGOLA
export async function export_csv(event){
  var c = document.createElement("a");
  let typeBTN = event.target.id;
  let nomePagina = (typeBTN == "export_EXT") ? "Uscita interni" : "Uscita contatti";
  let opzioni = (typeBTN == "export_EXT") ? { start: "A", end: "BH", rowend: null }: { start: "A", end: "N", rowend: null };
  c.download = `${nomePagina}.csv`;

  try {
    await Excel.run(async (context) => {
      let setUscitaInterni = await selezioneADV(context, opzioni);
      let uscita = context.workbook.worksheets.getItem(nomePagina);
      let rangeuscita = uscita.getRange(setUscitaInterni);
      rangeuscita.load("values");
      await context.sync();
      let rawData = rangeuscita.values;
      let dataOut = new Array();
      if(nomePagina == "Uscita interni"){
        dataOut.push(git_values.pagine[3].celle[0].join(","));
      }
      else{
        dataOut.push(git_values.pagine[4].celle[0].join(","));  
      }

      for (let i = 0; i < rawData.length; i++) {
        dataOut.push(rawData[i].join(","));
      }
      var t = new Blob([dataOut.join("\n")], {
        type: "text/plain"
      });
      c.href = window.URL.createObjectURL(t);
      c.click();
    });
  }
  catch(error){
    console.log(error);
  }
}

//IMPORTA CSV GENERATO DA 3CX 
export async function importCSV() {
  var contents;
  document.getElementById("btn_importEXT").style.display = "block";
  if(loadinPlace == "web"){
    contents = document.getElementById("webimport").value;
  }else{
    [fileHandle] = await window.showOpenFilePicker(opzioniImport);
    const file = await fileHandle.getFile();
    contents = await file.text();
  }
  csvData = CSVToArray(contents, ",");
  
  try {
    await Excel.run(async (context) => {

      //PREPARO SELEZIONE E IMPORT DEI INTERNI
      let selImpostazioni = await selezioneADV(context, { start: "A", end: "F", rowend: csvData.length - 1 });
      extImportData.settings_rowstart = ((selImpostazioni.replace(regexRangeRighe, "")).split(":"))[0];
      extImportData.totalrow = csvData.length - 1;

      //SELEZIONE E IMPORTO DATI
      const Impostazioni = context.workbook.worksheets.getItem("Impostazioni");
      let range = Impostazioni.getRange(selImpostazioni);
      let data = new Array();
      for (let i = 1; i < csvData.length; i++) {
        let Confg_impostazioni = git_values.conf.impostazioni;
        let row = [...Confg_impostazioni.template];
        row[0] = csvData[i][0];
        for (let j = 0; j < Confg_impostazioni.import.length; j++) {
          row[Confg_impostazioni.import[j]] = csvData[i][Confg_impostazioni.offsets[j]];
        }
        data.push(row);
      }
      range.load("values");
      await context.sync();
      range.values = data;
    })
  } catch (error) {
    genPagine();
  }
}

export async function importEXT() {
  document.getElementById("btn_importEXT").style.display = "none";
  if (csvData.length != null) {
    try {
      await Excel.run(async (context) => {
        //USCITA INTERNI
        let selInterni = await selezioneADV(context, { start: "A", end: "I", rowend: extImportData.totalrow });
        let Interni = context.workbook.worksheets.getItem("Interni");
        const range = Interni.getRange(selInterni);
        let data = new Array();
        for (let i = 1; i < csvData.length; i++) {
          let cfg = git_values.conf.interni;
          let row = [...cfg.template];
          row[0] = csvData[i][0];
          for (let j = 0; j < cfg.offsets.length; j++) {
            row[j] = csvData[i][cfg.offsets[j] - 1];
          }
          row.push(Number(extImportData.settings_rowstart) + i - 1);
          data.push(row);
        }
        range.load("values");
        await context.sync();
        range.values = data;
      })
    } catch (error) {
      genPagine();
    }
  }
}

export async function genInterni() {
  try {
    Excel.run(async (context) => {
      let ext = await selezioneADV(context, { start: "A", end: "H", rowend: null });                 // PAGINA INTERNI 
      let setSel = await selezioneADV(context, { start: "I", end: "I", rowend: null });             //PAGINA INTERNI RIGA IMPOSTAZIONI
      let setUscitaInterni = await selezioneADV(context, { start: "A", end: "BH", rowend: null }); //PAGINA USCITA INTERNI

      //IMPORTO INTERNI DA PAGINA IMPOSTAZIONI
      const Interni = context.workbook.worksheets.getItem("Interni");
      let rangeInterni = Interni.getRange(ext);
      rangeInterni.load("values");
      await context.sync();
      let dataInterni = rangeInterni.values;
      let dataOut = combineTemplate(dataInterni, git_values.conf.uscitaInterni.template, git_values.conf.interni.offsets);

      //CONTROLLO RIGA IMPOSTAZIONI PER RECUPERO XML
      rangeInterni = Interni.getRange(setSel);
      rangeInterni.load("values");
      await context.sync();
      let tmpSettings = getMinMax(rangeInterni.values); 
      
      // IMPORTO XML DA PAGINA IMPOSTAZIONI
      if(tmpSettings != null){
        const Impostazioni = context.workbook.worksheets.getItem("Impostazioni");
        let rangeSettings = Impostazioni.getRange(`B${tmpSettings.min}:F${tmpSettings.max}`);
        rangeSettings.load("values");
        await context.sync();
        for (let i = 0; i < rangeInterni.values.length; i++) {
          let index = rangeInterni.values[i];
          if(index != ""){
            index = Number(index);
            index -= tmpSettings.min;
            dataOut[i][git_values.conf.impostazioni.offsets[0]] = rangeSettings.values[index][0];
            dataOut[i][git_values.conf.impostazioni.offsets[1]] = rangeSettings.values[index][3];
            dataOut[i][git_values.conf.impostazioni.offsets[2]] = `"${rangeSettings.values[index][4].replace(/\"/g, "\"\"")}"`;
          }
        }
      }
      
      //SETTO PWD
      for(let i = 0; i < dataOut.length; i++){
        for (let j = 0; j < git_values.conf.credenziali.pin.length; j++) {
          dataOut[i][git_values.conf.credenziali.pin[j]-1] = genPass(0, git_values.conf.credenziali.pinLength, null);
        }
        for (let j = 0; j < git_values.conf.credenziali.password.length; j++) {
          dataOut[i][git_values.conf.credenziali.password[j]-1] = genPass(1, git_values.conf.credenziali.passwordLength,  git_values.conf.credenziali.pattern);
        }
      }
      
      //POPOLO USCITA INTERNI
      let uscitaInterni = context.workbook.worksheets.getItem("Uscita interni");
      let rangeUscitaInterni = uscitaInterni.getRange(setUscitaInterni);
      rangeUscitaInterni.load("values");
      await context.sync();
      rangeUscitaInterni.values = dataOut;
    });
  } catch (error) {
    console.log(error);
    genPagine();
  }
}


async function gitVal() {
  fetch('https://raw.githubusercontent.com/ZSirag/3CX-Excel/main/addin.json')
    .then(res => res.json())
    .then(json => {
    let selElm = document.getElementById("produttore");
    git_values = json;
    for (var produttore in git_values.produttori) {
      const elm = document.createElement("option");
      elm.innerHTML = produttore;
      elm.value = produttore;
      elm.dataset.numRiga = git_values.produttori[produttore].riga;
      selElm.appendChild(elm);
    }
  })
}

//RITORNA IL NUMERO DI RIGHE SELEZIONATO CON LE COLENNE PASSATE COME ARGOMENTO
async function selezioneADV(context, colonne) {
  let tmpRange = context.workbook.getSelectedRange();
  tmpRange.load("address");
  await context.sync();
  let range = tmpRange.address;
  let paginaTolta = (range.split("!"))[1];
  let righe = paginaTolta.split(":");
  for (let i = 0; i < righe.length; i++) {
    righe[i] = righe[i].replace(regexRangeRighe, "");
    if (righe[i] == 1) {
      righe[i] = 2
    }
  }
  if (colonne.rowend != null) {
    righe[1] = (colonne.rowend + Number(righe[0])) - 1;
  }
  if (righe.length == 1) {
    return `${colonne.start}${righe[0]}:${colonne.end}${righe[0]}`;
  }
  return `${colonne.start}${righe[0]}:${colonne.end}${righe[1]}`;
}
  
