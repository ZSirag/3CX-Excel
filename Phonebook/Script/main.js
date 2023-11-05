const fs = new FileReader();
const xml = new DOMParser();
var phoneBook = [];

//ANALISI DATI ED ESTRAZIONE
fs.onload = function (e){
  document.querySelector("#dz-container").style.display = "none";
  document.querySelector("#pb-container").style.display = "grid";
  const doc = xml.parseFromString(e.target.result, "application/xml");
  const tenant = doc.querySelector("Tenants");
  const publicPhonebook = tenant.querySelector("PhoneBookEntries");
  console.log(publicPhonebook);
  if(publicPhonebook.hasChildNodes()){
    
    populatePB([publicPhonebook], "Rubrica generale centralino");
  }
  const exts = doc.documentElement.querySelectorAll("Extension");
  for(let i = 0; i < exts.length; i++){
    const interni = exts[i];
    const rubrica = interni.getElementsByTagName("PhoneBookEntries");
    if(rubrica[0].hasChildNodes()){
      const nomeFile = `${(interni.getElementsByTagName("Number"))[0].innerHTML} - ${(interni.getElementsByTagName("FirstName"))[0].innerHTML}`;  
      populatePB(rubrica, nomeFile)
    }
  }
  createItem();
}


function populatePB(rubrica, nomeFile) {
  const contatto = rubrica[0].getElementsByTagName("PhoneBookEntry");
  let outputData = `FirstName,LastName,Company,Mobile,Mobile2,Home,Home2,Business,Business2,Email,Other,BusinessFax,HomeFax,Pager`;
  for(let j = 0; j < contatto.length; j++){
    let data = new Array(14);
    contatto[j].childNodes.forEach((item, index) => {
    switch (item.tagName) {
      case "FirstName": data[0] = item.innerHTML;
      break;
      case "LastName": data[1] = item.innerHTML;
      break;
      case "CompanyName": data[2] = item.innerHTML;
      break;
      case "PhoneNumber": data[3] = item.innerHTML;
      break;
      default: {
        if(item.tagName != undefined){
          data[index+4] = item.innerHTML
          }
        };
        break;
      }
    });
    outputData += `\n${data.join(",")}`;
  }
  phoneBook.push({nome: nomeFile, data: outputData})
}

function createItem(){
  parent = document.querySelector("#pb-entrylist");
  phoneBook.forEach((contact, index) => {
    const template = document.querySelector("#entry-clone").cloneNode(true);
    template.querySelector("input").className = "cg-entry";
    template.querySelector("p").innerHTML = contact.nome;
    template.querySelector("button").setAttribute("onclick", `saveThis(${index})`);
    parent.appendChild(template);
  });
  document.querySelector("#entry-clone").style.display = "none";
}

function saveThis(index) {
  let file = document.createElement("a");
  file.download = `${phoneBook[index].nome}.csv`;
  let content = new Blob([phoneBook[index].data], {type: "text/plain"});
  file.href = window.URL.createObjectURL(content);
  file.click();
}

function saveAll() {
  const items = document.querySelectorAll(".cg-entry");
  items.forEach((input, index) => {
    if(input.checked){
      saveThis(index);
    }
  });
}


//RECUPERO E LETTURA FILE 
function dropHandler(ev) {
  ev.preventDefault();  
  if (ev.dataTransfer.items) {
    [...ev.dataTransfer.items].forEach((item, i) => {
      if (item.kind === "file") {
        fs.readAsText(ev.dataTransfer.files[0]);
      }
    });
  } 
}
  
function dragOverHandler(ev) {
  ev.preventDefault();
}

function selectAll(elm) {
  const inputs = document.querySelectorAll(".pb-entry input");
  inputs.forEach(input => {
    input.checked = elm.checked
  });
}

function fileInputClick() {
  document.querySelector("#fileElm").click();
}

function fileInputHandle(elm){
  fs.readAsText(elm.files[0]);
}