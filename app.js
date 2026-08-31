/*
FILE: app.js
VAZIFASI:
- Барча бизнес-логика шу ерда
- Филиаллар рўйхати
- Фактура турига қараб счёт автомат танланади
- Суммалар пропорция қилиб тақсимланади
- OFFLINE .xls файл генерация қилинади

MUHIM:
- Бу файлдаги логикани сабабсиз ЎЗГАРТИРМА
- Excel импорти айнан шу форматга боғлиқ
*/

// ===== СПРАВОЧНИК (ФАКТУРА ТУРИ → ФИЛИАЛ → СЧЕТ) =====
const ACCOUNTS = {
  "КАТМ": {
    "BALIQCHI FILIALI":"56796000204603324051",
    "BESHARIQ FILIALI":"56796000404603324046",
    "IZBOSKAN FILIALI":"56796000604603324045",
    "QUVA FILIALI":"56796000304603324041",
    "QO‘QON FILIALI":"56796000904603324035",
    "QO‘RG‘ОНTEPA FILIALI":"56796000504603324050",
    "MINGBULOQ FILIALI":"56796000904603324049",
    "NORIN FILIALI":"56796000204603324047",
    "OLTIARIQ FILIALI":"56796000704603324039",
    "POP FILIALI":"56796000004603324036",
    "TO‘RAQO‘RG‘ON FILIALI":"56796000504603324040",
    "UCHKO‘PRIK FILIALI":"56796000104603324042",
    "UCHQO‘RG‘ON FILIALI":"56796000004603324043",
    "XO‘JAOBOD FILIALI":"56796000904603324048",
    "CHUST FILIALI":"56796000804603324038",
    "YAYPAN FILIALI":"56796000904603324037",
    "YANGIQURGON FILIALI":"56796000804603324044"
  },

  "Гаров реестр": {
    "QO‘QON FILIALI":"56796000104603324017",
    "BALIQCHI FILIALI":"56796000104603324034",
    "YANGIQURGON FILIALI":"56796000504603324026",
    "IZBOSKAN FILIALI":"56796000504603324027",
    "BESHARIQ FILIALI":"56796000504603324028",
    "NORIN FILIALI":"56796000504603324029",
    "XO‘JAOBOD FILIALI":"56796000504603324030",
    "MINGBULOQ FILIALI":"56796000404603324031",
    "CHUST FILIALI":"56796000504603324020",
    "OLTIARIQ FILIALI":"56796000504603324021",
    "TO‘RAQO‘RG‘ON FILIALI":"56796000504603324022",
    "QUVA FILIALI":"56796000504603324023",
    "UCHKO‘PRIK FILIALI":"56796000504603324024",
    "UCHQO‘RG‘ON FILIALI":"56796000504603324025",
    "POP FILIALI":"56796000204603324018",
    "YAYPAN FILIALI":"56796000304603324019",
    "QO‘RG‘ОНTEPA FILIALI":"56796000304603324032"
  },

  "АдлияОАКТМ": {
    "POP FILIALI":"56796000004603324802",
    "TO‘RAQO‘RG‘ON FILIALI":"56796000704603324806",
    "YAYPAN FILIALI":"56796000104603324803",
    "QUVA FILIALI":"56796000904603324807",
    "CHUST FILIALI":"56796000304603324804",
    "OLTIARIQ FILIALI":"56796000504603324805",
    "QO‘QON FILIALI":"56796000804603324801",
    "MINGBULOQ FILIALI":"56796000304603324815",
    "QO‘RG‘ОНTEPA FILIALI":"56796000404603324816",
    "UCHQO‘RG‘ON FILIALI":"56796000204603324809",
    "BALIQCHI FILIALI":"56796000504603324817",
    "XO‘JAOBOD FILIALI":"56796000204603324814",
    "UCHKO‘PRIK FILIALI":"56796000904603324808",
    "NORIN FILIALI":"56796000104603324813",
    "YANGIQURGON FILIALI":"56796000904603324810",
    "IZBOSKAN FILIALI":"56796000004603324811",
    "BESHARIQ FILIALI":"56796000904603324812"
    },

  "TRANS-INSSURANT": {
    "POP FILIALI":"56710000204603324004",
    "TO‘RAQO‘RG‘ON FILIALI":"56710000004603324008",
    "YAYPAN FILIALI":"56710000404603324005",
    "QUVA FILIALI":"56710000104603324009",
    "CHUST FILIALI":"56710000604603324006",
    "OLTIARIQ FILIALI":"56710000804603324007",
    "QO‘QON FILIALI":"56710000904603324003",
    "MINGBULOQ FILIALI":"56710000104603324017",
    "QO‘RG‘ОНTEPA FILIALI":"56710000204603324018",
    "UCHQO‘RG‘ON FILIALI":"56710000604603324011",
    "BALIQCHI FILIALI":"56710000304603324019",
    "XO‘JAOBOD FILIALI":"56710000904603324016",
    "UCHKO‘PRIK FILIALI":"56710000504603324010",
    "NORIN FILIALI":"56710000004603324015",
    "YANGIQURGON FILIALI":"56710000704603324012",
    "IZBOSKAN FILIALI":"56710000804603324013",
    "BESHARIQ FILIALI":"56710000904603324014"
  }
};

// КРЕДИТ СЧЕТ (ўзгармас)
const CREDIT = "29802000804603324001";

// Филиаллар рўйхати (шаблон қатъий)
const FILIALS = Object.keys(ACCOUNTS["Гаров реестр"]);

const tbody = document.querySelector("#tbl tbody");

// === ЖАДВАЛНИ ЯРАТИШ ===
function buildTable(){
  tbody.innerHTML = "";
  const startNo = Number(document.getElementById("startNo").value || 1);

  FILIALS.forEach((f,i)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td class="rowNo">${startNo + i}</td>
      <td>${f}</td>
      <td><input type="number" class="count"></td>
      <td class="amt">0.00</td>`;
    tbody.appendChild(tr);
  });
}

// === ENTER → кейинги қатор ===
document.addEventListener("keydown",e=>{
  if(e.key==="Enter" && e.target.classList.contains("count")){
    e.preventDefault();
    const a=[...document.querySelectorAll(".count")];
    const i=a.indexOf(e.target);
    if(a[i+1]) a[i+1].focus();
    calc();
  }
});

// === СУММАНИ ТАҚСИМЛАШ ===
function calc(){
  const total=+document.getElementById("total").value;
  let sumCnt=0,sumAmt=0;

  document.querySelectorAll(".count").forEach(i=>sumCnt+=+i.value);

  document.querySelectorAll("#tbl tbody tr").forEach(r=>{
    const c=+r.querySelector(".count").value;
    const a=sumCnt?total*c/sumCnt:0;
    r.querySelector(".amt").innerText=a.toFixed(2);
    sumAmt+=a;
  });

  document.getElementById("sumCnt").innerText=sumCnt;
  document.getElementById("sumAmt").innerText=sumAmt.toFixed(2);
}

// === XLS ЭКСПОРТ ===
document.getElementById("exportBtn").onclick = exportXLS;
document.getElementById("total").oninput = calc;
document.getElementById("startNo").oninput = buildTable;

function exportXLS(){
  const dateRaw=document.getElementById("date").value;
  const inv=document.getElementById("invoiceNo").value;
  const type=document.getElementById("type").value;

  if(!dateRaw||!inv){
    alert("Sana ва faktura рақамини киритинг");
    return;
  }

  const date=dateRaw.split("-").reverse().join(".");

  let xml=`<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Worksheet ss:Name="Trans"><Table>`;

  xml+=`<Row><Cell/><Cell/><Cell><Data ss:Type="String">${date}</Data></Cell></Row>`;
  xml+=`<Row/>`;

  xml+=`<Row>
<Cell><Data ss:Type="String">N</Data></Cell>
<Cell><Data ss:Type="String">Account_debet</Data></Cell>
<Cell><Data ss:Type="String">Account_credit</Data></Cell>
<Cell><Data ss:Type="String">Amount</Data></Cell>
<Cell><Data ss:Type="String">KNP</Data></Cell>
<Cell><Data ss:Type="String">Purpose</Data></Cell>
<Cell><Data ss:Type="String">Status</Data></Cell>
<Cell><Data ss:Type="String">Filial</Data></Cell>
</Row>`;

  document.querySelectorAll("#tbl tbody tr").forEach(r=>{
    const cnt=+r.querySelector(".count").value;
    if(!cnt) return;

    const n=r.querySelector(".rowNo").innerText;
    const filial=r.children[1].innerText;
    const debit=ACCOUNTS[type][filial];
    if(!debit) return;

    const amt=r.querySelector(".amt").innerText;
    const purpose=`${filial} ${date} № ${inv} СФга асосан ${type} хизмати харажатга олинди. (Маълумот сони: ${cnt})`;

    xml+=`<Row>
<Cell><Data ss:Type="Number">${n}</Data></Cell>
<Cell><Data ss:Type="String">${debit}</Data></Cell>
<Cell><Data ss:Type="String">${CREDIT}</Data></Cell>
<Cell><Data ss:Type="Number">${amt}</Data></Cell>
<Cell><Data ss:Type="String"></Data></Cell>
<Cell><Data ss:Type="String">${purpose}</Data></Cell>
<Cell><Data ss:Type="String"></Data></Cell>
<Cell><Data ss:Type="String">${filial}</Data></Cell>
</Row>`;
  });

  xml+=`</Table></Worksheet></Workbook>`;

  const blob=new Blob([xml],{type:"application/vnd.ms-excel"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="Trans_import.xls";
  a.click();
}

// Бошланғич юкланиш
buildTable();
