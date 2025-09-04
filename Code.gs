const SHEET_ID = '1nd7ILniGFRTDcqs15QwoRKWyNabsz_6BAzSiXfa4skQ';
const TAB_NAME = 'OS';
const HEADERS = ['meta','cliente','itens','totais'];

function getSheet(){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB_NAME);
  if(!sh){
    sh = ss.insertSheet(TAB_NAME);
    sh.appendRow(HEADERS);
  }
  const firstRow = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if(firstRow.join() !== HEADERS.join()){
    sh.clear();
    sh.appendRow(HEADERS);
  }
  return sh;
}

function saveOS(payload){
  const sh = getSheet();
  const os = payload?.meta?.os;
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),HEADERS.length).getValues();
  let idx = -1;
  if(os){
    idx = rows.findIndex(r=>{
      try{ return JSON.parse(r[0]).os == os; }catch(e){ return false; }
    });
  }
  const row = [
    JSON.stringify(payload.meta || {}),
    JSON.stringify(payload.cliente || {}),
    JSON.stringify(payload.itens || []),
    JSON.stringify(payload.totais || {})
  ];
  if(idx >= 0){
    sh.getRange(idx+2,1,1,row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function loadOS(os){
  if(!os) return null;
  const sh = getSheet();
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),HEADERS.length).getValues();
  for(const r of rows){
    try{
      const meta = JSON.parse(r[0] || '{}');
      if(String(meta.os) === String(os)){
        return {
          meta,
          cliente: JSON.parse(r[1] || '{}'),
          itens: JSON.parse(r[2] || '[]'),
          totais: JSON.parse(r[3] || '{}')
        };
      }
    }catch(e){ }
  }
  return null;
}

function getNextOS(){
  const sh = getSheet();
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),1).getValues();
  let max = 0;
  rows.forEach(r=>{
    try{
      const meta = JSON.parse(r[0] || '{}');
      const n = parseInt(meta.os,10);
      if(n > max) max = n;
    }catch(e){}
  });
  return max + 1;
}

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index');
}
