const SHEET_ID = '1nd7ILniGFRTDcqs15QwoRKWyNabsz_6BAzSiXfa4skQ';
const TAB_NAME = 'OS';
const HEADERS = ['meta','veiculo','checklist','itens','totais'];
const DATA_SHEET_ID = '1fGx1JHUVqZNKtEdwwErZYegHaDm3hH4TQhXeMfd4MW4';

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
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2,1,last-1,HEADERS.length).getValues() : [];
  let idx = -1;
  if(os){
    idx = rows.findIndex(r=>{
      try{ return JSON.parse(r[0]).os == os; }catch(e){ return false; }
    });
  }
  const row = [
    JSON.stringify(payload.meta || {}),
    JSON.stringify(payload.veiculo || {}),
    JSON.stringify(payload.checklist || {}),
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
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2,1,last-1,HEADERS.length).getValues() : [];
  for(const r of rows){
    try{
      const meta = JSON.parse(r[0] || '{}');
      if(String(meta.os) === String(os)){
        return {
          meta,
          veiculo: JSON.parse(r[1] || '{}'),
          checklist: JSON.parse(r[2] || '{}'),
          itens: JSON.parse(r[3] || '[]'),
          totais: JSON.parse(r[4] || '{}')
        };
      }
    }catch(e){ }
  }
  return null;
}

function getNextOS(){
  const sh = getSheet();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2,1,last-1,1).getValues() : [];
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

/* ======== Hierarquia (Sistemas/Serviços/Subsistemas) ======== */
function readTable(tab){
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sh = ss.getSheetByName(tab);
  if(!sh) return {headers:[], rows:[]};
  const values = sh.getDataRange().getValues();
  if(values.length === 0) return {headers:[], rows:[]};
  const headers = values.shift().map(h=>String(h));
  const rows = values
    .filter(r=>r.some(c=>c!==''))
    .map(r=>{
      const o={};
      headers.forEach((h,i)=>{ o[h.toLowerCase()] = r[i]; });
      return o;
    });
  return {headers, rows};
}

function writeTable(tab, table){
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  let sh = ss.getSheetByName(tab);
  if(!sh) sh = ss.insertSheet(tab);
  sh.clearContents();
  const headers = table.headers || [];
  const rows = table.rows || [];
  if(headers.length){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  if(rows.length){
    const lower = headers.map(h=>h.toLowerCase());
    const data = rows.map(r=> lower.map(k=> r[k]||''));
    sh.getRange(2,1,data.length,headers.length).setValues(data);
  }
}

function getTables(){
  return {
    sistemas: readTable('SISTEMAS'),
    servicos: readTable('SERVIÇOS'),
    subsistemas: readTable('SUBSISTEMAS')
  };
}

function saveTables(data){
  if(data.sistemas) writeTable('SISTEMAS', data.sistemas);
  if(data.servicos) writeTable('SERVIÇOS', data.servicos);
  if(data.subsistemas) writeTable('SUBSISTEMAS', data.subsistemas);
  return true;
}

function getVeiculo(placa){
  if(!placa) return null;
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sh = ss.getSheetByName('VEICULOS');
  if(!sh) return null;
  const values = sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    const row = values[i];
    if(String(row[0]).toUpperCase() === placa.toUpperCase()){
      return {veiculo: row[1], ano: row[2], mod: row[3]};
    }
  }
  return null;
}

function listPlacas(){
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sh = ss.getSheetByName('VEICULOS');
  if(!sh) return [];
  const values = sh.getDataRange().getValues();
  const placas = [];
  for(let i=1;i<values.length;i++){
    const p = values[i][0];
    if(p) placas.push(p);
  }
  return placas;
}

function getDashboard(){
  const sh = getSheet();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2,1,last-1,1).getValues() : [];
  const c = {abertas:0, aguardando:0, concluidas:0};
  rows.forEach(r=>{
    try{
      const meta = JSON.parse(r[0]||'{}');
      const st = String(meta.status||'').toLowerCase();
      if(st==='aguardando peça') c.aguardando++;
      else if(st==='concluída'||st==='concluido'||st==='concluida') c.concluidas++;
      else c.abertas++;
    }catch(e){}
  });
  return c;
}

function doGet(){
  return HtmlService.createHtmlOutputFromFile("Index");
}

function getPage(page){
  return HtmlService.createHtmlOutputFromFile(page).getContent();
}
