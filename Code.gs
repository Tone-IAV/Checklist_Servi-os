const SHEET_ID = '1nd7ILniGFRTDcqs15QwoRKWyNabsz_6BAzSiXfa4skQ';
const TAB_NAME = 'OS';
const HEADERS = ['meta','veiculo','checklist','itens','totais'];

const VEHICLE_SHEET_ID = '1fGx1JHUVqZNKtEdwwErZYegHaDm3hH4TQhXeMfd4MW4';
const VEHICLE_TAB_NAME = 'VEICULOS';
const VEHICLE_COLUMN_MAP = {
  'placa':'placa',
  'veículos':'veiculo',
  'categoria':'categoria',
  'ano':'ano',
  'mod':'modelo',
  'placa anterior':'placaAnterior',
  'renavam':'renavam',
  'chassi':'chassi',
  'fabricante':'fabricante',
  'tipo de veículo':'tipoVeiculo',
  'cor':'cor',
  'combustivel':'combustivel',
  'dt. aquisição':'dataAquisicao',
  'dt. venda':'dataVenda',
  'centro de custo':'centroCusto',
  'situação atual':'situacaoAtual',
  'situação':'situacao',
  'km 2024':'km2024',
  'km 2023':'km2023',
  'ultimo checklist recebido':'ultimoChecklist',
  'categoria checklist':'categoriaChecklist',
  'total km rodado':'totalKmRodado',
  'responsável':'responsavel',
  'cidade':'cidade',
  'uf':'uf',
  'empresa':'empresa',
  'motorista':'motorista',
  'motorista 2':'motorista2',
  'telefone':'telefone',
  'telefone²':'telefone2'
};

function getSheet(){
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB_NAME);
  if(!sh){
    sh = ss.insertSheet(TAB_NAME);
    sh.appendRow(HEADERS);
  }
  const lastCol = Math.max(sh.getLastColumn(), HEADERS.length);
  const firstRow = sh.getRange(1,1,1,lastCol).getValues()[0];
  if(firstRow[1] === 'cliente'){
    const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0), Math.max(firstRow.length, HEADERS.length)).getValues();
    sh.clear();
    sh.appendRow(HEADERS);
    if(rows.length){
      const normalized = rows.map(r=>{
        const row = new Array(HEADERS.length).fill('');
        for(let i=0;i<HEADERS.length;i++){
          row[i] = r[i] || '';
        }
        return row;
      });
      sh.getRange(2,1,normalized.length,HEADERS.length).setValues(normalized);
    }
  } else if(firstRow.join() !== HEADERS.join()){
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
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),HEADERS.length).getValues();
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

function getVehicleHistory(placa){
  if(!placa) return {historico:[]};
  const sh = getSheet();
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),HEADERS.length).getValues();
  const historico=[];
  rows.forEach(r=>{
    try{
      const meta = JSON.parse(r[0]||'{}');
      const veiculo = JSON.parse(r[1]||'{}');
      if(String(veiculo.placa||'').toUpperCase() === String(placa||'').toUpperCase()){
        historico.push({meta});
      }
    }catch(e){}
  });
  return {historico};
}

function normalizeHeaderName(name){
  return String(name||'').trim().toLowerCase();
}

function formatVehicleValue(value){
  if(value instanceof Date){
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return value;
}

function getVehicles(){
  try{
    const ss = SpreadsheetApp.openById(VEHICLE_SHEET_ID);
    const sh = ss.getSheetByName(VEHICLE_TAB_NAME);
    if(!sh) return [];
    const values = sh.getDataRange().getValues();
    if(!values.length) return [];
    const headers = values.shift();
    const indexMap = {};
    headers.forEach((header, idx)=>{
      const key = VEHICLE_COLUMN_MAP[normalizeHeaderName(header)];
      if(key) indexMap[key] = idx;
    });
    if(indexMap.placa == null) return [];
    return values
      .filter(row => row[indexMap.placa])
      .map(row => {
        const obj = {};
        Object.entries(indexMap).forEach(([key, colIdx])=>{
          obj[key] = formatVehicleValue(row[colIdx]);
        });
        return obj;
      });
  }catch(err){
    return [];
  }
}

function getDashboard(){
  const sh = getSheet();
  const rows = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),1).getValues();
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
  return HtmlService.createHtmlOutputFromFile('Index');
}
