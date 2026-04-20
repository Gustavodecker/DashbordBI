// =====================================================================
//  CONFIGURAÇÃO
// =====================================================================
var FOLDER_ID      = "1BcnH3PktM90m4FXdAV-gHp1VL-vH_8n9";
var NOME_CONSINCO  = "Dashboard Consinco";
var NOME_WINTHOR   = "Dashboard Winthor";
var NOME_MODA      = "Dashboard TOTVSMODA";
var CCODES_EXCLUIR = ["CSEINF","CVAREJO","CSGTST","CDAVIP","CDBAAS","CDEVOPS","CDEVOP"];

// =====================================================================
//  ROTA PRINCIPAL
// =====================================================================
function doGet(e) {
  var acao = (e && e.parameter && e.parameter.acao) ? e.parameter.acao : "";
  if (acao==="dados_consinco") return json(processarConsinco());
  if (acao==="dados_winthor")  return json(processarWinthor());
  if (acao==="dados_moda")     return json(processarModa());
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Dashboard BI — Consinco, Winthor & TOTVS Moda")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function json(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
//  LER ARQUIVO DO DRIVE — detecta CSV ou XLSX automaticamente
// =====================================================================
function lerDoDrive(nome) {
  var pasta = DriveApp.getFolderById(FOLDER_ID);

  // Tenta cada extensão na ordem
  var extensoes = [".csv", ".xlsx", ""];
  var arq = null;
  var ehCsv = false;

  for (var i = 0; i < extensoes.length; i++) {
    var iter = pasta.getFilesByName(nome + extensoes[i]);
    if (iter.hasNext()) {
      arq = iter.next();
      // Detecta se é CSV pelo nome do arquivo ou pelo mimeType
      var mime = arq.getMimeType();
      ehCsv = (mime === "text/csv" || mime === "text/plain" ||
               arq.getName().toLowerCase().indexOf(".csv") !== -1);
      break;
    }
  }

  if (!arq) throw new Error("Arquivo não encontrado: " + nome);

  // ── CSV: lê direto como texto ──────────────────────────────────────
  if (ehCsv) {
    var texto = arq.getBlob().getDataAsString("UTF-8");
    var linhas = texto.split("\n").map(function(l){ return l.replace(/\r/,""); }).filter(function(l){ return l.trim(); });

    // Detecta separador — primeira linha pode ser "sep=,"
    var sep = ";";
    if (linhas[0].toLowerCase().trim() === "sep=,") {
      sep = ",";
      linhas = linhas.slice(1);
    } else if (linhas[0].toLowerCase().trim() === "sep=;") {
      sep = ";";
      linhas = linhas.slice(1);
    } else {
      // Auto-detect: conta vírgulas vs ponto-e-vírgula no header
      var nVirgulas = (linhas[0].match(/,/g)||[]).length;
      var nPontoVirgulas = (linhas[0].match(/;/g)||[]).length;
      sep = nVirgulas > nPontoVirgulas ? "," : ";";
    }

    return csvParaObjetos(linhas, sep);
  }

  // ── XLSX: converte para Google Sheets para leitura ─────────────────
  var copia = arq.makeCopy("_bi_tmp_" + nome, DriveApp.getRootFolder());
  var valores;

  try {
    valores = SpreadsheetApp.openById(copia.getId()).getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(copia.getId()).setTrashed(true);
  } catch(e) {
    // Limpa a cópia que falhou
    try { DriveApp.getFileById(copia.getId()).setTrashed(true); } catch(e2) {}

    // Fallback: converte via API do Drive
    var token    = ScriptApp.getOAuthToken();
    var boundary = "bi_b_" + new Date().getTime();
    var meta     = JSON.stringify({ name: "_bi_imp_" + new Date().getTime(), mimeType: "application/vnd.google-apps.spreadsheet" });
    var body     = "--"+boundary+"\r\nContent-Type: application/json\r\n\r\n"+meta+
                   "\r\n--"+boundary+"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n"+
                   "Content-Transfer-Encoding: base64\r\n\r\n"+
                   Utilities.base64Encode(arq.getBlob().getBytes())+"\r\n--"+boundary+"--";
    var resp     = UrlFetchApp.fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
      { method:"post", contentType:"multipart/related; boundary="+boundary,
        payload:body, headers:{ Authorization:"Bearer "+token }, muteHttpExceptions:true }
    );
    var fid = JSON.parse(resp.getContentText()).id;
    valores = SpreadsheetApp.openById(fid).getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(fid).setTrashed(true);
  }

  return xlsxParaObjetos(valores);
}

// =====================================================================
//  PARSERS
// =====================================================================
function xlsxParaObjetos(valores) {
  var headers = valores[0].map(function(h){ return String(h).trim().toLowerCase().replace(/"/g,""); });
  var rows = [];
  for (var i = 1; i < valores.length; i++) {
    if (!valores[i][0] && !valores[i][1]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = valores[i][j];
    rows.push(obj);
  }
  return rows;
}

function csvParaObjetos(linhas, sep) {
  function parseLinha(linha, s) {
    var result=[], cur="", inQ=false;
    for (var i=0; i<linha.length; i++) {
      var c=linha[i];
      if (c==='"') { inQ=!inQ; }
      else if (c===s && !inQ) { result.push(cur.replace(/^"|"$/g,"")); cur=""; }
      else cur+=c;
    }
    result.push(cur.replace(/^"|"$/g,""));
    return result;
  }

  var headers = parseLinha(linhas[0], sep).map(function(h){ return h.trim().toLowerCase().replace(/"/g,""); });
  var rows = [];
  for (var i = 1; i < linhas.length; i++) {
    if (!linhas[i].trim()) continue;
    var vals = parseLinha(linhas[i], sep);
    if (!vals[0] && !vals[1]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = vals[j] || "";
    rows.push(obj);
  }
  return rows;
}

// =====================================================================
//  HELPERS COMUNS
// =====================================================================
function filtrar(rows) {
  return rows.filter(function(r) {
    var cc = String(r["ccode"] || "").trim().replace(/"/g,"");
    return CCODES_EXCLUIR.indexOf(cc) === -1 && String(r["cliente"]||"").trim() !== "";
  });
}

function parseFlavor(f) {
  var m = String(f).match(/^(\d+)vcpu(\d+)memo$/);
  return m ? [parseInt(m[1]), parseInt(m[2])] : [0, 0];
}

function grupar(rows, envKey) {
  var g = {};
  rows.forEach(function(r) {
    var key = String(r["cliente"]).trim() + "|||" + String(r[envKey]||"unknown").trim();
    if (!g[key]) g[key] = { c: String(r["cliente"]).trim(), e: String(r[envKey]||"unknown").trim(), rows: [] };
    g[key].rows.push(r);
  });
  return g;
}

function montarDetalhe(rows, cpuKey, memKey, ggKey, flavorKey) {
  var m = {};
  rows.forEach(function(r) {
    var st = String(r["service_type"]).trim();
    var cpu = parseFloat(r[cpuKey]) || 0;
    var mem = parseFloat(r[memKey]) || 0;
    if (!m[st]) m[st] = { tipo:st, qtd:0, cpu_total:0, mem_total:0,
      cpu_unit: cpu, mem_unit: mem,
      gg: ggKey ? (parseInt(r[ggKey])||0) : 0,
      flavor: String(r[flavorKey]||"") };
    m[st].qtd++;
    m[st].cpu_total += cpu;
    m[st].mem_total += mem;
  });
  return Object.keys(m).map(function(k) {
    m[k].cpu_total = Math.round(m[k].cpu_total);
    m[k].mem_total = Math.round(m[k].mem_total);
    return m[k];
  });
}

// =====================================================================
//  PROCESSAR CONSINCO
// =====================================================================
function processarConsinco() {
  var rows = filtrar(lerDoDrive(NOME_CONSINCO));
  rows.forEach(function(r) { var cf = parseFlavor(r["flavor"]); r._cpu = cf[0]; r._mem = cf[1]; });
  var grupos = grupar(rows, "topology_env");

  return Object.keys(grupos).map(function(key) {
    var g       = grupos[key];
    var relay   = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="relay_instance"; });
    var desktop = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="desktop_instance"; });
    var pdv     = g.rows.filter(function(r){ return String(r["service_type"]).trim().indexOf("pdv_instance")===0; });
    var ndd     = g.rows.filter(function(r){ return String(r["service_type"]).trim().indexOf("ndd_instance")===0; });
    var core    = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="core_instance"; });

    var gg=0, regiao=String(g.rows[0]["region"]||"").trim();
    if (relay.length>0)        { gg=parseInt(relay[0]["gg_conections"])||0;   regiao=String(relay[0]["region"]||"").trim(); }
    else if (desktop.length>0) { gg=parseInt(desktop[0]["gg_conections"])||0; regiao=String(desktop[0]["region"]||"").trim(); }

    return { c:g.c, e:g.e, r:regiao, gg:gg,
      gg_pdv: pdv.length>0 ? (parseInt(pdv[0]["gg_conections"])||0) : 0,
      gg_ndd: ndd.length>0 ? (parseInt(ndd[0]["gg_conections"])||0) : 0,
      n_relay:relay.length, n_desktop:desktop.length, n_pdv:pdv.length, n_ndd:ndd.length, n_core:core.length,
      total_cpu: g.rows.reduce(function(a,r){ return a+r._cpu; }, 0),
      total_mem: g.rows.reduce(function(a,r){ return a+r._mem; }, 0),
      total_srv: g.rows.length,
      detalhe: montarDetalhe(g.rows, "_cpu", "_mem", "gg_conections", "flavor") };
  });
}

// =====================================================================
//  PROCESSAR WINTHOR
// =====================================================================
function processarWinthor() {
  var rows = filtrar(lerDoDrive(NOME_WINTHOR));
  var grupos = grupar(rows, "topology_env");

  return Object.keys(grupos).map(function(key) {
    var g          = grupos[key];
    var coreRelay  = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="core_instance" && String(r["gg_is_relay"]).trim()==="True"; });
    var coreAll    = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="core_instance"; });
    var elastic    = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="elastic_service"; });
    var pdv        = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="pdv_instance"; });
    var coletor    = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="elastic_coletor"; });

    var gg=0, regiao=String(g.rows[0]["region"]||"").trim();
    if (coreRelay.length>0)    { gg=parseInt(coreRelay[0]["gg_conections"])||0; regiao=String(coreRelay[0]["region"]||"").trim(); }
    else if (elastic.length>0) { gg=parseInt(elastic[0]["gg_conections"])||0;   regiao=String(elastic[0]["region"]||"").trim(); }

    // Memória Winthor vem de wta_memory_min/max
    var totMin = g.rows.reduce(function(a,r){ return a+(parseFloat(r["wta_memory_min"])||0); }, 0);
    var totMax = g.rows.reduce(function(a,r){ return a+(parseFloat(r["wta_memory_max"])||0); }, 0);

    var dm = {};
    g.rows.forEach(function(r) {
      var st = String(r["service_type"]).trim();
      var mm = parseFloat(r["wta_memory_min"])||0, mx = parseFloat(r["wta_memory_max"])||0, pg = parseFloat(r["pagefile"])||0;
      if (!dm[st]) dm[st] = { tipo:st, qtd:0, mem_min:mm, mem_max:mx, mem_total_min:0, mem_total_max:0, pagefile:pg, gg:parseInt(r["gg_conections"])||0 };
      dm[st].qtd++; dm[st].mem_total_min+=mm; dm[st].mem_total_max+=mx;
    });

    return { c:g.c, e:g.e, r:regiao, gg:gg,
      gg_pdv: pdv.length>0    ? (parseInt(pdv[0]["gg_conections"])||0)     : 0,
      gg_col: coletor.length>0 ? (parseInt(coletor[0]["gg_conections"])||0) : 0,
      n_relay:coreRelay.length, n_elastic:elastic.length, n_pdv:pdv.length, n_coletor:coletor.length, n_core:coreAll.length,
      total_mem_min: Math.round(totMin), total_mem_max: Math.round(totMax),
      total_srv: g.rows.length,
      detalhe: Object.keys(dm).map(function(k){ return dm[k]; }) };
  });
}

// =====================================================================
//  PROCESSAR TOTVS MODA
// =====================================================================
function processarModa() {
  var rows = filtrar(lerDoDrive(NOME_MODA));
  // Moda usa "environment" (não "topology_env") e "datacenter" (não "region")
  var grupos = grupar(rows, "environment");

  return Object.keys(grupos).map(function(key) {
    var g       = grupos[key];
    var uniface = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="uniface_instance"; });
    var core    = g.rows.filter(function(r){ return String(r["service_type"]).trim()==="core_instance"; });

    var regiao = String(g.rows[0]["datacenter"]||"").trim();
    if (uniface.length>0)    regiao = String(uniface[0]["datacenter"]||"").trim();
    else if (core.length>0)  regiao = String(core[0]["datacenter"]||"").trim();

    return { c:g.c, e:g.e, r:regiao,
      gg: 0,
      n_uniface: uniface.length, n_core: core.length,
      total_cpu: g.rows.reduce(function(a,r){ return a+(parseInt(r["cpu_cores"])||0); }, 0),
      total_mem: g.rows.reduce(function(a,r){ return a+(parseFloat(r["ram_gb"])||0); }, 0),
      total_srv: g.rows.length,
      detalhe: montarDetalhe(g.rows, "cpu_cores", "ram_gb", null, "flavor_name") };
  });
}
