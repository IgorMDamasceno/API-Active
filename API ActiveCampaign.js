function criarCampanhas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Automatização');
  var controleSheet = ss.getSheetByName('Controle_Campanhas');

  if (!controleSheet) {
    controleSheet = ss.insertSheet('Controle_Campanhas');
    controleSheet.appendRow(['Identificador_Unico', 'Message_ID', 'Campaign_ID']);
  }

  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Cabeçalhos
  Logger.log('[DEBUG] Cabeçalhos: ' + JSON.stringify(headers));

  data.forEach(function(row, index) {
  Logger.log('[DEBUG] Processando linha: ' + JSON.stringify(row));
  if (row.length != headers.length) {
    Logger.log('[DEBUG] Linha inconsistente, pulando...');
    return;
  }

  // Pular linhas que já foram marcadas como "Já criada anteriormente"
  var statusAtual = row[headers.indexOf('Status')];
  if (statusAtual === "Já criada anteriormente") {
    Logger.log('[DEBUG] Linha já marcada como "Já criada anteriormente", pulando...');
    return;
  }

  var rowData = {};
  sheet.getRange(index + 2, 1).setNumberFormat("dd/MM/yyyy");
  headers.forEach(function(header, i) {
    rowData[header] = row[i];
  });

  var spreadsheet_id = rowData['ID'];
  var conta = rowData['Conta'];
  var dataRaw = rowData['Data'];
  var urlActive = rowData['UrlActive'];

  if (!spreadsheet_id || !conta || !dataRaw || !urlActive) {
    Logger.log('[DEBUG] Dados incompletos para criação. Pulando linha.');
    atualizarStatus(sheet, index + 2, "Dados incompletos");
    return;
  }

  var dataObj;
  if (typeof dataRaw === 'string') {
    if (dataRaw.includes('/')) {
      // dd/MM/yyyy
      var partes = dataRaw.split('/');
      dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
    } else if (dataRaw.includes('-')) {
      // yyyy-MM-dd
      var partes = dataRaw.split('-');
      dataObj = new Date(partes[0], partes[1] - 1, partes[2]);
    } else {
      dataObj = new Date(dataRaw); // tentativa padrão
    }
  } else {
    dataObj = new Date(dataRaw); // já é Date
  }

var dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');


  if (!urlActive.startsWith('http')) {
    urlActive = 'https://' + urlActive;
  }

  var identificadorUnico = spreadsheet_id + '|' + dataFormatada + '|' + conta;

  var controleValues = controleSheet.getDataRange().getValues();
  var jaCriado = controleValues.some(function(c) {
    return c[0] == identificadorUnico;
  });
  if (jaCriado) {
    Logger.log('[DEBUG] Já processado, pulando...');
    atualizarStatus(sheet, index + 2, "Já criada anteriormente");
    return;
  }

  var mensagem = criarMensagem(rowData, urlActive);
  if (!mensagem || !mensagem.result_code) {
    Logger.log('[DEBUG] Erro ao criar mensagem: ' + JSON.stringify(mensagem));
    atualizarStatus(sheet, index + 2, "Erro ao criar mensagem");
    return;
  }

  var campanha = criarCampanha(mensagem.id, rowData, urlActive);
  if (!campanha || !campanha.result_code) {
    Logger.log('[DEBUG] Erro ao criar campanha: ' + JSON.stringify(campanha));
    atualizarStatus(sheet, index + 2, "Erro ao criar campanha");
    return;
  }

  controleSheet.appendRow([identificadorUnico, mensagem.id, campanha.id]);
  Logger.log('[DEBUG] Campanha criada e registrada.');
  atualizarStatus(sheet, index + 2, "Campanha criada com sucesso");
});

  Logger.log('[DEBUG] Processamento concluído.');
}

function criarMensagem(data, urlActive) {
  Logger.log('[DEBUG] Criando mensagem com dados: ' + JSON.stringify(data));
  
  var payload = {
    api_action: 'message_add',
    api_output: 'json',
    format: 'html',
    subject: data['Assunto'],
    fromemail: data['Sender'],
    fromname: data['Remetente'],
    reply2: data['Sender'],
    priority: 3,
    charset: 'utf-8',
    encoding: 'quoted-printable',
    htmlconstructor: 'editor',
    html: data['HtmlEmail'],
  };
  
  payload['p[' + data['Lista id'] + ']'] = data['Lista id'];

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'API-TOKEN': data['KeyActive']
    },
    payload: payload,
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(urlActive + '/admin/api.php?api_action=message_add&api_output=json', options);
  var result = JSON.parse(response.getContentText());
  Logger.log('[DEBUG] Resposta mensagem: ' + JSON.stringify(result));
  return result;
}

function criarCampanha(messageId, data, urlActive) {
  Logger.log('[DEBUG] Criando campanha com dados: ' + JSON.stringify(data));

  var dataBruta = data['Data'];
var horaBruta = data['Hora'];

// Conversão segura da data
var dataObj;
if (typeof dataBruta === 'string') {
  if (dataBruta.includes('/')) {
    var partes = dataBruta.split('/');
    dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
  } else if (dataBruta.includes('-')) {
    var partes = dataBruta.split('-');
    dataObj = new Date(partes[0], partes[1] - 1, partes[2]);
  } else {
    dataObj = new Date(dataBruta);
  }
} else {
  dataObj = new Date(dataBruta);
}

// Conversão segura da hora
var horaObj;
if (typeof horaBruta === 'string') {
  var hParts = horaBruta.split(':');
  if (hParts.length === 2) hParts.push('00'); // se vier "11:30", vira "11:30:00"
  horaObj = new Date();
  horaObj.setHours(parseInt(hParts[0], 10), parseInt(hParts[1], 10), parseInt(hParts[2], 10));
} else if (horaBruta instanceof Date) {
  horaObj = horaBruta;
} else {
  horaObj = new Date(); // fallback
}

// Combinar data + hora
dataObj.setHours(horaObj.getHours());
dataObj.setMinutes(horaObj.getMinutes());
dataObj.setSeconds(horaObj.getSeconds());

var dataHora = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  var payload = {
    api_action: 'campaign_create',
    api_output: 'json',
    type: 'single',
    segmentid: data['SegmentoId'],
    name: data['Nomenclatura'],
    sdate: dataHora,
    status: 1,
    public: 1,
    trackreads: 1,
    htmlunsub: 1,
    textunsub: 1,
  };
  
  payload['p[' + data['Lista id'] + ']'] = data['Lista id'];
  payload['m[' + messageId + ']'] = 100;

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'API-TOKEN': data['KeyActive']
    },
    payload: payload,
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(urlActive + '/admin/api.php?api_action=campaign_create&api_output=json', options);
  var result = JSON.parse(response.getContentText());
  Logger.log('[DEBUG] Resposta campanha: ' + JSON.stringify(result));
  return result;
}

function atualizarStatus(sheet, rowNumber, statusMessage) {
  var headers = sheet.getDataRange().getValues()[0];
  var statusColumnIndex = headers.indexOf('Status');
  
  if (statusColumnIndex !== -1) {
    sheet.getRange(rowNumber, statusColumnIndex + 1).setValue(statusMessage);
  } else {
    Logger.log("[DEBUG] Coluna 'Status' não encontrada.");
  }
}
