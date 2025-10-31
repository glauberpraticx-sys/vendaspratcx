function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action;
  
  try {
    switch(action) {
      case 'getVendas':
        return ContentService.createTextOutput(JSON.stringify(getVendas()))
          .setMimetype(ContentService.MimeType.JSON);
      
      case 'addVenda':
        return ContentService.createTextOutput(JSON.stringify(addVenda(e.parameter)))
          .setMimetype(ContentService.MimeType.JSON);
          
      case 'deleteVenda':
        return ContentService.createTextOutput(JSON.stringify(deleteVenda(e.parameter)))
          .setMimetype(ContentService.MimeType.JSON);
          
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Ação não reconhecida'
        })).setMimetype(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimetype(ContentService.MimeType.JSON);
  }
}

function getVendas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const vendasSheet = sheet.getSheetByName('Vendas');
  const parcelasSheet = sheet.getSheetByName('Parcelas');
  
  const vendasData = vendasSheet.getDataRange().getValues();
  const parcelasData = parcelasSheet.getDataRange().getValues();
  
  const headersVendas = vendasData[0];
  const headersParcelas = parcelasData[0];
  
  const vendas = [];
  
  for (let i = 1; i < vendasData.length; i++) {
    const vendaRow = vendasData[i];
    const venda = {};
    
    // Mapear dados da venda
    for (let j = 0; j < headersVendas.length; j++) {
      venda[headersVendas[j].toLowerCase()] = vendaRow[j];
    }
    
    // Buscar parcelas desta venda
    venda.parcelas = [];
    for (let k = 1; k < parcelasData.length; k++) {
      const parcelaRow = parcelasData[k];
      if (parcelaRow[0] === venda.id) {
        const parcela = {};
        for (let l = 0; l < headersParcelas.length; l++) {
          parcela[headersParcelas[l].toLowerCase()] = parcelaRow[l];
        }
        venda.parcelas.push(parcela);
      }
    }
    
    vendas.push(venda);
  }
  
  return {
    success: true,
    data: vendas
  };
}

function addVenda(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const vendasSheet = sheet.getSheetByName('Vendas');
  const parcelasSheet = sheet.getSheetByName('Parcelas');
  
  try {
    // Gerar ID único
    const id = Utilities.getUuid();
    
    // Adicionar venda
    vendasSheet.appendRow([
      id,
      params.clientName,
      params.startDate,
      parseFloat(params.totalLicense),
      parseFloat(params.entradaValue),
      parseInt(params.parcelasQtd),
      parseFloat(params.monthlyValue),
      new Date()
    ]);
    
    // Calcular parcelas
    const valorParcela = (params.totalLicense - params.entradaValue) / params.parcelasQtd;
    
    // Adicionar entrada
    parcelasSheet.appendRow([
      id,
      'ENTRADA',
      parseFloat(params.entradaValue),
      parseFloat(params.entradaValue) * 0.15,
      params.startDate,
      'PENDENTE',
      0
    ]);
    
    // Adicionar parcelas
    for (let i = 1; i <= params.parcelasQtd; i++) {
      const parcelaDate = new Date(params.startDate);
      parcelaDate.setMonth(parcelaDate.getMonth() + i);
      
      parcelasSheet.appendRow([
        id,
        `PARCELA ${i}/${params.parcelasQtd}`,
        valorParcela,
        valorParcela * 0.15,
        Utilities.formatDate(parcelaDate, 'America/Sao_Paulo', 'yyyy-MM-dd'),
        'PENDENTE',
        i
      ]);
    }
    
    // Adicionar mensalidade
    const mensalidadeDate = new Date(params.startDate);
    mensalidadeDate.setMonth(mensalidadeDate.getMonth() + 1);
    
    parcelasSheet.appendRow([
      id,
      'MENSALIDADE',
      parseFloat(params.monthlyValue),
      parseFloat(params.monthlyValue) * 0.15,
      Utilities.formatDate(mensalidadeDate, 'America/Sao_Paulo', 'yyyy-MM-dd'),
      'PENDENTE',
      999
    ]);
    
    return {
      success: true,
      message: 'Venda cadastrada com sucesso!',
      id: id
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

function deleteVenda(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const vendasSheet = sheet.getSheetByName('Vendas');
  const parcelasSheet = sheet.getSheetByName('Parcelas');
  
  try {
    const id = params.id;
    
    // Deletar parcelas
    const parcelasData = parcelasSheet.getDataRange().getValues();
    for (let i = parcelasData.length - 1; i >= 1; i--) {
      if (parcelasData[i][0] === id) {
        parcelasSheet.deleteRow(i + 1);
      }
    }
    
    // Deletar venda
    const vendasData = vendasSheet.getDataRange().getValues();
    for (let i = vendasData.length - 1; i >= 1; i--) {
      if (vendasData[i][0] === id) {
        vendasSheet.deleteRow(i + 1);
        break;
      }
    }
    
    return {
      success: true,
      message: 'Venda excluída com sucesso!'
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}