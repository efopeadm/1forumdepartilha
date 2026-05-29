/**
 * SISTEMA DE INSCRIÇÃO - FÓRUM EJA POLO 5
 * Configurado para a planilha: 1j16U_FB4OF7ep-w0Weq3MDNLqIGRJChrCI8X1_t6Svs
 */

const SPREADSHEET_ID = '1j16U_FB4OF7ep-w0Weq3MDNLqIGRJChrCI8X1_t6Svs';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inscrição - 1º Fórum EJA Polo 5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Retorna a URL do script para recarregar a página corretamente
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Função para buscar dados das listas suspensas e calcular vagas ocupadas
function getFormData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dadosSheet = ss.getSheetByName('Dados');
  const inscSheet = ss.getSheetByName('Inscricoes');
  
  const data = dadosSheet.getDataRange().getValues();
  const inscritos = inscSheet.getDataRange().getValues(); 
  
  // Remove o cabeçalho da aba Dados
  data.shift();
  
  const listas = {
    secretarias: [...new Set(data.map(r => r[0]).filter(String))],
    cargos: [...new Set(data.map(r => r[1]).filter(String))],
    datas: [...new Set(data.map(r => r[2]).filter(String))],
    eixos: [...new Set(data.map(r => r[3]).filter(String))]
  };

  // Contagem de vagas por Eixo (Limite de 30)
  // O Eixo está na coluna H (índice 7)
  const contagemEixos = {};
  listas.eixos.forEach(eixo => {
    const totalInscritosNoEixo = inscritos.filter(row => row[7] === eixo).length;
    const vagasRestantes = 30 - totalInscritosNoEixo;
    contagemEixos[eixo] = vagasRestantes > 0 ? vagasRestantes : 0;
  });

  return { listas, contagemEixos };
}

// Processa o envio do formulário com trava de concorrência e envio de e-mail
function processForm(formObject) {
  const lock = LockService.getScriptLock();
  
  try {
    // Tenta obter o bloqueio por até 10 segundos para evitar exceder vagas
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const inscSheet = ss.getSheetByName('Inscricoes');
    const dadosInscritos = inscSheet.getDataRange().getValues();
    
    // Verificação dupla
    const totalEixo = dadosInscritos.filter(row => row[7] === formObject.eixo).length;
    
    if (totalEixo >= 30) {
      return { sucesso: false, msg: "ERRO: As vagas para este eixo acabaram de se esgotar. Por favor, selecione outro eixo." };
    }

    // Adiciona a linha na planilha
    inscSheet.appendRow([
      new Date(),              
      formObject.nome,         
      "'" + formObject.cpf,    
      formObject.email,        
      formObject.secretaria,   
      formObject.cargo,        
      formObject.dataPart,     
      formObject.eixo          
    ]);

    // ENVIO DO E-MAIL DE CONFIRMAÇÃO
    const assuntoEmail = "Confirmação de Inscrição: 1º Fórum de Partilha de Práticas da EJA - Polo 5";
    const corpoEmail = `Olá, ${formObject.nome}!
    
Sua inscrição para o 1º Fórum de Partilha de Práticas da EJA - Polo 5 (Ouvintes) foi confirmada com sucesso.

📍 DETALHES DA SUA PARTICIPAÇÃO:
- Data: ${formObject.dataPart}
- Eixo Escolhido: ${formObject.eixo}
- Local: CENFORPE / Secretaria de Educação de São Bernardo do Campo
- Endereço: Av. Dom Jaime de Barros Câmara, 201 - Planalto, São Bernardo do Campo - SP
- Horário: 19h às 22h

Por favor, guarde este e-mail. A sua participação é fundamental para fortalecer a EJA em nossa região!

Em caso de dúvidas, entre em contato através do e-mail: pactoejapolo5@gmail.com

Abraços cordiais,
Comissão Organizadora`;

    MailApp.sendEmail(formObject.email, assuntoEmail, corpoEmail);
    
    return { sucesso: true, msg: "Sucesso! Sua inscrição foi confirmada e um e-mail de confirmação foi enviado." };

  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}