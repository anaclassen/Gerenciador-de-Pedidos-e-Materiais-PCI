function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Solicitações e Status - Placas PCI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Respostas ao formulário 1");
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

// ======================================================
// BUSCA O ESTOQUE NA PLANILHA EXTERNA PELO ID DO GOOGLE DRIVE
// ======================================================
// ======================================================
// BUSCA O ESTOQUE APENAS NAS ABAS ESPECÍFICAS (A PARTIR DA LINHA 3)
// ======================================================
// ======================================================
// COLUNA "A" E "D" PARA A ABA "Brocas Fresa TTP" E "A" e "D" PARA AS DEMAIS
// ======================================================
function buscarEstoque() {
  try {
    const ID_PLANILHA_ESTOQUE = "Coloque aqui o ID da planilha google";
    const ssEstoque = SpreadsheetApp.openById(ID_PLANILHA_ESTOQUE);
    
    const abasDesejadas = ["Brocas Fresa TTP", "Placas", "Fresa LPKF"];
    const itens = [];

    abasDesejadas.forEach(nomeAba => {
      const sheet = ssEstoque.getSheetByName(nomeAba);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        
        // Começa em i = 2 (Linha 3 da planilha)
        for (let i = 2; i < data.length; i++) {
          const row = data[i];
          const nome = row[0] ? row[0].toString().trim() : "";
          
          if (nome !== "" && nome.toLowerCase() !== "nome") {
            const qtdEstoque = row[1] !== undefined && row[1] !== "" ? row[1] : 0;
            
            let status = "Disponível";
            if (qtdEstoque <= 0) {
              status = "Esgotado";
            } else if (qtdEstoque <= 5) {
              status = "Baixo";
            }

            // Se for a aba "Placas", pega a Coluna E (índice 4). 
            // Para as demais, pega a Coluna D (índice 3).
            let descricaoReal = '-';
            if (nomeAba === "Brocas Fresa TTP") {
              descricaoReal = row[3] ? row[3].toString().trim() : '-';
            } else {              
              descricaoReal = row[4] ? row[4].toString().trim() : '-';
            }

            itens.push({
              nome: nome,                                   // Coluna A
              descricao: descricaoReal,                     // Coluna E (Placas) ou D (Outras)
              uso: row[2] ? row[2].toString().trim() : '-',  // Coluna C
              status: status
            });
          }
        }
      }
    });

    return itens;
  } catch (erro) {
    throw new Error("Erro ao carregar o estoque: " + erro.message);
  }
}

// ======================================================
// BUSCA SOLICITAÇÕES DE PEDIDOS
// ======================================================
function buscarSolicitacoes(busca) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const resultados = [];
    
    if (!busca) return [];
    
    const termo = busca.toString().trim().toLowerCase();
    
    // Ordem decrescente (da mais recente para a mais antiga)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      
      const item = row[0] !== undefined && row[0] !== null ? row[0].toString().trim() : "";
      const emailAluno = row[6] !== undefined && row[6] !== null ? row[6].toString().trim().toLowerCase() : "";
      const emailProf = row[9] !== undefined && row[9] !== null ? row[9].toString().trim().toLowerCase() : "";
      const nomeAluno = row[10] !== undefined && row[10] !== null ? row[10].toString().trim().toLowerCase() : "";

      if (item !== "" && (item.toLowerCase() === termo || emailAluno.includes(termo) || emailProf.includes(termo) || nomeAluno.includes(termo))) {
        resultados.push({
          item: row[0],
          status: row[1] || 'Pendente',
          dataAgendada: row[2] ? formatDate(row[2]) : '-',
          maquina: row[3] || '-',
          tecnico: row[4] || '-',
          carimbo: row[5] ? formatDate(row[5]) : '-',
          emailAluno: row[6] || '-',
          projeto: row[7] || '-',
          qtd: row[8] || '1',
          emailProfessor: row[9] || '-',
          nomeAluno: row[10] || '-',
          //observacoes: row[11] || '-',
          //retrabalho: row[13] || 'Não',
          largura: row[14] || '-',
          comprimento: row[15] || '-',
          //arquivoUrl: row[16] || ''
        });
      }
    }
    return resultados;
  } catch (erro) {
    throw new Error("Erro na busca: " + erro.message);
  }
}

// ======================================================
// CADASTRA NOVA SOLICITAÇÃO
// ======================================================
function cadastrarSolicitacao(dados) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    const novoItem = lastRow > 1 ? lastRow - 1 : 1; 
    const carimbo = new Date();
    let arquivoUrl = "";
    
    if (dados.fileData && dados.fileName) {
      const bytes = Utilities.base64Decode(dados.fileData.split(",")[1]);
      const blob = Utilities.newBlob(bytes, dados.fileMimeType, dados.fileName);
      
      const folders = DriveApp.getFoldersByName("Projetos PCI Uploads");
      let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Projetos PCI Uploads");
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      arquivoUrl = file.getUrl();
    }

    const novaLinha = [];
    novaLinha[0] = novoItem;           // Coluna A: Item
    novaLinha[1] = "Pendente";         // Coluna B: Status
    novaLinha[2] = "";                 // Coluna C: Data
    novaLinha[3] = "";                 // Coluna D: Máquina
    novaLinha[4] = "";                 // Coluna E: Técnico
    novaLinha[5] = carimbo;            // Coluna F: Carimbo
    novaLinha[6] = dados.emailAluno;   // Coluna G: E-mail Aluno
    novaLinha[7] = dados.projeto;      // Coluna H: Projeto
    novaLinha[8] = dados.qtd;          // Coluna I: Quantidade
    novaLinha[9] = dados.emailProf;    // Coluna J: E-mail Professor
    novaLinha[10] = dados.nomeAluno;   // Coluna K: Nome Completo do Aluno
    novaLinha[11] = dados.observacoes; // Coluna L: Observações
    novaLinha[13] = dados.retrabalho;  // Coluna N: Retrabalho
    novaLinha[14] = dados.largura;     // Coluna O: Largura
    novaLinha[15] = dados.comprimento; // Coluna P: Comprimento
    novaLinha[16] = arquivoUrl;        // Coluna Q: Anexo ZIP

    sheet.appendRow(novaLinha);
    
    return { sucesso: true, item: novoItem };
  } catch (erro) {
    throw new Error("Erro no cadastro: " + erro.message);
  }
}

function formatDate(dateObj) {
  if (dateObj instanceof Date) {
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return dateObj;
}
