// Substitua a função editarAssociado atual por esta versão no Apps Script.
// Ela mantém o e-mail como identificador, ignorando espaços e diferenças de maiúsculas/minúsculas.
function editarAssociado(email, nome, telefone, status) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Associados");

  if (!aba) {
    throw new Error('A aba "Associados" não foi encontrada.');
  }

  const emailBusca = String(email || "").trim().toLowerCase();
  const nomeAtualizado = String(nome || "").trim();
  const telefoneAtualizado = String(telefone || "").trim();
  const statusAtualizado = String(status || "").trim();

  if (!emailBusca) {
    throw new Error("E-mail do associado não informado.");
  }

  if (!nomeAtualizado) {
    throw new Error("Nome do associado não informado.");
  }

  if (!statusAtualizado) {
    throw new Error("Status do associado não informado.");
  }

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailPlanilha = String(dados[i][2] || "").trim().toLowerCase();

    if (emailPlanilha === emailBusca) {
      aba.getRange(i + 1, 2).setValue(nomeAtualizado);
      aba.getRange(i + 1, 4).setValue(telefoneAtualizado);
      aba.getRange(i + 1, 5).setValue(statusAtualizado);
      SpreadsheetApp.flush();
      return "Associado atualizado com sucesso";
    }
  }

  throw new Error("Associado não encontrado.");
}

// No doPost, mantenha este bloco antes de "Ação não encontrada".
if (dados.acao === "editarAssociado") {
  const retorno = editarAssociado(
    dados.email,
    dados.nome,
    dados.telefone,
    dados.status
  );

  return ContentService
    .createTextOutput(JSON.stringify({ status: retorno }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Opcional, mas recomendado: no início do doPost, troque o parse simples por:
// dados = JSON.parse(e.postData.contents || "{}");
// Assim o script trata corretamente uma requisição sem conteúdo.
