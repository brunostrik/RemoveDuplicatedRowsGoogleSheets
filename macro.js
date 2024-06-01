function RemoverDuplicados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var valoresColuna = aba.getRange('D:D').getValues();
  var linhasExcluir = [];

  // Percorre os valores da coluna E e identifica linhas duplicadas
  for (var i = 2; i < valoresColuna.length - 1; i++) {
    console.log("Analyzing "+i);
    for (var j = i + 1; j < valoresColuna.length; j++) {
      if (valoresColuna[i][0] !== '' && valoresColuna[i][0].replace(/\s/g, "") === valoresColuna[j][0].replace(/\s/g, "")) {
        linhasExcluir.push(j + 1); // Adiciona o número da linha à lista de linhas a serem excluídas
        console.log("Add "+j);
      }
    }
  }
  console.log("Time to kill "+linhasExcluir.length+" values");
  // Remove as linhas duplicadas, mantendo apenas uma ocorrência de cada valor
  linhasExcluir.sort(function(a, b){return b-a}); // Ordena a lista em ordem decrescente para garantir que as linhas sejam excluídas corretamente
  for (var k = 0; k < linhasExcluir.length; k++) {
    aba.deleteRow(linhasExcluir[k]);
    console.log(k+": "+linhasExcluir[k]+" killed");
  }
};
