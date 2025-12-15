let dados = [];

// Carrega o arquivo Excel (.xlsx)
fetch("vendedoras.xlsx")
  .then(response => response.arrayBuffer())
  .then(buffer => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Converte para JSON
    dados = XLSX.utils.sheet_to_json(sheet);
  })
  .catch(err => {
    console.error("Erro ao carregar o Excel:", err);
  });

function buscarRastreios() {
  const input = document.getElementById("vendedoraInput");
  const codigo = input.value.trim().toUpperCase();
  const resultado = document.getElementById("resultado");

  resultado.innerHTML = "";

  // Validação do formato V00000 ou S00000
  if (!/^[VS]\d{5}$/.test(codigo)) {
    resultado.innerHTML = "<li>Código inválido. Use V00000 ou S00000</li>";
    return;
  }

  // Filtra os rastreios
  const encontrados = dados.filter(
    item => String(item.vendedora).toUpperCase() === codigo
  );

  if (encontrados.length === 0) {
    resultado.innerHTML = "<li>Nenhum rastreio encontrado</li>";
    return;
  }

  // Exibe os códigos
  encontrados.forEach(item => {
    const li = document.createElement("li");
    li.textContent = item.rastreio;
    resultado.appendChild(li);
  });
}
