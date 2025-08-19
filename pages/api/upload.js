import multer from "multer";
import XLSX from "xlsx";
import nextConnect from "next-connect";

// Multer em memória (não salva no disco, compatível com Vercel)
const storage = multer.memoryStorage();
const upload = multer({ storage });

const apiRoute = nextConnect({
  onError(error, req, res) {
    console.error(error);
    res.status(501).json({ error: `Erro: ${error.message}` });
  },
  onNoMatch(req, res) {
    res.status(405).json({ error: `Método '${req.method}' não permitido` });
  },
});

apiRoute.use(upload.array("files"));

function normalizarColunas(linha) {
  const nova = {};
  for (let chave in linha) {
    const chaveFormatada = chave.toLowerCase().trim();

    if (chaveFormatada.includes("código")) nova["Código de Barras"] = linha[chave];
    else if (chaveFormatada.includes("descrição") || chaveFormatada.includes("produto")) nova["Descrição"] = linha[chave];
    else if (chaveFormatada.includes("custo") || chaveFormatada.includes("valor")) nova["Valor do Custo"] = linha[chave];
  }
  return nova;
}

apiRoute.post(async (req, res) => {
  try {
    let identificadores = req.body.identificadores;
    if (!Array.isArray(identificadores)) {
      identificadores = [identificadores];
    }

    const produtos = {};

    for (let i = 0; i < req.files.length; i++) {
      const file = req.files[i];
      const id = identificadores[i] || `ID_${i + 1}`;

      // Ler a planilha direto da memória
      const workbook = XLSX.read(file.buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      let data = XLSX.utils.sheet_to_json(sheet);

      data = data
        .map(normalizarColunas)
        .filter(row => {
          const custo = parseFloat(row["Valor do Custo"]);
          return row["Código de Barras"] && row["Descrição"] && !isNaN(custo);
        })
        .map(row => ({
          "Código de Barras": String(row["Código de Barras"]).trim(),
          "Descrição": row["Descrição"],
          "Valor do Custo": parseFloat(row["Valor do Custo"]),
          "Identificador": id,
        }));

      for (const produto of data) {
        const codigo = produto["Código de Barras"];
        if (!produtos[codigo] || produto["Valor do Custo"] < produtos[codigo]["Valor do Custo"]) {
          produtos[codigo] = produto;
        }
      }
    }

    const resultado = Object.values(produtos);

    if (resultado.length === 0) {
      return res.status(400).send("Nenhum produto válido encontrado nos arquivos.");
    }

    const newSheet = XLSX.utils.json_to_sheet(resultado);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Menor Custo");

    const buffer = XLSX.write(newWorkbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Disposition", "attachment; filename=resultado.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
  } catch (err) {
    console.error("Erro no processamento:", err);
    res.status(500).send("Erro ao processar arquivos.");
  }
});

export default apiRoute;

export const config = {
  api: {
    bodyParser: false, // necessário para usar multer
  },
};
