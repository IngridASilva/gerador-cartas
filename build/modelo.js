const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
        Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle } = require('docx');
const fs = require('fs');

const FONTE = "Arial";
const p = (texto, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { after: opts.after === undefined ? 200 : opts.after, line: 276 },
  children: [new TextRun({ text: texto, font: FONTE, size: opts.size || 22,
                            bold: !!opts.bold, color: opts.color || "000000" })],
});

const celula = (texto, largura, negrito, fundo) => new TableCell({
  width: { size: largura, type: WidthType.DXA },
  shading: fundo ? { type: ShadingType.CLEAR, fill: fundo } : undefined,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  children: [new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 0 },
    children: [new TextRun({ text: texto, font: FONTE, size: 20, bold: !!negrito })],
  })],
});

function construir(titulo, corpo, tabela) {
  return new Document({
    creator: "Gerador de Cartas",
    title: titulo,
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT, spacing: { after: 400 },
          children: [new TextRun({ text: "{{EMPRESA}}", font: FONTE, size: 20,
                                   bold: true, color: "1F3864" })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT, spacing: { after: 500 },
          children: [new TextRun({ text: "{{CIDADE}}, {{DATA_CARTA}}", font: FONTE, size: 20 })],
        }),
        p("{{TRATAMENTO}} {{NOME}},", { after: 300 }),
        ...corpo.map(t => p(t)),
        ...(tabela ? [tabela, p("", { after: 300 })] : []),
        p("{{PARAGRAFO_FINAL}}"),
        p("Atenciosamente,", { after: 600 }),
        new Paragraph({
          spacing: { after: 0 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "808080", space: 4 } },
          children: [new TextRun({ text: "", font: FONTE, size: 20 })],
        }),
        p("{{ASSINATURA_NOME}}", { after: 0, bold: true, align: AlignmentType.LEFT }),
        p("{{ASSINATURA_CARGO}}", { after: 400, align: AlignmentType.LEFT }),
        new Paragraph({
          alignment: AlignmentType.LEFT, spacing: { after: 0 },
          children: [new TextRun({ text: "Documento gerado em {{DATA_GERACAO}} · Referência {{REFERENCIA}}",
                                   font: FONTE, size: 14, color: "808080" })],
        }),
      ],
    }],
  });
}

const tabelaMerito = new Table({
  columnWidths: [4600, 4600],
  width: { size: 9200, type: WidthType.DXA },
  rows: [
    new TableRow({ children: [celula("Descrição", 4600, true, "E8EDF5"),
                              celula("Valor", 4600, true, "E8EDF5")] }),
    new TableRow({ children: [celula("Remuneração atual", 4600), celula("{{SALARIO_ATUAL}}", 4600)] }),
    new TableRow({ children: [celula("Nova remuneração", 4600), celula("{{SALARIO_NOVO}}", 4600, true)] }),
    new TableRow({ children: [celula("Percentual de ajuste", 4600), celula("{{PCT_AUMENTO}}", 4600)] }),
    new TableRow({ children: [celula("Vigência a partir de", 4600), celula("{{DATA_EFEITO}}", 4600)] }),
  ],
});

const merito = construir("Comunicado de mérito", [
  "É com satisfação que comunicamos o resultado do ciclo de mérito de {{CICLO}}. Sua contribuição ao longo do período foi reconhecida pela liderança da área de {{GERENCIA}}, e por isso sua remuneração será ajustada conforme o quadro abaixo.",
], tabelaMerito);

const promocao = construir("Comunicado de promoção", [
  "Temos o prazer de comunicar sua promoção para o cargo de {{CARGO_NOVO}}, na área de {{GERENCIA}}, a partir de {{DATA_EFEITO}}.",
  "A decisão reflete a evolução da sua entrega no cargo de {{CARGO_ATUAL}} e a confiança da liderança na sua capacidade de assumir o novo escopo. Seu gestor, {{GESTOR}}, conduzirá a conversa sobre as atribuições e os primeiros objetivos da nova posição.",
], tabelaMerito);

fs.mkdirSync("modelos", { recursive: true });
Packer.toBuffer(merito).then(b => fs.writeFileSync("modelos/Modelo_Merito.docx", b));
Packer.toBuffer(promocao).then(b => fs.writeFileSync("modelos/Modelo_Promocao.docx", b));
console.log("modelos gerados");
