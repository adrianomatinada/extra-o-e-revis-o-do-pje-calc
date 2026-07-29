import express from 'express';
import path from 'path';
import multer from 'multer';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI, Type } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, HeadingLevel, AlignmentType } from 'docx';

const app = express();
const PORT = 3000;

// Configuração do Multer para armazenamento na memória
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // limite de 25MB
});

// Express middlewares
app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));

// Inicialização segura do cliente Gemini SDK
function getGeminiClient() {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('A chave GEMINI_API_KEY não foi configurada nas variáveis de ambiente do servidor.');
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

// System Prompt especializado em Perícia Contábil e Legislação Trabalhista Brasileira
const SYSTEM_PROMPT_EXTRACTION = `
Você é um assistente sênior especialista em Perícia Contábil e Direito do Trabalho no Brasil (CLT, Súmulas e OJs do TST).
Sua missão é analisar minuciosamente o documento jurídico fornecido (Sentença, Acórdão, Petição Inicial, Laudo Pericial ou Termo de Conciliação) e extrair os dados necessários para alimentação automatizada do PJe-Calc.

Siga rigorosamente estas REGRAS DE NEGÓCIO TRABALHISTAS:
1. Identifique os dados do processo: Número único do CNJ (ex: 0010454-82.2024.5.15.0122), Reclamante (Autor/Trabalhador), Reclamado (Empresa/Réu), Tribunal/Região, Vara do Trabalho e UF.
2. Identifique os marcos temporais e datas:
   - Data de Admissão
   - Data de Demissão / Rescisão
   - Data de Ajuizamento da ação
   - Prescrição Quinquenal: Calcule exatamente 5 anos antes da data de ajuizamento (ex: Ajuizamento em 18/12/2025 -> Parcelas anteriores a 18/12/2020 estão prescritas). Se não houver parcelas prescritas, registre explicativamente.
3. Mapeie TODAS as Verbas Deferidas / Reclamadas:
   - Identifique o tipo: 'horas_extras', 'acumulo_funcao', 'insalubridade', 'periculosidade', 'adicional_noturno', 'diferenca_salarial', 'intervalo_intrajornada', 'danos_morais', 'multa_art_477', 'multa_art_467', 'valor_fixo' ou 'outro'.
   - Percentual (ex: 50% para horas extras, 20% para insalubridade média, 40% para máxima, 20% acúmulo de função).
   - Quantidade mensal estimada ou horas diárias/semanais (ex: 42.86 horas/mês).
   - Divisor aplicável (ex: 220, 180, 150).
   - Mapeie com precisão os REFLEXOS deferidos para cada verba (ex: "Férias + 1/3", "13º Salário", "Aviso Prévio", "FGTS + 40%", "DSR").
   - OJ 394 TST: Verifique se o juiz aplicou ou determinou a observância da OJ 394 da SDI-1 do TST (não repercussão do DSR majorado pelas horas extras nas demais parcelas). Marque true ou false.
4. Histórico Salarial: Se houver indicação de salário base ou evolução salarial, extraia as competências e valores.
5. Honorários e Custas:
   - Percentual dos honorários advocatícios sucumbenciais (ex: 15%, 10%, 5%).
   - Honorários periciais (R$).
   - Custas processuais (2% sobre o valor da condenação).
   - Concessão de Gratuidade de Justiça ao Reclamante.
6. Laudo Técnico Resumido: Elabore uma síntese clara dos motivos da condenação e fundamentação técnica pericial.
7. SE ALGUM DADO NÃO ESTIVER PRESENTE NO TEXTO: Deixe o campo nulo ou com texto explicativo "Não informado", JAMAIS invente dados fictícios.
`;

// Schema do JSON Estruturado de Resposta
const EXTRACTION_RESPONSE_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    processo: {
      type: Type.OBJECT,
      properties: {
        numero: { type: Type.STRING, description: 'Número do processo no formato CNJ' },
        reclamante: { type: Type.STRING, description: 'Nome completo do Reclamante (Trabalhador)' },
        reclamado: { type: Type.STRING, description: 'Nome da Reclamada (Empresa/Empregador)' },
        tribunal: { type: Type.STRING, description: 'Tribunal Regional do Trabalho (ex: TRT 15ª Região)' },
        vara: { type: Type.STRING, description: 'Vara do Trabalho de Origem' },
        uf: { type: Type.STRING, description: 'Estado (UF)' },
      },
      required: ['numero', 'reclamante', 'reclamado'],
    },
    datas: {
      type: Type.OBJECT,
      properties: {
        admissao: { type: Type.STRING, description: 'Data de Admissão (DD/MM/AAAA ou AAAA-MM-DD)' },
        demissao: { type: Type.STRING, description: 'Data de Demissão/Rescisão' },
        ajuizamento: { type: Type.STRING, description: 'Data de Ajuizamento da Ação' },
        prescricaoQuinquenal: { type: Type.STRING, description: 'Data limite da Prescrição Quinquenal (5 anos antes do ajuizamento)' },
        dataSentenca: { type: Type.STRING, description: 'Data da Sentença/Decisão se houver' },
      },
      required: ['admissao', 'demissao', 'ajuizamento'],
    },
    verbas: {
      type: Type.ARRAY,
      description: 'Lista de verbas deferidas ou pleiteadas com seus respectivos percentuais e reflexos',
      items: {
        type: Type.OBJECT,
        properties: {
          id: { type: Type.STRING, description: 'ID único da verba (ex: v1, v2)' },
          nome: { type: Type.STRING, description: 'Nome completo da verba trabalhista' },
          tipo: {
            type: Type.STRING,
            description: 'Tipo da verba: horas_extras, acumulo_funcao, insalubridade, periculosidade, adicional_noturno, diferenca_salarial, intervalo_intrajornada, danos_morais, multa_art_477, multa_art_467, valor_fixo, outro',
          },
          percentual: { type: Type.NUMBER, description: 'Percentual do adicional ou verba (ex: 50 para 50%, 20, 40)' },
          valorEstimado: { type: Type.NUMBER, description: 'Valor estimado em reais se mencionado' },
          quantidadeMensal: { type: Type.NUMBER, description: 'Quantidade média mensal (ex: horas extras/mês)' },
          divisor: { type: Type.NUMBER, description: 'Divisor da jornada (ex: 220, 180, 150)' },
          baseCalculo: { type: Type.STRING, description: 'Descrição da base de cálculo (ex: Salário Base, Evolução Salarial, Salário Mínimo)' },
          oj394: { type: Type.BOOLEAN, description: 'Indica se deve aplicar a OJ 394 TST (sem repetição de DSR majorado nos reflexos)' },
          reflexos: {
            type: Type.ARRAY,
            items: { type: Type.STRING },
            description: 'Lista de verbas reflexas (ex: ["Férias + 1/3", "13º Salário", "Aviso Prévio", "FGTS + 40%", "DSR"])',
          },
          observacoes: { type: Type.STRING, description: 'Observações específicas do juiz ou perito sobre a verba' },
        },
        required: ['nome', 'tipo', 'reflexos', 'oj394'],
      },
    },
    historicoSalarial: {
      type: Type.ARRAY,
      items: {
        type: Type.OBJECT,
        properties: {
          competencia: { type: Type.STRING, description: 'Competência no formato MM/AAAA' },
          salarioBase: { type: Type.NUMBER, description: 'Salário base no período' },
          gratificacao: { type: Type.NUMBER, description: 'Gratificações ou adicionais fixos' },
          total: { type: Type.NUMBER, description: 'Remuneração total' },
        },
        required: ['competencia', 'salarioBase'],
      },
    },
    honorarios: {
      type: Type.OBJECT,
      properties: {
        honorariosAdvocaticiosPercentual: { type: Type.NUMBER, description: 'Percentual de honorários sucumbenciais' },
        honorariosPericiaisValor: { type: Type.NUMBER, description: 'Valor dos honorários periciais em R$' },
        custasProcessuaisPercentual: { type: Type.NUMBER, description: 'Percentual de custas (padrão 2%)' },
        gratuidadeJustica: { type: Type.BOOLEAN, description: 'Indica se foi concedida a gratuidade de justiça' },
      },
    },
    laudo: {
      type: Type.OBJECT,
      properties: {
        resumoSentenca: { type: Type.STRING, description: 'Resumo executivo do julgamento' },
        fundamentacaoTecnica: { type: Type.STRING, description: 'Fundamentação técnica e legal para apuração no PJe-Calc' },
        conclusaoPericial: { type: Type.STRING, description: 'Síntese das orientações para liquidação de sentença' },
      },
      required: ['resumoSentenca', 'fundamentacaoTecnica', 'conclusaoPericial'],
    },
  },
  required: ['processo', 'datas', 'verbas', 'laudo'],
};

// API Endpoint de Health Check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', service: 'CalcPerito AI Engine', timestamp: new Date().toISOString() });
});

// API Endpoint para Extração Automatizada de Documento (PDF / Texto) via Gemini API
app.post('/api/extract-document', upload.single('file'), async (req, res) => {
  try {
    const ai = getGeminiClient();
    const modelName = req.body.useProModel === 'true' ? 'gemini-3.1-pro-preview' : 'gemini-3.6-flash';
    const textContent = req.body.textContent;
    let contentsParts: any[] = [];

    // Se o usuário enviou um arquivo PDF via multipart form-data
    if (req.file) {
      const fileBase64 = req.file.buffer.toString('base64');
      const mimeType = req.file.mimetype || 'application/pdf';
      
      contentsParts = [
        {
          inlineData: {
            mimeType: mimeType === 'application/octet-stream' ? 'application/pdf' : mimeType,
            data: fileBase64,
          },
        },
        {
          text: 'Análise o PDF da sentença/petição anexo e extraia a estrutura completa JSON para liquidação no PJe-Calc conforme o esquema.',
        },
      ];
    } else if (textContent && textContent.trim().length > 0) {
      // Se enviou o texto colado diretamente
      contentsParts = [
        {
          text: `Análise o texto da decisão judicial / petição a seguir e extraia os dados trabalhistas estruturados:\n\n${textContent}`,
        },
      ];
    } else {
      return res.status(400).json({ error: 'Nenhum arquivo PDF ou texto de sentença foi fornecido.' });
    }

    const response = await ai.models.generateContent({
      model: modelName,
      contents: { parts: contentsParts },
      config: {
        systemInstruction: SYSTEM_PROMPT_EXTRACTION,
        responseMimeType: 'application/json',
        responseSchema: EXTRACTION_RESPONSE_SCHEMA,
        temperature: 0.1, // Temperatura baixa para máxima fidelidade dos dados jurídicos
      },
    });

    const responseText = response.text;
    if (!responseText) {
      throw new Error('A resposta da API do Gemini não continha texto.');
    }

    const parsedJson = JSON.parse(responseText);

    // Complementar IDs de verbas se faltarem e adicionar status de revisão inicial
    if (parsedJson.verbas && Array.isArray(parsedJson.verbas)) {
      parsedJson.verbas = parsedJson.verbas.map((v: any, index: number) => ({
        ...v,
        id: v.id || `v-${Date.now()}-${index + 1}`,
        statusRevisao: 'pendente',
      }));
    }

    const result: any = {
      id: `ext-${Date.now()}`,
      dataExtracao: new Date().toISOString(),
      nomeArquivo: req.file ? req.file.originalname : 'Texto Inserido Manualmente',
      ...parsedJson,
    };

    return res.json({ success: true, data: result });
  } catch (error: any) {
    console.error('Erro na extração de documento trabalhista:', error);
    return res.status(500).json({
      success: false,
      error: error.message || 'Falha ao processar o documento com IA.',
    });
  }
});

// API Endpoint para Gerar Laudo Técnico Pericial em Formato Word (.docx)
app.post('/api/generate-docx-laudo', async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.processo) {
      return res.status(400).json({ error: 'Dados insuficientes para geração do laudo.' });
    }

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            // Cabeçalho Oficial
            new Paragraph({
              text: 'LAUDO TÉCNICO PERICIAL DE CÁLCULOS TRABALHISTAS',
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'SISTEMA AUXILIAR DE LIQUIDAÇÃO DE SENTENÇA - INTEGRADO AO PJE-CALC', italics: true }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 400 },
            }),

            // Dados do Processo
            new Paragraph({
              text: '1. DADOS IDENTIFICADORES DO PROCESSO',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Processo Nº: ', bold: true }),
                new TextRun(data.processo.numero || 'Não informado'),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Reclamante (Autor): ', bold: true }),
                new TextRun(data.processo.reclamante || 'Não informado'),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Reclamado (Réu): ', bold: true }),
                new TextRun(data.processo.reclamado || 'Não informado'),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Juízo / Vara: ', bold: true }),
                new TextRun(`${data.processo.vara || ''} - ${data.processo.tribunal || ''}`),
              ],
              spacing: { after: 300 },
            }),

            // Marcos Temporais
            new Paragraph({
              text: '2. MARCOS TEMPORAIS E PRESCRIÇÃO',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 },
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Data de Admissão: ', bold: true }),
                new TextRun(data.datas?.admissao || 'Não informada'),
                new TextRun({ text: '  |  Data de Demissão: ', bold: true }),
                new TextRun(data.datas?.demissao || 'Não informada'),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: 'Data de Ajuizamento: ', bold: true }),
                new TextRun(data.datas?.ajuizamento || 'Não informada'),
                new TextRun({ text: '  |  Marco Prescricional Quinquenal: ', bold: true }),
                new TextRun(data.datas?.prescricaoQuinquenal || 'Não configurada'),
              ],
              spacing: { after: 300 },
            }),

            // Verbas Deferidas
            new Paragraph({
              text: '3. VERBAS TRABALHISTAS DEFERIDAS E PARÂMETROS DE CÁLCULO',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 },
            }),

            ...(data.verbas || []).map((v: any, index: number) => {
              return new Paragraph({
                children: [
                  new TextRun({ text: `${index + 1}. ${v.nome}`, bold: true, size: 24 }),
                  new TextRun({ text: `\n- Adicional/Percentual: ${v.percentual ? v.percentual + '%' : 'N/A'}` }),
                  new TextRun({ text: `\n- Divisor: ${v.divisor || '220'} | Base: ${v.baseCalculo || 'Salário Base'}` }),
                  new TextRun({ text: `\n- Reflexos Deferidos: ${(v.reflexos || []).join(', ') || 'Nenhum'}` }),
                  new TextRun({ text: `\n- Aplicação da OJ 394 TST: ${v.oj394 ? 'SIM (Respeitada a tese vinculante)' : 'NÃO'}` }),
                  new TextRun({ text: v.observacoes ? `\n- Obs: ${v.observacoes}` : '' }),
                ],
                spacing: { after: 200 },
              });
            }),

            // Fundamentação e Conclusão
            new Paragraph({
              text: '4. SÍNTESE E ORIENTAÇÃO PARA LIQUIDAÇÃO NO PJE-CALC',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 300, after: 100 },
            }),
            new Paragraph({
              text: data.laudo?.fundamentacaoTecnica || 'Cálculos parametrizados conforme os títulos deferidos na r. Sentença.',
              spacing: { after: 200 },
            }),
            new Paragraph({
              text: '5. PARECER CONCLUSIVO',
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 },
            }),
            new Paragraph({
              text: data.laudo?.conclusaoPericial || 'Estrutura JSON gerada e validada para importação direta no assistente PJe-Calc.',
              spacing: { after: 400 },
            }),

            // Assinatura
            new Paragraph({
              text: '____________________________________________',
              alignment: AlignmentType.CENTER,
              spacing: { before: 600, after: 50 },
            }),
            new Paragraph({
              text: 'PERITO CONTÁBIL TRABALHISTA / CALCULISTA',
              alignment: AlignmentType.CENTER,
            }),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename=Laudo_Pericial_${data.processo.numero || 'Processo'}.docx`);
    return res.send(buffer);
  } catch (error: any) {
    console.error('Erro ao gerar documento Word:', error);
    return res.status(500).json({ error: 'Erro ao gerar arquivo .docx do laudo pericial.' });
  }
});

// API Endpoint para Gerar e Converter JSON em Arquivo Native .PJC do PJe-Calc
app.post('/api/generate-pjc', async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.processo) {
      return res.status(400).json({ error: 'Payload JSON inválido para geração do arquivo .PJC.' });
    }

    const numProcesso = (data.processo?.numero || 'Processo').replace(/[^a-zA-Z0-9-]/g, '_');
    const escapeXml = (unsafe: any) => {
      if (unsafe === null || unsafe === undefined) return '';
      return String(unsafe)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    };

    const verbasXml = (data.verbas || [])
      .map((v: any, index: number) => {
        const reflexosXml = (v.reflexos || [])
          .map((r: string) => `        <reflexo>${escapeXml(r)}</reflexo>`)
          .join('\n');

        return `    <verba id="v_${index + 1}">
      <nome>${escapeXml(v.nome)}</nome>
      <tipo>${escapeXml(v.tipo)}</tipo>
      <percentual>${v.percentual ?? ''}</percentual>
      <quantidadeMensal>${v.quantidadeMensal ?? ''}</quantidadeMensal>
      <divisor>${v.divisor || 220}</divisor>
      <baseCalculo>${escapeXml(v.baseCalculo || 'Salário Base')}</baseCalculo>
      <aplicarOj394>${v.oj394 ? 'true' : 'false'}</aplicarOj394>
      <reflexos>
${reflexosXml}
      </reflexos>
    </verba>`;
      })
      .join('\n');

    const xmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pjeCalc versao="2.14.0" dataGeracao="${new Date().toISOString()}">
  <calculo id="${escapeXml(data.id || 'calc_' + Date.now())}">
    <processo>
      <numero>${escapeXml(data.processo.numero || '')}</numero>
      <reclamante>${escapeXml(data.processo.reclamante || '')}</reclamante>
      <reclamado>${escapeXml(data.processo.reclamado || '')}</reclamado>
      <tribunal>${escapeXml(data.processo.tribunal || '')}</tribunal>
      <vara>${escapeXml(data.processo.vara || '')}</vara>
      <uf>${escapeXml(data.processo.uf || '')}</uf>
    </processo>
    <parametros>
      <dataAdmissao>${escapeXml(data.datas?.admissao || '')}</dataAdmissao>
      <dataDemissao>${escapeXml(data.datas?.demissao || '')}</dataDemissao>
      <dataAjuizamento>${escapeXml(data.datas?.ajuizamento || '')}</dataAjuizamento>
      <prescricaoQuinquenal>${escapeXml(data.datas?.prescricaoQuinquenal || '')}</prescricaoQuinquenal>
    </parametros>
    <verbas>
${verbasXml}
    </verbas>
  </calculo>
</pjeCalc>`;

    res.setHeader('Content-Type', 'application/x-pje-calc;charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename=Calculo_PJe_${numProcesso}.pjc`);
    return res.send(xmlContent);
  } catch (error: any) {
    console.error('Erro ao gerar arquivo .pjc:', error);
    return res.status(500).json({ error: 'Erro interno ao processar o arquivo .PJC.' });
  }
});

// Inicialização do Servidor com Vite Middleware no desenvolvimento
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[CalcPerito AI] Servidor rodando em http://0.0.0.0:${PORT}`);
  });
}

startServer();
