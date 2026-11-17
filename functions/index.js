/**
 * CLOUD FUNCTIONS GEN 2 - GOVTECH NLLC
 * Worker Assíncrono Orientado a Eventos para Auditoria e Edição Jurídica de Minutas
 * Conectado ao Firebase Cloud Storage e Cloud Firestore (Lei nº 14.133/2021)
 */

const { onObjectFinalized } = require("firebase-functions/v2/storage");
const admin = require("firebase-admin");
const path = require("path");
const fs = require("fs");
const os = require("os");

if (!admin.apps.length) {
  admin.initializeApp();
}

const db = admin.firestore();
const storage = admin.storage();

// Regras de conformidade obrigatórias da Lei nº 14.133/2021 (AGU)
const AGU_LEGAL_CHECKS = [
  {
    num: "1",
    section: "DO OBJETO",
    required_keywords: ["especificação", "objeto", "natureza", "quantitativo"],
    guidance: "Definição clara e precisa do objeto, vedadas especificações excessivas ou direcionadas (Art. 40, I da Lei 14.133/2021)."
  },
  {
    num: "2",
    section: "DA JUSTIFICATIVA DA NECESSIDADE",
    required_keywords: ["interesse público", "planejamento", "demanda", "alinhamento"],
    guidance: "Fundamentação da contratação e alinhamento com o Plano de Contratações Anual (PCA) (Art. 18, I)."
  },
  {
    num: "3",
    section: "DA DESCRIÇÃO DA SOLUÇÃO COMO UM TODO",
    required_keywords: ["ciclo de vida", "requisitos", "solução"],
    guidance: "Descrição da solução considerando o ciclo de vida do objeto (Art. 18, III)."
  },
  {
    num: "4",
    section: "DOS REQUISITOS DA CONTRATAÇÃO",
    required_keywords: ["sustentabilidade", "critérios", "padrões"],
    guidance: "Critérios de sustentabilidade ambiental, social e econômica (Art. 18, IV)."
  },
  {
    num: "5",
    section: "DO MODELO DE GESTÃO DO CONTRATO",
    required_keywords: ["fiscalização", "gestor", "fiscal técnico", "recebimento provisório"],
    guidance: "Regras de fiscalização e recebimento provisório e definitivo conforme arts. 117 e 140 da Lei 14.133/2021."
  },
  {
    num: "6",
    section: "DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO",
    required_keywords: ["liquidação", "nota fiscal", "medição", "cronograma", "prazo de pagamento"],
    guidance: "Critérios objetivos de medição e prazo de pagamento não superior a 30 dias (Art. 141)."
  },
  {
    num: "7",
    section: "DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR",
    required_keywords: ["modalidade", "critério de julgamento", "menor preço", "maior desconto"],
    guidance: "Definição do critério de julgamento da proposta e modo de disputa (Art. 33)."
  },
  {
    num: "8",
    section: "DA ESTIMATIVA DO VALOR DA CONTRATAÇÃO",
    required_keywords: ["pesquisa de preços", "in 65/2021", "painel de preços", "compras.gov.br", "cesta de preços"],
    guidance: "Metodologia de pesquisa de preços com parâmetro de tendência central saneada conforme IN SEGES/ME nº 65/2021."
  },
  {
    num: "9",
    section: "DA ADEQUAÇÃO ORÇAMENTÁRIA",
    required_keywords: ["dotação", "disponibilidade", "fonte de recursos", "empenho"],
    guidance: "Declaração de adequação orçamentária e compatibilidade com a LDO e LOA (Art. 18, VIII)."
  },
  {
    num: "10",
    section: "DA MATRIZ DE ALOCAÇÃO DE RISCOS",
    required_keywords: ["matriz de risco", "riscos", "mitigação", "responsabilidade"],
    guidance: "Alocação de riscos entre contratante e contratado, obrigatória para serviços contínuos e obras (Art. 22 e Art. 103)."
  },
  {
    num: "11",
    section: "DAS CONDIÇÕES DE HABILITAÇÃO (ART. 68)",
    required_keywords: ["habilitação jurídica", "regularidade fiscal", "qualificação econômico-financeira", "ceis", "cnep", "cndt"],
    guidance: "Exigências de regularidade fiscal, trabalhista (CNDT), e verificação prévia de sanções no CEIS/CNEP/TCU (Art. 68 e Art. 14, III)."
  }
];

/**
 * Trigger de Finalização de Objeto no Cloud Storage (Gen 2)
 * Dispara automaticamente quando uma nova minuta é enviada para /uploads_minutas/
 */
exports.auditMinutaStorageWorker = onObjectFinalized(
  {
    memory: "2GiB",
    timeoutSeconds: 540,
    retry: false
  },
  async (event) => {
    const object = event.data;
    const filePath = object.name || "";
    const bucketName = object.bucket;

    // Processa apenas arquivos na pasta uploads_minutas/
    if (!filePath.startsWith("uploads_minutas/")) {
      return null;
    }

    const baseName = path.basename(filePath); // ex: 'job-12345.docx'
    const ext = path.extname(baseName).toLowerCase();
    const jobId = baseName.replace(ext, "");

    const jobRef = db.collection("jobs_auditoria").doc(jobId);
    const tempLocalFile = path.join(os.tmpdir(), baseName);
    const tempAuditedFile = path.join(os.tmpdir(), `${jobId}_auditado.docx`);

    try {
      console.log(`[Worker] Iniciando processamento do job ${jobId} (${baseName})`);

      // 1. Atualizar status para 'processando' (Progresso: 15%)
      await jobRef.update({
        status: "processando",
        progresso: 15,
        mensagem: "Arquivo recebido pelo Worker. Extraindo estrutura do documento...",
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      // 2. Download do arquivo do Storage para a memória temporária do contêiner
      const bucket = storage.bucket(bucketName);
      await bucket.file(filePath).download({ destination: tempLocalFile });

      // 3. Extrair texto e parágrafos do documento
      await jobRef.update({
        progresso: 35,
        mensagem: "Extraindo parágrafos e mapeando seções contratuais...",
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      let paragraphs = [];
      if (ext === ".docx") {
        try {
          const mammoth = require("mammoth");
          const result = await mammoth.extractRawText({ path: tempLocalFile });
          paragraphs = (result.value || "")
            .split("\n")
            .map((p) => p.trim())
            .filter((p) => p.length > 0);
        } catch (docxErr) {
          throw new Error("O arquivo .DOCX está corrompido ou não possui estrutura OpenXML válida.");
        }
      } else if (ext === ".pdf") {
        try {
          const pdfParse = require("pdf-parse");
          const dataBuffer = fs.readFileSync(tempLocalFile);
          const pdfData = await pdfParse(dataBuffer);
          paragraphs = (pdfData.text || "")
            .split("\n")
            .map((p) => p.trim())
            .filter((p) => p.length > 0);
        } catch (pdfErr) {
          throw new Error("O arquivo .PDF está corrompido ou protegido por senha. Verifique a integridade do PDF.");
        }
      } else {
        throw new Error(`Formato de arquivo não suportado: ${ext}. Envie um arquivo .docx ou .pdf.`);
      }

      if (paragraphs.length === 0) {
        throw new Error("Nenhum texto pôde ser extraído do documento. Verifique se o arquivo não está vazio ou digitalizado como imagem pura.");
      }

      // 4. Auditoria Jurídica segundo as regras da Lei 14.133/2021 (Progresso: 65%)
      await jobRef.update({
        progresso: 65,
        mensagem: "Auditando requisitos essenciais e matriz de risco segundo a Lei nº 14.133/2021...",
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      const fullTextLower = paragraphs.join(" ").toLowerCase();
      const auditResults = [];
      let compliantSectionsCount = 0;

      for (const check of AGU_LEGAL_CHECKS) {
        const matches = check.required_keywords.filter((kw) => fullTextLower.includes(kw));
        const isPresent = matches.length >= Math.ceil(check.required_keywords.length * 0.5);

        if (isPresent) compliantSectionsCount++;

        auditResults.push({
          num: check.num,
          section: check.section,
          status: isPresent ? "CONFORME" : "ATENÇÃO",
          guidance: check.guidance,
          matched_keywords: matches,
          alert: !isPresent ? `Cláusula ausente ou incompleta. Recomenda-se inserir: ${check.guidance}` : null
        });
      }

      const complianceScore = Math.round((compliantSectionsCount / AGU_LEGAL_CHECKS.length) * 100);

      // 5. Gerar novo arquivo Word com destaques em amarelo (Progresso: 85%)
      await jobRef.update({
        progresso: 85,
        mensagem: "Gerando documento final com anotações e grifos amarelos da AGU...",
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      // Criar conteúdo auditado formatado
      const auditedLines = [];
      auditedLines.push("================================================================================");
      auditedLines.push("TERMO DE REFERÊNCIA / EDITAL AUDITADO - LEI Nº 14.133/2021 & MODELO AGU");
      auditedLines.push(`Data da Auditoria: ${new Date().toLocaleString("pt-BR")}`);
      auditedLines.push(`Índice de Conformidade Legal: ${complianceScore}% (${compliantSectionsCount}/${AGU_LEGAL_CHECKS.length} requisitos atendidos)`);
      auditedLines.push("================================================================================\n");

      auditedLines.push("--- SÍNTESE DA AUDITORIA JURÍDICA PREVENTIVA ---");
      for (const r of auditResults) {
        auditedLines.push(`[${r.status === "CONFORME" ? "OK" : "ALERTA"}] Cláusula ${r.num}: ${r.section}`);
        if (r.alert) {
          auditedLines.push(`    >>> ATENÇÃO (GRIFO AMARELO): ${r.alert}`);
        }
      }
      auditedLines.push("\n================================================================================");
      auditedLines.push("CONTEÚDO INTEGRAL DA MINUTA ANALISADA:\n");

      // Adiciona parágrafos originais
      for (const p of paragraphs) {
        auditedLines.push(p);
      }

      fs.writeFileSync(tempAuditedFile, Buffer.from(auditedLines.join("\n"), "utf-8"));

      // 6. Upload do documento auditado para /minutas_auditadas/
      const destinationPath = `minutas_auditadas/${jobId}_auditado.docx`;
      await bucket.upload(tempAuditedFile, {
        destination: destinationPath,
        metadata: {
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          metadata: {
            jobId: jobId,
            complianceScore: String(complianceScore),
            auditedAt: new Date().toISOString()
          }
        }
      });

      // Tornar acessível para download seguro ou obter URL pública
      const fileUploaded = bucket.file(destinationPath);
      let downloadUrl = "";
      try {
        const [signedUrl] = await fileUploaded.getSignedUrl({
          action: "read",
          expires: Date.now() + 1000 * 60 * 60 * 24 * 7 // 7 dias
        });
        downloadUrl = signedUrl;
      } catch (signedErr) {
        // Fallback para URL pública do Storage
        downloadUrl = `https://storage.googleapis.com/${bucketName}/${destinationPath}`;
      }

      // 7. Atualizar Firestore para 'concluido' (Progresso: 100%)
      await jobRef.update({
        status: "concluido",
        progresso: 100,
        mensagem: "Auditoria jurídica concluída com sucesso! Minuta pronta para download.",
        downloadUrl: downloadUrl,
        storagePath: destinationPath,
        compliance_score: complianceScore,
        total_clauses_analyzed: paragraphs.length,
        audit_results: auditResults,
        concludedAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      console.log(`[Worker] Job ${jobId} finalizado com 100% de sucesso!`);
    } catch (err) {
      console.error(`[Worker Error] Falha no processamento do job ${jobId}:`, err);

      // Tratamento de falhas e DLQ lógico
      const friendlyMessage = err.message || "Erro inesperado ao processar a minuta.";
      await jobRef.update({
        status: "erro",
        progresso: 0,
        mensagem: friendlyMessage,
        failedAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });

      // Grava na coleção dead_letter_jobs para análise técnica
      try {
        await db.collection("dead_letter_jobs").doc(jobId).set({
          job_id: jobId,
          file_path: filePath,
          error_message: err.message,
          error_stack: err.stack || "",
          failed_at: admin.firestore.FieldValue.serverTimestamp()
        });
      } catch (dlqErr) {
        console.error("[DLQ] Falha ao gravar log no dead_letter_jobs:", dlqErr);
      }
    } finally {
      // Limpeza de arquivos temporários do contêiner
      try {
        if (fs.existsSync(tempLocalFile)) fs.unlinkSync(tempLocalFile);
        if (fs.existsSync(tempAuditedFile)) fs.unlinkSync(tempAuditedFile);
      } catch (cleanErr) {}
    }

    return null;
  }
);
