# 🏛️ GUIA DE ARQUITETURA ENTERPRISE: OBSERVABILIDADE, DEVOPS & DATA LAKE (BIGQUERY)

**Sistema GovTech NLLC** • Inteligência e Auditoria em Licitações Públicas  
**Norma Regente:** Lei nº 14.133/2021 & IN SEGES/ME nº 65/2021  
**Ecossistema:** Google Cloud Platform, Firebase & Render  

---

## 🛠️ ÉPICO 1: TELEMETRIA E OBSERVABILIDADE (HEALTH CHECK)

### 1. HTTP Interceptor & Alerta Visual Graceful Degradation (Frontend)
* **Localização no Código:** `frontend/index.html` e `frontend/preview_app.html`.
* **Mecanismo:** Interceptador nativo envelopando `window.fetch`. Monitora chamadas a `/api/v1/prices/*`, `/api/v1/catalog/*`, `compras.dados.gov.br` e `pncp.gov.br`.
* **Regra de Transição de Estado:**
  * **Ativação:** 2 erros sequenciais com código `HTTP 502`, `HTTP 503`, `HTTP 504` ou falhas de rede (`TypeError: Failed to fetch`).
  * **Evento Global:** Dispara `window.dispatchEvent(new CustomEvent('gov-api-health-change', { detail: { status: 'degraded' } }))`.
  * **Banner Dinâmico:** Exibe `<aside id="govApiHealthBanner">` no topo da tela com `role="alert"` e `aria-live="assertive"`, informando o servidor instável sem travar a navegação do operador.
  * **Auto-Recuperação:** Ao detectar 2 respostas bem-sucedidas (`200 OK`) consecutivas, oculta suavemente o banner e restaura o estado `healthy`.

### 2. Structured Logging no Google Cloud (Backend)
* **Localização no Código:**
  * `backend/core/gov_api_client.py`: Enriquecido com `time.perf_counter()` e JSON estruturado.
  * `functions/index.js`: Função `logExternalApiTelemetry`.
* **Formato do Log Estruturado:**
```json
{
  "severity": "WARNING",
  "service": "compras-gov-api",
  "latency_ms": 3420,
  "endpoint": "https://dadosabertos.compras.gov.br/modulo-pesquisa-preco/1_consultarMaterial",
  "status_code": 503,
  "attempt": 2,
  "message": "Instabilidade em gateway governamental: HTTP 503",
  "timestamp": "2026-11-18T01:40:00.000Z"
}
```

---

## 🛠️ ÉPICO 2: DEVOPS, CI/CD E AMBIENTES EFÊMEROS

### 1. Bloqueio de Deploy Manual & Isolamento de Permissões (IAM)
Para evitar deploys manuais de máquinas locais direto para a produção:

```bash
# 1. Remover permissão de deploy direto aos desenvolvedores humanos:
gcloud projects remove-iam-policy-binding admin-cfj \
  --member="user:dev@suaempresa.gov.br" \
  --role="roles/firebasehosting.admin"

# 2. Conceder permissão somente leitura (Viewer) aos desenvolvedores:
gcloud projects add-iam-policy-binding admin-cfj \
  --member="user:dev@suaempresa.gov.br" \
  --role="roles/firebasehosting.viewer"

# 3. Criar Conta de Serviço exclusiva para a automação do GitHub Actions:
gcloud iam service-accounts create github-actions-hosting \
  --display-name="GitHub Actions Firebase Hosting CI/CD"

# 4. Atribuir a role de Hosting Admin exclusivamente à Conta de Serviço:
gcloud projects add-iam-policy-binding admin-cfj \
  --member="serviceAccount:github-actions-hosting@admin-cfj.iam.gserviceaccount.com" \
  --role="roles/firebasehosting.admin"

# 5. Gerar chave privada da Service Account e cadastrar como Secret no GitHub:
# Nome da Secret no GitHub: FIREBASE_SERVICE_ACCOUNT_ADMIN_CFJ
gcloud iam service-accounts keys create sa-key.json \
  --iam-account=github-actions-hosting@admin-cfj.iam.gserviceaccount.com
```

* **Pre-deploy Safeguard (`scripts/check_ci_deploy_lock.js`):** Integrado ao `firebase.json` no hook `predeploy`. Se o comando `firebase deploy` for executado fora do GitHub Actions (`CI != true`), a execução é imediatamente interrompida.

### 2. Suíte de Testes Automatizados (IN 65/2021)
* **Testes Node.js (executáveis no CI):** `tests/in65_math_engine.test.js`
  * Execução: `npm test`
* **Testes Python (executáveis no CI):** `backend/tests/test_in65_stat_engine.py`
  * Execução: `npm run test:python`
* **Cenários Cobertos:**
  1. *Amostra Homogênea ($CV \le 25\%$):* Garante a seleção obrigatória da **MÉDIA**.
  2. *Amostra Heterogênea ($CV > 25\%$):* Garante a seleção obrigatória da **MEDIANA**.
  3. *Expurgo de Outliers:* Preços inexequíveis e com sobrepreço recebem `is_outlier: true` e `purge_status: EXPURGADA_...`.
  4. *Cesta Restrita ($< 3$ cotações):* Valida a emissão da nota de auditoria do Art. 6º, § 5º da IN 65/2021.

### 3. Pipeline de Homologação em Canais Efêmeros (`.github/workflows/pr-preview.yml`)
* Disparado em qualquer Pull Request para a branch `main`.
* Roda os testes unitários.
* Cria um Preview Channel com URL única: `https://admin-cfj--pr-<PR_NUMBER>-<HASH>.web.app`.
* Posta automaticamente um comentário no PR com a URL de homologação.

### 4. Pipeline de Produção (`.github/workflows/deploy-prod.yml`)
* Disparado no merge para a branch `main`.
* Reexecuta os testes e realiza o deploy no canal oficial `live`.

---

## 🛠️ ÉPICO 3: DATA LAKE E INTELIGÊNCIA PREDITIVA (BIGQUERY)

### 1. Inicialização da Extensão Oficial `ext-firestore-bigquery-export`
Execute no terminal:

```bash
# Instalação das 3 instâncias para as coleções do GovTech
firebase ext:install firebase/firestore-bigquery-export --save --env=extensions/firestore-bigquery-export-lotes.env
firebase ext:install firebase/firestore-bigquery-export --save --env=extensions/firestore-bigquery-export-itens.env
firebase ext:install firebase/firestore-bigquery-export --save --env=extensions/firestore-bigquery-export-saneados.env

# Deploy das extensões no projeto Google Cloud
firebase deploy --only extensions
```

### 2. Mapeamento de Coleções & Tabelas Espelho

| Coleção Firestore | Dataset BigQuery | Tabela Raw Changelog | View Raw Latest |
| :--- | :--- | :--- | :--- |
| `lotes_pesquisa` | `govtech_analytics` | `lotes_pesquisa_raw_changelog` | `lotes_pesquisa_raw_latest` |
| `itens_lote` | `govtech_analytics` | `itens_lote_raw_changelog` | `itens_lote_raw_latest` |
| `resultados_saneados` | `govtech_analytics` | `resultados_saneados_raw_changelog` | `resultados_saneados_raw_latest` |

### 3. View Analítica Consolidada (`vw_sazonalidade_compras`)
* **Código SQL:** `database/bigquery_vw_sazonalidade_compras.sql`
* **Campos Gerados:**
  * `ano_mes`: Mês de referência (`YYYY-MM`) para eixos contínuos.
  * `catmat_codigo` & `catmat_descricao`: Identificação inequívoca do item.
  * `preco_mediana_saneada`: Preço de referência auditado sob a IN 65/2021.
  * `variacao_mensal_preco_pct`: Variação mês a mês (MoM).
  * `media_movel_3_meses`: Linha de tendência para eliminar ruídos sazonais.
  * `ranking_economia_anual`: Rank (1º ao 12º) do mês mais vantajoso para licitar o item.

### 4. Como Conectar no Google Looker Studio
1. Acesse [lookerstudio.google.com](https://lookerstudio.google.com).
2. Clique em **Criar** > **Origem de Dados** > Selecione o conector **BigQuery**.
3. Escolha o Projeto `admin-cfj` > Conjunto de Dados `govtech_analytics` > Visualização `vw_sazonalidade_compras`.
4. Crie um gráfico de **Linhas**:
   * **Dimensão:** `ano_mes`
   * **Dimensão de Detalhamento:** `catmat_codigo` (ou `catmat_descricao`)
   * **Métrica:** `preco_mediana_saneada` (e opcionalmente `media_movel_3_meses`).
5. **Resultado:** Gráfico dinâmico de flutuação de preços nos 12 meses do ano para planejamento do Plano de Contratações Anual (PCA).
