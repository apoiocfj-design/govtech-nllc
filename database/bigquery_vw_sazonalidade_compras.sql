-- ============================================================================
-- BIGQUERY DATA LAKE: INTELIGÊNCIA PREDITIVA E SAZONALIDADE DE PREÇOS PÚBLICOS
-- Sistema GovTech • Instrução Normativa SEGES/ME nº 65/2021 & Lei nº 14.133/2021
-- View Consolidada: `vw_sazonalidade_compras`
-- Compatibilidade: Google Looker Studio, BigQuery BI Engine e Python DataFrames
-- ============================================================================

CREATE OR REPLACE VIEW `admin-cfj.govtech_analytics.vw_sazonalidade_compras` AS

WITH parsed_lotes AS (
  SELECT
    document_id AS lote_id,
    timestamp AS firestore_timestamp,
    COALESCE(
      JSON_VALUE(data, '$.orgao_nome'),
      JSON_VALUE(data, '$.om_nome'),
      'MARINHA DO BRASIL'
    ) AS orgao_comprador,
    COALESCE(
      JSON_VALUE(data, '$.uasg'),
      JSON_VALUE(data, '$.uasg_codigo'),
      '783000'
    ) AS uasg_codigo,
    SAFE_CAST(JSON_VALUE(data, '$.created_at') AS TIMESTAMP) AS data_criacao_lote
  FROM
    `admin-cfj.govtech_analytics.lotes_pesquisa_raw_latest`
  WHERE
    operation != 'DELETE'
),

parsed_itens AS (
  SELECT
    document_id AS item_id,
    COALESCE(JSON_VALUE(data, '$.lote_id'), document_id) AS lote_id_ref,
    SAFE_CAST(JSON_VALUE(data, '$.catmat_codigo') AS INT64) AS catmat_codigo,
    UPPER(TRIM(COALESCE(JSON_VALUE(data, '$.descricao'), JSON_VALUE(data, '$.item_description')))) AS catmat_descricao,
    UPPER(TRIM(COALESCE(JSON_VALUE(data, '$.unidade'), JSON_VALUE(data, '$.target_supply_unit'), 'UN'))) AS unidade_fornecimento,
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.quantidade'), '1') AS FLOAT64) AS quantidade_solicitada
  FROM
    `admin-cfj.govtech_analytics.itens_lote_raw_latest`
  WHERE
    operation != 'DELETE'
),

parsed_resultados AS (
  SELECT
    document_id AS resultado_id,
    COALESCE(JSON_VALUE(data, '$.item_id'), document_id) AS item_id_ref,
    COALESCE(JSON_VALUE(data, '$.lote_id'), '') AS lote_id_ref,
    SAFE_CAST(JSON_VALUE(data, '$.catmat_codigo') AS INT64) AS catmat_codigo_res,
    
    -- Valores monetários saneados (Art. 6º IN SEGES/ME nº 65/2021)
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.mediana_saneada'), JSON_VALUE(data, '$.final_estimated_unit_price')) AS FLOAT64) AS preco_mediana_saneada,
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.media_saneada'), JSON_VALUE(data, '$.mean')) AS FLOAT64) AS preco_media_saneada,
    SAFE_CAST(JSON_VALUE(data, '$.final_estimated_unit_price') AS FLOAT64) AS preco_referencia_homologado,
    
    -- Métricas estatísticas de dispersão e robustez
    COALESCE(JSON_VALUE(data, '$.parametro_adotado'), JSON_VALUE(data, '$.adopted_parameter'), 'MEDIANA') AS parametro_adotado,
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.cv_saneado'), JSON_VALUE(data, '$.cv_percent')) AS FLOAT64) AS coeficiente_variacao_pct,
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.qtd_amostras_validas'), JSON_VALUE(data, '$.valid_quotes_count'), '3') AS INT64) AS qtd_cotacoes_validas,
    SAFE_CAST(COALESCE(JSON_VALUE(data, '$.qtd_outliers_expurgados'), JSON_VALUE(data, '$.purged_quotes_count'), '0') AS INT64) AS qtd_outliers_expurgados,
    
    -- Data de referência da pesquisa
    COALESCE(
      SAFE.PARSE_DATE('%Y-%m-%d', SUBSTR(JSON_VALUE(data, '$.data_pesquisa'), 1, 10)),
      SAFE.PARSE_DATE('%Y-%m-%d', SUBSTR(JSON_VALUE(data, '$.created_at'), 1, 10)),
      DATE(timestamp)
    ) AS data_pesquisa_base
  FROM
    `admin-cfj.govtech_analytics.resultados_saneados_raw_latest`
  WHERE
    operation != 'DELETE'
),

consolidado_base AS (
  SELECT
    r.resultado_id,
    COALESCE(i.catmat_codigo, r.catmat_codigo_res) AS catmat_codigo,
    COALESCE(i.catmat_descricao, CONCAT('CATMAT ', CAST(r.catmat_codigo_res AS STRING))) AS catmat_descricao,
    COALESCE(i.unidade_fornecimento, 'UN') AS unidade_fornecimento,
    COALESCE(l.orgao_comprador, 'MARINHA DO BRASIL') AS orgao_comprador,
    COALESCE(l.uasg_codigo, '783000') AS uasg_codigo,
    r.data_pesquisa_base AS data_pesquisa,
    
    -- Eixos temporais para gráficos do Looker Studio
    FORMAT_DATE('%Y-%m', r.data_pesquisa_base) AS ano_mes,
    EXTRACT(YEAR FROM r.data_pesquisa_base) AS ano,
    EXTRACT(MONTH FROM r.data_pesquisa_base) AS mes,
    CASE EXTRACT(MONTH FROM r.data_pesquisa_base)
      WHEN 1 THEN '01 - Janeiro'
      WHEN 2 THEN '02 - Fevereiro'
      WHEN 3 THEN '03 - Março'
      WHEN 4 THEN '04 - Abril'
      WHEN 5 THEN '05 - Maio'
      WHEN 6 THEN '06 - Junho'
      WHEN 7 THEN '07 - Julho'
      WHEN 8 THEN '08 - Agosto'
      WHEN 9 THEN '09 - Setembro'
      WHEN 10 THEN '10 - Outubro'
      WHEN 11 THEN '11 - Novembro'
      WHEN 12 THEN '12 - Dezembro'
    END AS mes_nome,
    CONCAT('T', CAST(EXTRACT(QUARTER FROM r.data_pesquisa_base) AS STRING)) AS trimestre,
    
    -- Métricas de Preço
    ROUND(r.preco_mediana_saneada, 2) AS preco_mediana_saneada,
    ROUND(r.preco_media_saneada, 2) AS preco_media_saneada,
    ROUND(COALESCE(r.preco_referencia_homologado, r.preco_mediana_saneada), 2) AS preco_referencia_homologado,
    r.parametro_adotado,
    ROUND(r.coeficiente_variacao_pct, 2) AS cv_percentual,
    CASE WHEN r.cv_percentual <= 25.0 THEN 'HOMOGÊNEA (CV <= 25%)' ELSE 'HETEROGÊNEA (CV > 25%)' END AS classificacao_dispersao,
    r.qtd_cotacoes_validas,
    r.qtd_outliers_expurgados
  FROM
    parsed_resultados r
  LEFT JOIN
    parsed_itens i ON r.item_id_ref = i.item_id
  LEFT JOIN
    parsed_lotes l ON COALESCE(i.lote_id_ref, r.lote_id_ref) = l.lote_id
  WHERE
    COALESCE(i.catmat_codigo, r.catmat_codigo_res) IS NOT NULL
    AND r.preco_mediana_saneada > 0
)

-- Camada Final: Agregações Mensais e Análise de Sazonalidade (Looker Studio Ready)
SELECT
  ano_mes,
  ano,
  mes,
  mes_nome,
  trimestre,
  catmat_codigo,
  catmat_descricao,
  unidade_fornecimento,
  orgao_comprador,
  uasg_codigo,
  
  -- Preço Mediano Mensal do CATMAT (Métrica Central de Inteligência Preditiva)
  ROUND(AVG(preco_mediana_saneada), 2) AS preco_mediana_saneada,
  ROUND(AVG(preco_media_saneada), 2) AS preco_media_saneada,
  ROUND(MIN(preco_mediana_saneada), 2) AS menor_preco_registrado_mes,
  ROUND(MAX(preco_mediana_saneada), 2) AS maior_preco_registrado_mes,
  ROUND(AVG(cv_percentual), 2) AS cv_medio_percentual,
  SUM(qtd_cotacoes_validas) AS total_amostras_auditadas,
  COUNT(resultado_id) AS total_pesquisas_concluidas,
  
  -- Flutuação Percentual em Relação ao Mês Anterior (MoM)
  ROUND(
    SAFE_DIVIDE(
      AVG(preco_mediana_saneada) - LAG(AVG(preco_mediana_saneada)) OVER (PARTITION BY catmat_codigo, orgao_comprador ORDER BY ano_mes),
      LAG(AVG(preco_mediana_saneada)) OVER (PARTITION BY catmat_codigo, orgao_comprador ORDER BY ano_mes)
    ) * 100.0,
    2
  ) AS variacao_mensal_preco_pct,

  -- Média Móvel de 3 Meses para Suavização de Tendência de Mercado
  ROUND(
    AVG(AVG(preco_mediana_saneada)) OVER (
      PARTITION BY catmat_codigo, orgao_comprador 
      ORDER BY ano_mes 
      ROWS BETWEEN 2 PRECEDING AND CURRENT ROW
    ),
    2
  ) AS media_movel_3_meses,

  -- Indicador de Janela Ótima de Compra (Rank do mês mais barato do ano para o item)
  DENSE_RANK() OVER (
    PARTITION BY catmat_codigo, ano 
    ORDER BY AVG(preco_mediana_saneada) ASC
  ) AS ranking_economia_anual

FROM
  consolidado_base
GROUP BY
  ano_mes,
  ano,
  mes,
  mes_nome,
  trimestre,
  catmat_codigo,
  catmat_descricao,
  unidade_fornecimento,
  orgao_comprador,
  uasg_codigo;
