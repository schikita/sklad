-- Сквозная аналитика РемМонтаж.
-- Проект remmontazh-eu, регион EU. Песочница BigQuery, без оплаты.
-- Рекламный шлюз BigQuery Data Transfer сюда не входит: его включает сам Google
-- после появления Customer ID. Таблица ads_campaign_stats пока пустая.

CREATE SCHEMA IF NOT EXISTS `remmontazh-eu.raw_data`
OPTIONS(location = "EU");

CREATE SCHEMA IF NOT EXISTS `remmontazh-eu.analytics`
OPTIONS(location = "EU");

CREATE TABLE IF NOT EXISTS `remmontazh-eu.raw_data.raw_google_sheets` (
  franchise_id STRING NOT NULL,
  source_name STRING,
  sheet_name STRING,
  loaded_at TIMESTAMP,
  report_date DATE,
  name STRING,
  money_in NUMERIC,
  money_out NUMERIC,
  campaign_id STRING,
  link STRING,
  payload STRING,
  custom_string_field_1 STRING,
  custom_string_field_2 STRING,
  custom_string_field_3 STRING,
  custom_string_field_4 STRING,
  custom_string_field_5 STRING
)
PARTITION BY report_date
CLUSTER BY franchise_id;

CREATE TABLE IF NOT EXISTS `remmontazh-eu.raw_data.raw_moysklad_json` (
  franchise_id STRING NOT NULL,
  entity STRING,
  loaded_at TIMESTAMP,
  report_date DATE,
  payload STRING,
  custom_string_field_1 STRING,
  custom_string_field_2 STRING,
  custom_string_field_3 STRING,
  custom_string_field_4 STRING,
  custom_string_field_5 STRING
)
PARTITION BY report_date
CLUSTER BY franchise_id;

CREATE TABLE IF NOT EXISTS `remmontazh-eu.raw_data.raw_1c_incoming` (
  franchise_id STRING,
  loaded_at TIMESTAMP,
  report_date DATE,
  payload STRING
)
PARTITION BY report_date
CLUSTER BY franchise_id;

CREATE TABLE IF NOT EXISTS `remmontazh-eu.raw_data.ads_campaign_stats` (
  report_date DATE NOT NULL,
  campaign_id STRING,
  campaign_name STRING,
  spend NUMERIC,
  impressions INT64,
  clicks INT64
)
PARTITION BY report_date
CLUSTER BY campaign_id;

CREATE TABLE IF NOT EXISTS `remmontazh-eu.raw_data.master_registry` (
  franchise_id STRING,
  franchise_name STRING,
  status STRING,
  source_name STRING,
  sheet_id STRING,
  sheet_name STRING,
  source_status STRING
);

CREATE OR REPLACE VIEW `remmontazh-eu.analytics.active_franchises` AS
SELECT franchise_id, ANY_VALUE(franchise_name) AS franchise_name
FROM `remmontazh-eu.raw_data.master_registry`
WHERE status = 'АКТИВЕН'
GROUP BY franchise_id;

CREATE OR REPLACE VIEW `remmontazh-eu.analytics.v_marketing_roas` AS
WITH sales AS (
  SELECT
    franchise_id,
    report_date,
    campaign_id,
    SUM(money_in) AS revenue,
    SUM(money_out) AS cost_out
  FROM `remmontazh-eu.raw_data.raw_google_sheets`
  WHERE report_date IS NOT NULL
  GROUP BY 1, 2, 3
),
ads AS (
  SELECT
    report_date,
    campaign_id,
    SUM(spend) AS spend,
    SUM(impressions) AS impressions,
    SUM(clicks) AS clicks
  FROM `remmontazh-eu.raw_data.ads_campaign_stats`
  GROUP BY 1, 2
)
SELECT
  s.franchise_id,
  COALESCE(s.report_date, a.report_date) AS report_date,
  COALESCE(s.campaign_id, a.campaign_id) AS campaign_id,
  IFNULL(s.revenue, 0) AS revenue,
  IFNULL(a.spend, 0) AS ads_spend,
  IFNULL(s.revenue, 0) - IFNULL(a.spend, 0) AS profit_before_cogs,
  SAFE_DIVIDE(s.revenue, NULLIF(a.spend, 0)) AS roas
FROM sales s
FULL OUTER JOIN ads a
  ON s.report_date = a.report_date
 AND s.campaign_id = a.campaign_id;

-- Приоритет: факт МойСклад, если за день и кампанию он есть. Иначе лист.
CREATE OR REPLACE VIEW `remmontazh-eu.analytics.global_financial_summary` AS
SELECT r.*
FROM `remmontazh-eu.analytics.v_marketing_roas` r
JOIN `remmontazh-eu.analytics.active_franchises` f
  ON r.franchise_id = f.franchise_id;

CREATE OR REPLACE VIEW `remmontazh-eu.analytics.view_inventory_abc_xyz` AS
WITH sales AS (
  SELECT
    franchise_id,
    JSON_VALUE(payload, '$.name') AS product_name,
    JSON_VALUE(payload, '$.article') AS product_sku,
    DATE(JSON_VALUE(payload, '$.updated')) AS movement_date,
    CAST(JSON_VALUE(payload, '$.quantity') AS NUMERIC) AS qty,
    SAFE_DIVIDE(CAST(JSON_VALUE(payload, '$.sum') AS NUMERIC), 100) AS amount
  FROM `remmontazh-eu.raw_data.raw_moysklad_json`
  WHERE entity IN ('demand', 'customerorder')
),
rolled AS (
  SELECT
    franchise_id,
    product_name,
    product_sku,
    SUM(amount) AS product_revenue,
    AVG(qty) AS avg_qty,
    STDDEV(qty) AS stddev_qty,
    DATE_DIFF(CURRENT_DATE(), MAX(movement_date), DAY) AS days_since_last_sale
  FROM sales
  GROUP BY 1, 2, 3
),
shares AS (
  SELECT
    *,
    SAFE_DIVIDE(
      SUM(product_revenue) OVER (PARTITION BY franchise_id ORDER BY product_revenue DESC),
      NULLIF(SUM(product_revenue) OVER (PARTITION BY franchise_id), 0)
    ) AS cumulative_share,
    ROUND(IF(avg_qty > 0, stddev_qty / avg_qty * 100, NULL), 2) AS variation_coefficient
  FROM rolled
)
SELECT
  *,
  CASE
    WHEN days_since_last_sale > 90 THEN 'НЕЛИКВИД'
    WHEN cumulative_share <= 0.80 THEN 'ХОДОВОЙ'
    WHEN cumulative_share <= 0.95 THEN 'СРЕДНЕХОДОВОЙ'
    ELSE 'НЕХОДОВОЙ'
  END AS abc_status,
  CASE
    WHEN variation_coefficient IS NULL THEN 'НЕТ РЯДА'
    WHEN variation_coefficient <= 10 THEN 'X'
    WHEN variation_coefficient <= 25 THEN 'Y'
    ELSE 'Z'
  END AS xyz_status
FROM shares;

CREATE OR REPLACE VIEW `remmontazh-eu.analytics.view_staff` AS
SELECT
  franchise_id,
  sheet_name AS role_name,
  name AS person_name,
  COUNT(1) AS rows_count
FROM `remmontazh-eu.raw_data.raw_google_sheets`
WHERE source_name IN ('Ген. Директор', 'Бухгалтер', 'Слесарь-ремонтник № 1')
GROUP BY 1, 2, 3;
