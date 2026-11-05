-- =============================================================================
-- Higher Education SIS & Student Financials Cloud Modernization
-- SQL Crosswalk Conversion, Trial-Balance Reconciliation & UAT Validation Suite
-- =============================================================================

-- 1. Crosswalk Lookup & Orphan Key Audit (VLOOKUP / INNER JOIN Parity Gate)
SELECT
    l.legacy_term_code,
    l.legacy_fund_code,
    COUNT(*) AS unmapped_record_count
FROM legacy_student_financials_extract l
LEFT JOIN legacy_sis_Crosswalk c
    ON TRIM(l.legacy_term_code) = c.legacy_term_code
   AND TRIM(l.legacy_fund_code) = c.legacy_fund_code
WHERE c.cloud_ledger_account IS NULL
GROUP BY l.legacy_term_code, l.legacy_fund_code;

-- 2. Tuition & Financial Aid Ledger Conservation Check ($0.00 Variance Gate)
WITH legacy_totals AS (
    SELECT
        COUNT(*) AS legacy_row_count,
        ROUND(SUM(tuition_charge_usd), 2) AS legacy_tuition_usd,
        ROUND(SUM(aid_disbursement_usd), 2) AS legacy_aid_usd
    FROM legacy_student_financials_extract
),
cloud_totals AS (
    SELECT
        COUNT(*) AS cloud_row_count,
        ROUND(SUM(tuition_charge_usd), 2) AS cloud_tuition_usd,
        ROUND(SUM(aid_disbursement_usd), 2) AS cloud_aid_usd,
        ROUND(SUM(net_student_balance_usd), 2) AS cloud_net_balance_usd
    FROM cloud_sis_student_ledger
)
SELECT
    l.legacy_row_count,
    c.cloud_row_count,
    ROUND(ABS(l.legacy_tuition_usd - c.cloud_tuition_usd), 2) AS tuition_variance_usd,
    ROUND(ABS(l.legacy_aid_usd - c.cloud_aid_usd), 2) AS aid_variance_usd,
    ROUND(ABS((c.cloud_tuition_usd - c.cloud_aid_usd) - c.cloud_net_balance_usd), 2) AS net_identity_delta_usd
FROM legacy_totals l
CROSS JOIN cloud_totals c;

-- 3. Executive Pivot Summary by Academic Unit and Cloud Academic Period
SELECT
    academic_unit,
    cloud_academic_period,
    COUNT(DISTINCT normalized_student_uid) AS headcount,
    SUM(enrollment_credits) AS total_credit_hours,
    ROUND(SUM(tuition_charge_usd), 2) AS gross_tuition_usd,
    ROUND(SUM(aid_disbursement_usd), 2) AS total_financial_aid_usd,
    ROUND(SUM(net_student_balance_usd), 2) AS net_receivable_usd
FROM cloud_sis_student_ledger
GROUP BY academic_unit, cloud_academic_period
ORDER BY academic_unit, cloud_academic_period;
