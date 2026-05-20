from __future__ import annotations

from app.services.credit_diagnostics import (
    build_credit_diagnostic_datasets,
    build_credit_diagnostic_period_comparison_dataset,
    build_credit_diagnostic_period_detail_comparison_dataset,
    build_credit_diagnostic_product_rows,
    build_credit_diagnostic_product_sheet_payload,
    classify_credit_base_reason,
    classify_debit_base_reason,
    classify_icms_diagnostic_reason,
    format_grouped_line_numbers,
    merge_credit_diagnostic_detail_rows,
)
from app.services.period_comparisons import (
    build_entry_period_comparison_rows,
    build_multi_sped_entry_analysis,
    build_nfce_note_snapshot,
    build_sale_period_comparison_rows,
    summarize_entry_analysis,
    summarize_sale_analysis,
)
from app.services.monthly_reports import (
    build_contrib_operation_launch_details_map,
    build_contrib_operation_summary_rows,
    build_contrib_product_monthly_linear_dataset,
    build_contrib_product_monthly_summary_rows,
    build_monthly_aliquota_divergence_rows,
    build_product_monthly_linear_dataset,
    build_product_monthly_summary_rows,
    build_summary_operation_totals_by_period,
)
from app.services.entry_exit_reports import (
    build_entry_exit_analysis_rows,
    build_entry_exit_footer_rows,
    classify_entry_exit_category,
    classify_entry_exit_marker,
    sum_entry_exit_analysis_rows,
    write_entry_exit_analysis_excel,
)
from app.services.analysis_utils import period_label_sort_key
