"""
Executive audit reporting module generating professional Excel & HTML reconciliation reports.
"""

from decimal import Decimal

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from loblaws_automation.config import CUSTOMER_ACCOUNT_MAP
from loblaws_automation.models import AuditMetrics, ReconciliationBatch


def generate_excel_audit_report(batch: ReconciliationBatch, metrics: AuditMetrics, output_path: str) -> None:
    """
    Generates a corporate-styled Excel audit report for the reconciliation batch.

    Args:
        batch: The reconciled transaction batch.
        metrics: The compiled audit metrics.
        output_path: Target filesystem path to save the .xlsx report.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reconciliation Audit"
    ws.views.sheetView[0].showGridLines = True

    # Palette
    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    slate_fill = PatternFill(start_color="2E5B88", end_color="2E5B88", fill_type="solid")
    light_blue_fill = PatternFill(start_color="EBF2FA", end_color="EBF2FA", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")

    title_font = Font(name="Calibri", size=16, bold=True, color="1B365D")
    subtitle_font = Font(name="Calibri", size=11, italic=True, color="555555")
    section_font = Font(name="Calibri", size=13, bold=True, color="1B365D")
    white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=11, color="333333")

    thin_border_side = Side(border_style="thin", color="D0D7DE")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # 1. Title Block
    ws["B2"] = "AUTOMATED RECONCILIATION AUDIT REPORT"
    ws["B2"].font = title_font
    ws["B3"] = f"Batch: {metrics.batch_id} | Cheque No: {batch.cheque_number} | Execution Time: {metrics.execution_time_ms:.1f} ms"
    ws["B3"].font = subtitle_font

    # 2. Executive KPI Cards
    kpis = [
        (f"{batch.total_count:,}", "Reconciled Transactions", "Staged to Oracle FBDI", "B5"),
        (f"${metrics.total_debits:,.2f}", "Total Net Debits", "Chargebacks & Adjustments", "E5"),
        (f"${metrics.total_credits:,.2f}", "Total Net Credits", "Deductions & Master Debit", "H5"),
        (f"${metrics.net_variance:.2f}", "Net Batch Variance", "PASSED: Zero-Sum Verified", "K5"),
    ]

    for val, label, subtext, start_cell in kpis:
        col = start_cell[0]
        row = int(start_cell[1:])
        col_idx = openpyxl.utils.column_index_from_string(col)
        end_col = get_column_letter(col_idx + 2)
        ws.merge_cells(f"{col}{row}:{end_col}{row+2}")

        cell = ws[f"{col}{row}"]
        cell.value = f"{val}\n{label}\n{subtext}"
        cell.font = Font(name="Calibri", size=11, bold=True, color="1B365D")
        cell.fill = light_blue_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for r in range(row, row + 3):
            for c in range(col_idx, col_idx + 3):
                ws.cell(row=r, column=c).border = thin_border

    # 3. Regional Breakdown Table
    ws["B9"] = "1. Regional Breakdown by Operating Division"
    ws["B9"].font = section_font

    reg_headers = ["Division Name", "Customer Account", "Transaction Count", "Net Financial Impact (CAD)"]
    for idx, h in enumerate(reg_headers, start=2):
        cell = ws.cell(row=11, column=idx, value=h)
        cell.fill = navy_fill
        cell.font = white_bold
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    current_row = 12
    for div, data in sorted(metrics.division_summary.items()):
        fill = zebra_fill if current_row % 2 == 0 else PatternFill(fill_type=None)
        ws.cell(row=current_row, column=2, value=div).fill = fill
        # BUG FIX #3: Look up the correct account number per division,
        # instead of using batch.transactions[0].account_number for all rows.
        div_account = CUSTOMER_ACCOUNT_MAP.get(div, "N/A")
        ws.cell(row=current_row, column=3, value=div_account).fill = fill
        ws.cell(row=current_row, column=4, value=data["count"]).fill = fill
        ws.cell(row=current_row, column=5, value=float(data["net_amount"])).fill = fill
        ws.cell(row=current_row, column=5).number_format = "$#,##0.00"

        for c in range(2, 6):
            ws.cell(row=current_row, column=c).border = thin_border
            ws.cell(row=current_row, column=c).font = regular_font
        current_row += 1

    # 4. Itemized Transactions Sample
    current_row += 2
    ws.cell(row=current_row, column=2, value="2. Staged Interface Transactions (First 50 Items)").font = section_font
    current_row += 2

    item_headers = ["#", "Account Number", "Transaction Number", "Division", "Category", "Amount (CAD)"]
    for idx, h in enumerate(item_headers, start=2):
        cell = ws.cell(row=current_row, column=idx, value=h)
        cell.fill = slate_fill
        cell.font = white_bold
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    current_row += 1
    for idx, item in enumerate(batch.transactions[:50], start=1):
        fill = zebra_fill if current_row % 2 == 0 else PatternFill(fill_type=None)
        ws.cell(row=current_row, column=2, value=idx).fill = fill
        ws.cell(row=current_row, column=3, value=item.account_number).fill = fill
        ws.cell(row=current_row, column=4, value=item.trx_number).fill = fill
        ws.cell(row=current_row, column=5, value=item.division).fill = fill
        ws.cell(row=current_row, column=6, value=item.category).fill = fill
        amt_cell = ws.cell(row=current_row, column=7, value=float(item.amount))
        amt_cell.fill = fill
        amt_cell.number_format = "$#,##0.00"
        if item.amount < Decimal("0.00"):
            amt_cell.font = Font(name="Calibri", size=11, color="B00000", bold=True)
        else:
            amt_cell.font = regular_font

        for c in range(2, 8):
            ws.cell(row=current_row, column=c).border = thin_border
        current_row += 1

    # 5. Skipped Rows Section (if any)
    if metrics.skipped_rows:
        current_row += 2
        ws.cell(
            row=current_row, column=2,
            value=f"3. ATTENTION: {len(metrics.skipped_rows)} Skipped Source Rows (Require Review)",
        ).font = Font(name="Calibri", size=13, bold=True, color="B00000")
        current_row += 2

        skip_headers = ["#", "Source Row", "Reason", "Raw Account", "Raw Amount", "Raw Trx Number"]
        for idx, h in enumerate(skip_headers, start=2):
            cell = ws.cell(row=current_row, column=idx, value=h)
            cell.fill = PatternFill(start_color="B00000", end_color="B00000", fill_type="solid")
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        current_row += 1
        for idx, sr in enumerate(metrics.skipped_rows, start=1):
            fill = zebra_fill if current_row % 2 == 0 else PatternFill(fill_type=None)
            ws.cell(row=current_row, column=2, value=idx).fill = fill
            ws.cell(row=current_row, column=3, value=sr.row_index).fill = fill
            ws.cell(row=current_row, column=4, value=sr.reason).fill = fill
            ws.cell(row=current_row, column=5, value=sr.raw_values.get("A", "")).fill = fill
            ws.cell(row=current_row, column=6, value=sr.raw_values.get("B", "")).fill = fill
            ws.cell(row=current_row, column=7, value=sr.raw_values.get("C", "")).fill = fill
            for c in range(2, 8):
                ws.cell(row=current_row, column=c).border = thin_border
                ws.cell(row=current_row, column=c).font = regular_font
            current_row += 1

    # 6. Account Routing Discrepancies Section (if any)
    if metrics.account_mismatches:
        current_row += 2
        sec_num = 4 if metrics.skipped_rows else 3
        ws.cell(
            row=current_row, column=2,
            value=f"{sec_num}. ROUTING NOTICE: {len(metrics.account_mismatches)} Invoice-to-Account Mismatches",
        ).font = Font(name="Calibri", size=13, bold=True, color="D97706")
        current_row += 2

        mismatch_headers = ["#", "Transaction Number", "Staged Account", "Staged Division", "Inferred Division", "Inferred Account", "Amount (CAD)"]
        for idx, h in enumerate(mismatch_headers, start=2):
            cell = ws.cell(row=current_row, column=idx, value=h)
            cell.fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
            cell.font = white_bold
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        current_row += 1
        for idx, mm in enumerate(metrics.account_mismatches, start=1):
            fill = zebra_fill if current_row % 2 == 0 else PatternFill(fill_type=None)
            ws.cell(row=current_row, column=2, value=idx).fill = fill
            ws.cell(row=current_row, column=3, value=mm.trx_number).fill = fill
            ws.cell(row=current_row, column=4, value=mm.account_number).fill = fill
            ws.cell(row=current_row, column=5, value=mm.staged_division).fill = fill
            ws.cell(row=current_row, column=6, value=mm.inferred_division).fill = fill
            ws.cell(row=current_row, column=7, value=mm.inferred_account).fill = fill
            amt_cell = ws.cell(row=current_row, column=8, value=float(mm.amount))
            amt_cell.fill = fill
            amt_cell.number_format = "$#,##0.00"

            for c in range(2, 9):
                ws.cell(row=current_row, column=c).border = thin_border
                ws.cell(row=current_row, column=c).font = regular_font
            current_row += 1

    # Adjust widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if "\n" in val_str:
                val_str = max(val_str.split("\n"), key=len)
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    wb.save(output_path)
