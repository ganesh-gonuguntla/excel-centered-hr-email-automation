"""Generate results Excel report and display Rich stats table."""

from __future__ import annotations

import csv
import logging
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Font, PatternFill
from rich.console import Console
from rich.table import Table

logger = logging.getLogger(__name__)
console = Console()

# Color fills for the Excel report
FILL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FILL_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FILL_BLUE = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

FONT_GREEN = Font(color="006100")
FONT_RED = Font(color="9C0006")
FONT_YELLOW = Font(color="9C6500")
FONT_BLUE = Font(color="1F4E79")


class RunStats:
    """Accumulates statistics during a run."""

    def __init__(self):
        self.total: int = 0
        self.sent: int = 0
        self.failed: int = 0
        self.skipped_invalid: int = 0
        self.skipped_no_mx: int = 0
        self.skipped_duplicate: int = 0
        self.skipped_already_sent: int = 0
        self.dry_run: int = 0
        self.error_counts: Counter = Counter()
        self.failed_list: list[dict[str, str]] = []
        self.skipped_list: list[dict[str, str]] = []
        self.start_time: datetime = datetime.now()
        self.end_time: datetime | None = None

    def record(self, email: str, status: str, reason: str = "") -> None:
        """Record the outcome of one email."""
        self.total += 1
        if status == "SENT":
            self.sent += 1
        elif status == "DRY_RUN":
            self.dry_run += 1
        elif status == "FAILED":
            self.failed += 1
            self.error_counts[reason] += 1
            self.failed_list.append({"email": email, "reason": reason})
        elif status in ("SKIPPED_ALREADY_SENT", "SKIPPED_VISITED"):
            self.skipped_already_sent += 1
            self.skipped_list.append({"email": email, "reason": reason or "Already visited / sent in previous run"})
        elif status == "SKIPPED_INVALID_EMAIL":
            self.skipped_invalid += 1
            self.skipped_list.append({"email": email, "reason": reason or "Invalid email"})
        elif status == "SKIPPED_NO_MX":
            self.skipped_no_mx += 1
            self.skipped_list.append({"email": email, "reason": reason or "No MX record"})
        elif status == "SKIPPED_DUPLICATE":
            self.skipped_duplicate += 1
            self.skipped_list.append({"email": email, "reason": reason or "Duplicate"})

    def finalize(self) -> None:
        self.end_time = datetime.now()

    @property
    def success_rate(self) -> float:
        """Calculate success rate as a percentage."""
        total_attempted = self.sent + self.failed
        if total_attempted == 0:
            return 0.0
        return (self.sent / total_attempted) * 100

    @property
    def elapsed(self) -> str:
        """Human-readable elapsed time."""
        end = self.end_time or datetime.now()
        delta = end - self.start_time
        minutes, seconds = divmod(int(delta.total_seconds()), 60)
        return f"{minutes}m {seconds}s"

    @property
    def total_skipped(self) -> int:
        return (
            self.skipped_invalid
            + self.skipped_no_mx
            + self.skipped_duplicate
            + self.skipped_already_sent
        )


def print_stats(stats: RunStats, mask: bool = False) -> None:
    """Print a rich summary table to the console."""
    table = Table(title="📊 Email Campaign Results", show_header=True)
    table.add_column("Metric", style="bold cyan", min_width=30)
    table.add_column("Value", style="bold white", min_width=15)

    table.add_row("Total rows processed", str(stats.total))
    table.add_row("✅ Sent successfully", f"[green]{stats.sent}[/green]")
    if stats.dry_run > 0:
        table.add_row("🔵 Dry run (not sent)", f"[blue]{stats.dry_run}[/blue]")
    table.add_row("❌ Failed to send", f"[red]{stats.failed}[/red]")
    if stats.skipped_already_sent > 0:
        table.add_row("⏭️  Skipped: already sent/visited", f"[yellow]{stats.skipped_already_sent}[/yellow]")
    table.add_row("⏭️  Skipped: invalid email", f"[yellow]{stats.skipped_invalid}[/yellow]")
    table.add_row("⏭️  Skipped: no MX record", f"[yellow]{stats.skipped_no_mx}[/yellow]")
    table.add_row("⏭️  Skipped: duplicate in file", f"[yellow]{stats.skipped_duplicate}[/yellow]")
    table.add_row("📈 Success rate", f"{stats.success_rate:.1f}%")
    table.add_row("⏱️  Time taken", stats.elapsed)

    console.print()
    console.print(table)

    # Error breakdown
    if stats.error_counts:
        err_table = Table(title="❌ Failure Breakdown", show_header=True)
        err_table.add_column("Error", style="red")
        err_table.add_column("Count", style="bold")
        for error, count in stats.error_counts.most_common():
            err_table.add_row(error[:80], str(count))
        console.print(err_table)

    # Failed emails
    if stats.failed_list:
        fail_table = Table(title="📋 Failed Emails", show_header=True)
        fail_table.add_column("Email", style="red")
        fail_table.add_column("Reason")
        for item in stats.failed_list:
            email_display = _mask_email(item["email"]) if mask else item["email"]
            fail_table.add_row(email_display, item["reason"][:60])
        console.print(fail_table)

    # Skipped emails
    if stats.skipped_list:
        skip_table = Table(title="📋 Skipped Emails", show_header=True)
        skip_table.add_column("Email", style="yellow")
        skip_table.add_column("Reason")
        for item in stats.skipped_list:
            email_display = _mask_email(item["email"]) if mask else item["email"]
            skip_table.add_row(email_display, item["reason"])
        console.print(skip_table)


def _mask_email(email: str) -> str:
    """Partially mask an email address: a***z@domain.com"""
    if not email or "@" not in email:
        return email
    local, domain = email.split("@", 1)
    if len(local) <= 2:
        masked_local = local[0] + "***"
    else:
        masked_local = local[0] + "***" + local[-1]
    return f"{masked_local}@{domain}"


# ---------- sent_log.csv ----------


def load_visited_records(log_path: str | Path) -> dict[str, dict[str, Any]]:
    """Load all visited email records from sent_log.csv.

    Returns:
        dict mapping email.lower() -> {
            'status': status,
            'reason': reason,
            'subject': subject,
            'timestamp': timestamp,
            'is_sent': bool,
        }
    """
    path = Path(log_path)
    visited: dict[str, dict[str, Any]] = {}
    if not path.exists():
        return visited
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            raw_email = row.get("email", "")
            email = raw_email.lower().strip()
            if not email:
                continue
            status = row.get("status", "")
            # If already marked SENT in any previous row, preserve is_sent=True
            was_sent = visited.get(email, {}).get("is_sent", False) or (status == "SENT")
            visited[email] = {
                "status": status,
                "reason": row.get("reason", ""),
                "subject": row.get("subject", ""),
                "timestamp": row.get("timestamp", ""),
                "is_sent": was_sent,
            }
    return visited


def load_sent_log(log_path: str | Path) -> set[str]:
    """Load emails previously sent successfully from sent_log.csv."""
    visited = load_visited_records(log_path)
    return {email for email, r in visited.items() if r.get("is_sent", False)}


def load_failed_emails(log_path: str | Path) -> set[str]:
    """Load emails that previously FAILED from sent_log.csv."""
    path = Path(log_path)
    failed: set[str] = set()
    if not path.exists():
        return failed
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "FAILED":
                email = row.get("email", "").lower().strip()
                if email:
                    failed.add(email)
    # Remove any that were later successfully sent
    sent = load_sent_log(log_path)
    failed -= sent
    return failed


def append_to_sent_log(
    log_path: str | Path,
    email: str,
    status: str,
    reason: str,
    subject: str,
    timestamp: str,
) -> None:
    """Append one result row to sent_log.csv (creates the file if needed)."""
    path = Path(log_path)
    file_exists = path.exists() and path.stat().st_size > 0
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["email", "status", "reason", "subject", "timestamp"],
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "email": email,
            "status": status,
            "reason": reason,
            "subject": subject,
            "timestamp": timestamp,
        })
        f.flush()


def get_today_send_count(log_path: str | Path) -> int:
    """Count how many emails were SENT today (for daily cap tracking)."""
    path = Path(log_path)
    if not path.exists():
        return 0
    today_str = datetime.now().strftime("%Y-%m-%d")
    count = 0
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "SENT":
                ts = row.get("timestamp", "")
                if ts.startswith(today_str):
                    count += 1
    return count


# ---------- Results Excel ----------


def write_results_excel(
    results: list[dict[str, Any]],
    original_df: pd.DataFrame,
    output_dir: str | Path = "output",
) -> Path:
    """Write the results report to an Excel file with color-coded status."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"results_{timestamp}.xlsx"
    output_path = output_dir / filename

    # Build results DataFrame
    result_df = original_df.copy()

    # Add result columns
    statuses = [r.get("status", "") for r in results]
    reasons = [r.get("reason", "") for r in results]
    subjects = [r.get("subject", "") for r in results]
    timestamps = [r.get("timestamp", "") for r in results]

    # Pad if results are shorter than df (e.g., interrupted run)
    while len(statuses) < len(result_df):
        statuses.append("")
        reasons.append("")
        subjects.append("")
        timestamps.append("")

    result_df["Status"] = statuses[: len(result_df)]
    result_df["Reason"] = reasons[: len(result_df)]
    result_df["Subject Used"] = subjects[: len(result_df)]
    result_df["Timestamp"] = timestamps[: len(result_df)]

    # Write to Excel
    result_df.to_excel(output_path, index=False, engine="openpyxl")

    # Apply color coding
    from openpyxl import load_workbook

    wb = load_workbook(output_path)
    ws = wb.active

    # Find the Status column index
    status_col = None
    for col_idx, cell in enumerate(ws[1], 1):
        if cell.value == "Status":
            status_col = col_idx
            break

    if status_col:
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=status_col)
            status = str(cell.value).upper() if cell.value else ""
            if status == "SENT":
                cell.fill = FILL_GREEN
                cell.font = FONT_GREEN
            elif status == "FAILED":
                cell.fill = FILL_RED
                cell.font = FONT_RED
            elif status.startswith("SKIPPED"):
                cell.fill = FILL_YELLOW
                cell.font = FONT_YELLOW
            elif status == "DRY_RUN":
                cell.fill = FILL_BLUE
                cell.font = FONT_BLUE

    wb.save(output_path)
    logger.info(f"Results written to {output_path}")
    return output_path


def summarize_sent_log(log_path: str | Path) -> None:
    """Print summary statistics from sent_log.csv."""
    path = Path(log_path)
    if not path.exists():
        console.print("[yellow]No sent_log.csv found. Run the mailer first.[/yellow]")
        return

    stats = RunStats()
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            email = row.get("email", "")
            status = row.get("status", "")
            reason = row.get("reason", "")
            stats.record(email, status, reason)

    stats.finalize()
    console.print(f"\n[bold]Sent log: {path}[/bold]")
    print_stats(stats)
