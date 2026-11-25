"""
Internship Mailer – CLI entry point.

A production-quality CLI tool that reads HR/recruiter contacts from an Excel
file and sends personalized internship-request emails via SMTP.

Usage:
    python main.py run --file contacts.xlsx [--sheet Ranked] [--limit 50] ...
    python main.py preview --file contacts.xlsx
    python main.py test-send --file contacts.xlsx
    python main.py stats
    python main.py followup --file contacts.xlsx
"""

from __future__ import annotations

import logging
import os
import signal
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import typer
import yaml
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn

from mailer.personalizer import EmailPersonalizer
from mailer.reader import get_score_for_row, load_contacts, normalize_row
from mailer.reporter import (
    RunStats,
    append_to_sent_log,
    get_today_send_count,
    load_failed_emails,
    load_sent_log,
    load_visited_records,
    print_stats,
    summarize_sent_log,
    write_results_excel,
)
from mailer.sender import MailSender
from mailer.validator import validate_email_full

console = Console()
app = typer.Typer(
    name="internship-mailer",
    help="Excel-driven cold email automation for internship outreach.",
    add_completion=False,
)

# ----- Globals for graceful shutdown -----
_shutdown_requested = False


def _signal_handler(signum, frame):
    """Handle Ctrl+C gracefully."""
    global _shutdown_requested
    _shutdown_requested = True
    console.print("\n[yellow]⚠️  Ctrl+C received. Finishing current email and saving...[/yellow]")


# ----- Config loading helpers -----

BASE_DIR = Path(__file__).parent
SENT_LOG = BASE_DIR / "sent_log.csv"
LOG_DIR = BASE_DIR / "logs"
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE_DIR = BASE_DIR / "templates"


def _load_yaml(path: Path, name: str) -> dict:
    """Load a YAML file, fail fast with a clear message if missing."""
    if not path.exists():
        console.print(f"[red]❌ {name} not found: {path}[/red]")
        console.print(f"[dim]Copy the example file and fill in your details.[/dim]")
        raise typer.Exit(1)
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return data


def _check_placeholders(data: dict, name: str) -> None:
    """Check for any REPLACE_ME placeholders still present."""
    for key, value in data.items():
        if isinstance(value, str) and value.strip() == "REPLACE_ME":
            console.print(
                f"[red]❌ {name}: '{key}' still has the placeholder 'REPLACE_ME'. "
                f"Please fill in your details.[/red]"
            )
            raise typer.Exit(1)
        if isinstance(value, list):
            for item in value:
                if isinstance(item, str) and item.strip() == "REPLACE_ME":
                    console.print(
                        f"[red]❌ {name}: '{key}' list contains 'REPLACE_ME'. "
                        f"Please fill in all values.[/red]"
                    )
                    raise typer.Exit(1)


def _setup_logging() -> None:
    """Configure logging to file and console."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def _load_all_config() -> tuple[dict, dict, str, str]:
    """Load config.yaml, profile.yaml, and .env. Returns (config, profile, smtp_user, smtp_pass)."""
    load_dotenv(BASE_DIR / ".env")

    config = _load_yaml(BASE_DIR / "config.yaml", "config.yaml")
    profile = _load_yaml(BASE_DIR / "profile.yaml", "profile.yaml")
    _check_placeholders(profile, "profile.yaml")

    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_APP_PASSWORD", "")

    if not smtp_user or not smtp_pass:
        console.print(
            "[red]❌ SMTP_USER and SMTP_APP_PASSWORD must be set in .env[/red]"
        )
        raise typer.Exit(1)

    # Verify email matches
    if profile.get("email") and profile["email"] != smtp_user:
        console.print(
            "[yellow]⚠️  Warning: profile.yaml 'email' differs from SMTP_USER in .env. "
            "Using SMTP_USER for sending.[/yellow]"
        )

    return config, profile, smtp_user, smtp_pass


# ----- CLI Commands -----


@app.command()
def run(
    file: str = typer.Option(..., "--file", "-f", help="Path to contacts Excel/CSV file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name (default: first sheet)"),
    limit: Optional[int] = typer.Option(None, "--limit", "-l", help="Max number of rows to process"),
    min_score: Optional[int] = typer.Option(None, "--min-score", help="Minimum Intern Interest Score"),
    start_row: int = typer.Option(1, "--start-row", help="Start from this row (1-based)"),
    dry_run: bool = typer.Option(False, "--dry-run", help="Validate and log, but don't send emails"),
    retry_failed: bool = typer.Option(False, "--retry-failed", help="Only retry previously FAILED rows"),
    mask: bool = typer.Option(False, "--mask", help="Mask email addresses in console output"),
    resume: Optional[str] = typer.Option(None, "--resume", "-r", help="Path to resume PDF to attach for high-scored contacts"),
):
    """Send personalized internship emails to contacts in the Excel file."""
    _setup_logging()
    logger = logging.getLogger("main.run")

    signal.signal(signal.SIGINT, _signal_handler)
    if sys.platform == "win32":
        signal.signal(signal.SIGBREAK, _signal_handler)

    config, profile, smtp_user, smtp_pass = _load_all_config()

    # Override resume path from CLI if provided
    if resume:
        profile["resume_path"] = resume

    smtp_config = config.get("smtp", {})
    send_config = config.get("sending", {})
    column_mapping = config.get("column_mapping", None)

    # Load contacts
    console.print(Panel(f"📂 Loading contacts from: [bold]{file}[/bold]"))
    try:
        df, col_map = load_contacts(
            file,
            sheet=sheet,
            column_mapping=column_mapping,
            min_score=min_score,
            limit=limit,
            start_row=start_row,
        )
    except Exception as e:
        console.print(f"[red]❌ Failed to load contacts: {e}[/red]")
        raise typer.Exit(1)

    console.print(f"[green]✅ Loaded {len(df)} rows[/green]")

    # Load sent log and visited history for resume / dedup
    visited_records = load_visited_records(SENT_LOG)
    sent_emails = load_sent_log(SENT_LOG)
    failed_emails = load_failed_emails(SENT_LOG) if retry_failed else set()

    console.print(
        f"[dim]History record: {len(visited_records)} emails previously visited "
        f"({len(sent_emails)} successfully sent)[/dim]"
    )

    if retry_failed:
        console.print(f"[cyan]🔄 Retry mode: {len(failed_emails)} previously failed emails[/cyan]")

    # Daily cap check
    daily_cap = send_config.get("daily_cap", 100)
    today_count = get_today_send_count(SENT_LOG)
    remaining_cap = max(0, daily_cap - today_count)
    console.print(f"[dim]Daily cap: {daily_cap} | Sent today: {today_count} | Remaining: {remaining_cap}[/dim]")

    if remaining_cap == 0 and not dry_run:
        console.print(
            "[yellow]⚠️  Daily cap reached! Re-run tomorrow. The tool will resume where it left off.[/yellow]"
        )
        raise typer.Exit(0)

    # Init personalizer and sender
    personalizer = EmailPersonalizer(TEMPLATE_DIR, profile)

    sender = MailSender(
        smtp_host=smtp_config.get("host", "smtp.gmail.com"),
        smtp_port=smtp_config.get("port", 587),
        smtp_user=smtp_user,
        smtp_password=smtp_pass,
        use_ssl=smtp_config.get("use_ssl", False),
        sender_name=profile.get("sender_name", ""),
        sender_email=smtp_user,
        reply_to=smtp_user,
        resume_path=profile.get("resume_path"),
        delay_range=tuple(send_config.get("delay_range", [25, 60])),
        max_retries=send_config.get("max_retries", 2),
    )

    if not dry_run:
        try:
            sender.connect()
        except Exception as e:
            console.print(f"[red]❌ SMTP connection failed: {e}[/red]")
            console.print("[dim]Check your .env credentials. For Gmail, use an App Password.[/dim]")
            raise typer.Exit(1)

    # Process rows
    stats = RunStats()
    results: list[dict] = []
    seen_emails: set[str] = set()
    sends_this_run = 0

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Sending emails...", total=len(df))

        for idx, row in df.iterrows():
            if _shutdown_requested:
                console.print("[yellow]⚠️  Shutdown requested. Saving progress...[/yellow]")
                break

            normalized = normalize_row(row)
            email = normalized["email"]
            name = normalized["name"]
            company = normalized["company"]
            designation = normalized["designation"]
            email_lower = email.lower().strip()

            # In retry mode, skip emails that didn't previously fail
            if retry_failed and email_lower not in failed_emails:
                progress.update(task, advance=1)
                results.append({
                    "status": "",
                    "reason": "Not in retry set",
                    "subject": "",
                    "timestamp": "",
                })
                continue

            # -------------------------------------------------------------
            # PRE-PIPELINE CHECK 1: Already visited & sent in a previous run
            # -------------------------------------------------------------
            if not retry_failed and (email_lower in sent_emails):
                prev_info = visited_records.get(email_lower, {})
                prev_ts = prev_info.get("timestamp", "")
                reason = f"Already sent in previous run ({prev_ts})" if prev_ts else "Already sent in previous run"
                logger.info(f"Row {idx + 1}: SKIPPED_ALREADY_SENT - {email if not mask else '***'} ({reason})")
                stats.record(email, "SKIPPED_ALREADY_SENT", reason)
                now = datetime.now().isoformat()
                results.append({
                    "status": "SKIPPED_ALREADY_SENT",
                    "reason": reason,
                    "subject": prev_info.get("subject", ""),
                    "timestamp": now,
                })
                progress.update(task, advance=1)
                continue

            # -------------------------------------------------------------
            # PRE-PIPELINE CHECK 2: Duplicate within the current file
            # -------------------------------------------------------------
            if email_lower in seen_emails:
                reason = "Duplicate email in current file"
                logger.info(f"Row {idx + 1}: SKIPPED_DUPLICATE - {email if not mask else '***'} ({reason})")
                stats.record(email, "SKIPPED_DUPLICATE", reason)
                now = datetime.now().isoformat()
                results.append({
                    "status": "SKIPPED_DUPLICATE",
                    "reason": reason,
                    "subject": "",
                    "timestamp": now,
                })
                progress.update(task, advance=1)
                continue

            # -------------------------------------------------------------
            # PRE-PIPELINE CHECK 3: Previously verified invalid or no-MX
            # -------------------------------------------------------------
            if not retry_failed and email_lower in visited_records:
                prev_info = visited_records[email_lower]
                prev_status = prev_info.get("status")
                if prev_status == "SKIPPED_NO_MX":
                    prev_ts = prev_info.get("timestamp", "")
                    reason = f"Previously verified no MX record ({prev_ts})" if prev_ts else "Previously verified no MX record"
                    logger.info(f"Row {idx + 1}: SKIPPED_NO_MX - {email if not mask else '***'} ({reason})")
                    stats.record(email, "SKIPPED_NO_MX", reason)
                    now = datetime.now().isoformat()
                    results.append({
                        "status": "SKIPPED_NO_MX",
                        "reason": reason,
                        "subject": "",
                        "timestamp": now,
                    })
                    progress.update(task, advance=1)
                    continue
                elif prev_status == "SKIPPED_INVALID_EMAIL":
                    reason = "Previously verified invalid email syntax"
                    logger.info(f"Row {idx + 1}: SKIPPED_INVALID_EMAIL - {email if not mask else '***'} ({reason})")
                    stats.record(email, "SKIPPED_INVALID_EMAIL", reason)
                    now = datetime.now().isoformat()
                    results.append({
                        "status": "SKIPPED_INVALID_EMAIL",
                        "reason": reason,
                        "subject": "",
                        "timestamp": now,
                    })
                    progress.update(task, advance=1)
                    continue

            # Validate
            is_valid, validation_status = validate_email_full(
                email, sent_emails, seen_emails
            )

            if not is_valid:
                logger.info(f"Row {idx + 1}: {validation_status} - {email if not mask else '***'}")
                stats.record(email, validation_status)
                now = datetime.now().isoformat()
                append_to_sent_log(SENT_LOG, email, validation_status, validation_status, "", now)
                results.append({
                    "status": validation_status,
                    "reason": validation_status,
                    "subject": "",
                    "timestamp": now,
                })
                progress.update(task, advance=1)
                continue

            seen_emails.add(email.lower())

            # Check daily cap
            if sends_this_run >= remaining_cap and not dry_run:
                console.print(
                    f"\n[yellow]⚠️  Daily cap of {daily_cap} reached! "
                    f"Save progress and re-run tomorrow.[/yellow]"
                )
                # Record remaining rows
                results.append({
                    "status": "SKIPPED_CAP",
                    "reason": "Daily cap reached",
                    "subject": "",
                    "timestamp": datetime.now().isoformat(),
                })
                break

            # Personalize
            try:
                rendered = personalizer.render(name, email, company, designation)
            except Exception as e:
                logger.error(f"Template rendering error for row {idx + 1}: {e}")
                stats.record(email, "FAILED", f"Template error: {e}")
                now = datetime.now().isoformat()
                append_to_sent_log(SENT_LOG, email, "FAILED", f"Template error: {e}", "", now)
                results.append({
                    "status": "FAILED",
                    "reason": f"Template error: {e}",
                    "subject": "",
                    "timestamp": now,
                })
                progress.update(task, advance=1)
                continue

            # Determine if resume should be attached (score 3-5)
            score = get_score_for_row(row, df.columns)
            attach_resume = False
            if score is not None and score >= 3:
                attach_resume = True

            # Send
            email_display = email if not mask else "***"
            progress.update(task, description=f"Sending to {email_display}...")

            result = sender.send(
                to_email=email,
                subject=rendered["subject"],
                body_text=rendered["body_text"],
                body_html=rendered["body_html"],
                attach_resume=attach_resume,
                dry_run=dry_run,
            )

            status = result["status"]
            error = result.get("error", "")
            now = datetime.now().isoformat()

            stats.record(email, status, error or "")
            append_to_sent_log(SENT_LOG, email, status, error or "", rendered["subject"], now)
            results.append({
                "status": status,
                "reason": error or "",
                "subject": rendered["subject"],
                "timestamp": now,
            })

            if status == "SENT":
                sends_this_run += 1
                logger.info(f"✅ Row {idx + 1}: SENT to {email}")

            if status in ("SENT", "DRY_RUN"):
                # Throttle only after successful sends
                if not dry_run and not _shutdown_requested:
                    sender.throttle()

            progress.update(task, advance=1)

    # Finalize
    stats.finalize()
    sender.close()

    # Write results Excel
    try:
        output_path = write_results_excel(results, df, OUTPUT_DIR)
        console.print(f"\n[green]📄 Results saved to: {output_path}[/green]")
    except Exception as e:
        console.print(f"[red]⚠️  Failed to write results Excel: {e}[/red]")

    # Print stats
    print_stats(stats, mask=mask)
    console.print("\n[bold green]✅ Run complete![/bold green]")


@app.command()
def preview(
    file: str = typer.Option(..., "--file", "-f", help="Path to contacts Excel/CSV file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name"),
    count: int = typer.Option(3, "--count", "-n", help="Number of sample emails to preview"),
):
    """Render sample emails to the terminal for proofreading."""
    _setup_logging()
    config, profile, _, _ = _load_all_config()
    column_mapping = config.get("column_mapping", None)

    df, _ = load_contacts(file, sheet=sheet, column_mapping=column_mapping)
    personalizer = EmailPersonalizer(TEMPLATE_DIR, profile)

    previewed = 0
    for idx, row in df.iterrows():
        if previewed >= count:
            break

        normalized = normalize_row(row)
        email = normalized["email"]
        if not email or not email.strip():
            continue

        rendered = personalizer.render(
            normalized["name"],
            email,
            normalized["company"],
            normalized["designation"],
        )

        score = get_score_for_row(row, df.columns)
        attach_note = ""
        if score is not None and score >= 3:
            attach_note = " [📎 Resume attached]"

        console.print(Panel(
            f"[bold]To:[/bold] {email}\n"
            f"[bold]Subject:[/bold] {rendered['subject']}\n"
            f"[bold]Template:[/bold] {rendered['template_type'].upper()}{attach_note}\n"
            f"[bold]Greeting:[/bold] {rendered['greeting']}\n\n"
            f"{rendered['body_text']}",
            title=f"📧 Preview {previewed + 1}",
            border_style="cyan",
        ))
        previewed += 1

    if previewed == 0:
        console.print("[yellow]No valid rows found for preview.[/yellow]")
    else:
        console.print(f"\n[green]Previewed {previewed} sample email(s).[/green]")


@app.command(name="test-send")
def test_send(
    file: str = typer.Option(..., "--file", "-f", help="Path to contacts Excel/CSV file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name"),
    resume: Optional[str] = typer.Option(None, "--resume", "-r", help="Path to resume PDF"),
):
    """Send ONE test email to yourself using the first valid row's data."""
    _setup_logging()
    logger = logging.getLogger("main.test-send")

    config, profile, smtp_user, smtp_pass = _load_all_config()

    if resume:
        profile["resume_path"] = resume

    smtp_config = config.get("smtp", {})
    column_mapping = config.get("column_mapping", None)

    df, _ = load_contacts(file, sheet=sheet, column_mapping=column_mapping, limit=10)
    personalizer = EmailPersonalizer(TEMPLATE_DIR, profile)

    # Find first valid row
    for idx, row in df.iterrows():
        normalized = normalize_row(row)
        email = normalized["email"]
        if email and "@" in email:
            rendered = personalizer.render(
                normalized["name"],
                email,
                normalized["company"],
                normalized["designation"],
            )

            sender = MailSender(
                smtp_host=smtp_config.get("host", "smtp.gmail.com"),
                smtp_port=smtp_config.get("port", 587),
                smtp_user=smtp_user,
                smtp_password=smtp_pass,
                use_ssl=smtp_config.get("use_ssl", False),
                sender_name=profile.get("sender_name", ""),
                sender_email=smtp_user,
                reply_to=smtp_user,
                resume_path=profile.get("resume_path"),
            )

            console.print(
                f"[cyan]📧 Sending test email to [bold]{smtp_user}[/bold] "
                f"using data from row {idx + 1}...[/cyan]"
            )

            sender.connect()
            result = sender.send(
                to_email=smtp_user,  # Send to self!
                subject=f"[TEST] {rendered['subject']}",
                body_text=rendered["body_text"],
                body_html=rendered["body_html"],
                attach_resume=True,
            )
            sender.close()

            if result["status"] == "SENT":
                console.print("[green]✅ Test email sent! Check your inbox.[/green]")
            else:
                console.print(f"[red]❌ Test email failed: {result['error']}[/red]")

            return

    console.print("[red]No valid rows found for test send.[/red]")


@app.command()
def stats():
    """Summarize sent_log.csv statistics."""
    _setup_logging()
    summarize_sent_log(SENT_LOG)


@app.command()
def followup(
    file: str = typer.Option(..., "--file", "-f", help="Path to contacts Excel/CSV file"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s", help="Sheet name"),
    days: int = typer.Option(7, "--days", "-d", help="Min days since original email"),
    dry_run: bool = typer.Option(False, "--dry-run", help="Don't actually send"),
    mask: bool = typer.Option(False, "--mask", help="Mask emails in console"),
    resume: Optional[str] = typer.Option(None, "--resume", "-r", help="Path to resume PDF"),
):
    """Send follow-up emails to contacts who were SENT at least N days ago."""
    import csv

    _setup_logging()
    logger = logging.getLogger("main.followup")

    signal.signal(signal.SIGINT, _signal_handler)
    if sys.platform == "win32":
        signal.signal(signal.SIGBREAK, _signal_handler)

    config, profile, smtp_user, smtp_pass = _load_all_config()

    if resume:
        profile["resume_path"] = resume

    smtp_config = config.get("smtp", {})
    send_config = config.get("sending", {})
    column_mapping = config.get("column_mapping", None)

    followup_log = BASE_DIR / "followup_log.csv"

    # Load contacts who were SENT
    if not SENT_LOG.exists():
        console.print("[yellow]No sent_log.csv found. Run the mailer first.[/yellow]")
        raise typer.Exit(0)

    # Build set of emails eligible for follow-up
    eligible: dict[str, dict] = {}  # email -> {subject, timestamp}
    already_followed_up: set[str] = set()

    # Load followup log
    if followup_log.exists():
        with open(followup_log, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("status") in ("SENT", "DRY_RUN"):
                    already_followed_up.add(row.get("email", "").lower())

    # Find eligible SENT contacts
    with open(SENT_LOG, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "SENT":
                email = row.get("email", "").lower()
                ts_str = row.get("timestamp", "")
                subject = row.get("subject", "")

                if email in already_followed_up:
                    continue

                try:
                    sent_dt = datetime.fromisoformat(ts_str)
                    delta = (datetime.now() - sent_dt).days
                    if delta >= days:
                        eligible[email] = {"subject": subject, "timestamp": ts_str}
                except (ValueError, TypeError):
                    continue

    if not eligible:
        console.print(f"[yellow]No contacts eligible for follow-up (min {days} days).[/yellow]")
        raise typer.Exit(0)

    console.print(f"[green]Found {len(eligible)} contact(s) eligible for follow-up.[/green]")

    # Load the original contacts file for name/company info
    df, _ = load_contacts(file, sheet=sheet, column_mapping=column_mapping)
    personalizer = EmailPersonalizer(TEMPLATE_DIR, profile)

    # Setup sender
    sender = MailSender(
        smtp_host=smtp_config.get("host", "smtp.gmail.com"),
        smtp_port=smtp_config.get("port", 587),
        smtp_user=smtp_user,
        smtp_password=smtp_pass,
        use_ssl=smtp_config.get("use_ssl", False),
        sender_name=profile.get("sender_name", ""),
        sender_email=smtp_user,
        reply_to=smtp_user,
        resume_path=profile.get("resume_path"),
        delay_range=tuple(send_config.get("delay_range", [25, 60])),
    )

    if not dry_run:
        sender.connect()

    stats = RunStats()
    from jinja2 import Environment, FileSystemLoader, select_autoescape

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=select_autoescape(["html", "htm", "xml", "j2"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )

    for idx, row in df.iterrows():
        if _shutdown_requested:
            break

        normalized = normalize_row(row)
        email = normalized["email"].lower()

        if email not in eligible:
            continue

        orig_data = eligible[email]
        from mailer.personalizer import build_greeting

        greeting = build_greeting(normalized["name"])
        ctx = {
            "greeting": greeting,
            "company": normalized["company"],
            "sender_name": profile.get("sender_name", ""),
            "internship_interest": profile.get("internship_interest", ""),
            "phone": profile.get("phone", ""),
            "email": profile.get("email", ""),
        }

        body_html = env.get_template("followup.html.j2").render(**ctx)
        body_text = env.get_template("followup.txt.j2").render(**ctx)
        subject = f"Re: {orig_data['subject']}"

        result = sender.send(
            to_email=email,
            subject=subject,
            body_text=body_text,
            body_html=body_html,
            attach_resume=True,
            dry_run=dry_run,
        )

        status = result["status"]
        error = result.get("error", "")
        now = datetime.now().isoformat()

        stats.record(email, status, error or "")
        append_to_sent_log(followup_log, email, status, error or "", subject, now)

        if status in ("SENT", "DRY_RUN") and not _shutdown_requested and not dry_run:
            sender.throttle()

    stats.finalize()
    sender.close()
    print_stats(stats, mask=mask)
    console.print("[bold green]✅ Follow-up run complete![/bold green]")


if __name__ == "__main__":
    app()
