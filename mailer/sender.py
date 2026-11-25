"""SMTP connection management, email sending with retry and throttling."""

from __future__ import annotations

import logging
import mimetypes
import random
import smtplib
import ssl
import time
from email.message import EmailMessage
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class SMTPSendError(Exception):
    """Wraps an SMTP send failure with error details."""

    def __init__(self, message: str, is_permanent: bool = False):
        super().__init__(message)
        self.is_permanent = is_permanent


class MailSender:
    """Manages an SMTP connection and sends emails with retry and throttling."""

    def __init__(
        self,
        smtp_host: str = "smtp.gmail.com",
        smtp_port: int = 587,
        smtp_user: str = "",
        smtp_password: str = "",
        use_ssl: bool = False,
        sender_name: str = "",
        sender_email: str = "",
        reply_to: str = "",
        resume_path: str | None = None,
        delay_range: tuple[int, int] = (25, 60),
        max_retries: int = 2,
    ):
        self.smtp_host = smtp_host
        self.smtp_port = smtp_port
        self.smtp_user = smtp_user
        self.smtp_password = smtp_password
        self.use_ssl = use_ssl
        self.sender_name = sender_name
        self.sender_email = sender_email
        self.reply_to = reply_to or sender_email
        self.resume_path = resume_path
        self.delay_range = delay_range
        self.max_retries = max_retries

        self._connection: smtplib.SMTP | smtplib.SMTP_SSL | None = None
        self._resume_data: bytes | None = None
        self._resume_filename: str | None = None

        # Pre-load resume if path is set
        if self.resume_path:
            rp = Path(self.resume_path)
            if rp.exists():
                self._resume_data = rp.read_bytes()
                self._resume_filename = rp.name
                logger.info(f"Resume loaded: {rp.name} ({len(self._resume_data)} bytes)")
            else:
                logger.warning(f"Resume file not found: {self.resume_path}")

    def connect(self) -> None:
        """Establish SMTP connection with TLS/SSL."""
        try:
            if self.use_ssl:
                context = ssl.create_default_context()
                self._connection = smtplib.SMTP_SSL(
                    self.smtp_host, self.smtp_port, context=context, timeout=30
                )
            else:
                self._connection = smtplib.SMTP(
                    self.smtp_host, self.smtp_port, timeout=30
                )
                self._connection.ehlo()
                self._connection.starttls(context=ssl.create_default_context())
                self._connection.ehlo()

            self._connection.login(self.smtp_user, self.smtp_password)
            logger.info(f"SMTP connected to {self.smtp_host}:{self.smtp_port}")
        except Exception as e:
            logger.error(f"SMTP connection failed: {e}")
            raise

    def _ensure_connection(self) -> None:
        """Reconnect if the connection has been dropped."""
        if self._connection is None:
            self.connect()
            return
        try:
            status = self._connection.noop()[0]
            if status != 250:
                raise smtplib.SMTPServerDisconnected("NOOP failed")
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPException, OSError):
            logger.info("SMTP connection dropped, reconnecting...")
            self.close()
            self.connect()

    def close(self) -> None:
        """Close the SMTP connection gracefully."""
        if self._connection:
            try:
                self._connection.quit()
            except Exception:
                pass
            self._connection = None

    def build_message(
        self,
        to_email: str,
        subject: str,
        body_text: str,
        body_html: str,
        attach_resume: bool = False,
    ) -> EmailMessage:
        """Build a multipart/alternative email message."""
        msg = EmailMessage()
        msg["From"] = f"{self.sender_name} <{self.sender_email}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        msg["Reply-To"] = self.reply_to

        # Set plain text as the main content
        msg.set_content(body_text)
        # Add HTML alternative
        msg.add_alternative(body_html, subtype="html")

        # Attach resume if requested and available
        if attach_resume and self._resume_data and self._resume_filename:
            maintype, subtype = ("application", "pdf")
            msg.add_attachment(
                self._resume_data,
                maintype=maintype,
                subtype=subtype,
                filename=self._resume_filename,
            )

        return msg

    def send(
        self,
        to_email: str,
        subject: str,
        body_text: str,
        body_html: str,
        attach_resume: bool = False,
        dry_run: bool = False,
    ) -> dict[str, Any]:
        """Send one email with retry logic.

        Returns:
            dict with keys: status ('SENT', 'FAILED', 'DRY_RUN'),
                            error (str or None), attempts (int)
        """
        msg = self.build_message(to_email, subject, body_text, body_html, attach_resume)

        if dry_run:
            logger.info(f"[DRY RUN] Would send to: {to_email} | Subject: {subject}")
            return {"status": "DRY_RUN", "error": None, "attempts": 0}

        last_error = ""
        for attempt in range(1, self.max_retries + 2):  # 1 initial + max_retries
            try:
                self._ensure_connection()
                self._connection.send_message(msg)  # type: ignore[union-attr]
                logger.info(f"Email sent to {to_email} (attempt {attempt})")
                return {"status": "SENT", "error": None, "attempts": attempt}

            except smtplib.SMTPRecipientsRefused as e:
                error_msg = str(e)
                logger.error(f"Recipient refused {to_email}: {error_msg}")
                return {
                    "status": "FAILED",
                    "error": f"Recipient refused: {error_msg}",
                    "attempts": attempt,
                }

            except smtplib.SMTPResponseException as e:
                last_error = f"SMTP {e.smtp_code}: {e.smtp_error}"
                logger.warning(
                    f"SMTP error sending to {to_email} (attempt {attempt}): {last_error}"
                )
                # 5xx = permanent failure, don't retry
                if 500 <= e.smtp_code < 600:
                    return {
                        "status": "FAILED",
                        "error": last_error,
                        "attempts": attempt,
                    }
                # 4xx or other = transient, retry
                if attempt <= self.max_retries:
                    backoff = 2**attempt + random.uniform(0, 1)
                    logger.info(f"Retrying in {backoff:.1f}s...")
                    time.sleep(backoff)
                    self._connection = None  # Force reconnect

            except (
                smtplib.SMTPServerDisconnected,
                ConnectionResetError,
                TimeoutError,
                OSError,
            ) as e:
                last_error = f"Connection error: {e}"
                logger.warning(
                    f"Connection error sending to {to_email} (attempt {attempt}): {e}"
                )
                self._connection = None
                if attempt <= self.max_retries:
                    backoff = 2**attempt + random.uniform(0, 1)
                    logger.info(f"Retrying in {backoff:.1f}s...")
                    time.sleep(backoff)

            except Exception as e:
                last_error = str(e)
                logger.error(f"Unexpected error sending to {to_email}: {e}")
                return {
                    "status": "FAILED",
                    "error": last_error,
                    "attempts": attempt,
                }

        return {
            "status": "FAILED",
            "error": f"Max retries exceeded. Last error: {last_error}",
            "attempts": self.max_retries + 1,
        }

    def throttle(self) -> None:
        """Wait a randomized delay between sends."""
        delay = random.uniform(*self.delay_range)
        logger.debug(f"Throttling: waiting {delay:.1f}s before next send")
        time.sleep(delay)
