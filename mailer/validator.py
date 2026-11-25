"""Email validation: syntax check + DNS MX record lookup."""

from __future__ import annotations

import logging
import socket
from functools import lru_cache
from typing import Literal

import dns.resolver
from email_validator import EmailNotValidError, validate_email

logger = logging.getLogger(__name__)

# Validation result type
ValidationResult = tuple[
    bool,  # is_valid
    Literal[
        "VALID",
        "SKIPPED_INVALID_EMAIL",
        "SKIPPED_NO_MX",
        "SKIPPED_DUPLICATE",
    ],
]


def check_syntax(email: str) -> tuple[bool, str | None]:
    """Check email syntax using email-validator (no deliverability check).

    Returns (is_valid, normalized_email_or_None).
    """
    if not email or not email.strip():
        return False, None
    try:
        result = validate_email(email.strip(), check_deliverability=False)
        return True, result.normalized
    except EmailNotValidError:
        return False, None


@lru_cache(maxsize=4096)
def check_mx(domain: str, timeout: float = 5.0) -> bool:
    """Check if domain has MX or A records. Results are cached per domain."""
    try:
        resolver = dns.resolver.Resolver()
        resolver.timeout = timeout
        resolver.lifetime = timeout
        try:
            answers = resolver.resolve(domain, "MX")
            if answers:
                return True
        except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.resolver.NoNameservers):
            pass

        # Fallback: check for A record
        try:
            answers = resolver.resolve(domain, "A")
            if answers:
                return True
        except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, dns.resolver.NoNameservers):
            pass

        return False
    except dns.exception.Timeout:
        logger.warning(f"DNS timeout for domain: {domain}")
        return False
    except Exception as e:
        logger.warning(f"DNS lookup error for {domain}: {e}")
        return False


def validate_email_full(
    email: str,
    sent_emails: set[str],
    seen_emails: set[str],
) -> ValidationResult:
    """Full validation pipeline for a single email address.

    Args:
        email: The email address to validate.
        sent_emails: Set of emails already sent successfully (from sent_log.csv).
        seen_emails: Set of emails already processed in this run.

    Returns:
        (is_valid, status_string)
    """
    # 1. Syntax check
    is_valid_syntax, normalized = check_syntax(email)
    if not is_valid_syntax or normalized is None:
        return False, "SKIPPED_INVALID_EMAIL"

    email_lower = normalized.lower()

    # 2. Duplicate check (already sent in previous run OR earlier in this file)
    if email_lower in sent_emails or email_lower in seen_emails:
        return False, "SKIPPED_DUPLICATE"

    # 3. MX record check
    domain = email_lower.split("@")[1]
    if not check_mx(domain):
        return False, "SKIPPED_NO_MX"

    return True, "VALID"
