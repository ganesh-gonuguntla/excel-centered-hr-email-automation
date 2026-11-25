"""Name/company cleaning, template selection, and Jinja2 rendering."""

from __future__ import annotations

import hashlib
import logging
import re
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape

logger = logging.getLogger(__name__)

# Honorifics / prefixes to strip when extracting first name
DEFAULT_HONORIFICS = {
    "mr", "mrs", "ms", "dr", "prof", "capt", "captain",
    "sir", "madam", "crp", "cpl", "sgt", "lt", "col",
    "gen", "rev", "hon",
}

# Tokens that look like a company / placeholder, not a real first name
PLACEHOLDER_TOKENS = {
    "hr", "hiring", "team", "admin", "info", "contact",
    "recruitment", "support", "office", "manager", "friend",
    "na", "n/a", "none", "null", "unknown", "-", ".",
}

# Keywords in Designation that trigger the TA template
TA_KEYWORDS = {
    "talent acquisition", "recruit", "campus", "staffing",
    "hiring", " ta", "ta ",
}


def extract_first_name(
    name: str,
    honorifics: set[str] | None = None,
) -> str | None:
    """Extract a usable first name from the contact's name.

    Returns the first name in Title Case, or None if no usable name is found.
    """
    if honorifics is None:
        honorifics = DEFAULT_HONORIFICS

    if not name or not name.strip():
        return None

    # Split on whitespace
    tokens = name.strip().split()

    # Filter out honorifics (with or without trailing period)
    cleaned = []
    for t in tokens:
        bare = t.rstrip(".").lower()
        if bare not in honorifics:
            cleaned.append(t)

    if not cleaned:
        return None

    # If only one token, check if it looks like a placeholder or company name
    if len(cleaned) == 1:
        single = cleaned[0].lower().rstrip(".,;:")
        if single in PLACEHOLDER_TOKENS:
            return None
        # If it's all uppercase and short (like an abbreviation), might be odd
        # But a single real name is fine (e.g. "Anita")
        if len(single) <= 1:
            return None

    # If ALL tokens are placeholders, fall back
    if all(t.lower().rstrip(".,;:") in PLACEHOLDER_TOKENS for t in cleaned):
        return None

    first = cleaned[0].strip(".,;:")

    # Final check: if the first token is in placeholders, skip to next usable
    if first.lower() in PLACEHOLDER_TOKENS:
        for t in cleaned[1:]:
            candidate = t.strip(".,;:")
            if candidate.lower() not in PLACEHOLDER_TOKENS and len(candidate) > 1:
                return candidate.title()
        return None

    return first.title()


def build_greeting(name: str, honorifics: set[str] | None = None) -> str:
    """Build the greeting line for an email."""
    first = extract_first_name(name, honorifics)
    if first:
        return f"Dear {first},"
    return "Dear Hiring Team,"


def select_template_type(designation: str) -> str:
    """Choose 'ta' or 'hr' template based on the designation."""
    desig_lower = f" {designation.lower()} "
    for keyword in TA_KEYWORDS:
        if keyword in desig_lower:
            return "ta"
    return "hr"


def select_subject(
    subjects: list[str],
    email: str,
) -> str:
    """Deterministically select a subject line for an email address.

    Uses a hash of the email so the same contact always gets the same subject.
    """
    if not subjects:
        raise ValueError("No subject lines available in subjects.txt")
    idx = int(hashlib.md5(email.lower().encode()).hexdigest(), 16) % len(subjects)
    return subjects[idx]


def load_subjects(template_dir: str | Path) -> list[str]:
    """Load subject templates from subjects.txt (one per line, skip blanks)."""
    path = Path(template_dir) / "subjects.txt"
    if not path.exists():
        raise FileNotFoundError(f"Subject file not found: {path}")
    lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines()]
    subjects = [line for line in lines if line and not line.startswith("#")]
    if not subjects:
        raise ValueError(f"No subject lines found in {path}")
    return subjects


class EmailPersonalizer:
    """Renders personalized email content using Jinja2 templates."""

    def __init__(
        self,
        template_dir: str | Path,
        profile: dict[str, Any],
        honorifics: set[str] | None = None,
    ):
        self.template_dir = Path(template_dir)
        self.profile = profile
        self.honorifics = honorifics or DEFAULT_HONORIFICS
        self.subjects = load_subjects(self.template_dir)

        self.env = Environment(
            loader=FileSystemLoader(str(self.template_dir)),
            autoescape=select_autoescape(["html", "htm", "xml", "j2"]),
            trim_blocks=True,
            lstrip_blocks=True,
        )

        # Build skills string
        skills_list = profile.get("skills", [])
        self.skills_str = ", ".join(
            s for s in skills_list if s and s != "REPLACE_ME"
        )

    def render(
        self,
        name: str,
        email: str,
        company: str,
        designation: str,
    ) -> dict[str, str]:
        """Render a personalized email for one contact.

        Returns dict with keys: subject, body_html, body_text, greeting,
        template_type, first_name.
        """
        greeting = build_greeting(name, self.honorifics)
        first_name = extract_first_name(name, self.honorifics)
        template_type = select_template_type(designation)
        subject_template = select_subject(self.subjects, email)

        # Template context
        ctx: dict[str, Any] = {
            "greeting": greeting,
            "first_name": first_name or "Hiring Team",
            "company": company,
            "sender_name": self.profile.get("sender_name", ""),
            "degree": self.profile.get("degree", ""),
            "college": self.profile.get("college", ""),
            "graduation_year": self.profile.get("graduation_year", ""),
            "phone": self.profile.get("phone", ""),
            "email": self.profile.get("email", ""),
            "linkedin": self.profile.get("linkedin", ""),
            "github": self.profile.get("github", ""),
            "portfolio": self.profile.get("portfolio", ""),
            "internship_interest": self.profile.get("internship_interest", ""),
            "skills": self.skills_str,
            "highlight_project": self.profile.get("highlight_project", ""),
        }

        # Render subject
        subject = self.env.from_string(subject_template).render(**ctx)

        # Render body
        html_template = f"body_{template_type}.html.j2"
        text_template = f"body_{template_type}.txt.j2"

        body_html = self.env.get_template(html_template).render(**ctx)
        body_text = self.env.get_template(text_template).render(**ctx)

        return {
            "subject": subject,
            "body_html": body_html,
            "body_text": body_text,
            "greeting": greeting,
            "template_type": template_type,
            "first_name": first_name or "Hiring Team",
        }
