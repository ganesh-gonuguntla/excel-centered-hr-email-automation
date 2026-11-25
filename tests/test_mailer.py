"""Comprehensive unit tests for the Internship Mailer."""

from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

# ----- reader tests -----

from mailer.reader import (
    _clean_company,
    _normalize_header,
    load_contacts,
    normalize_row,
    get_score_for_row,
)


class TestNormalizeHeader:
    def test_basic(self):
        assert _normalize_header("Name") == "name"
        assert _normalize_header("  Email  ") == "email"

    def test_underscores_and_spaces(self):
        assert _normalize_header("Company_Name") == "company name"
        assert _normalize_header("  Intern  Interest  Score  ") == "intern interest score"


class TestCleanCompany:
    def test_trailing_comma(self):
        assert _clean_company("Trellissoft,") == "Trellissoft"

    def test_trailing_period(self):
        assert _clean_company("Infosys.") == "Infosys"

    def test_trailing_multiple(self):
        assert _clean_company("Acme Corp.,;") == "Acme Corp"

    def test_no_trailing(self):
        assert _clean_company("Google") == "Google"

    def test_whitespace(self):
        assert _clean_company("  Some  Company  ") == "Some Company"


class TestNormalizeRow:
    def test_basic(self):
        row = pd.Series({
            "name": "  Anita  Sharma  ",
            "email": " Anita@Example.COM ",
            "company": "Trellissoft,",
            "designation": "Talent Acquisition  Lead",
        })
        result = normalize_row(row)
        assert result["name"] == "Anita Sharma"
        assert result["email"] == "anita@example.com"
        assert result["company"] == "Trellissoft"
        assert result["designation"] == "Talent Acquisition Lead"

    def test_empty_fields(self):
        row = pd.Series({
            "name": "",
            "email": "",
            "company": "",
            "designation": "",
        })
        result = normalize_row(row)
        assert result["name"] == ""
        assert result["email"] == ""


class TestLoadContacts:
    def test_load_xlsx(self, tmp_path):
        # Create a test Excel file
        df = pd.DataFrame({
            "Name": ["Alice", "Bob"],
            "Email": ["alice@test.com", "bob@test.com"],
            "Designation": ["HR Manager", "Recruiter"],
            "Company": ["Acme", "Beta Inc"],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        result_df, col_map = load_contacts(path)
        assert len(result_df) == 2
        assert "name" in col_map

    def test_load_csv(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["Alice"],
            "Email": ["alice@test.com"],
            "Designation": ["HR"],
            "Company": ["Acme"],
        })
        path = tmp_path / "test.csv"
        df.to_csv(path, index=False)

        result_df, _ = load_contacts(path)
        assert len(result_df) == 1

    def test_start_row(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["A", "B", "C"],
            "Email": ["a@t.com", "b@t.com", "c@t.com"],
            "Designation": ["HR", "HR", "HR"],
            "Company": ["X", "Y", "Z"],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        result_df, _ = load_contacts(path, start_row=2)
        assert len(result_df) == 2

    def test_limit(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["A", "B", "C"],
            "Email": ["a@t.com", "b@t.com", "c@t.com"],
            "Designation": ["HR", "HR", "HR"],
            "Company": ["X", "Y", "Z"],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        result_df, _ = load_contacts(path, limit=1)
        assert len(result_df) == 1

    def test_min_score(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["A", "B", "C"],
            "Email": ["a@t.com", "b@t.com", "c@t.com"],
            "Designation": ["HR", "HR", "HR"],
            "Company": ["X", "Y", "Z"],
            "Intern Interest Score (1-5)": [2, 4, 5],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        result_df, _ = load_contacts(path, min_score=4)
        assert len(result_df) == 2

    def test_missing_column(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["A"],
            "Designation": ["HR"],
            "Company": ["X"],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        with pytest.raises(ValueError, match="Could not find required columns"):
            load_contacts(path)

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            load_contacts("nonexistent.xlsx")

    def test_extra_columns_preserved(self, tmp_path):
        df = pd.DataFrame({
            "Name": ["A"],
            "Email": ["a@t.com"],
            "Designation": ["HR"],
            "Company": ["X"],
            "Rank": [1],
            "Interest Level": ["High"],
        })
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")

        result_df, _ = load_contacts(path)
        assert "Rank" in result_df.columns
        assert "Interest Level" in result_df.columns


class TestGetScoreForRow:
    def test_with_score(self):
        row = pd.Series({"Intern Interest Score (1-5)": 4})
        columns = pd.Index(["Intern Interest Score (1-5)"])
        assert get_score_for_row(row, columns) == 4.0

    def test_without_score(self):
        row = pd.Series({"Name": "Alice"})
        columns = pd.Index(["Name"])
        assert get_score_for_row(row, columns) is None

    def test_nan_score(self):
        row = pd.Series({"Intern Interest Score (1-5)": float("nan")})
        columns = pd.Index(["Intern Interest Score (1-5)"])
        assert get_score_for_row(row, columns) is None


# ----- validator tests -----

from mailer.validator import check_syntax, validate_email_full


class TestCheckSyntax:
    def test_valid(self):
        is_valid, normalized = check_syntax("test@example.com")
        assert is_valid
        assert normalized == "test@example.com"

    def test_valid_with_spaces(self):
        is_valid, _ = check_syntax("  test@example.com  ")
        assert is_valid

    def test_malformed(self):
        is_valid, _ = check_syntax("not-an-email")
        assert not is_valid

    def test_empty(self):
        is_valid, _ = check_syntax("")
        assert not is_valid

    def test_none_like(self):
        is_valid, _ = check_syntax("   ")
        assert not is_valid

    def test_truncated_domain(self):
        is_valid, _ = check_syntax("x@excelenciaconsulting.c")
        # email-validator may accept or reject short TLDs
        # The important thing is it doesn't crash
        assert isinstance(is_valid, bool)

    def test_missing_at(self):
        is_valid, _ = check_syntax("testexample.com")
        assert not is_valid

    def test_double_at(self):
        is_valid, _ = check_syntax("test@@example.com")
        assert not is_valid


class TestValidateEmailFull:
    @patch("mailer.validator.check_mx", return_value=True)
    def test_valid_email(self, mock_mx):
        result = validate_email_full("test@example.com", set(), set())
        assert result == (True, "VALID")

    @patch("mailer.validator.check_mx", return_value=True)
    def test_duplicate_in_sent(self, mock_mx):
        sent = {"test@example.com"}
        result = validate_email_full("test@example.com", sent, set())
        assert result == (False, "SKIPPED_DUPLICATE")

    @patch("mailer.validator.check_mx", return_value=True)
    def test_duplicate_in_seen(self, mock_mx):
        seen = {"test@example.com"}
        result = validate_email_full("test@example.com", set(), seen)
        assert result == (False, "SKIPPED_DUPLICATE")

    @patch("mailer.validator.check_mx", return_value=False)
    def test_no_mx(self, mock_mx):
        result = validate_email_full("test@nonexistentdomain12345.com", set(), set())
        assert result == (False, "SKIPPED_NO_MX")

    def test_invalid_syntax(self):
        result = validate_email_full("not-an-email", set(), set())
        assert result == (False, "SKIPPED_INVALID_EMAIL")

    def test_empty_email(self):
        result = validate_email_full("", set(), set())
        assert result == (False, "SKIPPED_INVALID_EMAIL")


# ----- personalizer tests -----

from mailer.personalizer import (
    build_greeting,
    extract_first_name,
    select_subject,
    select_template_type,
)


class TestExtractFirstName:
    def test_normal(self):
        assert extract_first_name("Anita Sharma") == "Anita"

    def test_single_name(self):
        assert extract_first_name("Anita") == "Anita"

    def test_empty(self):
        assert extract_first_name("") is None

    def test_none_input(self):
        assert extract_first_name("") is None

    def test_placeholder_hr(self):
        assert extract_first_name("HR") is None

    def test_placeholder_friend(self):
        # "Friend Friend" - both tokens are placeholder words, falls back to None
        assert extract_first_name("Friend Friend") is None

    def test_mixed_placeholder_and_real(self):
        # First token is placeholder, second is real -> use the real one
        assert extract_first_name("HR Anita") == "Anita"

    def test_single_odd_token(self):
        assert extract_first_name("hr") is None
        assert extract_first_name("Hiring") is None
        assert extract_first_name("Team") is None
        assert extract_first_name("NA") is None

    def test_honorific_prefix(self):
        assert extract_first_name("Dr. Anita Sharma") == "Anita"
        assert extract_first_name("Capt Ravi Kumar") == "Ravi"
        assert extract_first_name("Mr. John Smith") == "John"

    def test_all_honorifics(self):
        assert extract_first_name("Dr.") is None
        assert extract_first_name("Mr.") is None

    def test_crp_prefix(self):
        assert extract_first_name("Crp John Doe") == "John"

    def test_whitespace(self):
        assert extract_first_name("  Anita  Sharma  ") == "Anita"

    def test_single_char(self):
        assert extract_first_name("A") is None


class TestBuildGreeting:
    def test_normal(self):
        assert build_greeting("Anita Sharma") == "Dear Anita,"

    def test_empty(self):
        assert build_greeting("") == "Dear Hiring Team,"

    def test_placeholder(self):
        assert build_greeting("HR") == "Dear Hiring Team,"

    def test_honorific(self):
        assert build_greeting("Dr. Anita") == "Dear Anita,"


class TestSelectTemplateType:
    def test_talent_acquisition(self):
        assert select_template_type("Talent Acquisition Manager") == "ta"

    def test_recruiter(self):
        assert select_template_type("Senior Recruiter") == "ta"

    def test_campus(self):
        assert select_template_type("Campus Relations") == "ta"

    def test_hiring(self):
        assert select_template_type("Hiring Manager") == "ta"

    def test_hr_general(self):
        assert select_template_type("HR Manager") == "hr"

    def test_people_head(self):
        assert select_template_type("Head of People") == "hr"

    def test_empty(self):
        assert select_template_type("") == "hr"

    def test_staffing(self):
        assert select_template_type("Staffing Coordinator") == "ta"


class TestSelectSubject:
    def test_deterministic(self):
        subjects = ["Sub A", "Sub B", "Sub C", "Sub D", "Sub E"]
        s1 = select_subject(subjects, "test@example.com")
        s2 = select_subject(subjects, "test@example.com")
        assert s1 == s2  # Same email -> same subject

    def test_different_emails(self):
        subjects = ["Sub A", "Sub B", "Sub C", "Sub D", "Sub E"]
        s1 = select_subject(subjects, "alice@example.com")
        s2 = select_subject(subjects, "bob@example.com")
        # They might be the same by chance, but test they don't crash
        assert s1 in subjects
        assert s2 in subjects

    def test_empty_subjects(self):
        with pytest.raises(ValueError):
            select_subject([], "test@example.com")


# ----- reporter tests -----

from mailer.reporter import (
    RunStats,
    append_to_sent_log,
    load_sent_log,
    load_failed_emails,
    get_today_send_count,
)


class TestRunStats:
    def test_record_sent(self):
        stats = RunStats()
        stats.record("a@b.com", "SENT")
        assert stats.sent == 1
        assert stats.total == 1

    def test_record_failed(self):
        stats = RunStats()
        stats.record("a@b.com", "FAILED", "SMTP 550")
        assert stats.failed == 1
        assert len(stats.failed_list) == 1

    def test_record_skipped(self):
        stats = RunStats()
        stats.record("a@b.com", "SKIPPED_INVALID_EMAIL")
        stats.record("b@c.com", "SKIPPED_NO_MX")
        stats.record("c@d.com", "SKIPPED_DUPLICATE")
        assert stats.skipped_invalid == 1
        assert stats.skipped_no_mx == 1
        assert stats.skipped_duplicate == 1
        assert stats.total_skipped == 3

    def test_success_rate(self):
        stats = RunStats()
        stats.record("a@b.com", "SENT")
        stats.record("b@c.com", "SENT")
        stats.record("c@d.com", "FAILED", "err")
        assert stats.success_rate == pytest.approx(66.666, abs=0.01)

    def test_success_rate_no_attempts(self):
        stats = RunStats()
        stats.record("a@b.com", "SKIPPED_INVALID_EMAIL")
        assert stats.success_rate == 0.0


class TestSentLog:
    def test_append_and_load(self, tmp_path):
        log_path = tmp_path / "sent_log.csv"
        append_to_sent_log(log_path, "a@b.com", "SENT", "", "Subject 1", "2026-01-01T00:00:00")
        append_to_sent_log(log_path, "c@d.com", "FAILED", "err", "Subject 2", "2026-01-01T00:01:00")

        sent = load_sent_log(log_path)
        assert "a@b.com" in sent
        assert "c@d.com" not in sent

    def test_load_failed(self, tmp_path):
        log_path = tmp_path / "sent_log.csv"
        append_to_sent_log(log_path, "a@b.com", "FAILED", "err", "S1", "2026-01-01T00:00:00")
        append_to_sent_log(log_path, "b@c.com", "SENT", "", "S2", "2026-01-01T00:01:00")

        failed = load_failed_emails(log_path)
        assert "a@b.com" in failed
        assert "b@c.com" not in failed

    def test_failed_then_sent(self, tmp_path):
        """If an email was FAILED then later SENT, it should not be in failed set."""
        log_path = tmp_path / "sent_log.csv"
        append_to_sent_log(log_path, "a@b.com", "FAILED", "err", "S1", "2026-01-01T00:00:00")
        append_to_sent_log(log_path, "a@b.com", "SENT", "", "S1", "2026-01-02T00:00:00")

        failed = load_failed_emails(log_path)
        assert "a@b.com" not in failed

    def test_today_count(self, tmp_path):
        from datetime import datetime

        log_path = tmp_path / "sent_log.csv"
        today = datetime.now().strftime("%Y-%m-%dT12:00:00")
        append_to_sent_log(log_path, "a@b.com", "SENT", "", "S1", today)
        append_to_sent_log(log_path, "b@c.com", "SENT", "", "S2", today)
        append_to_sent_log(log_path, "c@d.com", "FAILED", "err", "S3", today)

        count = get_today_send_count(log_path)
        assert count == 2

    def test_no_log_file(self, tmp_path):
        log_path = tmp_path / "nonexistent.csv"
        assert load_sent_log(log_path) == set()
        assert load_failed_emails(log_path) == set()
        assert get_today_send_count(log_path) == 0


# ----- sender tests (mocked SMTP) -----

from mailer.sender import MailSender


class TestMailSender:
    def test_build_message(self):
        sender = MailSender(
            sender_name="Test User",
            sender_email="test@gmail.com",
            reply_to="test@gmail.com",
        )
        msg = sender.build_message(
            to_email="hr@company.com",
            subject="Test Subject",
            body_text="Plain text body",
            body_html="<p>HTML body</p>",
        )
        assert msg["To"] == "hr@company.com"
        assert msg["Subject"] == "Test Subject"
        assert msg["Reply-To"] == "test@gmail.com"
        assert "Test User" in msg["From"]

    def test_build_message_with_resume(self, tmp_path):
        resume = tmp_path / "resume.pdf"
        resume.write_bytes(b"%PDF-1.4 fake pdf content")

        sender = MailSender(
            sender_name="Test",
            sender_email="test@gmail.com",
            resume_path=str(resume),
        )
        msg = sender.build_message(
            to_email="hr@co.com",
            subject="Sub",
            body_text="Text",
            body_html="<p>HTML</p>",
            attach_resume=True,
        )
        # Check that there's an attachment
        attachments = [p for p in msg.iter_attachments()]
        assert len(attachments) == 1
        assert attachments[0].get_filename() == "resume.pdf"

    def test_dry_run(self):
        sender = MailSender(sender_email="test@gmail.com")
        result = sender.send(
            to_email="hr@co.com",
            subject="Sub",
            body_text="Text",
            body_html="<p>HTML</p>",
            dry_run=True,
        )
        assert result["status"] == "DRY_RUN"
        assert result["attempts"] == 0

    @patch("mailer.sender.smtplib.SMTP")
    def test_send_success(self, mock_smtp_cls):
        mock_conn = MagicMock()
        mock_conn.noop.return_value = (250, b"OK")
        mock_smtp_cls.return_value = mock_conn

        sender = MailSender(
            smtp_user="test@gmail.com",
            smtp_password="pass",
            sender_name="Test",
            sender_email="test@gmail.com",
        )
        sender._connection = mock_conn

        result = sender.send(
            to_email="hr@co.com",
            subject="Sub",
            body_text="Text",
            body_html="<p>HTML</p>",
        )
        assert result["status"] == "SENT"
        mock_conn.send_message.assert_called_once()

    @patch("mailer.sender.smtplib.SMTP")
    def test_send_permanent_failure(self, mock_smtp_cls):
        import smtplib

        mock_conn = MagicMock()
        mock_conn.noop.return_value = (250, b"OK")
        mock_conn.send_message.side_effect = smtplib.SMTPResponseException(
            550, b"User not found"
        )
        mock_smtp_cls.return_value = mock_conn

        sender = MailSender(
            smtp_user="test@gmail.com",
            smtp_password="pass",
            sender_email="test@gmail.com",
        )
        sender._connection = mock_conn

        result = sender.send(
            to_email="hr@co.com",
            subject="Sub",
            body_text="Text",
            body_html="<p>HTML</p>",
        )
        assert result["status"] == "FAILED"
        assert "550" in result["error"]
        assert result["attempts"] == 1  # No retry for 5xx


# ----- personalizer integration test -----


class TestEmailPersonalizer:
    def test_render(self, tmp_path):
        """Test full template rendering with real template files."""
        from mailer.personalizer import EmailPersonalizer

        # Create minimal template files
        (tmp_path / "subjects.txt").write_text(
            "Internship at {{company}} | {{sender_name}}\n"
        )
        (tmp_path / "body_ta.html.j2").write_text(
            "<p>{{greeting}}</p><p>I am {{sender_name}} from {{college}}.</p>"
        )
        (tmp_path / "body_ta.txt.j2").write_text(
            "{{greeting}}\nI am {{sender_name}} from {{college}}."
        )
        (tmp_path / "body_hr.html.j2").write_text(
            "<p>{{greeting}}</p><p>I am {{sender_name}}, writing formally.</p>"
        )
        (tmp_path / "body_hr.txt.j2").write_text(
            "{{greeting}}\nI am {{sender_name}}, writing formally."
        )

        profile = {
            "sender_name": "John Doe",
            "degree": "B.Tech CS",
            "college": "MIT",
            "graduation_year": "2026",
            "phone": "+1-234-567",
            "email": "john@test.com",
            "linkedin": "https://linkedin.com/in/john",
            "github": "https://github.com/john",
            "portfolio": "",
            "internship_interest": "AI/ML",
            "skills": ["Python", "ML"],
            "highlight_project": "Built a cool thing.",
        }

        personalizer = EmailPersonalizer(tmp_path, profile)

        # Test TA template
        result = personalizer.render(
            "Anita Sharma", "anita@acme.com", "Acme Corp", "Talent Acquisition Lead"
        )
        assert result["template_type"] == "ta"
        assert "Dear Anita," in result["body_text"]
        assert "Acme Corp" in result["subject"]

        # Test HR template
        result = personalizer.render(
            "Bob Smith", "bob@beta.com", "Beta Inc", "HR Manager"
        )
        assert result["template_type"] == "hr"
        assert "Dear Bob," in result["body_text"]


# ----- Resume-from-log integration test -----


class TestResumeFromLog:
    def test_skip_already_sent(self, tmp_path):
        """Verify that emails already in sent_log as SENT are skipped as duplicates."""
        log_path = tmp_path / "sent_log.csv"
        append_to_sent_log(log_path, "alice@test.com", "SENT", "", "S1", "2026-01-01T00:00:00")

        sent_emails = load_sent_log(log_path)
        assert "alice@test.com" in sent_emails

        # Validating the same email should return SKIPPED_DUPLICATE
        with patch("mailer.validator.check_mx", return_value=True):
            is_valid, status = validate_email_full("alice@test.com", sent_emails, set())
        assert not is_valid
        assert status == "SKIPPED_DUPLICATE"

    def test_new_email_passes(self, tmp_path):
        log_path = tmp_path / "sent_log.csv"
        append_to_sent_log(log_path, "alice@test.com", "SENT", "", "S1", "2026-01-01T00:00:00")

        sent_emails = load_sent_log(log_path)
        with patch("mailer.validator.check_mx", return_value=True):
            is_valid, status = validate_email_full("bob@test.com", sent_emails, set())
        assert is_valid
        assert status == "VALID"


# ----- Duplicate detection within a single run -----


class TestDuplicateDetectionInRun:
    def test_seen_set_prevents_duplicates(self):
        seen: set[str] = set()
        with patch("mailer.validator.check_mx", return_value=True):
            # First occurrence
            is_valid, status = validate_email_full("test@example.com", set(), seen)
            assert is_valid
            seen.add("test@example.com")

            # Second occurrence in same run
            is_valid2, status2 = validate_email_full("test@example.com", set(), seen)
            assert not is_valid2
            assert status2 == "SKIPPED_DUPLICATE"
