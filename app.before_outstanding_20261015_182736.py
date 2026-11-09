from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    abort,
    send_file,
    jsonify,
    session,
    g,
    make_response,
)

import sqlite3
import secrets
import io
import qrcode
import calendar
import os
import re
import json
import time
import hashlib
import hmac
import urllib.request as urllib_request
import urllib.error as urllib_error
from urllib.parse import urlsplit

from pathlib import Path
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from PIL import Image, ImageOps, UnidentifiedImageError
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import generate_password_hash, check_password_hash


# =========================================================
# APP CONFIG
# =========================================================

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "sri_sumana_demo.db"
SECRET_FILE = BASE_DIR / ".flask_secret"


def load_secret_key():
    env_key = os.environ.get("SSC_SECRET_KEY", "").strip()

    if env_key:
        return env_key

    if SECRET_FILE.exists():
        key = SECRET_FILE.read_text(encoding="utf-8").strip()

        if key:
            return key

    key = secrets.token_hex(32)
    SECRET_FILE.write_text(key, encoding="utf-8")
    return key


HTTPS_ENABLED = (
    os.environ.get("SSC_HTTPS", "0") == "1"
)

TRUST_PROXY = (
    os.environ.get("SSC_TRUST_PROXY", "0") == "1"
)

if TRUST_PROXY:
    app.wsgi_app = ProxyFix(
        app.wsgi_app,
        x_for=1,
        x_proto=1,
        x_host=1,
    )

app.config.update(
    SECRET_KEY=load_secret_key(),
    SESSION_COOKIE_NAME=(
        "__Host-sri_sumana_session"
        if HTTPS_ENABLED
        else "sri_sumana_session"
    ),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=HTTPS_ENABLED,
    SESSION_COOKIE_PATH="/",
    SESSION_REFRESH_EACH_REQUEST=True,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
    MAX_CONTENT_LENGTH=5 * 1024 * 1024,
    PREFERRED_URL_SCHEME=(
        "https"
        if HTTPS_ENABLED
        else "http"
    ),
)

BACKUP_DIR = BASE_DIR / "backups"
BACKUP_RETENTION_COUNT = 14

STUDENT_PHOTO_DIR = BASE_DIR / "uploads" / "student_photos"
STUDENT_PHOTO_DIR.mkdir(parents=True, exist_ok=True)
STUDENT_PHOTO_MAX_SIZE = (1000, 1000)
STUDENT_PHOTO_MAX_PIXELS = 20_000_000

# =========================================================
# PUBLIC REGISTRATION SMS OTP
# =========================================================

SMS_MODE = os.environ.get(
    "SSC_SMS_MODE",
    "demo",
).strip().lower()

TEXTLK_API_URL = "https://app.text.lk/api/v3/sms/send"
TEXTLK_API_TOKEN = os.environ.get(
    "TEXTLK_API_TOKEN",
    "",
).strip()
TEXTLK_SENDER_ID = os.environ.get(
    "TEXTLK_SENDER_ID",
    "",
).strip()

SIMPAPP_API_URL = (
    "https://europe-west1-sms-gateway-api-simpapp"
    ".cloudfunctions.net/api_v2_sms_send"
)
SIMPAPP_SMS_API_KEY = os.environ.get(
    "SIMPAPP_SMS_API_KEY",
    "",
).strip()

REGISTRATION_OTP_TARGET = os.environ.get(
    "SSC_REGISTRATION_OTP_TARGET",
    "student",
).strip().lower()

REGISTRATION_OTP_TTL_SECONDS = 5 * 60
REGISTRATION_OTP_RESEND_SECONDS = 60
REGISTRATION_OTP_MAX_ATTEMPTS = 5
REGISTRATION_OTP_MAX_SENDS = 5
REGISTRATION_OTP_DRAFT_HOURS = 24


DAY_NAMES = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]


# =========================================================
# DATABASE
# =========================================================

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    conn.execute(
        "PRAGMA foreign_keys = ON"
    )

    return conn


def column_exists(
    conn,
    table_name,
    column_name,
):
    rows = conn.execute(
        f"PRAGMA table_info({table_name})"
    ).fetchall()

    return any(
        row["name"] == column_name
        for row in rows
    )


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():

    conn = get_db()

    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS student_registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            registration_ref TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            parent_name TEXT NOT NULL,
            parent_phone TEXT NOT NULL,
            school TEXT,
            grade TEXT NOT NULL,
            address TEXT,
            selected_classes TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'PENDING',
            created_at TEXT NOT NULL
        );


        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_code TEXT UNIQUE NOT NULL,
            registration_id INTEGER UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            parent_name TEXT NOT NULL,
            parent_phone TEXT NOT NULL,
            school TEXT,
            grade TEXT NOT NULL,
            address TEXT,
            status TEXT NOT NULL DEFAULT 'ACTIVE',
            created_at TEXT NOT NULL,

            FOREIGN KEY (registration_id)
            REFERENCES student_registrations(id)
            ON DELETE RESTRICT
        );


        CREATE TABLE IF NOT EXISTS student_classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            class_code TEXT NOT NULL,
            subject TEXT NOT NULL,
            teacher TEXT NOT NULL,
            created_at TEXT NOT NULL,

            UNIQUE (
                student_id,
                class_code
            ),

            FOREIGN KEY (student_id)
            REFERENCES students(id)
            ON DELETE CASCADE
        );


        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            class_code TEXT NOT NULL,
            attendance_date TEXT NOT NULL,
            check_in_time TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'PRESENT',
            source TEXT NOT NULL DEFAULT 'QR',
            created_at TEXT NOT NULL,

            UNIQUE (
                student_id,
                class_code,
                attendance_date
            ),

            FOREIGN KEY (student_id)
            REFERENCES students(id)
            ON DELETE RESTRICT
        );


        CREATE TABLE IF NOT EXISTS teachers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_code TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            subject TEXT,
            phone TEXT,
            status TEXT NOT NULL DEFAULT 'ACTIVE',
            created_at TEXT NOT NULL
        );


        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_code TEXT UNIQUE NOT NULL,
            class_name TEXT NOT NULL,
            grade TEXT NOT NULL,
            subject TEXT NOT NULL,
            teacher_id INTEGER NOT NULL,
            day_of_week INTEGER NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL,
            room TEXT,
            monthly_fee REAL NOT NULL DEFAULT 0,
            teacher_share_percent REAL NOT NULL DEFAULT 70,
            status TEXT NOT NULL DEFAULT 'ACTIVE',
            created_at TEXT NOT NULL,

            FOREIGN KEY (teacher_id)
            REFERENCES teachers(id)
            ON DELETE RESTRICT
        );


        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            class_code TEXT NOT NULL,
            fee_month TEXT NOT NULL,
            amount REAL NOT NULL,
            status TEXT NOT NULL DEFAULT 'CONFIRMED',
            payment_date TEXT NOT NULL,
            created_at TEXT NOT NULL,

            FOREIGN KEY (student_id)
            REFERENCES students(id)
            ON DELETE RESTRICT
        );
        """
    )


    migrations = {

        "student_registrations": {
            "reviewed_at": "TEXT",
            "review_note": "TEXT",
            "rejection_reason": "TEXT",
            "student_code": "TEXT",
            "public_status_token": "TEXT",
            "photo_filename": "TEXT",
        },

        "students": {
            "qr_token": "TEXT",
            "qr_created_at": "TEXT",
            "qr_status": "TEXT DEFAULT 'ACTIVE'",
            "photo_filename": "TEXT",
        },

        "student_classes": {
            "enrollment_status": "TEXT DEFAULT 'ACTIVE'",
            "ended_at": "TEXT",
        },

        "payments": {
            "receipt_no": "TEXT",
            "payment_method": "TEXT DEFAULT 'CASH'",
            "notes": "TEXT",
            "recorded_by": "TEXT DEFAULT 'Owner Demo'",
            "voided_at": "TEXT",
            "void_reason": "TEXT",
        },
    }


    for table_name, columns in migrations.items():

        for column_name, column_type in columns.items():

            if not column_exists(
                conn,
                table_name,
                column_name,
            ):

                conn.execute(
                    f"""
                    ALTER TABLE {table_name}
                    ADD COLUMN {column_name} {column_type}
                    """
                )


    # =========================================================
    # PUBLIC REGISTRATION STATUS TOKENS
    # =========================================================

    registration_rows = conn.execute(
        """
        SELECT
            id,
            public_status_token

        FROM student_registrations
        """
    ).fetchall()

    for registration_row in registration_rows:

        if registration_row["public_status_token"]:
            continue

        while True:

            public_status_token = (
                secrets.token_urlsafe(
                    32
                )
            )

            token_exists = conn.execute(
                """
                SELECT id

                FROM student_registrations

                WHERE public_status_token = ?
                """,
                (
                    public_status_token,
                ),
            ).fetchone()

            if not token_exists:
                break

        conn.execute(
            """
            UPDATE student_registrations

            SET public_status_token = ?

            WHERE id = ?
            """,
            (
                public_status_token,
                registration_row["id"],
            ),
        )

    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS
        idx_registration_public_status_token

        ON student_registrations(
            public_status_token
        )
        """
    )


    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL COLLATE NOCASE UNIQUE,
            full_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL
                CHECK (role IN ('OWNER', 'STAFF')),
            status TEXT NOT NULL DEFAULT 'ACTIVE'
                CHECK (status IN ('ACTIVE', 'INACTIVE')),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            last_login_at TEXT
        )
        """
    )


    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_app_users_role_status

        ON app_users(role, status)
        """
    )


    # =========================================================
    # TEACHER PORTAL ACCOUNTS
    # Separate from Owner/Staff accounts so Teacher access
    # can never inherit Staff/Owner permissions.
    # =========================================================

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS teacher_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL UNIQUE,
            username TEXT NOT NULL COLLATE NOCASE UNIQUE,
            password_hash TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'ACTIVE'
                CHECK (status IN ('ACTIVE', 'INACTIVE')),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            last_login_at TEXT,

            FOREIGN KEY (teacher_id)
                REFERENCES teachers(id)
                ON DELETE CASCADE
        )
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_teacher_accounts_status
        ON teacher_accounts(status)
        """
    )


    # =========================================================
    # AUDIT LOGS
    # =========================================================

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor_user_id INTEGER,
            actor_username TEXT,
            actor_full_name TEXT,
            actor_role TEXT,
            action TEXT NOT NULL,
            entity_type TEXT,
            entity_id TEXT,
            request_method TEXT,
            request_path TEXT,
            ip_address TEXT,
            user_agent TEXT,
            status_code INTEGER,
            metadata_json TEXT,
            created_at TEXT NOT NULL,

            FOREIGN KEY (actor_user_id)
            REFERENCES app_users(id)
            ON DELETE SET NULL
        )
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_audit_created_at
        ON audit_logs(created_at)
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_audit_actor
        ON audit_logs(actor_user_id)
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_audit_action
        ON audit_logs(action)
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_audit_entity
        ON audit_logs(entity_type, entity_id)
        """
    )


    # =========================================================
    # SECURITY RATE LIMITS
    # =========================================================

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS rate_limit_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scope TEXT NOT NULL,
            rate_key TEXT NOT NULL,
            created_at_epoch INTEGER NOT NULL
        )
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_rate_limit_scope_key_time
        ON rate_limit_events(
            scope,
            rate_key,
            created_at_epoch
        )
        """
    )


    # =========================================================
    # PUBLIC REGISTRATION OTP DRAFTS
    # Unverified website registrations live here temporarily.
    # They do NOT appear in Pending Approvals until OTP passes.
    # =========================================================

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS registration_otp_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            verify_token TEXT NOT NULL UNIQUE,
            otp_hash TEXT NOT NULL,
            otp_phone TEXT NOT NULL,
            full_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            parent_name TEXT NOT NULL,
            parent_phone TEXT NOT NULL,
            school TEXT,
            grade TEXT NOT NULL,
            address TEXT,
            photo_filename TEXT,
            selected_classes TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'PENDING_OTP'
                CHECK (status IN ('PENDING_OTP', 'VERIFIED', 'EXPIRED')),
            attempt_count INTEGER NOT NULL DEFAULT 0,
            send_count INTEGER NOT NULL DEFAULT 1,
            otp_expires_at TEXT NOT NULL,
            resend_available_at TEXT NOT NULL,
            registration_ref TEXT,
            public_status_token TEXT,
            created_at TEXT NOT NULL,
            verified_at TEXT
        )
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_registration_otp_status_created
        ON registration_otp_sessions(status, created_at)
        """
    )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_registration_otp_phone_created
        ON registration_otp_sessions(otp_phone, created_at)
        """
    )


    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS free_cards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            class_code TEXT NOT NULL,
            fee_month TEXT NOT NULL,
            reason TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'PENDING'
                CHECK (status IN (
                    'PENDING',
                    'APPROVED',
                    'REJECTED',
                    'REVOKED'
                )),
            waived_amount REAL NOT NULL DEFAULT 0,
            requested_by_user_id INTEGER NOT NULL,
            requested_by_name TEXT NOT NULL,
            requested_at TEXT NOT NULL,
            reviewed_by_user_id INTEGER,
            reviewed_by_name TEXT,
            reviewed_at TEXT,
            review_note TEXT,
            rejected_reason TEXT,
            revoked_by_user_id INTEGER,
            revoked_by_name TEXT,
            revoked_at TEXT,
            revoke_reason TEXT,

            FOREIGN KEY (student_id)
                REFERENCES students(id)
                ON DELETE RESTRICT,

            FOREIGN KEY (class_code)
                REFERENCES classes(class_code)
                ON DELETE RESTRICT,

            FOREIGN KEY (requested_by_user_id)
                REFERENCES app_users(id)
                ON DELETE RESTRICT,

            FOREIGN KEY (reviewed_by_user_id)
                REFERENCES app_users(id)
                ON DELETE RESTRICT,

            FOREIGN KEY (revoked_by_user_id)
                REFERENCES app_users(id)
                ON DELETE RESTRICT
        )
        """
    )


    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS
        idx_free_cards_open_unique

        ON free_cards(
            student_id,
            class_code,
            fee_month
        )

        WHERE status IN (
            'PENDING',
            'APPROVED'
        )
        """
    )


    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_free_cards_month_status

        ON free_cards(
            fee_month,
            status
        )
        """
    )


    # Teacher Free Card requests reuse the existing Owner review
    # workflow. These columns keep the original Teacher identity
    # without changing or recreating the free_cards table.
    teacher_free_card_columns = {
        "request_source": "TEXT DEFAULT 'SYSTEM'",
        "requested_by_teacher_id": "INTEGER",
        "requested_by_teacher_name": "TEXT",
    }

    for column_name, column_type in teacher_free_card_columns.items():
        if not column_exists(
            conn,
            "free_cards",
            column_name,
        ):
            conn.execute(
                f"""
                ALTER TABLE free_cards
                ADD COLUMN {column_name} {column_type}
                """
            )

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS
        idx_free_cards_teacher_request
        ON free_cards(
            requested_by_teacher_id,
            fee_month,
            status
        )
        """
    )


    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS
        idx_students_qr_token

        ON students(qr_token)

        WHERE qr_token IS NOT NULL
        """
    )


    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS
        idx_payment_receipt

        ON payments(receipt_no)

        WHERE receipt_no IS NOT NULL
        """
    )


    conn.commit()

    seed_teachers_and_classes(
        conn
    )

    conn.close()


# =========================================================
# DEFAULT DEMO DATA
# =========================================================

def seed_teachers_and_classes(
    conn
):

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


    teacher_rows = [
        (
            "T001",
            "Rusiru",
            "ICT",
            "",
        ),
        (
            "T002",
            "Thihari Miss",
            "Sinhala",
            "",
        ),
        (
            "T003",
            "Nimal Sir",
            "Mathematics",
            "",
        ),
    ]


    for row in teacher_rows:

        conn.execute(
            """
            INSERT OR IGNORE INTO teachers
            (
                teacher_code,
                full_name,
                subject,
                phone,
                status,
                created_at
            )

            VALUES
            (
                ?, ?, ?, ?,
                'ACTIVE',
                ?
            )
            """,
            (
                *row,
                now,
            ),
        )


    conn.commit()


    teacher_ids = {
        row["teacher_code"]: row["id"]

        for row in conn.execute(
            """
            SELECT
                id,
                teacher_code

            FROM teachers
            """
        ).fetchall()
    }


    class_rows = [
        (
            "C001",
            "Grade 11 ICT",
            "Grade 11",
            "ICT",
            teacher_ids["T001"],
            5,
            "15:00",
            "17:00",
            "Room 01",
            2500,
            70,
        ),
        (
            "C002",
            "Grade 11 Sinhala",
            "Grade 11",
            "Sinhala",
            teacher_ids["T002"],
            6,
            "09:00",
            "11:00",
            "Room 01",
            2000,
            70,
        ),
        (
            "C003",
            "Grade 11 Mathematics",
            "Grade 11",
            "Mathematics",
            teacher_ids["T003"],
            6,
            "14:00",
            "16:00",
            "Room 02",
            2500,
            70,
        ),
        (
            "C004",
            "Grade 10 ICT",
            "Grade 10",
            "ICT",
            teacher_ids["T001"],
            5,
            "13:00",
            "15:00",
            "Room 01",
            2200,
            70,
        ),
    ]


    for row in class_rows:

        conn.execute(
            """
            INSERT OR IGNORE INTO classes
            (
                class_code,
                class_name,
                grade,
                subject,
                teacher_id,
                day_of_week,
                start_time,
                end_time,
                room,
                monthly_fee,
                teacher_share_percent,
                status,
                created_at
            )

            VALUES
            (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?,
                'ACTIVE',
                ?
            )
            """,
            (
                *row,
                now,
            ),
        )


    conn.commit()


# =========================================================
# STUDENT PHOTO HELPERS
# =========================================================

def save_student_photo(upload):
    # Keep the backend compatible with existing staff/manual
    # registration flows. The public form can still mark the
    # photo field as required in the browser.
    if upload is None or not getattr(upload, "filename", ""):
        return None

    try:
        upload.stream.seek(0)

        with Image.open(upload.stream) as check_image:
            detected_format = (check_image.format or "").upper()

            if detected_format not in {"JPEG", "PNG", "WEBP"}:
                raise ValueError(
                    "Student photo must be a JPG, PNG, or WEBP image."
                )

            width, height = check_image.size

            if (width * height) > STUDENT_PHOTO_MAX_PIXELS:
                raise ValueError(
                    "Student photo is too large. Please use a smaller image."
                )

            check_image.verify()

        upload.stream.seek(0)

        with Image.open(upload.stream) as image:
            image = ImageOps.exif_transpose(image)
            image = image.convert("RGB")
            image.thumbnail(
                STUDENT_PHOTO_MAX_SIZE,
                Image.Resampling.LANCZOS,
            )

            filename = f"student_{secrets.token_hex(16)}.jpg"
            output_path = STUDENT_PHOTO_DIR / filename

            image.save(
                output_path,
                format="JPEG",
                quality=85,
                optimize=True,
            )

        return filename

    except ValueError:
        raise

    except (UnidentifiedImageError, OSError, Image.DecompressionBombError):
        raise ValueError(
            "The selected student photo is not a valid image."
        )


# =========================================================
# PUBLIC REGISTRATION SMS OTP HELPERS
# =========================================================

def delete_student_photo_file(filename):
    if not filename:
        return

    if not re.fullmatch(
        r"student_[0-9a-f]{32}\.jpg",
        filename,
    ):
        return

    path = STUDENT_PHOTO_DIR / filename

    try:
        if path.is_file():
            path.unlink()
    except OSError:
        pass


def normalize_sri_lanka_mobile(value):
    raw = (value or "").strip()
    digits = re.sub(r"\D", "", raw)

    if (
        len(digits) == 10
        and digits.startswith("07")
    ):
        return "94" + digits[1:]

    if (
        len(digits) == 11
        and digits.startswith("947")
    ):
        return digits

    if (
        len(digits) == 9
        and digits.startswith("7")
    ):
        return "94" + digits

    raise ValueError(
        "Please enter a valid Sri Lankan mobile number, for example 07XXXXXXXX."
    )


def masked_mobile_number(value):
    digits = re.sub(r"\D", "", value or "")

    if len(digits) >= 4:
        return "+94 7X XXX " + digits[-4:]

    return "the registered mobile number"


def registration_otp_hash(verify_token, otp_code):
    secret = str(app.config["SECRET_KEY"]).encode("utf-8")
    message = f"{verify_token}:{otp_code}".encode("utf-8")

    return hmac.new(
        secret,
        message,
        hashlib.sha256,
    ).hexdigest()


def make_registration_otp_code():
    return f"{secrets.randbelow(1_000_000):06d}"


def make_demo_registration_otp_code(verify_token):
    """
    Return a deterministic 6-digit demo OTP for this verification token.

    The demo OTP is not stored as plain text in the database. It is derived
    from the application's secret key and the private verification token.
    """
    secret = str(app.config["SECRET_KEY"]).encode("utf-8")
    message = f"demo-registration-otp:{verify_token}".encode("utf-8")

    digest = hmac.new(
        secret,
        message,
        hashlib.sha256,
    ).digest()

    number = int.from_bytes(
        digest[:8],
        "big",
    ) % 1_000_000

    return f"{number:06d}"


def send_registration_sms(recipient, message):
    # -----------------------------------------------------
    # DEMO OTP MODE
    # No external SMS is sent. The OTP is shown only on the
    # private verification page for short-term client demos.
    # -----------------------------------------------------
    if SMS_MODE == "demo":
        print(
            f"[SSC DEMO OTP] To {recipient}: {message}",
            flush=True,
        )
        return True, "demo"

    # -----------------------------------------------------
    # LOCAL / DEVELOPMENT CONSOLE MODE
    # -----------------------------------------------------
    if SMS_MODE == "console":
        print(
            f"[SSC DEV SMS] To {recipient}: {message}",
            flush=True,
        )
        return True, "console"

    # -----------------------------------------------------
    # ANDROID SIM GATEWAY - SIMPAPP
    # Temporary/demo option. The Android app must be online,
    # running, and have an active SIM that can send SMS.
    # -----------------------------------------------------
    if SMS_MODE == "simpapp":
        if not SIMPAPP_SMS_API_KEY:
            print(
                "[SSC SMS ERROR] Simpapp API key is not configured.",
                flush=True,
            )
            return False, "Simpapp API key is not configured."

        normalized_recipient = str(recipient or "").strip()

        if not normalized_recipient:
            print(
                "[SSC SMS ERROR] Simpapp recipient is empty.",
                flush=True,
            )
            return False, "SMS recipient is empty."

        # registration phone normalization returns 947XXXXXXXX.
        # Simpapp V2 expects E.164, so add the + prefix.
        if not normalized_recipient.startswith("+"):
            normalized_recipient = "+" + normalized_recipient

        payload = json.dumps(
            {
                "phoneNumber": normalized_recipient,
                "message": message,
            }
        ).encode("utf-8")

        api_request = urllib_request.Request(
            SIMPAPP_API_URL,
            data=payload,
            method="POST",
            headers={
                "Authorization": (
                    f"Bearer {SIMPAPP_SMS_API_KEY}"
                ),
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )

        try:
            with urllib_request.urlopen(
                api_request,
                timeout=15,
            ) as response:
                body = response.read().decode(
                    "utf-8",
                    errors="replace",
                )

        except urllib_error.HTTPError as exc:
            try:
                error_body = exc.read().decode(
                    "utf-8",
                    errors="replace",
                )
            except Exception:
                error_body = ""

            print(
                "[SSC SMS ERROR] Simpapp HTTP "
                f"{exc.code}: {error_body or exc.reason}",
                flush=True,
            )
            return False, (
                f"Simpapp HTTP {exc.code}."
            )

        except (
            urllib_error.URLError,
            TimeoutError,
            OSError,
        ) as exc:
            print(
                f"[SSC SMS ERROR] Simpapp request failed: {exc}",
                flush=True,
            )
            return False, "Simpapp request failed."

        try:
            result = json.loads(body)
        except json.JSONDecodeError:
            print(
                "[SSC SMS ERROR] Simpapp returned invalid JSON: "
                f"{body[:500]}",
                flush=True,
            )
            return False, "Simpapp returned an invalid response."

        if result.get("success") is not True:
            provider_error = (
                result.get("error")
                or result.get("message")
                or "Simpapp rejected the SMS."
            )

            print(
                "[SSC SMS ERROR] Simpapp rejected SMS: "
                f"{provider_error}",
                flush=True,
            )
            return False, str(provider_error)

        message_id = (
            result.get("messageId")
            or result.get("id")
            or result.get("status")
            or "queued"
        )

        print(
            "[SSC SMS] Simpapp queued OTP SMS to "
            f"{normalized_recipient} ({message_id})",
            flush=True,
        )

        return True, str(message_id)

    # -----------------------------------------------------
    # TEXT.LK
    # -----------------------------------------------------
    if SMS_MODE == "textlk":
        if not TEXTLK_API_TOKEN or not TEXTLK_SENDER_ID:
            print(
                "[SSC SMS ERROR] Text.lk credentials are not configured.",
                flush=True,
            )
            return False, "SMS credentials are not configured."

        payload = json.dumps(
            {
                "recipient": recipient,
                "sender_id": TEXTLK_SENDER_ID,
                "type": "plain",
                "message": message,
            }
        ).encode("utf-8")

        api_request = urllib_request.Request(
            TEXTLK_API_URL,
            data=payload,
            method="POST",
            headers={
                "Authorization": f"Bearer {TEXTLK_API_TOKEN}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )

        try:
            with urllib_request.urlopen(
                api_request,
                timeout=12,
            ) as response:
                body = response.read().decode(
                    "utf-8",
                    errors="replace",
                )

        except urllib_error.HTTPError as exc:
            try:
                error_body = exc.read().decode(
                    "utf-8",
                    errors="replace",
                )
            except Exception:
                error_body = ""

            print(
                "[SSC SMS ERROR] Text.lk HTTP "
                f"{exc.code}: {error_body or exc.reason}",
                flush=True,
            )
            return False, f"Text.lk HTTP {exc.code}."

        except (
            urllib_error.URLError,
            TimeoutError,
            OSError,
        ) as exc:
            print(
                f"[SSC SMS ERROR] Text.lk request failed: {exc}",
                flush=True,
            )
            return False, "SMS gateway request failed."

        try:
            result = json.loads(body)
        except json.JSONDecodeError:
            print(
                "[SSC SMS ERROR] Text.lk returned invalid JSON: "
                f"{body[:500]}",
                flush=True,
            )
            return False, "SMS gateway returned an invalid response."

        status = result.get("status")
        success = (
            status is True
            or str(status).lower() == "success"
        )

        if not success:
            provider_error = (
                result.get("message")
                or result.get("error")
                or "SMS gateway rejected the message."
            )

            print(
                "[SSC SMS ERROR] Text.lk rejected SMS: "
                f"{provider_error}",
                flush=True,
            )
            return False, str(provider_error)

        data = result.get("data") or {}
        sms_id = (
            data.get("sms_id")
            if isinstance(data, dict)
            else None
        )

        return True, sms_id or "queued"

    print(
        f"[SSC SMS ERROR] Unsupported SSC_SMS_MODE: {SMS_MODE}",
        flush=True,
    )
    return False, "SMS provider is not configured."


def cleanup_registration_otp_sessions():
    cutoff = (
        datetime.now()
        - timedelta(hours=REGISTRATION_OTP_DRAFT_HOURS)
    ).strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()

    try:
        stale_rows = conn.execute(
            """
            SELECT photo_filename
            FROM registration_otp_sessions
            WHERE status != 'VERIFIED'
            AND created_at < ?
            """,
            (cutoff,),
        ).fetchall()

        conn.execute(
            """
            DELETE FROM registration_otp_sessions
            WHERE created_at < ?
            """,
            (cutoff,),
        )

        conn.commit()

    finally:
        conn.close()

    for row in stale_rows:
        delete_student_photo_file(
            row["photo_filename"]
        )


def registration_otp_destination(phone, parent_phone):
    if REGISTRATION_OTP_TARGET == "student":
        return (
            normalize_sri_lanka_mobile(phone),
            "Student Phone",
        )

    return (
        normalize_sri_lanka_mobile(parent_phone),
        "Parent / Guardian Phone",
    )


def registration_otp_message(code):
    return (
        "Sri Sumana verification code: "
        f"{code}. "
        "Valid for 5 minutes. "
        "Do not share this code."
    )


# =========================================================
# GENERAL HELPERS
# =========================================================

def format_time(
    value
):

    try:

        return datetime.strptime(
            value,
            "%H:%M"
        ).strftime(
            "%I:%M %p"
        )

    except ValueError:

        return value


def valid_month(
    value
):

    try:

        datetime.strptime(
            value,
            "%Y-%m"
        )

        return True

    except ValueError:

        return False


def month_bounds(
    month_key
):

    year, month = map(
        int,
        month_key.split("-")
    )

    last_day = calendar.monthrange(
        year,
        month,
    )[1]

    start = (
        f"{year:04d}-"
        f"{month:02d}-01 00:00:00"
    )

    end = (
        f"{year:04d}-"
        f"{month:02d}-"
        f"{last_day:02d} 23:59:59"
    )

    return start, end


def shift_month(
    month_key,
    offset,
):

    date = datetime.strptime(
        month_key,
        "%Y-%m"
    )

    total = (
        date.year * 12
        + date.month - 1
        + offset
    )

    year = total // 12
    month = total % 12 + 1

    return (
        f"{year:04d}-"
        f"{month:02d}"
    )


def make_registration_ref(
    conn
):

    today = datetime.now().strftime(
        "%Y%m%d"
    )

    count = conn.execute(
        """
        SELECT COUNT(*)

        FROM student_registrations

        WHERE registration_ref LIKE ?
        """,
        (
            f"REG-{today}-%",
        ),
    ).fetchone()[0]

    return (
        f"REG-{today}-"
        f"{count + 1:03d}"
    )


def make_student_code(
    conn
):

    year = datetime.now().strftime(
        "%Y"
    )

    rows = conn.execute(
        """
        SELECT student_code

        FROM students

        WHERE student_code LIKE ?
        """,
        (
            f"SSC-{year}-%",
        ),
    ).fetchall()

    max_number = 0

    for row in rows:

        try:

            number = int(
                row[
                    "student_code"
                ].split("-")[-1]
            )

            max_number = max(
                max_number,
                number,
            )

        except (
            ValueError,
            IndexError,
        ):

            pass

    return (
        f"SSC-{year}-"
        f"{max_number + 1:04d}"
    )


def make_qr_token(
    conn
):

    while True:

        token = secrets.token_urlsafe(
            32
        )

        exists = conn.execute(
            """
            SELECT id

            FROM students

            WHERE qr_token = ?
            """,
            (
                token,
            ),
        ).fetchone()

        if not exists:
            return token


def make_receipt_no(
    conn
):

    today = datetime.now().strftime(
        "%Y%m%d"
    )

    rows = conn.execute(
        """
        SELECT receipt_no

        FROM payments

        WHERE receipt_no LIKE ?
        """,
        (
            f"PAY-{today}-%",
        ),
    ).fetchall()

    max_number = 0

    for row in rows:

        receipt_no = (
            row["receipt_no"]
        )

        if not receipt_no:
            continue

        try:

            number = int(
                receipt_no.split("-")[-1]
            )

            max_number = max(
                max_number,
                number,
            )

        except (
            ValueError,
            IndexError,
        ):

            pass

    return (
        f"PAY-{today}-"
        f"{max_number + 1:04d}"
    )


def make_class_code(
    conn
):

    rows = conn.execute(
        """
        SELECT class_code

        FROM classes

        WHERE class_code LIKE 'C%'
        """
    ).fetchall()

    max_number = 0

    for row in rows:

        code = (
            row["class_code"]
            or ""
        )

        try:

            number = int(
                code.replace(
                    "C",
                    ""
                )
            )

            max_number = max(
                max_number,
                number,
            )

        except ValueError:

            pass

    return (
        f"C{max_number + 1:03d}"
    )


# =========================================================
# TEACHER CODE HELPER
# =========================================================

def make_teacher_code(
    conn
):

    rows = conn.execute(
        """
        SELECT teacher_code

        FROM teachers

        WHERE teacher_code LIKE 'T%'
        """
    ).fetchall()

    max_number = 0

    for row in rows:

        code = (
            row["teacher_code"]
            or ""
        )

        try:

            number = int(
                code.replace(
                    "T",
                    ""
                )
            )

            max_number = max(
                max_number,
                number,
            )

        except ValueError:

            pass

    return (
        f"T{max_number + 1:03d}"
    )


def parse_selected_classes(
    summary
):

    classes = []

    if not summary:
        return classes

    for part in summary.split(
        " | "
    ):

        bits = part.split(":")

        if len(bits) >= 3:

            classes.append(
                {
                    "id":
                        bits[0],

                    "subject":
                        bits[1],

                    "teacher":
                        ":".join(
                            bits[2:]
                        ),
                }
            )

    return classes


# =========================================================
# CLASS HELPERS
# =========================================================

def get_student_classes(
    conn,
    student_id,
):

    rows = conn.execute(
        """
        SELECT
            student_classes.class_code,

            classes.subject,

            teachers.full_name
                AS teacher

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            student_classes.student_id = ?

        AND (
            student_classes.enrollment_status = 'ACTIVE'

            OR student_classes.enrollment_status IS NULL
        )

        AND classes.status = 'ACTIVE'

        ORDER BY
            classes.day_of_week,
            classes.start_time
        """,
        (
            student_id,
        ),
    ).fetchall()

    return [
        dict(row)
        for row in rows
    ]


def get_active_registration_classes(
    conn
):

    rows = conn.execute(
        """
        SELECT
            classes.class_code,
            classes.grade,
            classes.subject,
            classes.day_of_week,
            classes.start_time,
            classes.end_time,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            classes.status = 'ACTIVE'

        AND teachers.status = 'ACTIVE'

        ORDER BY
            classes.grade,
            classes.subject,
            classes.day_of_week,
            classes.start_time
        """
    ).fetchall()

    result = []

    for row in rows:

        schedule = (
            f"{DAY_NAMES[row['day_of_week']]} "
            f"{format_time(row['start_time'])}"
            " - "
            f"{format_time(row['end_time'])}"
        )

        result.append(
            {
                "id":
                    row["class_code"],

                "grade":
                    row["grade"],

                "subject":
                    row["subject"],

                "teacher":
                    row["teacher_name"],

                "schedule":
                    schedule,

                "fee":
                    float(
                        row["monthly_fee"]
                        or 0
                    ),
            }
        )

    return result


def count_students_at_month_end(
    conn,
    class_code,
    month_key,
):

    _, month_end = (
        month_bounds(
            month_key
        )
    )

    return conn.execute(
        """
        SELECT COUNT(
            DISTINCT student_id
        )

        FROM student_classes

        WHERE class_code = ?

        AND created_at <= ?

        AND (
            ended_at IS NULL
            OR ended_at > ?
        )
        """,
        (
            class_code,
            month_end,
            month_end,
        ),
    ).fetchone()[0]


def find_class_conflict(
    conn,
    teacher_id,
    day_of_week,
    start_time,
    end_time,
    room,
    exclude_class_code=None,
):

    params = [
        teacher_id,
        day_of_week,
        start_time,
        end_time,
    ]

    sql = """
        SELECT
            class_code,
            class_name

        FROM classes

        WHERE
            teacher_id = ?

        AND day_of_week = ?

        AND status = 'ACTIVE'

        AND NOT (
            end_time <= ?
            OR start_time >= ?
        )
    """

    if exclude_class_code:

        sql += """
            AND class_code != ?
        """

        params.append(
            exclude_class_code
        )

    conflict = conn.execute(
        sql,
        params,
    ).fetchone()

    if conflict:

        return (
            "Teacher already has another class "
            f"({conflict['class_code']} - "
            f"{conflict['class_name']}) "
            "during this time."
        )


    if room:

        room_params = [
            room,
            day_of_week,
            start_time,
            end_time,
        ]

        room_sql = """
            SELECT
                class_code,
                class_name

            FROM classes

            WHERE
                room = ?

            AND day_of_week = ?

            AND status = 'ACTIVE'

            AND NOT (
                end_time <= ?
                OR start_time >= ?
            )
        """

        if exclude_class_code:

            room_sql += """
                AND class_code != ?
            """

            room_params.append(
                exclude_class_code
            )

        room_conflict = conn.execute(
            room_sql,
            room_params,
        ).fetchone()

        if room_conflict:

            return (
                f"{room} is already used by "
                f"{room_conflict['class_code']} - "
                f"{room_conflict['class_name']} "
                "during this time."
            )

    return None


# =========================================================
# FREE CARD HELPERS + PAYMENT SUMMARY
# =========================================================

def get_approved_free_card(
    conn,
    student_id,
    class_code,
    month_key,
):
    return conn.execute(
        """
        SELECT *
        FROM free_cards
        WHERE student_id = ?
        AND class_code = ?
        AND fee_month = ?
        AND status = 'APPROVED'
        ORDER BY id DESC
        LIMIT 1
        """,
        (
            student_id,
            class_code,
            month_key,
        ),
    ).fetchone()


def get_free_card_waiver(
    conn,
    student_id,
    class_code,
    month_key,
    default_fee=0,
):
    row = get_approved_free_card(
        conn,
        student_id,
        class_code,
        month_key,
    )

    if not row:
        return 0.0

    waived = float(
        row["waived_amount"]
        or 0
    )

    if waived <= 0:
        waived = float(
            default_fee
            or 0
        )

    return max(
        waived,
        0,
    )


def get_student_class_payment_status(
    conn,
    student_id,
    class_code,
    fee_month,
):

    class_row = conn.execute(
        """
        SELECT
            class_code,
            class_name,
            monthly_fee

        FROM classes

        WHERE class_code = ?
        """,
        (
            class_code,
        ),
    ).fetchone()

    if not class_row:

        return {
            "status": "UNKNOWN",
            "fee": 0.0,
            "paid": 0.0,
            "balance": 0.0,
            "free_card": False,
            "waived": 0.0,
            "class_name": class_code,
        }

    fee = float(
        class_row["monthly_fee"]
        or 0
    )

    paid = conn.execute(
        """
        SELECT COALESCE(
            SUM(amount),
            0
        )

        FROM payments

        WHERE student_id = ?

        AND class_code = ?

        AND fee_month = ?

        AND status = 'CONFIRMED'
        """,
        (
            student_id,
            class_code,
            fee_month,
        ),
    ).fetchone()[0]

    paid = float(
        paid
        or 0
    )

    approved_free_card = (
        get_approved_free_card(
            conn,
            student_id,
            class_code,
            fee_month,
        )
    )

    if approved_free_card:

        waived = get_free_card_waiver(
            conn,
            student_id,
            class_code,
            fee_month,
            fee,
        )

        return {
            "status": "FREE CARD",
            "fee": fee,
            "paid": paid,
            "balance": 0.0,
            "free_card": True,
            "waived": float(
                waived
                or 0
            ),
            "class_name":
                class_row["class_name"],
        }

    balance = max(
        fee - paid,
        0,
    )

    if paid <= 0:

        status = "UNPAID"

    elif paid < fee:

        status = "PARTIAL"

    else:

        status = "PAID"

    return {
        "status": status,
        "fee": fee,
        "paid": paid,
        "balance": balance,
        "free_card": False,
        "waived": 0.0,
        "class_name":
            class_row["class_name"],
    }




def get_student_class_outstanding_summary(
    conn,
    student_id,
    class_code,
    through_month=None,
):
    """Return unpaid/partial fee months for one student and class.

    The calculation follows the system's existing enrollment rules:
    a month is chargeable when the enrollment existed through that
    month's end. Paid months and approved Free Card months are skipped.
    """

    through_month = (
        through_month
        or datetime.now().strftime("%Y-%m")
    )

    if not valid_month(through_month):
        through_month = datetime.now().strftime("%Y-%m")

    enrollment = conn.execute(
        """
        SELECT
            created_at,
            ended_at
        FROM student_classes
        WHERE student_id = ?
        AND class_code = ?
        LIMIT 1
        """,
        (
            student_id,
            class_code,
        ),
    ).fetchone()

    if not enrollment:
        return {
            "months": [],
            "total_outstanding": 0.0,
            "previous_outstanding": 0.0,
            "total_count": 0,
            "previous_count": 0,
            "oldest_due_month": None,
        }

    created_at = str(
        enrollment["created_at"]
        or ""
    )

    start_month = created_at[:7]

    if not valid_month(start_month):
        start_month = through_month

    if start_month > through_month:
        return {
            "months": [],
            "total_outstanding": 0.0,
            "previous_outstanding": 0.0,
            "total_count": 0,
            "previous_count": 0,
            "oldest_due_month": None,
        }

    ended_at = (
        str(enrollment["ended_at"])
        if enrollment["ended_at"]
        else None
    )

    due_months = []
    current_month = start_month

    # Hard guard protects against malformed legacy dates while still
    # allowing many years of legitimate payment history.
    for _ in range(240):
        if current_month > through_month:
            break

        _, month_end = month_bounds(
            current_month
        )

        enrollment_active_for_month = (
            created_at <= month_end
            and (
                ended_at is None
                or ended_at > month_end
            )
        )

        if enrollment_active_for_month:
            payment = get_student_class_payment_status(
                conn,
                student_id,
                class_code,
                current_month,
            )

            balance = float(
                payment.get("balance", 0)
                or 0
            )

            status = str(
                payment.get("status", "UNKNOWN")
                or "UNKNOWN"
            ).upper()

            if (
                status in {"UNPAID", "PARTIAL"}
                and balance > 0
            ):
                due_months.append(
                    {
                        "fee_month": current_month,
                        "month_label": datetime.strptime(
                            current_month,
                            "%Y-%m",
                        ).strftime("%B %Y"),
                        "status": status,
                        "fee": float(
                            payment.get("fee", 0)
                            or 0
                        ),
                        "paid": float(
                            payment.get("paid", 0)
                            or 0
                        ),
                        "balance": balance,
                        "is_current": (
                            current_month == through_month
                        ),
                    }
                )

        current_month = shift_month(
            current_month,
            1,
        )

    total_outstanding = sum(
        item["balance"]
        for item in due_months
    )

    previous_due_months = [
        item
        for item in due_months
        if not item["is_current"]
    ]

    previous_outstanding = sum(
        item["balance"]
        for item in previous_due_months
    )

    return {
        "months": due_months,
        "total_outstanding": total_outstanding,
        "previous_outstanding": previous_outstanding,
        "total_count": len(due_months),
        "previous_count": len(previous_due_months),
        "oldest_due_month": (
            due_months[0]["fee_month"]
            if due_months
            else None
        ),
    }

def payment_status_text(
    payment,
):

    status = payment.get(
        "status",
        "UNKNOWN",
    )

    if status == "FREE CARD":

        return (
            "FREE CARD | "
            "Fee Waived"
        )

    return (
        f"{status} | "
        f"Paid Rs. "
        f"{payment.get('paid', 0):,.2f} | "
        f"Balance Rs. "
        f"{payment.get('balance', 0):,.2f}"
    )


def get_class_free_card_summary(
    conn,
    class_code,
    month_key,
):
    row = conn.execute(
        """
        SELECT
            COUNT(*) AS free_card_count,
            COALESCE(
                SUM(waived_amount),
                0
            ) AS waived_amount
        FROM free_cards
        WHERE class_code = ?
        AND fee_month = ?
        AND status = 'APPROVED'
        """,
        (
            class_code,
            month_key,
        ),
    ).fetchone()

    return {
        "free_card_count": int(
            row["free_card_count"]
            or 0
        ),
        "waived_amount": float(
            row["waived_amount"]
            or 0
        ),
    }


def get_payment_summary(
    conn,
    month_key,
):

    month_start, month_end = (
        month_bounds(
            month_key
        )
    )

    enrollments = conn.execute(
        """
        SELECT
            student_classes.student_id,
            student_classes.class_code,
            classes.monthly_fee

        FROM student_classes

        JOIN classes
            ON classes.class_code =
                student_classes.class_code

        WHERE
            student_classes.created_at <= ?

        AND (
            student_classes.ended_at IS NULL
            OR student_classes.ended_at > ?
        )
        """,
        (
            month_end,
            month_end,
        ),
    ).fetchall()

    expected = 0.0
    collected = 0.0
    waived_amount = 0.0

    paid_count = 0
    partial_count = 0
    unpaid_count = 0
    free_card_count = 0

    for enrollment in enrollments:

        fee = float(
            enrollment[
                "monthly_fee"
            ]
            or 0
        )

        waiver = get_free_card_waiver(
            conn,
            enrollment[
                "student_id"
            ],
            enrollment[
                "class_code"
            ],
            month_key,
            fee,
        )

        waiver = min(
            waiver,
            fee,
        )

        payable_fee = max(
            fee - waiver,
            0,
        )

        total_paid = conn.execute(
            """
            SELECT COALESCE(
                SUM(amount),
                0
            )

            FROM payments

            WHERE
                student_id = ?

            AND class_code = ?

            AND fee_month = ?

            AND status = 'CONFIRMED'
            """,
            (
                enrollment[
                    "student_id"
                ],
                enrollment[
                    "class_code"
                ],
                month_key,
            ),
        ).fetchone()[0]

        total_paid = float(
            total_paid
            or 0
        )

        expected += payable_fee
        collected += total_paid
        waived_amount += waiver

        if waiver > 0 and payable_fee <= 0:
            free_card_count += 1
            continue

        if total_paid <= 0:
            unpaid_count += 1

        elif total_paid < payable_fee:
            partial_count += 1

        else:
            paid_count += 1

    outstanding = max(
        expected
        - collected,
        0,
    )

    rate = (
        round(
            collected
            /
            expected
            *
            100,
            1,
        )
        if expected > 0
        else 0
    )

    return {
        "expected": expected,
        "collected": collected,
        "outstanding": outstanding,
        "rate": rate,
        "paid_count": paid_count,
        "partial_count": partial_count,
        "unpaid_count": unpaid_count,
        "free_card_count": free_card_count,
        "waived_amount": waived_amount,
    }


# =========================================================
# TIMETABLE HELPERS
# =========================================================

def class_status(
    row,
    now=None,
):

    now = (
        now
        or datetime.now()
    )

    if (
        row["day_of_week"]
        != now.weekday()
    ):
        return "SCHEDULED"

    current_time = now.strftime(
        "%H:%M"
    )

    if current_time < row[
        "start_time"
    ]:
        return "UPCOMING"

    if (
        row["start_time"]
        <= current_time
        < row["end_time"]
    ):
        return "IN PROGRESS"

    return "COMPLETED"


def get_today_remaining_classes(
    conn
):

    now = datetime.now()

    rows = conn.execute(
        """
        SELECT
            classes.*,

            teachers.full_name
                AS teacher_name

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            classes.status = 'ACTIVE'

        AND classes.day_of_week = ?

        AND classes.end_time > ?

        ORDER BY
            classes.start_time
        """,
        (
            now.weekday(),

            now.strftime(
                "%H:%M"
            ),
        ),
    ).fetchall()

    result = []

    for row in rows:

        item = dict(row)

        item[
            "status_label"
        ] = class_status(
            row,
            now,
        )

        item[
            "start_display"
        ] = format_time(
            row["start_time"]
        )

        item[
            "end_display"
        ] = format_time(
            row["end_time"]
        )

        result.append(
            item
        )

    return result


def find_next_class(
    conn
):

    now = datetime.now()

    rows = conn.execute(
        """
        SELECT
            classes.*,

            teachers.full_name
                AS teacher_name

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            classes.status = 'ACTIVE'
        """
    ).fetchall()

    possible = []

    for row in rows:

        days_ahead = (
            row[
                "day_of_week"
            ]
            - now.weekday()
        ) % 7

        hour, minute = map(
            int,
            row[
                "start_time"
            ].split(":")
        )

        target = (
            now
            + timedelta(
                days=
                    days_ahead
            )
        ).replace(
            hour=hour,
            minute=minute,
            second=0,
            microsecond=0,
        )

        if target <= now:

            target += timedelta(
                days=7
            )

        possible.append(
            (
                target,
                row,
            )
        )

    if not possible:
        return None

    possible.sort(
        key=lambda item:
            item[0]
    )

    target, row = (
        possible[0]
    )

    item = dict(row)

    item[
        "next_date"
    ] = target.strftime(
        "%A, %d %B"
    )

    item[
        "start_display"
    ] = format_time(
        row["start_time"]
    )

    item[
        "end_display"
    ] = format_time(
        row["end_time"]
    )

    return item


# =========================================================
# AUDIT HELPERS
# =========================================================

SENSITIVE_AUDIT_FIELDS = {
    "password",
    "confirm_password",
    "current_password",
    "new_password",
    "confirm_new_password",
    "password_hash",
}


def sanitize_audit_data(
    data
):

    if not data:
        return {}

    clean = {}

    for key, value in data.items():

        key_text = str(key)

        if key_text.lower() in SENSITIVE_AUDIT_FIELDS:
            clean[key_text] = "[REDACTED]"
            continue

        if isinstance(
            value,
            (
                str,
                int,
                float,
                bool,
            ),
        ):
            value_text = str(value)

            if len(value_text) > 500:
                value_text = value_text[:500] + "..."

            clean[key_text] = value_text

        elif isinstance(value, list):
            clean[key_text] = [
                str(item)[:200]
                for item in value[:20]
            ]

        else:
            clean[key_text] = str(value)[:500]

    return clean


def get_request_audit_metadata():

    metadata = {}

    if request.form:
        form_data = {}

        for key in request.form.keys():
            values = request.form.getlist(key)

            if len(values) == 1:
                form_data[key] = values[0]
            else:
                form_data[key] = values

        metadata["form"] = sanitize_audit_data(
            form_data
        )

    json_data = request.get_json(
        silent=True
    )

    if isinstance(json_data, dict):
        metadata["json"] = sanitize_audit_data(
            json_data
        )

    if request.args:
        metadata["query"] = sanitize_audit_data(
            request.args.to_dict()
        )

    if request.endpoint:
        metadata["endpoint"] = request.endpoint

    return metadata


def detect_audit_entity():

    path = request.path.lower()

    if path.startswith("/approvals"):
        return "STUDENT_APPROVAL"

    if path.startswith("/students"):
        return "STUDENT"

    if path.startswith("/teachers"):
        return "TEACHER"

    if path.startswith("/classes"):
        return "CLASS"

    if path.startswith("/payments"):
        return "PAYMENT"

    if path.startswith("/free-cards"):
        return "FREE_CARD"

    if path.startswith("/api/attendance"):
        return "ATTENDANCE"

    if path.startswith("/attendance"):
        return "ATTENDANCE"

    if path.startswith("/access"):
        return "STAFF_ACCESS"

    if path.startswith("/staff"):
        return "STAFF_ACCESS"

    if path.startswith("/account"):
        return "ACCOUNT"

    if path.startswith("/logout"):
        return "AUTH"

    return "SYSTEM"


def detect_audit_entity_id():

    view_args = request.view_args or {}

    preferred_keys = [
        "student_id",
        "teacher_id",
        "registration_id",
        "payment_id",
        "free_card_id",
        "class_code",
        "user_id",
    ]

    for key in preferred_keys:
        value = view_args.get(key)

        if value is not None:
            return str(value)

    return None


def build_audit_action():

    endpoint = request.endpoint or "unknown"

    return endpoint.replace(
        "_",
        " "
    ).upper()


def write_audit_log(
    action=None,
    entity_type=None,
    entity_id=None,
    metadata=None,
    status_code=None,
    actor=None,
):

    user = actor or getattr(
        g,
        "current_user",
        None,
    )

    if user is None:
        return

    if metadata is None:
        metadata = get_request_audit_metadata()

    ip_address = get_client_ip()

    user_agent = request.headers.get(
        "User-Agent",
        "",
    )

    if len(user_agent) > 500:
        user_agent = user_agent[:500]

    conn = get_db()

    try:
        conn.execute(
            """
            INSERT INTO audit_logs
            (
                actor_user_id,
                actor_username,
                actor_full_name,
                actor_role,
                action,
                entity_type,
                entity_id,
                request_method,
                request_path,
                ip_address,
                user_agent,
                status_code,
                metadata_json,
                created_at
            )
            VALUES
            (
                ?, ?, ?, ?,
                ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?
            )
            """,
            (
                user["id"],
                user["username"],
                user["full_name"],
                user["role"],
                action or build_audit_action(),
                entity_type or detect_audit_entity(),
                entity_id or detect_audit_entity_id(),
                request.method,
                request.path,
                ip_address,
                user_agent,
                status_code,
                json.dumps(
                    metadata,
                    ensure_ascii=False,
                ),
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            ),
        )

        conn.commit()

    finally:
        conn.close()



def write_teacher_audit_log(
    action,
    entity_type=None,
    entity_id=None,
    metadata=None,
    status_code=None,
    teacher=None,
):
    """Write an audit record for a Teacher Portal action.

    Teacher accounts are intentionally stored outside app_users, so
    actor_user_id remains NULL while the Teacher identity is preserved.
    """

    actor = teacher or getattr(
        g,
        "current_teacher",
        None,
    )

    if actor is None:
        return

    if metadata is None:
        metadata = get_request_audit_metadata()

    username = actor["username"]
    full_name = actor["full_name"]
    teacher_id = actor["teacher_id"]

    ip_address = get_client_ip()
    user_agent = request.headers.get(
        "User-Agent",
        "",
    )

    if len(user_agent) > 500:
        user_agent = user_agent[:500]

    conn = get_db()

    try:
        conn.execute(
            """
            INSERT INTO audit_logs
            (
                actor_user_id,
                actor_username,
                actor_full_name,
                actor_role,
                action,
                entity_type,
                entity_id,
                request_method,
                request_path,
                ip_address,
                user_agent,
                status_code,
                metadata_json,
                created_at
            )
            VALUES
            (
                NULL, ?, ?, 'TEACHER',
                ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?, ?
            )
            """,
            (
                username,
                full_name,
                action,
                entity_type or "TEACHER",
                entity_id or str(teacher_id),
                request.method,
                request.path,
                ip_address,
                user_agent,
                status_code,
                json.dumps(
                    metadata,
                    ensure_ascii=False,
                ),
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            ),
        )

        conn.commit()

    finally:
        conn.close()


# =========================================================
# SECURITY HELPERS
# =========================================================

def get_client_ip():
    # Trust X-Forwarded-For only when the deployment explicitly enables it.
    if TRUST_PROXY:
        forwarded = request.headers.get(
            "X-Forwarded-For",
            "",
        ).strip()

        if forwarded:
            return forwarded.split(",")[0].strip()

    return request.remote_addr or "unknown"


def same_origin_url(value):
    if not value:
        return False

    try:
        candidate = urlsplit(value)

        if TRUST_PROXY:
            forwarded_proto = request.headers.get(
                "X-Forwarded-Proto",
                "",
            ).split(",")[0].strip()

            forwarded_host = request.headers.get(
                "X-Forwarded-Host",
                "",
            ).split(",")[0].strip()

            current_scheme = (
                forwarded_proto
                or request.scheme
            )

            current_host = (
                forwarded_host
                or request.host
            )

        else:
            current_scheme = request.scheme
            current_host = request.host

        candidate_scheme = (
            candidate.scheme
            or ""
        ).lower()

        candidate_host = (
            candidate.netloc
            or ""
        ).lower()

        current_scheme = (
            current_scheme
            or ""
        ).lower()

        current_host = (
            current_host
            or ""
        ).lower()

        return (
            candidate_scheme == current_scheme
            and candidate_host == current_host
        )

    except (ValueError, TypeError):
        return False


def csrf_request_is_valid():
    # SameSite=Lax is already enabled on the session cookie.
    # Unsafe browser requests must additionally come from the same origin.
    if request.method in {
        "GET",
        "HEAD",
        "OPTIONS",
        "TRACE",
    }:
        return True

    origin = request.headers.get(
        "Origin",
        "",
    ).strip()

    if origin:
        return same_origin_url(origin)

    referer = request.headers.get(
        "Referer",
        "",
    ).strip()

    if referer:
        return same_origin_url(referer)

    return False


def cleanup_rate_limit_events(conn):
    cutoff = int(time.time()) - 86400

    conn.execute(
        """
        DELETE FROM rate_limit_events
        WHERE created_at_epoch < ?
        """,
        (cutoff,),
    )


def rate_limit_count(
    scope,
    rate_key,
    window_seconds,
):
    cutoff = (
        int(time.time())
        - int(window_seconds)
    )

    conn = get_db()

    try:
        cleanup_rate_limit_events(conn)

        count = conn.execute(
            """
            SELECT COUNT(*)
            FROM rate_limit_events
            WHERE scope = ?
            AND rate_key = ?
            AND created_at_epoch >= ?
            """,
            (
                scope,
                rate_key,
                cutoff,
            ),
        ).fetchone()[0]

        conn.commit()

        return int(count)

    finally:
        conn.close()


def record_rate_limit_event(
    scope,
    rate_key,
):
    conn = get_db()

    try:
        conn.execute(
            """
            INSERT INTO rate_limit_events
            (
                scope,
                rate_key,
                created_at_epoch
            )
            VALUES (?, ?, ?)
            """,
            (
                scope,
                rate_key,
                int(time.time()),
            ),
        )

        conn.commit()

    finally:
        conn.close()


def clear_rate_limit_events(
    scope,
    rate_key,
):
    conn = get_db()

    try:
        conn.execute(
            """
            DELETE FROM rate_limit_events
            WHERE scope = ?
            AND rate_key = ?
            """,
            (
                scope,
                rate_key,
            ),
        )

        conn.commit()

    finally:
        conn.close()


def rate_limit_response(
    message,
    retry_after,
):
    if request.path.startswith("/api/"):
        response = jsonify(
            {
                "success": False,
                "message": message,
            }
        )

        response.status_code = 429
        response.headers["Retry-After"] = str(
            retry_after
        )
        return response

    flash(
        message,
        "error",
    )

    target = (
        request.referrer
        if same_origin_url(
            request.referrer or ""
        )
        else url_for("login")
    )

    return redirect(target)


def create_database_backup(
    force=False,
):
    if not DB_PATH.exists():
        return None

    BACKUP_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    now = datetime.now()

    if force:
        filename = (
            "sri_sumana_"
            + now.strftime("%Y%m%d_%H%M%S")
            + ".db"
        )
    else:
        filename = (
            "sri_sumana_"
            + now.strftime("%Y%m%d")
            + ".db"
        )

    backup_path = BACKUP_DIR / filename

    if backup_path.exists() and not force:
        return backup_path

    source = sqlite3.connect(DB_PATH)
    destination = sqlite3.connect(backup_path)

    try:
        source.backup(destination)
        destination.commit()
    finally:
        destination.close()
        source.close()

    backups = sorted(
        BACKUP_DIR.glob("sri_sumana_*.db"),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )

    for old_backup in backups[
        BACKUP_RETENTION_COUNT:
    ]:
        try:
            old_backup.unlink()
        except OSError:
            pass

    return backup_path


# =========================================================
# AUTHENTICATION + ACCESS CONTROL
# =========================================================

PUBLIC_ENDPOINTS = {
    "static",
    "public_home",
    "login",
    "logout",
    "teacher_login",
    "setup_owner",
    "public_register_student",
    "registration_otp_verify",
    "registration_otp_resend",
    "registration_success",
    "registration_status",
    "registration_status_photo",
    "registration_status_qr",
}


TEACHER_ENDPOINTS = {
    "teacher_portal_home",
    "teacher_dashboard",
    "teacher_classes",
    "teacher_finance",
    "teacher_students",
    "teacher_attendance",
    "teacher_class_profile",
    "teacher_class_attendance_save",
    "teacher_free_card_request_page",
    "teacher_free_card_history",
    "teacher_free_card_request",
    "teacher_logout",
}


STAFF_ENDPOINTS = {
    "staff_dashboard",
    "front_desk",
    "change_password",
    "register_student",
    "registration_success",
    "pending_approvals",
    "approval_detail",
    "approve_registration",
    "reject_registration",
    "students",
    "student_profile",
    "edit_student",
    "student_photo_file",
    "student_qr_pass",
    "qr_image",
    "scan_qr",
    "verify_qr",
    "mark_attendance",
    "attendance_list",
    "attendance_today",
    "class_attendance",
    "timetable",
    "payments",
    "new_payment",
    "payment_balance",
    "payment_receipt",
    "finance_overview",
    "finance_outstanding",
    "finance_history",
    "free_cards",
    "new_free_card_request",
    "free_card_detail",
}


def is_safe_next_url(value):
    return bool(
        value
        and value.startswith("/")
        and not value.startswith("//")
    )


def username_error(username):
    if len(username) < 3 or len(username) > 40:
        return "Username must be between 3 and 40 characters."

    if not re.fullmatch(r"[A-Za-z0-9._-]+", username):
        return (
            "Username can contain only letters, numbers, "
            "dot, underscore and hyphen."
        )

    return None


def password_error(password):
    if len(password) < 10:
        return "Password must contain at least 10 characters."

    if not any(char.isupper() for char in password):
        return "Password must contain an uppercase letter."

    if not any(char.islower() for char in password):
        return "Password must contain a lowercase letter."

    if not any(char.isdigit() for char in password):
        return "Password must contain a number."

    return None


def json_access_error(message, status_code):
    return jsonify(
        {
            "success": False,
            "message": message,
        }
    ), status_code


@app.before_request
def enforce_access_control():
    g.current_user = None
    g.current_teacher = None

    # -----------------------------------------------------
    # OWNER / STAFF SESSION
    # -----------------------------------------------------
    user_id = session.get("user_id")

    if user_id:
        conn = get_db()

        user = conn.execute(
            """
            SELECT
                id,
                username,
                full_name,
                role,
                status,
                created_at,
                updated_at,
                last_login_at
            FROM app_users
            WHERE id = ?
            """,
            (user_id,),
        ).fetchone()

        conn.close()

        if user and user["status"] == "ACTIVE":
            g.current_user = user
        else:
            session.pop("user_id", None)

    # -----------------------------------------------------
    # TEACHER SESSION
    # -----------------------------------------------------
    teacher_account_id = session.get(
        "teacher_account_id"
    )

    if teacher_account_id:
        conn = get_db()

        teacher = conn.execute(
            """
            SELECT
                teacher_accounts.id AS account_id,
                teacher_accounts.username,
                teacher_accounts.status AS account_status,
                teacher_accounts.created_at AS account_created_at,
                teacher_accounts.updated_at AS account_updated_at,
                teacher_accounts.last_login_at,
                teachers.id AS teacher_id,
                teachers.teacher_code,
                teachers.full_name,
                teachers.subject,
                teachers.phone,
                teachers.status AS teacher_status
            FROM teacher_accounts
            JOIN teachers
                ON teachers.id = teacher_accounts.teacher_id
            WHERE teacher_accounts.id = ?
            """,
            (teacher_account_id,),
        ).fetchone()

        conn.close()

        if (
            teacher
            and teacher["account_status"] == "ACTIVE"
            and teacher["teacher_status"] == "ACTIVE"
        ):
            g.current_teacher = teacher
        else:
            session.pop(
                "teacher_account_id",
                None,
            )

    endpoint = request.endpoint

    if endpoint is None:
        return None

    if endpoint in PUBLIC_ENDPOINTS:
        return None

    # Teacher portal pages are isolated from Owner/Staff pages.
    if endpoint in TEACHER_ENDPOINTS:
        if g.current_teacher is not None:
            return None

        next_url = request.full_path

        if next_url.endswith("?"):
            next_url = next_url[:-1]

        return redirect(
            url_for(
                "teacher_login",
                next=next_url,
            )
        )

    # A Teacher-only session must never be treated as an
    # Owner/Staff session. If an Owner/Staff page is requested,
    # send that tab to the Owner/Staff login instead of bouncing
    # it back to the Teacher Dashboard.
    if (
        g.current_teacher is not None
        and g.current_user is None
    ):
        if request.path.startswith("/api/"):
            return json_access_error(
                "Owner/Staff login required for this endpoint.",
                401,
            )

        next_url = request.full_path

        if next_url.endswith("?"):
            next_url = next_url[:-1]

        return redirect(
            url_for(
                "login",
                next=next_url,
            )
        )

    if g.current_user is None:
        if request.path.startswith("/api/"):
            return json_access_error(
                "Login required.",
                401,
            )

        next_url = request.full_path

        if next_url.endswith("?"):
            next_url = next_url[:-1]

        return redirect(
            url_for(
                "login",
                next=next_url,
            )
        )

    if g.current_user["role"] == "OWNER":
        return None

    if endpoint in STAFF_ENDPOINTS:
        return None

    if request.path.startswith("/api/"):
        return json_access_error(
            "Owner access required.",
            403,
        )

    flash(
        "Owner access is required for that page.",
        "error",
    )

    return redirect(
        url_for(
            "staff_dashboard"
        )
    )


@app.context_processor
def inject_current_user():
    return {
        "current_user": getattr(
            g,
            "current_user",
            None,
        ),
        "current_teacher": getattr(
            g,
            "current_teacher",
            None,
        ),
    }


@app.before_request
def security_csrf_guard():
    if csrf_request_is_valid():
        return None

    if request.path.startswith("/api/"):
        return jsonify(
            {
                "success": False,
                "message": (
                    "Security check failed. "
                    "Refresh the page and try again."
                ),
            }
        ), 403

    flash(
        (
            "Security check failed. "
            "Please refresh the page and try again."
        ),
        "error",
    )

    if same_origin_url(
        request.referrer or ""
    ):
        return redirect(
            request.referrer
        )

    if getattr(
        g,
        "current_teacher",
        None,
    ) is not None:
        return redirect(
            url_for("teacher_dashboard")
        )

    if getattr(
        g,
        "current_user",
        None,
    ) is not None:
        return redirect(
            url_for("dashboard")
            if g.current_user["role"] == "OWNER"
            else url_for("staff_dashboard")
        )

    return redirect(
        url_for("login")
    )


@app.before_request
def operational_rate_limits():
    if request.method != "POST":
        return None

    endpoint = request.endpoint or ""
    ip_address = get_client_ip()

    rules = {
        "public_register_student": (
            "registration_ip",
            10,
            3600,
        ),
        "registration_otp_verify": (
            "registration_otp_verify_ip",
            30,
            3600,
        ),
        "registration_otp_resend": (
            "registration_otp_resend_ip",
            10,
            3600,
        ),
        "setup_owner": (
            "setup_owner_ip",
            10,
            3600,
        ),
        "verify_qr": (
            "qr_verify_ip",
            120,
            60,
        ),
        "mark_attendance": (
            "attendance_ip",
            120,
            60,
        ),
        "teacher_class_attendance_save": (
            "teacher_manual_attendance_ip",
            60,
            60,
        ),
    }

    rule = rules.get(endpoint)

    if not rule:
        return None

    scope, limit, window_seconds = rule

    if rate_limit_count(
        scope,
        ip_address,
        window_seconds,
    ) >= limit:
        return rate_limit_response(
            "Too many requests. Please try again later.",
            window_seconds,
        )

    record_rate_limit_event(
        scope,
        ip_address,
    )

    return None


# =========================================================
# AUTOMATIC AUDIT LOGGING
# =========================================================

AUDIT_METHODS = {
    "POST",
    "PUT",
    "PATCH",
    "DELETE",
}

AUDIT_EXCLUDED_ENDPOINTS = {
    "static",
    "audit_logs",
}


@app.after_request
def automatic_audit_logger(
    response
):

    try:
        if (
            request.method in AUDIT_METHODS
            and request.endpoint not in AUDIT_EXCLUDED_ENDPOINTS
            and getattr(
                g,
                "current_user",
                None,
            ) is not None
        ):
            metadata = get_request_audit_metadata()
            metadata["response_status"] = response.status_code

            write_audit_log(
                metadata=metadata,
                status_code=response.status_code,
            )

    except Exception as error:
        # Audit logging must never break the real
        # business transaction or response.
        print(
            "AUDIT LOG ERROR:",
            error,
        )

    return response


@app.after_request
def add_security_headers(response):
    response.headers.setdefault(
        "X-Content-Type-Options",
        "nosniff",
    )

    response.headers.setdefault(
        "X-Frame-Options",
        "DENY",
    )

    response.headers.setdefault(
        "Referrer-Policy",
        "same-origin",
    )

    response.headers.setdefault(
        "Permissions-Policy",
        "camera=(self), microphone=(), geolocation=()",
    )

    if (
        getattr(
            g,
            "current_user",
            None,
        ) is not None
        and response.mimetype == "text/html"
    ):
        response.headers.setdefault(
            "Cache-Control",
            "no-store, private",
        )

    if HTTPS_ENABLED:
        response.headers.setdefault(
            "Strict-Transport-Security",
            "max-age=31536000; includeSubDomains",
        )

    return response


# =========================================================
# FIRST OWNER SETUP
# =========================================================

@app.route(
    "/setup-owner",
    methods=["GET", "POST"],
)
def setup_owner():
    conn = get_db()

    owner_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'OWNER'
        """
    ).fetchone()[0]

    if owner_count > 0:
        conn.close()

        return redirect(
            url_for("login")
        )

    if request.method == "POST":
        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        username = request.form.get(
            "username",
            "",
        ).strip().lower()

        password = request.form.get(
            "password",
            "",
        )

        confirm_password = request.form.get(
            "confirm_password",
            "",
        )

        if not full_name:
            conn.close()
            flash(
                "Owner name is required.",
                "error",
            )
            return redirect(
                url_for("setup_owner")
            )

        error = username_error(
            username
        )

        if error:
            conn.close()
            flash(error, "error")
            return redirect(
                url_for("setup_owner")
            )

        error = password_error(
            password
        )

        if error:
            conn.close()
            flash(error, "error")
            return redirect(
                url_for("setup_owner")
            )

        if password != confirm_password:
            conn.close()
            flash(
                "Passwords do not match.",
                "error",
            )
            return redirect(
                url_for("setup_owner")
            )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        try:
            conn.execute(
                """
                INSERT INTO app_users
                (
                    username,
                    full_name,
                    password_hash,
                    role,
                    status,
                    created_at,
                    updated_at
                )
                VALUES
                (
                    ?, ?, ?,
                    'OWNER',
                    'ACTIVE',
                    ?, ?
                )
                """,
                (
                    username,
                    full_name,
                    generate_password_hash(
                        password
                    ),
                    now,
                    now,
                ),
            )

            conn.commit()

        except sqlite3.IntegrityError:
            conn.close()
            flash(
                "That username is already in use.",
                "error",
            )
            return redirect(
                url_for("setup_owner")
            )

        conn.close()

        flash(
            "Owner account created. Please sign in.",
            "success",
        )

        return redirect(
            url_for("login")
        )

    conn.close()

    return render_template(
        "setup_owner.html"
    )


# =========================================================
# LOGIN / LOGOUT
# =========================================================

@app.route(
    "/login",
    methods=["GET", "POST"],
)
def login():
    conn = get_db()

    owner_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'OWNER'
        """
    ).fetchone()[0]

    if owner_count == 0:
        conn.close()
        return redirect(
            url_for("setup_owner")
        )

    if g.current_user is not None:
        conn.close()

        return redirect(
            url_for("front_desk")
        )

    if request.method == "POST":
        username = request.form.get(
            "username",
            "",
        ).strip().lower()

        password = request.form.get(
            "password",
            "",
        )

        ip_address = get_client_ip()
        username_key = username or "<blank>"

        if (
            rate_limit_count(
                "login_failure_ip",
                ip_address,
                900,
            ) >= 20
            or rate_limit_count(
                "login_failure_user",
                username_key,
                900,
            ) >= 8
        ):
            conn.close()

            return rate_limit_response(
                (
                    "Too many failed login attempts. "
                    "Please try again in about 15 minutes."
                ),
                900,
            )

        user = conn.execute(
            """
            SELECT *
            FROM app_users
            WHERE username = ?
            """,
            (username,),
        ).fetchone()

        valid = bool(
            user
            and user["status"] == "ACTIVE"
            and check_password_hash(
                user["password_hash"],
                password,
            )
        )

        if not valid:
            conn.close()

            record_rate_limit_event(
                "login_failure_ip",
                ip_address,
            )

            record_rate_limit_event(
                "login_failure_user",
                username_key,
            )

            flash(
                "Invalid username or password.",
                "error",
            )
            return render_template(
                "login.html",
                next_url=(
                    request.form.get(
                        "next",
                        "",
                    )
                ),
            )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute(
            """
            UPDATE app_users
            SET last_login_at = ?
            WHERE id = ?
            """,
            (
                now,
                user["id"],
            ),
        )

        conn.commit()
        conn.close()

        # Keep the Teacher Portal session independent.
        # Signing in as Owner/Staff must not sign out a Teacher
        # who is already signed in in another tab.
        session.permanent = True
        session["user_id"] = user["id"]

        clear_rate_limit_events(
            "login_failure_user",
            username_key,
        )

        try:
            write_audit_log(
                action="LOGIN SUCCESS",
                entity_type="AUTH",
                entity_id=str(user["id"]),
                metadata={
                    "username": user["username"],
                },
                status_code=200,
                actor=user,
            )
        except Exception as error:
            print(
                "AUDIT LOG ERROR:",
                error,
            )

        next_url = request.form.get(
            "next",
            "",
        ).strip()

        if is_safe_next_url(
            next_url
        ):
            return redirect(
                next_url
            )

        return redirect(
            url_for("front_desk")
        )

    conn.close()

    next_url = request.args.get(
        "next",
        "",
    )

    return render_template(
        "login.html",
        next_url=next_url,
    )


@app.route("/logout")
def logout():
    if g.current_user is not None:
        try:
            write_audit_log(
                action="LOGOUT",
                entity_type="AUTH",
                entity_id=str(
                    g.current_user["id"]
                ),
                metadata={},
                status_code=200,
            )
        except Exception as error:
            print(
                "AUDIT LOG ERROR:",
                error,
            )

    # Sign out only the Owner/Staff side.
    # Do not clear an active Teacher Portal session.
    session.pop("user_id", None)

    flash(
        "You have been signed out.",
        "success",
    )

    return redirect(
        url_for("login")
    )


# =========================================================
# CHANGE OWN PASSWORD
# =========================================================

@app.route(
    "/account/password",
    methods=["GET", "POST"],
)
def change_password():
    if request.method == "POST":
        current_password = request.form.get(
            "current_password",
            "",
        )

        new_password = request.form.get(
            "new_password",
            "",
        )

        confirm_password = request.form.get(
            "confirm_password",
            "",
        )

        conn = get_db()

        user = conn.execute(
            """
            SELECT *
            FROM app_users
            WHERE id = ?
            """,
            (
                g.current_user["id"],
            ),
        ).fetchone()

        if not user or not check_password_hash(
            user["password_hash"],
            current_password,
        ):
            conn.close()
            flash(
                "Current password is incorrect.",
                "error",
            )
            return redirect(
                url_for("change_password")
            )

        error = password_error(
            new_password
        )

        if error:
            conn.close()
            flash(error, "error")
            return redirect(
                url_for("change_password")
            )

        if new_password != confirm_password:
            conn.close()
            flash(
                "New passwords do not match.",
                "error",
            )
            return redirect(
                url_for("change_password")
            )

        conn.execute(
            """
            UPDATE app_users
            SET
                password_hash = ?,
                updated_at = ?
            WHERE id = ?
            """,
            (
                generate_password_hash(
                    new_password
                ),
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                user["id"],
            ),
        )

        conn.commit()
        conn.close()

        flash(
            "Password changed successfully.",
            "success",
        )

        if g.current_user["role"] == "OWNER":
            return redirect(
                url_for("dashboard")
            )

        return redirect(
            url_for("staff_dashboard")
        )

    return render_template(
        "change_password.html"
    )


# =========================================================
# TEACHER PORTAL
# =========================================================

@app.route(
    "/teacher/login",
    methods=["GET", "POST"],
)
def teacher_login():
    if g.current_teacher is not None:
        return redirect(
            url_for("teacher_dashboard")
        )

    # Owner/Staff and Teacher sessions are intentionally separate.
    # An active Owner/Staff session must not block the Teacher login page.
    if request.method == "POST":
        username = request.form.get(
            "username",
            "",
        ).strip().lower()

        password = request.form.get(
            "password",
            "",
        )

        ip_address = get_client_ip()
        username_key = username or "<blank>"

        if (
            rate_limit_count(
                "teacher_login_failure_ip",
                ip_address,
                900,
            ) >= 20
            or rate_limit_count(
                "teacher_login_failure_user",
                username_key,
                900,
            ) >= 8
        ):
            return rate_limit_response(
                (
                    "Too many failed Teacher login attempts. "
                    "Please try again in about 15 minutes."
                ),
                900,
            )

        conn = get_db()

        account = conn.execute(
            """
            SELECT
                teacher_accounts.*,
                teachers.teacher_code,
                teachers.full_name,
                teachers.subject,
                teachers.phone,
                teachers.status AS teacher_status
            FROM teacher_accounts
            JOIN teachers
                ON teachers.id = teacher_accounts.teacher_id
            WHERE teacher_accounts.username = ?
            """,
            (username,),
        ).fetchone()

        valid = bool(
            account
            and account["status"] == "ACTIVE"
            and account["teacher_status"] == "ACTIVE"
            and check_password_hash(
                account["password_hash"],
                password,
            )
        )

        if not valid:
            conn.close()

            record_rate_limit_event(
                "teacher_login_failure_ip",
                ip_address,
            )

            record_rate_limit_event(
                "teacher_login_failure_user",
                username_key,
            )

            flash(
                "Invalid Teacher username or password.",
                "error",
            )

            return render_template(
                "teacher_login.html",
                next_url=request.form.get(
                    "next",
                    "",
                ),
            )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute(
            """
            UPDATE teacher_accounts
            SET last_login_at = ?
            WHERE id = ?
            """,
            (
                now,
                account["id"],
            ),
        )

        conn.commit()
        conn.close()

        # Keep the Owner/Staff session independent.
        # Signing in as a Teacher must not sign out Owner/Staff
        # who may already be signed in in another tab.
        session.permanent = True
        session["teacher_account_id"] = (
            account["id"]
        )

        clear_rate_limit_events(
            "teacher_login_failure_user",
            username_key,
        )

        try:
            teacher_actor = {
                "username": account["username"],
                "full_name": account["full_name"],
                "teacher_id": account["teacher_id"],
            }
            write_teacher_audit_log(
                action="TEACHER LOGIN SUCCESS",
                entity_type="TEACHER",
                entity_id=str(
                    account["teacher_id"]
                ),
                metadata={
                    "teacher_code": account[
                        "teacher_code"
                    ],
                },
                status_code=200,
                teacher=teacher_actor,
            )
        except Exception as error:
            print(
                "TEACHER AUDIT LOG ERROR:",
                error,
            )

        next_url = request.form.get(
            "next",
            "",
        ).strip()

        if is_safe_next_url(next_url):
            return redirect(next_url)

        return redirect(
            url_for("teacher_dashboard")
        )

    next_url = request.args.get(
        "next",
        "",
    )

    return render_template(
        "teacher_login.html",
        next_url=next_url,
    )


@app.route("/teacher/logout")
def teacher_logout():
    if g.current_teacher is not None:
        try:
            write_teacher_audit_log(
                action="TEACHER LOGOUT",
                entity_type="TEACHER",
                entity_id=str(
                    g.current_teacher["teacher_id"]
                ),
                metadata={},
                status_code=200,
            )
        except Exception as error:
            print(
                "TEACHER AUDIT LOG ERROR:",
                error,
            )

    # Sign out only the Teacher Portal side.
    # Do not clear an active Owner/Staff session.
    session.pop("teacher_account_id", None)

    flash(
        "Teacher account signed out.",
        "success",
    )

    return redirect(
        url_for("teacher_login")
    )


def build_teacher_dashboard_data(
    conn,
    teacher_id,
    selected_month,
):
    month_start, month_end = month_bounds(
        selected_month
    )

    class_rows = conn.execute(
        """
        SELECT *
        FROM classes
        WHERE teacher_id = ?
        AND status = 'ACTIVE'
        ORDER BY day_of_week, start_time
        """,
        (teacher_id,),
    ).fetchall()

    class_summaries = []
    total_expected = 0.0
    total_collected = 0.0
    total_outstanding = 0.0
    total_teacher_income = 0.0
    total_free_cards = 0

    for class_row in class_rows:
        class_code = class_row["class_code"]
        monthly_fee = float(
            class_row["monthly_fee"] or 0
        )

        student_count = count_students_at_month_end(
            conn,
            class_code,
            selected_month,
        )

        gross_expected = (
            student_count * monthly_fee
        )

        free_card_summary = (
            get_class_free_card_summary(
                conn,
                class_code,
                selected_month,
            )
        )

        waived_amount = min(
            float(
                free_card_summary[
                    "waived_amount"
                ]
                or 0
            ),
            gross_expected,
        )

        expected = max(
            gross_expected - waived_amount,
            0,
        )

        collected = float(
            conn.execute(
                """
                SELECT COALESCE(SUM(amount), 0)
                FROM payments
                WHERE class_code = ?
                AND fee_month = ?
                AND status = 'CONFIRMED'
                """,
                (
                    class_code,
                    selected_month,
                ),
            ).fetchone()[0]
            or 0
        )

        outstanding = max(
            expected - collected,
            0,
        )

        teacher_share_percent = float(
            class_row[
                "teacher_share_percent"
            ]
            or 0
        )

        teacher_income = (
            collected
            * teacher_share_percent
            / 100
        )

        status_counts = {
            "PAID": 0,
            "PARTIAL": 0,
            "UNPAID": 0,
            "FREE CARD": 0,
        }

        enrollment_rows = conn.execute(
            """
            SELECT DISTINCT student_id
            FROM student_classes
            WHERE class_code = ?
            AND created_at <= ?
            AND (
                ended_at IS NULL
                OR ended_at > ?
            )
            """,
            (
                class_code,
                month_end,
                month_end,
            ),
        ).fetchall()

        for enrollment in enrollment_rows:
            payment = (
                get_student_class_payment_status(
                    conn,
                    enrollment["student_id"],
                    class_code,
                    selected_month,
                )
            )

            status = payment.get(
                "status",
                "UNPAID",
            )

            if status in status_counts:
                status_counts[status] += 1

        class_summaries.append(
            {
                "class_code": class_code,
                "class_name": class_row[
                    "class_name"
                ],
                "subject": class_row[
                    "subject"
                ],
                "grade": class_row["grade"],
                "day": DAY_NAMES[
                    class_row["day_of_week"]
                ],
                "start_display": format_time(
                    class_row["start_time"]
                ),
                "end_display": format_time(
                    class_row["end_time"]
                ),
                "room": class_row["room"],
                "monthly_fee": monthly_fee,
                "student_count": student_count,
                "expected": expected,
                "collected": collected,
                "outstanding": outstanding,
                "teacher_share_percent": (
                    teacher_share_percent
                ),
                "teacher_income": (
                    teacher_income
                ),
                "free_card_count": int(
                    free_card_summary[
                        "free_card_count"
                    ]
                    or 0
                ),
                "paid_count": status_counts[
                    "PAID"
                ],
                "partial_count": status_counts[
                    "PARTIAL"
                ],
                "unpaid_count": status_counts[
                    "UNPAID"
                ],
            }
        )

        total_expected += expected
        total_collected += collected
        total_outstanding += outstanding
        total_teacher_income += teacher_income
        total_free_cards += int(
            free_card_summary[
                "free_card_count"
            ]
            or 0
        )

    return {
        "classes": class_summaries,
        "total_expected": total_expected,
        "total_collected": total_collected,
        "total_outstanding": total_outstanding,
        "total_teacher_income": (
            total_teacher_income
        ),
        "total_free_cards": total_free_cards,
    }


@app.route("/teacher")
def teacher_portal_home():
    return redirect(
        url_for("teacher_dashboard")
    )


def build_teacher_portal_context(
    teacher_id,
    selected_month=None,
    search_query="",
):
    selected_month = (
        selected_month
        or datetime.now().strftime("%Y-%m")
    )

    if not valid_month(selected_month):
        selected_month = datetime.now().strftime(
            "%Y-%m"
        )

    search_query = (search_query or "").strip()

    conn = get_db()

    summary = build_teacher_dashboard_data(
        conn,
        teacher_id,
        selected_month,
    )

    today = datetime.now()
    today_date = today.strftime("%Y-%m-%d")
    today_day = today.weekday()

    today_rows = conn.execute(
        """
        SELECT *
        FROM classes
        WHERE teacher_id = ?
        AND day_of_week = ?
        AND status = 'ACTIVE'
        ORDER BY start_time
        """,
        (
            teacher_id,
            today_day,
        ),
    ).fetchall()

    today_classes = []

    for row in today_rows:
        enrolled = conn.execute(
            """
            SELECT COUNT(DISTINCT student_id)
            FROM student_classes
            WHERE class_code = ?
            AND (
                enrollment_status = 'ACTIVE'
                OR enrollment_status IS NULL
            )
            """,
            (row["class_code"],),
        ).fetchone()[0]

        attended = conn.execute(
            """
            SELECT COUNT(*)
            FROM attendance
            WHERE class_code = ?
            AND attendance_date = ?
            """,
            (
                row["class_code"],
                today_date,
            ),
        ).fetchone()[0]

        item = dict(row)
        item["start_display"] = format_time(
            row["start_time"]
        )
        item["end_display"] = format_time(
            row["end_time"]
        )
        item["enrolled_count"] = enrolled
        item["attendance_count"] = attended
        item["status_label"] = class_status(
            row,
            today,
        )
        today_classes.append(item)

    today_attendance_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM attendance
        JOIN classes
            ON classes.class_code = attendance.class_code
        WHERE classes.teacher_id = ?
        AND attendance.attendance_date = ?
        """,
        (
            teacher_id,
            today_date,
        ),
    ).fetchone()[0]

    total_students = conn.execute(
        """
        SELECT COUNT(DISTINCT student_classes.student_id)
        FROM student_classes
        JOIN students
            ON students.id = student_classes.student_id
        JOIN classes
            ON classes.class_code = student_classes.class_code
        WHERE classes.teacher_id = ?
        AND classes.status = 'ACTIVE'
        AND students.status = 'ACTIVE'
        AND (
            student_classes.enrollment_status = 'ACTIVE'
            OR student_classes.enrollment_status IS NULL
        )
        """,
        (teacher_id,),
    ).fetchone()[0]

    query_like = f"%{search_query}%"

    student_rows = conn.execute(
        """
        SELECT
            students.id AS student_id,
            students.student_code,
            students.full_name,
            students.phone,
            students.parent_phone,
            students.grade,
            students.school,
            classes.class_code,
            classes.class_name,
            classes.monthly_fee
        FROM student_classes
        JOIN students
            ON students.id = student_classes.student_id
        JOIN classes
            ON classes.class_code = student_classes.class_code
        WHERE classes.teacher_id = ?
        AND classes.status = 'ACTIVE'
        AND students.status = 'ACTIVE'
        AND (
            student_classes.enrollment_status = 'ACTIVE'
            OR student_classes.enrollment_status IS NULL
        )
        AND (
            ? = ''
            OR students.student_code LIKE ?
            OR students.full_name LIKE ?
            OR students.phone LIKE ?
            OR students.parent_phone LIKE ?
        )
        ORDER BY students.full_name, classes.class_name
        LIMIT 200
        """,
        (
            teacher_id,
            search_query,
            query_like,
            query_like,
            query_like,
            query_like,
        ),
    ).fetchall()

    student_payments = []

    for row in student_rows:
        item = dict(row)
        payment = get_student_class_payment_status(
            conn,
            row["student_id"],
            row["class_code"],
            selected_month,
        )
        item["payment"] = payment
        student_payments.append(item)

    recent_attendance = conn.execute(
        """
        SELECT
            attendance.attendance_date,
            attendance.check_in_time,
            attendance.class_code,
            students.student_code,
            students.full_name,
            classes.class_name
        FROM attendance
        JOIN students
            ON students.id = attendance.student_id
        JOIN classes
            ON classes.class_code = attendance.class_code
        WHERE classes.teacher_id = ?
        ORDER BY attendance.attendance_date DESC,
                 attendance.check_in_time DESC
        LIMIT 50
        """,
        (teacher_id,),
    ).fetchall()

    my_free_cards = conn.execute(
        """
        SELECT
            free_cards.id,
            free_cards.fee_month,
            free_cards.reason,
            free_cards.status,
            free_cards.waived_amount,
            free_cards.requested_at,
            free_cards.reviewed_at,
            free_cards.rejected_reason,
            students.student_code,
            students.full_name AS student_name,
            classes.class_code,
            classes.class_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        WHERE free_cards.request_source = 'TEACHER'
        AND free_cards.requested_by_teacher_id = ?
        ORDER BY free_cards.id DESC
        LIMIT 20
        """,
        (teacher_id,),
    ).fetchall()

    conn.close()

    return {
        "teacher": g.current_teacher,
        "selected_month": selected_month,
        "month_label": datetime.strptime(
            selected_month,
            "%Y-%m",
        ).strftime("%B %Y"),
        "previous_month": shift_month(
            selected_month,
            -1,
        ),
        "next_month": shift_month(
            selected_month,
            1,
        ),
        "today_label": today.strftime(
            "%A, %d %B %Y"
        ),
        "today_day_name": today.strftime("%A"),
        "today_classes": today_classes,
        "today_attendance_count": today_attendance_count,
        "total_students": total_students,
        "summary": summary,
        "student_payments": student_payments,
        "search_query": search_query,
        "recent_attendance": recent_attendance,
        "my_free_cards": my_free_cards,
    }


def get_teacher_page_context():
    return build_teacher_portal_context(
        g.current_teacher["teacher_id"],
        request.args.get("month"),
        request.args.get("q", ""),
    )


@app.route("/teacher/dashboard")
def teacher_dashboard():
    return render_template(
        "teacher_dashboard.html",
        active_page="overview",
        **get_teacher_page_context(),
    )


@app.route("/teacher/classes")
def teacher_classes():
    return render_template(
        "teacher_classes.html",
        active_page="classes",
        **get_teacher_page_context(),
    )


@app.route("/teacher/finance")
def teacher_finance():
    return render_template(
        "teacher_finance.html",
        active_page="finance",
        **get_teacher_page_context(),
    )


@app.route("/teacher/students")
def teacher_students():
    return render_template(
        "teacher_students.html",
        active_page="students",
        **get_teacher_page_context(),
    )


@app.route("/teacher/attendance")
def teacher_attendance():
    return render_template(
        "teacher_attendance.html",
        active_page="attendance",
        **get_teacher_page_context(),
    )



def get_teacher_owned_class(
    conn,
    teacher_id,
    class_code,
):
    return conn.execute(
        """
        SELECT
            classes.*,
            teachers.full_name AS teacher_name
        FROM classes
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE classes.class_code = ?
        AND classes.teacher_id = ?
        AND classes.status = 'ACTIVE'
        """,
        (
            class_code,
            teacher_id,
        ),
    ).fetchone()


def build_teacher_class_profile_context(
    teacher_id,
    class_code,
):
    base_context = build_teacher_portal_context(
        teacher_id,
        request.args.get("month"),
        request.args.get("q", ""),
    )

    allowed_tabs = {
        "overview",
        "students",
        "attendance",
        "payments",
    }

    active_tab = request.args.get(
        "tab",
        "overview",
    ).strip().lower()

    if active_tab not in allowed_tabs:
        active_tab = "overview"

    selected_date = request.args.get(
        "date",
        "",
    ).strip()

    today = datetime.now()
    today_date = today.strftime("%Y-%m-%d")

    if not selected_date:
        selected_date = today_date

    try:
        selected_date_object = datetime.strptime(
            selected_date,
            "%Y-%m-%d",
        )
    except ValueError:
        selected_date = today_date
        selected_date_object = datetime.strptime(
            selected_date,
            "%Y-%m-%d",
        )

    selected_month = base_context[
        "selected_month"
    ]

    conn = get_db()

    class_row = get_teacher_owned_class(
        conn,
        teacher_id,
        class_code,
    )

    if class_row is None:
        conn.close()
        abort(404)

    student_rows = conn.execute(
        """
        SELECT
            students.id AS student_id,
            students.student_code,
            students.full_name,
            students.phone,
            students.parent_phone,
            students.grade,
            students.school,
            students.status,
            student_classes.created_at AS enrolled_at
        FROM student_classes
        JOIN students
            ON students.id = student_classes.student_id
        WHERE student_classes.class_code = ?
        AND students.status = 'ACTIVE'
        AND (
            student_classes.enrollment_status = 'ACTIVE'
            OR student_classes.enrollment_status IS NULL
        )
        ORDER BY students.full_name
        """,
        (class_code,),
    ).fetchall()

    attendance_rows = conn.execute(
        """
        SELECT
            student_id,
            check_in_time,
            source
        FROM attendance
        WHERE class_code = ?
        AND attendance_date = ?
        AND status = 'PRESENT'
        """,
        (
            class_code,
            selected_date,
        ),
    ).fetchall()

    attendance_map = {
        row["student_id"]: {
            "check_in_time": row["check_in_time"],
            "source": row["source"],
        }
        for row in attendance_rows
    }

    students = []
    present_students = []
    absent_students = []
    payment_status_counts = {
        "PAID": 0,
        "PARTIAL": 0,
        "UNPAID": 0,
        "FREE CARD": 0,
    }

    for row in student_rows:
        item = dict(row)

        attendance_info = attendance_map.get(
            row["student_id"]
        )

        item["is_present"] = bool(
            attendance_info
        )

        item["attendance_time"] = (
            attendance_info["check_in_time"]
            if attendance_info
            else ""
        )

        item["attendance_source"] = (
            attendance_info["source"]
            if attendance_info
            else ""
        )

        payment = get_student_class_payment_status(
            conn,
            row["student_id"],
            class_code,
            selected_month,
        )

        item["payment"] = payment

        payment_status = payment.get(
            "status",
            "UNPAID",
        )

        if payment_status in payment_status_counts:
            payment_status_counts[
                payment_status
            ] += 1

        students.append(item)

        if item["is_present"]:
            present_students.append(item)
        else:
            absent_students.append(item)

    total_students = len(students)
    present_count = len(present_students)
    absent_count = len(absent_students)

    attendance_rate = (
        round(
            present_count
            / total_students
            * 100,
            1,
        )
        if total_students
        else 0
    )

    attendance_history = conn.execute(
        """
        SELECT
            attendance_date,
            COUNT(DISTINCT student_id) AS present_count
        FROM attendance
        WHERE class_code = ?
        AND status = 'PRESENT'
        GROUP BY attendance_date
        ORDER BY attendance_date DESC
        LIMIT 12
        """,
        (class_code,),
    ).fetchall()

    class_item = dict(class_row)
    class_item["day"] = DAY_NAMES[
        class_row["day_of_week"]
    ]
    class_item["start_display"] = format_time(
        class_row["start_time"]
    )
    class_item["end_display"] = format_time(
        class_row["end_time"]
    )

    class_finance = None

    for summary_item in base_context[
        "summary"
    ]["classes"]:
        if summary_item["class_code"] == class_code:
            class_finance = summary_item
            break

    if class_finance is None:
        class_finance = {
            "student_count": total_students,
            "expected": 0.0,
            "collected": 0.0,
            "outstanding": 0.0,
            "teacher_income": 0.0,
            "teacher_share_percent": float(
                class_row[
                    "teacher_share_percent"
                ]
                or 0
            ),
            "free_card_count": 0,
            "paid_count": payment_status_counts[
                "PAID"
            ],
            "partial_count": payment_status_counts[
                "PARTIAL"
            ],
            "unpaid_count": payment_status_counts[
                "UNPAID"
            ],
        }

    conn.close()

    base_context.update(
        {
            "active_page": "classes",
            "class_item": class_item,
            "active_tab": active_tab,
            "selected_date": selected_date,
            "today_date": today_date,
            "selected_date_label": (
                selected_date_object.strftime(
                    "%A, %d %B %Y"
                )
            ),
            "selected_date_is_today": (
                selected_date == today_date
            ),
            "selected_date_is_future": (
                selected_date > today_date
            ),
            "selected_date_is_scheduled": (
                selected_date_object.weekday()
                == class_row["day_of_week"]
            ),
            "students": students,
            "present_students": present_students,
            "absent_students": absent_students,
            "present_count": present_count,
            "absent_count": absent_count,
            "attendance_rate": attendance_rate,
            "attendance_history": attendance_history,
            "class_finance": class_finance,
            "payment_status_counts": payment_status_counts,
        }
    )

    return base_context


@app.route(
    "/teacher/classes/<class_code>"
)
def teacher_class_profile(
    class_code,
):
    return render_template(
        "teacher_class_profile.html",
        **build_teacher_class_profile_context(
            g.current_teacher["teacher_id"],
            class_code,
        ),
    )


@app.post(
    "/teacher/classes/<class_code>/attendance"
)
def teacher_class_attendance_save(
    class_code,
):
    teacher_id = g.current_teacher[
        "teacher_id"
    ]

    selected_date = request.form.get(
        "attendance_date",
        "",
    ).strip()

    selected_month = request.form.get(
        "month",
        "",
    ).strip()

    today_date = datetime.now().strftime(
        "%Y-%m-%d"
    )

    try:
        datetime.strptime(
            selected_date,
            "%Y-%m-%d",
        )
    except ValueError:
        flash(
            "Please select a valid attendance date.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_class_profile",
                class_code=class_code,
                tab="attendance",
            )
        )

    if selected_date > today_date:
        flash(
            "Attendance cannot be marked for a future date.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_class_profile",
                class_code=class_code,
                tab="attendance",
                date=selected_date,
                month=selected_month,
            )
        )

    conn = get_db()

    class_row = get_teacher_owned_class(
        conn,
        teacher_id,
        class_code,
    )

    if class_row is None:
        conn.close()
        abort(404)

    student_rows = conn.execute(
        """
        SELECT DISTINCT student_id
        FROM student_classes
        JOIN students
            ON students.id = student_classes.student_id
        WHERE student_classes.class_code = ?
        AND students.status = 'ACTIVE'
        AND (
            student_classes.enrollment_status = 'ACTIVE'
            OR student_classes.enrollment_status IS NULL
        )
        """,
        (class_code,),
    ).fetchall()

    allowed_student_ids = {
        int(row["student_id"])
        for row in student_rows
    }

    submitted_student_ids = set()

    for raw_value in request.form.getlist(
        "present_student_ids"
    ):
        try:
            student_id = int(raw_value)
        except (TypeError, ValueError):
            continue

        if student_id in allowed_student_ids:
            submitted_student_ids.add(
                student_id
            )

    existing_rows = conn.execute(
        """
        SELECT
            id,
            student_id,
            source
        FROM attendance
        WHERE class_code = ?
        AND attendance_date = ?
        AND status = 'PRESENT'
        """,
        (
            class_code,
            selected_date,
        ),
    ).fetchall()

    existing_by_student = {
        int(row["student_id"]): row
        for row in existing_rows
    }

    now = datetime.now()
    now_time = now.strftime("%H:%M:%S")

    class_start_time = str(
        class_row["start_time"]
        or "00:00"
    )

    if len(class_start_time) == 5:
        class_start_time += ":00"

    manual_check_in_time = (
        now_time
        if selected_date == today_date
        else class_start_time
    )

    created_at = now.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    added_count = 0
    removed_count = 0

    try:
        for student_id in sorted(
            allowed_student_ids
        ):
            existing = existing_by_student.get(
                student_id
            )

            if student_id in submitted_student_ids:
                if existing is None:
                    conn.execute(
                        """
                        INSERT INTO attendance
                        (
                            student_id,
                            class_code,
                            attendance_date,
                            check_in_time,
                            status,
                            source,
                            created_at
                        )
                        VALUES
                        (
                            ?, ?, ?, ?,
                            'PRESENT',
                            'TEACHER_MANUAL',
                            ?
                        )
                        """,
                        (
                            student_id,
                            class_code,
                            selected_date,
                            manual_check_in_time,
                            created_at,
                        ),
                    )
                    added_count += 1
                else:
                    conn.execute(
                        """
                        UPDATE attendance
                        SET status = 'PRESENT'
                        WHERE id = ?
                        """,
                        (existing["id"],),
                    )
            elif existing is not None:
                conn.execute(
                    """
                    DELETE FROM attendance
                    WHERE id = ?
                    """,
                    (existing["id"],),
                )
                removed_count += 1

        conn.commit()

    except Exception:
        conn.rollback()
        conn.close()
        raise

    conn.close()

    write_teacher_audit_log(
        action="TEACHER_CLASS_ATTENDANCE_SAVE",
        entity_type="CLASS",
        entity_id=class_code,
        metadata={
            "class_code": class_code,
            "attendance_date": selected_date,
            "present_count": len(
                submitted_student_ids
            ),
            "total_students": len(
                allowed_student_ids
            ),
            "added_count": added_count,
            "removed_count": removed_count,
        },
        status_code=302,
    )

    flash(
        (
            f"Attendance saved for {selected_date}. "
            f"Present: {len(submitted_student_ids)} · "
            f"Absent: "
            f"{max(len(allowed_student_ids) - len(submitted_student_ids), 0)}."
        ),
        "success",
    )

    return redirect(
        url_for(
            "teacher_class_profile",
            class_code=class_code,
            tab="attendance",
            date=selected_date,
            month=(
                selected_month
                if valid_month(selected_month)
                else None
            ),
        )
    )


@app.route("/teacher/free-card/request")
def teacher_free_card_request_page():
    return render_template(
        "teacher_free_card_request.html",
        active_page="free-card-request",
        **get_teacher_page_context(),
    )


@app.route("/teacher/free-card/history")
def teacher_free_card_history():
    return render_template(
        "teacher_free_card_history.html",
        active_page="free-card-history",
        **get_teacher_page_context(),
    )


@app.post("/teacher/free-card-request")
def teacher_free_card_request():
    teacher_id = g.current_teacher[
        "teacher_id"
    ]

    student_lookup = request.form.get(
        "student_lookup",
        "",
    ).strip()

    class_code = request.form.get(
        "class_code",
        "",
    ).strip()

    fee_month = request.form.get(
        "fee_month",
        "",
    ).strip()

    reason = request.form.get(
        "reason",
        "",
    ).strip()

    redirect_month = (
        fee_month
        if valid_month(fee_month)
        else datetime.now().strftime("%Y-%m")
    )

    if (
        not student_lookup
        or not class_code
        or not valid_month(fee_month)
        or len(reason) < 3
    ):
        flash(
            "Student number, class, month and reason are required.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    if len(reason) > 500:
        flash(
            "Reason must be 500 characters or less.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    conn = get_db()

    class_row = conn.execute(
        """
        SELECT *
        FROM classes
        WHERE class_code = ?
        AND teacher_id = ?
        AND status = 'ACTIVE'
        """,
        (
            class_code,
            teacher_id,
        ),
    ).fetchone()

    if not class_row:
        conn.close()
        flash(
            "You can request a Free Card only for your own active class.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    student_matches = conn.execute(
        """
        SELECT *
        FROM students
        WHERE status = 'ACTIVE'
        AND (
            student_code = ?
            OR phone = ?
            OR parent_phone = ?
        )
        ORDER BY id
        LIMIT 2
        """,
        (
            student_lookup,
            student_lookup,
            student_lookup,
        ),
    ).fetchall()

    if not student_matches:
        conn.close()
        flash(
            "Student not found. Use the exact Student ID or phone number.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    if len(student_matches) > 1:
        conn.close()
        flash(
            "More than one student uses that phone number. Please use the Student ID.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    student = student_matches[0]
    month_start, month_end = month_bounds(
        fee_month
    )

    enrollment = conn.execute(
        """
        SELECT id
        FROM student_classes
        WHERE student_id = ?
        AND class_code = ?
        AND created_at <= ?
        AND (
            ended_at IS NULL
            OR ended_at > ?
        )
        """,
        (
            student["id"],
            class_code,
            month_end,
            month_start,
        ),
    ).fetchone()

    if not enrollment:
        conn.close()
        flash(
            "That student is not enrolled in your selected class for this month.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    paid = float(
        conn.execute(
            """
            SELECT COALESCE(SUM(amount), 0)
            FROM payments
            WHERE student_id = ?
            AND class_code = ?
            AND fee_month = ?
            AND status = 'CONFIRMED'
            """,
            (
                student["id"],
                class_code,
                fee_month,
            ),
        ).fetchone()[0]
        or 0
    )

    if paid > 0:
        conn.close()
        flash(
            (
                "A confirmed payment already exists for this "
                "student/class/month. The Owner must resolve it "
                "before a full Free Card can be requested."
            ),
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    duplicate = conn.execute(
        """
        SELECT id, status
        FROM free_cards
        WHERE student_id = ?
        AND class_code = ?
        AND fee_month = ?
        AND status IN ('PENDING', 'APPROVED')
        LIMIT 1
        """,
        (
            student["id"],
            class_code,
            fee_month,
        ),
    ).fetchone()

    if duplicate:
        conn.close()
        flash(
            "A Pending or Approved Free Card already exists for this student, class and month.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    owner = conn.execute(
        """
        SELECT id
        FROM app_users
        WHERE role = 'OWNER'
        AND status = 'ACTIVE'
        ORDER BY id
        LIMIT 1
        """
    ).fetchone()

    if not owner:
        conn.close()
        flash(
            "Owner account is unavailable. Please contact the center administrator.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    try:
        cursor = conn.execute(
            """
            INSERT INTO free_cards
            (
                student_id,
                class_code,
                fee_month,
                reason,
                status,
                waived_amount,
                requested_by_user_id,
                requested_by_name,
                requested_at,
                request_source,
                requested_by_teacher_id,
                requested_by_teacher_name
            )
            VALUES
            (
                ?, ?, ?, ?,
                'PENDING',
                0,
                ?, ?, ?,
                'TEACHER',
                ?, ?
            )
            """,
            (
                student["id"],
                class_code,
                fee_month,
                reason,
                owner["id"],
                (
                    g.current_teacher[
                        "full_name"
                    ]
                    + " (Teacher)"
                ),
                now,
                teacher_id,
                g.current_teacher[
                    "full_name"
                ],
            ),
        )

        free_card_id = cursor.lastrowid
        conn.commit()

    except sqlite3.IntegrityError:
        conn.rollback()
        conn.close()
        flash(
            "A Free Card request already exists for this period.",
            "error",
        )
        return redirect(
            url_for(
                "teacher_free_card_request_page",
                month=redirect_month,
            )
        )

    conn.close()

    try:
        write_teacher_audit_log(
            action="TEACHER FREE CARD REQUEST",
            entity_type="FREE_CARD",
            entity_id=str(free_card_id),
            metadata={
                "student_id": student["id"],
                "student_code": student[
                    "student_code"
                ],
                "class_code": class_code,
                "fee_month": fee_month,
                "reason": reason,
            },
            status_code=200,
        )
    except Exception as error:
        print(
            "TEACHER AUDIT LOG ERROR:",
            error,
        )

    flash(
        "Free Card request submitted. Waiting for Owner approval.",
        "success",
    )

    return redirect(
        url_for(
            "teacher_free_card_history",
            month=redirect_month,
        )
    )


# =========================================================
# OWNER: TEACHER PORTAL ACCOUNT
# =========================================================

@app.route(
    "/teachers/<int:teacher_id>/portal-account",
    methods=["GET", "POST"],
)
def teacher_portal_account(
    teacher_id,
):
    conn = get_db()

    teacher = conn.execute(
        """
        SELECT *
        FROM teachers
        WHERE id = ?
        """,
        (teacher_id,),
    ).fetchone()

    if not teacher:
        conn.close()
        abort(404)

    account = conn.execute(
        """
        SELECT *
        FROM teacher_accounts
        WHERE teacher_id = ?
        """,
        (teacher_id,),
    ).fetchone()

    if request.method == "POST":
        action = request.form.get(
            "action",
            "save",
        ).strip().lower()

        if action == "toggle" and account:
            new_status = (
                "INACTIVE"
                if account["status"] == "ACTIVE"
                else "ACTIVE"
            )

            conn.execute(
                """
                UPDATE teacher_accounts
                SET status = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    new_status,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    account["id"],
                ),
            )
            conn.commit()
            conn.close()

            write_audit_log(
                action=(
                    "TEACHER PORTAL ACCOUNT "
                    + new_status
                ),
                entity_type="TEACHER",
                entity_id=str(teacher_id),
                metadata={
                    "teacher_code": teacher[
                        "teacher_code"
                    ],
                },
                status_code=200,
            )

            flash(
                f"Teacher Portal account set to {new_status}.",
                "success",
            )

            return redirect(
                url_for(
                    "teacher_portal_account",
                    teacher_id=teacher_id,
                )
            )

        username = request.form.get(
            "username",
            "",
        ).strip().lower()

        password = request.form.get(
            "password",
            "",
        )

        error = username_error(username)

        if error:
            conn.close()
            flash(error, "error")
            return redirect(
                url_for(
                    "teacher_portal_account",
                    teacher_id=teacher_id,
                )
            )

        if not account and not password:
            conn.close()
            flash(
                "Password is required when creating a Teacher account.",
                "error",
            )
            return redirect(
                url_for(
                    "teacher_portal_account",
                    teacher_id=teacher_id,
                )
            )

        if password:
            error = password_error(password)

            if error:
                conn.close()
                flash(error, "error")
                return redirect(
                    url_for(
                        "teacher_portal_account",
                        teacher_id=teacher_id,
                    )
                )

        duplicate = conn.execute(
            """
            SELECT id
            FROM teacher_accounts
            WHERE username = ?
            AND teacher_id != ?
            """,
            (
                username,
                teacher_id,
            ),
        ).fetchone()

        if duplicate:
            conn.close()
            flash(
                "That Teacher username is already in use.",
                "error",
            )
            return redirect(
                url_for(
                    "teacher_portal_account",
                    teacher_id=teacher_id,
                )
            )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        if account:
            if password:
                conn.execute(
                    """
                    UPDATE teacher_accounts
                    SET
                        username = ?,
                        password_hash = ?,
                        updated_at = ?
                    WHERE teacher_id = ?
                    """,
                    (
                        username,
                        generate_password_hash(
                            password
                        ),
                        now,
                        teacher_id,
                    ),
                )
            else:
                conn.execute(
                    """
                    UPDATE teacher_accounts
                    SET
                        username = ?,
                        updated_at = ?
                    WHERE teacher_id = ?
                    """,
                    (
                        username,
                        now,
                        teacher_id,
                    ),
                )
        else:
            conn.execute(
                """
                INSERT INTO teacher_accounts
                (
                    teacher_id,
                    username,
                    password_hash,
                    status,
                    created_at,
                    updated_at
                )
                VALUES
                (
                    ?, ?, ?,
                    'ACTIVE',
                    ?, ?
                )
                """,
                (
                    teacher_id,
                    username,
                    generate_password_hash(
                        password
                    ),
                    now,
                    now,
                ),
            )

        conn.commit()
        conn.close()

        write_audit_log(
            action=(
                "UPDATE TEACHER PORTAL ACCOUNT"
                if account
                else "CREATE TEACHER PORTAL ACCOUNT"
            ),
            entity_type="TEACHER",
            entity_id=str(teacher_id),
            metadata={
                "teacher_code": teacher[
                    "teacher_code"
                ],
                "username": username,
            },
            status_code=200,
        )

        flash(
            "Teacher Portal account saved successfully.",
            "success",
        )

        return redirect(
            url_for(
                "teacher_portal_account",
                teacher_id=teacher_id,
            )
        )

    conn.close()

    return render_template(
        "teacher_account.html",
        teacher=teacher,
        account=account,
    )


# =========================================================
# STAFF DASHBOARD
# =========================================================

@app.route("/staff")
def staff_dashboard():
    conn = get_db()

    pending_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM student_registrations
        WHERE status = 'PENDING'
        """
    ).fetchone()[0]

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    attendance_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM attendance
        WHERE attendance_date = ?
        """,
        (today,),
    ).fetchone()[0]

    month_key = datetime.now().strftime(
        "%Y-%m"
    )

    collection = conn.execute(
        """
        SELECT COALESCE(SUM(amount), 0)
        FROM payments
        WHERE fee_month = ?
        AND status = 'CONFIRMED'
        """,
        (month_key,),
    ).fetchone()[0]

    active_students = conn.execute(
        """
        SELECT COUNT(*)
        FROM students
        WHERE status = 'ACTIVE'
        """
    ).fetchone()[0]

    my_pending_free_cards = conn.execute(
        """
        SELECT COUNT(*)
        FROM free_cards
        WHERE requested_by_user_id = ?
        AND status = 'PENDING'
        """,
        (
            g.current_user["id"],
        ),
    ).fetchone()[0]

    conn.close()

    return render_template(
        "staff_dashboard.html",
        pending_count=pending_count,
        attendance_count=attendance_count,
        collection=float(
            collection or 0
        ),
        active_students=active_students,
        my_pending_free_cards=my_pending_free_cards,
    )


# =========================================================
# ADMINISTRATION OVERVIEW
# OWNER ONLY
# =========================================================

@app.route("/administration")
def administration_overview():

    conn = get_db()

    # -----------------------------------------------------
    # STAFF COUNTS
    # -----------------------------------------------------

    total_staff = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'STAFF'
        """
    ).fetchone()[0]

    active_staff = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'STAFF'
        AND status = 'ACTIVE'
        """
    ).fetchone()[0]

    inactive_staff = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'STAFF'
        AND status != 'ACTIVE'
        """
    ).fetchone()[0]

    # -----------------------------------------------------
    # AUDIT COUNTS
    # -----------------------------------------------------

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    audit_today = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        WHERE substr(created_at, 1, 10) = ?
        """,
        (today,),
    ).fetchone()[0]

    total_audit = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        """
    ).fetchone()[0]

    # -----------------------------------------------------
    # FAILED LOGIN EVENTS - LAST 15 MINUTES
    # -----------------------------------------------------

    cutoff_epoch = (
        int(time.time()) - 900
    )

    failed_login_events = conn.execute(
        """
        SELECT COUNT(*)
        FROM rate_limit_events
        WHERE scope IN (
            'login_failure_ip',
            'login_failure_user'
        )
        AND created_at_epoch >= ?
        """,
        (cutoff_epoch,),
    ).fetchone()[0]

    # -----------------------------------------------------
    # RECENT AUDIT ACTIVITY
    # -----------------------------------------------------

    recent_rows = conn.execute(
        """
        SELECT
            id,
            actor_username,
            actor_full_name,
            actor_role,
            action,
            entity_type,
            entity_id,
            request_method,
            request_path,
            ip_address,
            status_code,
            created_at
        FROM audit_logs
        ORDER BY id DESC
        LIMIT 8
        """
    ).fetchall()

    recent_activity = [
        dict(row)
        for row in recent_rows
    ]

    conn.close()

    # -----------------------------------------------------
    # BACKUP STATUS
    # -----------------------------------------------------

    backup_items = []

    if BACKUP_DIR.exists():

        backup_paths = sorted(
            BACKUP_DIR.glob(
                "sri_sumana_*.db"
            ),
            key=lambda item:
                item.stat().st_mtime,
            reverse=True,
        )

        for backup_path in backup_paths[:5]:

            stat = backup_path.stat()

            backup_items.append(
                {
                    "name":
                        backup_path.name,

                    "size_mb":
                        round(
                            stat.st_size
                            / (1024 * 1024),
                            2,
                        ),

                    "modified_at":
                        datetime.fromtimestamp(
                            stat.st_mtime
                        ).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),
                }
            )

    latest_backup = (
        backup_items[0]
        if backup_items
        else None
    )

    # -----------------------------------------------------
    # SECURITY STATUS
    # -----------------------------------------------------

    security_summary = {

        "https_enabled":
            bool(
                HTTPS_ENABLED
            ),

        "trust_proxy":
            bool(
                TRUST_PROXY
            ),

        "secure_cookie":
            bool(
                app.config.get(
                    "SESSION_COOKIE_SECURE"
                )
            ),

        "httponly_cookie":
            bool(
                app.config.get(
                    "SESSION_COOKIE_HTTPONLY"
                )
            ),

        "same_site":
            app.config.get(
                "SESSION_COOKIE_SAMESITE",
                "Not Set",
            ),

        "rate_limiting": True,
        "audit_logging": True,
        "security_headers": True,
    }

    stats = {

        "total_staff":
            int(total_staff or 0),

        "active_staff":
            int(active_staff or 0),

        "inactive_staff":
            int(inactive_staff or 0),

        "audit_today":
            int(audit_today or 0),

        "total_audit":
            int(total_audit or 0),

        "failed_login_events":
            int(
                failed_login_events
                or 0
            ),
    }

    return render_template(
        "admin_overview.html",
        stats=stats,
        security_summary=security_summary,
        latest_backup=latest_backup,
        backup_items=backup_items,
        recent_activity=recent_activity,
    )


# =========================================================
# OWNER: STAFF & ACCESS
# =========================================================

@app.route("/access")
def staff_access():
    conn = get_db()

    staff_rows = conn.execute(
        """
        SELECT
            id,
            username,
            full_name,
            role,
            status,
            created_at,
            updated_at,
            last_login_at
        FROM app_users
        WHERE role = 'STAFF'
        ORDER BY
            CASE
                WHEN status = 'ACTIVE'
                THEN 0
                ELSE 1
            END,
            full_name
        """
    ).fetchall()

    conn.close()

    return render_template(
        "staff_access.html",
        staff_members=staff_rows,
    )


@app.route(
    "/access/staff/new",
    methods=["GET", "POST"],
)
def new_staff():
    if request.method == "POST":
        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        username = request.form.get(
            "username",
            "",
        ).strip().lower()

        password = request.form.get(
            "password",
            "",
        )

        confirm_password = request.form.get(
            "confirm_password",
            "",
        )

        if not full_name:
            flash(
                "Staff name is required.",
                "error",
            )
            return redirect(
                url_for("new_staff")
            )

        error = username_error(
            username
        )

        if error:
            flash(error, "error")
            return redirect(
                url_for("new_staff")
            )

        error = password_error(
            password
        )

        if error:
            flash(error, "error")
            return redirect(
                url_for("new_staff")
            )

        if password != confirm_password:
            flash(
                "Passwords do not match.",
                "error",
            )
            return redirect(
                url_for("new_staff")
            )

        conn = get_db()
        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        try:
            conn.execute(
                """
                INSERT INTO app_users
                (
                    username,
                    full_name,
                    password_hash,
                    role,
                    status,
                    created_at,
                    updated_at
                )
                VALUES
                (
                    ?, ?, ?,
                    'STAFF',
                    'ACTIVE',
                    ?, ?
                )
                """,
                (
                    username,
                    full_name,
                    generate_password_hash(
                        password
                    ),
                    now,
                    now,
                ),
            )

            conn.commit()

        except sqlite3.IntegrityError:
            conn.close()
            flash(
                "That username is already in use.",
                "error",
            )
            return redirect(
                url_for("new_staff")
            )

        conn.close()

        flash(
            "Staff account created successfully.",
            "success",
        )

        return redirect(
            url_for("staff_access")
        )

    return render_template(
        "staff_form.html"
    )


@app.post(
    "/access/staff/<int:user_id>/toggle"
)
def toggle_staff(user_id):
    conn = get_db()

    user = conn.execute(
        """
        SELECT *
        FROM app_users
        WHERE id = ?
        AND role = 'STAFF'
        """,
        (user_id,),
    ).fetchone()

    if not user:
        conn.close()
        abort(404)

    new_status = (
        "INACTIVE"
        if user["status"] == "ACTIVE"
        else "ACTIVE"
    )

    conn.execute(
        """
        UPDATE app_users
        SET
            status = ?,
            updated_at = ?
        WHERE id = ?
        """,
        (
            new_status,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            user_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        f"Staff account is now {new_status}.",
        "success",
    )

    return redirect(
        url_for("staff_access")
    )


@app.post(
    "/access/staff/<int:user_id>/reset-password"
)
def reset_staff_password(user_id):
    new_password = request.form.get(
        "new_password",
        "",
    )

    confirm_password = request.form.get(
        "confirm_password",
        "",
    )

    error = password_error(
        new_password
    )

    if error:
        flash(error, "error")
        return redirect(
            url_for("staff_access")
        )

    if new_password != confirm_password:
        flash(
            "Passwords do not match.",
            "error",
        )
        return redirect(
            url_for("staff_access")
        )

    conn = get_db()

    user = conn.execute(
        """
        SELECT id
        FROM app_users
        WHERE id = ?
        AND role = 'STAFF'
        """,
        (user_id,),
    ).fetchone()

    if not user:
        conn.close()
        abort(404)

    conn.execute(
        """
        UPDATE app_users
        SET
            password_hash = ?,
            updated_at = ?
        WHERE id = ?
        """,
        (
            generate_password_hash(
                new_password
            ),
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            user_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        "Staff password reset successfully.",
        "success",
    )

    return redirect(
        url_for("staff_access")
    )


# =========================================================
# FRONT DESK / QUICK ACCESS
# =========================================================

@app.route("/front-desk")
def front_desk():

    conn = get_db()

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    month_key = datetime.now().strftime(
        "%Y-%m"
    )

    attendance_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM attendance
        WHERE attendance_date = ?
        """,
        (
            today,
        ),
    ).fetchone()[0]

    pending_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM student_registrations
        WHERE status = 'PENDING'
        """
    ).fetchone()[0]

    finance = get_payment_summary(
        conn,
        month_key,
    )

    today_classes = (
        get_today_remaining_classes(
            conn
        )
    )

    payment_followup_count = (
        finance["partial_count"]
        +
        finance["unpaid_count"]
    )

    conn.close()

    return render_template(
        "front_desk.html",

        today_name=
            datetime.now().strftime(
                "%A"
            ),

        today_date=
            datetime.now().strftime(
                "%Y-%m-%d"
            ),

        attendance_count=
            attendance_count,

        pending_count=
            pending_count,

        payment_followup_count=
            payment_followup_count,

        today_class_count=
            len(
                today_classes
            ),
    )


# =========================================================
# PUBLIC WEBSITE
# =========================================================

@app.route("/")
def public_home():
    conn = get_db()

    available_classes = get_active_registration_classes(
        conn
    )

    conn.close()

    return render_template(
        "public_home.html",
        classes=available_classes,
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
def dashboard():

    conn = get_db()

    pending_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM student_registrations

        WHERE status = 'PENDING'
        """
    ).fetchone()[0]

    approved_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM students

        WHERE status = 'ACTIVE'
        """
    ).fetchone()[0]

    teacher_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM teachers

        WHERE status = 'ACTIVE'
        """
    ).fetchone()[0]

    class_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM classes

        WHERE status = 'ACTIVE'
        """
    ).fetchone()[0]

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    attendance_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM attendance

        WHERE attendance_date = ?
        """,
        (
            today,
        ),
    ).fetchone()[0]

    month_key = datetime.now().strftime(
        "%Y-%m"
    )

    pending_free_card_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM free_cards
        WHERE status = 'PENDING'
        AND request_source = 'TEACHER'
        """
    ).fetchone()[0]

    active_free_card_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM free_cards
        WHERE status = 'APPROVED'
        AND fee_month = ?
        """,
        (
            month_key,
        ),
    ).fetchone()[0]

    finance = get_payment_summary(
        conn,
        month_key,
    )

    today_classes = (
        get_today_remaining_classes(
            conn
        )
    )

    next_class = (
        find_next_class(
            conn
        )
    )

    conn.close()

    stats = [
        {
            "title":
                "Total Students",
            "value":
                str(
                    approved_count
                ),
            "note":
                "Approved students",
        },
        {
            "title":
                "Total Teachers",
            "value":
                str(
                    teacher_count
                ),
            "note":
                "Active teachers",
        },
        {
            "title":
                "Active Classes",
            "value":
                str(
                    class_count
                ),
            "note":
                "Running classes",
        },
        {
            "title":
                "Pending Approvals",
            "value":
                str(
                    pending_count
                ),
            "note":
                "Need review",
        },
        {
            "title":
                "Active Free Cards",
            "value":
                str(
                    active_free_card_count
                ),
            "note":
                "Approved this month",
        },
        {
            "title":
                "This Month Collection",
            "value":
                f"Rs. {finance['collected']:,.0f}",
            "note":
                "Confirmed payments",
        },
        {
            "title":
                "Outstanding",
            "value":
                f"Rs. {finance['outstanding']:,.0f}",
            "note":
                "Current fee month",
        },
    ]

    pending_actions = [
        {
            "name":
                "Student Approvals",
            "count":
                pending_count,
            "url":
                url_for(
                    "pending_approvals"
                ),
        },
        {
            "name":
                "Free Card Approvals",
            "count":
                pending_free_card_count,
            "url":
                url_for(
                    "teacher_free_card_approvals"
                ),
        },
        {
            "name":
                "Scan QR",
            "count":
                attendance_count,
            "url":
                url_for(
                    "scan_qr"
                ),
        },
        {
            "name":
                "Record Payment",
            "count":
                finance[
                    "partial_count"
                ]
                +
                finance[
                    "unpaid_count"
                ],
            "url":
                url_for(
                    "new_payment"
                ),
        },
    ]

    return render_template(
        "dashboard.html",

        stats=
            stats,

        pending_actions=
            pending_actions,

        today_classes=
            today_classes,

        next_class=
            next_class,

        today_name=
            datetime.now().strftime(
                "%A"
            ),

        finance=
            finance,

        finance_month=
            datetime.now().strftime(
                "%B %Y"
            ),
    )


def make_public_status_token(
    conn
):

    while True:

        token = secrets.token_urlsafe(
            32
        )

        exists = conn.execute(
            """
            SELECT id

            FROM student_registrations

            WHERE public_status_token = ?
            """,
            (
                token,
            ),
        ).fetchone()

        if not exists:
            return token


# =========================================================
# PRIVATE STUDENT PHOTOS
# =========================================================

@app.get("/student-photos/<filename>")
def student_photo_file(filename):

    if not re.fullmatch(
        r"student_[0-9a-f]{32}\.jpg",
        filename or "",
    ):
        abort(404)

    conn = get_db()

    photo_exists = conn.execute(
        """
        SELECT 1
        FROM student_registrations
        WHERE photo_filename = ?

        UNION

        SELECT 1
        FROM students
        WHERE photo_filename = ?

        LIMIT 1
        """,
        (filename, filename),
    ).fetchone()

    conn.close()

    if not photo_exists:
        abort(404)

    photo_path = STUDENT_PHOTO_DIR / filename

    if not photo_path.is_file():
        abort(404)

    response = send_file(
        photo_path,
        mimetype="image/jpeg",
        conditional=True,
    )

    response.headers["Cache-Control"] = (
        "private, no-store, max-age=0"
    )

    return response


# =========================================================
# STUDENT REGISTRATION
# =========================================================

@app.route(
    "/admin/students/new",
    methods=[
        "GET",
        "POST",
    ],
)
@app.route(
    "/register",
    methods=[
        "GET",
        "POST",
    ],
    endpoint="public_register_student",
)
def register_student():

    conn = get_db()

    available_classes = (
        get_active_registration_classes(
            conn
        )
    )

    public_registration = (
        request.endpoint == "public_register_student"
    )

    template_name = (
        "public_register.html"
        if public_registration
        else "register.html"
    )

    if request.method == "POST":

        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        phone = request.form.get(
            "phone",
            "",
        ).strip()

        parent_name = request.form.get(
            "parent_name",
            "",
        ).strip()

        parent_phone = request.form.get(
            "parent_phone",
            "",
        ).strip()

        school = request.form.get(
            "school",
            "",
        ).strip()

        grade = request.form.get(
            "grade",
            "",
        ).strip()

        address = request.form.get(
            "address",
            "",
        ).strip()

        student_photo = request.files.get(
            "student_photo"
        )

        selected_class_ids = (
            request.form.getlist(
                "classes"
            )
        )

        if (
            not full_name
            or not phone
            or not parent_name
            or not parent_phone
            or not grade
        ):

            conn.close()

            flash(
                "Please fill all required fields.",
                "error",
            )

            return render_template(
                template_name,
                classes=available_classes,
                form=request.form,
            )

        if not selected_class_ids:

            conn.close()

            flash(
                "Please select at least one class.",
                "error",
            )

            return render_template(
                template_name,
                classes=available_classes,
                form=request.form,
            )

        selected_classes = [
            item
            for item in available_classes
            if (
                item["id"] in selected_class_ids
                and item["grade"] == grade
            )
        ]

        if not selected_classes:

            conn.close()

            flash(
                "Please select a class that matches the selected grade.",
                "error",
            )

            return render_template(
                template_name,
                classes=available_classes,
                form=request.form,
            )

        # Public /register ALWAYS requires OTP, even if an Owner/Staff
        # session is also active in this browser. The admin route
        # /admin/students/new ALWAYS keeps the original manual flow.
        otp_phone = None
        otp_phone_label = None

        if public_registration:
            try:
                otp_phone, otp_phone_label = (
                    registration_otp_destination(
                        phone,
                        parent_phone,
                    )
                )
            except ValueError as exc:
                conn.close()

                flash(
                    str(exc),
                    "error",
                )

                return render_template(
                    template_name,
                    classes=available_classes,
                    form=request.form,
                )

            ip_address = get_client_ip()

            if rate_limit_count(
                "registration_otp_phone",
                otp_phone,
                3600,
            ) >= REGISTRATION_OTP_MAX_SENDS:
                conn.close()

                flash(
                    "Too many verification codes were requested for this phone number. Please try again later.",
                    "error",
                )

                return render_template(
                    template_name,
                    classes=available_classes,
                    form=request.form,
                )

            if rate_limit_count(
                "registration_otp_send_ip",
                ip_address,
                3600,
            ) >= 10:
                conn.close()

                flash(
                    "Too many verification requests. Please try again later.",
                    "error",
                )

                return render_template(
                    template_name,
                    classes=available_classes,
                    form=request.form,
                )

        try:
            photo_filename = save_student_photo(
                student_photo
            )
        except ValueError as exc:
            conn.close()

            flash(
                str(exc),
                "error",
            )

            return render_template(
                template_name,
                classes=available_classes,
                form=request.form,
            )

        class_summary = " | ".join(
            f'{item["id"]}:'
            f'{item["subject"]}:'
            f'{item["teacher"]}'
            for item in selected_classes
        )

        # -----------------------------------------------------
        # OWNER / STAFF MANUAL REGISTRATION
        # Keep the original flow: save directly as PENDING.
        # -----------------------------------------------------
        if not public_registration:

            registration_ref = (
                make_registration_ref(
                    conn
                )
            )

            public_status_token = (
                make_public_status_token(
                    conn
                )
            )

            conn.execute(
                """
                INSERT INTO student_registrations
                (
                    registration_ref,
                    full_name,
                    phone,
                    parent_name,
                    parent_phone,
                    school,
                    grade,
                    address,
                    photo_filename,
                    selected_classes,
                    public_status_token,
                    status,
                    created_at
                )
                VALUES
                (
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?,
                    ?,
                    'PENDING',
                    ?
                )
                """,
                (
                    registration_ref,
                    full_name,
                    phone,
                    parent_name,
                    parent_phone,
                    school,
                    grade,
                    address,
                    photo_filename,
                    class_summary,
                    public_status_token,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                ),
            )

            conn.commit()
            conn.close()

            return redirect(
                url_for(
                    "registration_success",
                    ref=registration_ref,
                    token=public_status_token,
                )
            )

        # -----------------------------------------------------
        # PUBLIC WEBSITE REGISTRATION
        # Hold the application as an OTP draft first.
        # It does not enter Pending Approvals until verified.
        # -----------------------------------------------------
        conn.close()

        cleanup_registration_otp_sessions()

        verify_token = secrets.token_urlsafe(32)

        if SMS_MODE == "demo":
            otp_code = make_demo_registration_otp_code(
                verify_token
            )
        else:
            otp_code = make_registration_otp_code()

        now = datetime.now()
        otp_expires_at = (
            now
            + timedelta(
                seconds=REGISTRATION_OTP_TTL_SECONDS
            )
        )
        resend_available_at = (
            now
            + timedelta(
                seconds=REGISTRATION_OTP_RESEND_SECONDS
            )
        )

        conn = get_db()

        try:
            conn.execute(
                """
                INSERT INTO registration_otp_sessions
                (
                    verify_token,
                    otp_hash,
                    otp_phone,
                    full_name,
                    phone,
                    parent_name,
                    parent_phone,
                    school,
                    grade,
                    address,
                    photo_filename,
                    selected_classes,
                    status,
                    attempt_count,
                    send_count,
                    otp_expires_at,
                    resend_available_at,
                    created_at
                )
                VALUES
                (
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?,
                    ?, ?,
                    'PENDING_OTP',
                    0,
                    1,
                    ?, ?, ?
                )
                """,
                (
                    verify_token,
                    registration_otp_hash(
                        verify_token,
                        otp_code,
                    ),
                    otp_phone,
                    full_name,
                    phone,
                    parent_name,
                    parent_phone,
                    school,
                    grade,
                    address,
                    photo_filename,
                    class_summary,
                    otp_expires_at.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    resend_available_at.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    now.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                ),
            )

            conn.commit()

        finally:
            conn.close()

        sent, _ = send_registration_sms(
            otp_phone,
            registration_otp_message(
                otp_code
            ),
        )

        if not sent:
            cleanup_conn = get_db()

            try:
                cleanup_conn.execute(
                    """
                    DELETE FROM registration_otp_sessions
                    WHERE verify_token = ?
                    """,
                    (verify_token,),
                )
                cleanup_conn.commit()
            finally:
                cleanup_conn.close()

            delete_student_photo_file(
                photo_filename
            )

            flash(
                "We could not send the verification SMS. Please check the phone number and try again.",
                "error",
            )

            return render_template(
                template_name,
                classes=available_classes,
                form=request.form,
            )

        record_rate_limit_event(
            "registration_otp_phone",
            otp_phone,
        )
        record_rate_limit_event(
            "registration_otp_send_ip",
            get_client_ip(),
        )

        return redirect(
            url_for(
                "registration_otp_verify",
                token=verify_token,
            )
        )

    conn.close()

    return render_template(
        template_name,
        classes=available_classes,
        form={},
    )


@app.route(
    "/registration/verify/<token>",
    methods=["GET", "POST"],
)
def registration_otp_verify(token):

    if (
        not token
        or len(token) < 30
        or len(token) > 128
    ):
        abort(404)

    conn = get_db()

    draft = conn.execute(
        """
        SELECT *
        FROM registration_otp_sessions
        WHERE verify_token = ?
        """,
        (token,),
    ).fetchone()

    if not draft:
        conn.close()
        abort(404)

    if (
        draft["status"] == "VERIFIED"
        and draft["registration_ref"]
        and draft["public_status_token"]
    ):
        conn.close()

        return redirect(
            url_for(
                "registration_success",
                ref=draft["registration_ref"],
                token=draft["public_status_token"],
            )
        )

    now = datetime.now()

    try:
        expires_at = datetime.strptime(
            draft["otp_expires_at"],
            "%Y-%m-%d %H:%M:%S",
        )
    except (TypeError, ValueError):
        expires_at = now

    try:
        resend_at = datetime.strptime(
            draft["resend_available_at"],
            "%Y-%m-%d %H:%M:%S",
        )
    except (TypeError, ValueError):
        resend_at = now

    expired = now > expires_at

    if request.method == "POST":

        if expired:
            conn.close()

            flash(
                "This verification code has expired. Please request a new code.",
                "error",
            )

            return redirect(
                url_for(
                    "registration_otp_verify",
                    token=token,
                )
            )

        if (
            int(draft["attempt_count"] or 0)
            >= REGISTRATION_OTP_MAX_ATTEMPTS
        ):
            conn.close()

            flash(
                "Too many incorrect attempts. Please request a new verification code.",
                "error",
            )

            return redirect(
                url_for(
                    "registration_otp_verify",
                    token=token,
                )
            )

        otp_code = re.sub(
            r"\D",
            "",
            request.form.get(
                "otp_code",
                "",
            ),
        )

        correct_hash = registration_otp_hash(
            token,
            otp_code,
        )

        if (
            len(otp_code) != 6
            or not hmac.compare_digest(
                correct_hash,
                draft["otp_hash"],
            )
        ):
            conn.execute(
                """
                UPDATE registration_otp_sessions
                SET attempt_count = attempt_count + 1
                WHERE id = ?
                """,
                (draft["id"],),
            )
            conn.commit()
            conn.close()

            flash(
                "The verification code is incorrect.",
                "error",
            )

            return redirect(
                url_for(
                    "registration_otp_verify",
                    token=token,
                )
            )

        # OTP is valid. Only now create the real PENDING registration.
        registration_ref = make_registration_ref(
            conn
        )
        public_status_token = make_public_status_token(
            conn
        )

        conn.execute(
            """
            INSERT INTO student_registrations
            (
                registration_ref,
                full_name,
                phone,
                parent_name,
                parent_phone,
                school,
                grade,
                address,
                photo_filename,
                selected_classes,
                public_status_token,
                status,
                created_at
            )
            VALUES
            (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?,
                'PENDING',
                ?
            )
            """,
            (
                registration_ref,
                draft["full_name"],
                draft["phone"],
                draft["parent_name"],
                draft["parent_phone"],
                draft["school"],
                draft["grade"],
                draft["address"],
                draft["photo_filename"],
                draft["selected_classes"],
                public_status_token,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            ),
        )

        conn.execute(
            """
            UPDATE registration_otp_sessions
            SET
                status = 'VERIFIED',
                registration_ref = ?,
                public_status_token = ?,
                verified_at = ?
            WHERE id = ?
            """,
            (
                registration_ref,
                public_status_token,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                draft["id"],
            ),
        )

        conn.commit()
        conn.close()

        return redirect(
            url_for(
                "registration_success",
                ref=registration_ref,
                token=public_status_token,
            )
        )

    remaining_attempts = max(
        0,
        REGISTRATION_OTP_MAX_ATTEMPTS
        - int(draft["attempt_count"] or 0),
    )

    resend_seconds = max(
        0,
        int((resend_at - now).total_seconds()),
    )

    conn.close()

    return render_template(
        "registration_otp.html",
        verify_token=token,
        full_name=draft["full_name"],
        masked_phone=masked_mobile_number(
            draft["otp_phone"]
        ),
        expired=expired,
        remaining_attempts=remaining_attempts,
        resend_seconds=resend_seconds,
        send_count=int(draft["send_count"] or 1),
        max_sends=REGISTRATION_OTP_MAX_SENDS,
        demo_mode=(SMS_MODE == "demo"),
        demo_otp_code=(
            make_demo_registration_otp_code(token)
            if SMS_MODE == "demo"
            else None
        ),
    )


@app.post(
    "/registration/verify/<token>/resend"
)
def registration_otp_resend(token):

    if (
        not token
        or len(token) < 30
        or len(token) > 128
    ):
        abort(404)

    conn = get_db()

    draft = conn.execute(
        """
        SELECT *
        FROM registration_otp_sessions
        WHERE verify_token = ?
        """,
        (token,),
    ).fetchone()

    if not draft:
        conn.close()
        abort(404)

    if draft["status"] == "VERIFIED":
        conn.close()

        return redirect(
            url_for(
                "registration_otp_verify",
                token=token,
            )
        )

    now = datetime.now()

    try:
        created_at = datetime.strptime(
            draft["created_at"],
            "%Y-%m-%d %H:%M:%S",
        )
    except (TypeError, ValueError):
        created_at = now - timedelta(days=2)

    if (
        now - created_at
        > timedelta(hours=REGISTRATION_OTP_DRAFT_HOURS)
    ):
        conn.execute(
            """
            UPDATE registration_otp_sessions
            SET status = 'EXPIRED'
            WHERE id = ?
            """,
            (draft["id"],),
        )
        conn.commit()
        conn.close()

        flash(
            "This registration verification session has expired. Please submit the registration form again.",
            "error",
        )

        return redirect(
            url_for("public_home")
        )

    try:
        resend_at = datetime.strptime(
            draft["resend_available_at"],
            "%Y-%m-%d %H:%M:%S",
        )
    except (TypeError, ValueError):
        resend_at = now

    if now < resend_at:
        seconds = max(
            1,
            int((resend_at - now).total_seconds()),
        )
        conn.close()

        flash(
            f"Please wait {seconds} seconds before requesting another code.",
            "error",
        )

        return redirect(
            url_for(
                "registration_otp_verify",
                token=token,
            )
        )

    if (
        int(draft["send_count"] or 0)
        >= REGISTRATION_OTP_MAX_SENDS
    ):
        conn.close()

        flash(
            "The maximum number of verification codes has been reached. Please submit the registration form again later.",
            "error",
        )

        return redirect(
            url_for(
                "registration_otp_verify",
                token=token,
            )
        )

    otp_phone = draft["otp_phone"]
    ip_address = get_client_ip()

    if rate_limit_count(
        "registration_otp_phone",
        otp_phone,
        3600,
    ) >= REGISTRATION_OTP_MAX_SENDS:
        conn.close()

        flash(
            "Too many verification codes were requested for this phone number. Please try again later.",
            "error",
        )

        return redirect(
            url_for(
                "registration_otp_verify",
                token=token,
            )
        )

    if SMS_MODE == "demo":
        otp_code = make_demo_registration_otp_code(
            token
        )
    else:
        otp_code = make_registration_otp_code()

    sent, _ = send_registration_sms(
        otp_phone,
        registration_otp_message(
            otp_code
        ),
    )

    if not sent:
        conn.close()

        flash(
            "We could not resend the verification SMS. Please try again.",
            "error",
        )

        return redirect(
            url_for(
                "registration_otp_verify",
                token=token,
            )
        )

    new_expires_at = (
        now
        + timedelta(
            seconds=REGISTRATION_OTP_TTL_SECONDS
        )
    )
    new_resend_at = (
        now
        + timedelta(
            seconds=REGISTRATION_OTP_RESEND_SECONDS
        )
    )

    conn.execute(
        """
        UPDATE registration_otp_sessions
        SET
            otp_hash = ?,
            attempt_count = 0,
            send_count = send_count + 1,
            otp_expires_at = ?,
            resend_available_at = ?,
            status = 'PENDING_OTP'
        WHERE id = ?
        """,
        (
            registration_otp_hash(
                token,
                otp_code,
            ),
            new_expires_at.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            new_resend_at.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            draft["id"],
        ),
    )

    conn.commit()
    conn.close()

    record_rate_limit_event(
        "registration_otp_phone",
        otp_phone,
    )
    record_rate_limit_event(
        "registration_otp_send_ip",
        ip_address,
    )

    flash(
        "A new verification code was sent.",
        "success",
    )

    return redirect(
        url_for(
            "registration_otp_verify",
            token=token,
        )
    )


@app.route(
    "/registration-success"
)
def registration_success():

    token = request.args.get(
        "token",
        "",
    ).strip()

    ref = request.args.get(
        "ref",
        "",
    ).strip()

    if (
        not token
        or
        len(token) < 30
        or
        len(token) > 128
    ):
        abort(404)

    if not ref:
        abort(404)

    return render_template(
        "registration_success.html",

        ref=
            ref,

        full_name=
            "",

        status_path=
            url_for(
                "registration_status",
                token=token,
            ),
    )


@app.route(
    "/registration/status/<token>"
)
def registration_status(
    token
):

    if (
        not token
        or
        len(token) < 30
        or
        len(token) > 128
    ):
        abort(404)

    conn = get_db()

    registration_row = conn.execute(
        """
        SELECT *

        FROM student_registrations

        WHERE public_status_token = ?
        """,
        (
            token,
        ),
    ).fetchone()

    if not registration_row:

        conn.close()
        abort(404)

    registration = dict(
        registration_row
    )

    registration["classes"] = (
        parse_selected_classes(
            registration[
                "selected_classes"
            ]
        )
    )

    student = None
    enrolled_classes = []

    if (
        registration["status"]
        ==
        "APPROVED"
    ):

        student_row = conn.execute(
            """
            SELECT *

            FROM students

            WHERE registration_id = ?
            """,
            (
                registration["id"],
            ),
        ).fetchone()

        if student_row:

            student = dict(
                student_row
            )

            class_rows = conn.execute(
                """
                SELECT
                    class_code,
                    subject,
                    teacher,
                    enrollment_status

                FROM student_classes

                WHERE student_id = ?
                AND enrollment_status = 'ACTIVE'

                ORDER BY id ASC
                """,
                (
                    student["id"],
                ),
            ).fetchall()

            enrolled_classes = [
                dict(item)
                for item in class_rows
            ]

    conn.close()

    return render_template(
        "registration_status.html",

        registration=
            registration,

        student=
            student,

        classes=
            enrolled_classes,

        photo_path=(
            url_for(
                "registration_status_photo",
                token=token,
            )
            if (
                student
                and
                student.get(
                    "photo_filename"
                )
            )
            else
            None
        ),

        qr_path=(
            url_for(
                "registration_status_qr",
                token=token,
            )
            if (
                student
                and
                student.get(
                    "qr_status"
                )
                ==
                "ACTIVE"
            )
            else
            None
        ),
    )


@app.route(
    "/registration/status/<token>/photo.jpg"
)
def registration_status_photo(
    token
):

    if (
        not token
        or
        len(token) < 30
        or
        len(token) > 128
    ):
        abort(404)

    conn = get_db()

    row = conn.execute(
        """
        SELECT
            s.photo_filename,
            s.status AS student_status,
            r.status AS registration_status

        FROM student_registrations r

        JOIN students s
        ON s.registration_id = r.id

        WHERE r.public_status_token = ?
        """,
        (
            token,
        ),
    ).fetchone()

    conn.close()

    if not row:
        abort(404)

    if (
        row[
            "registration_status"
        ]
        !=
        "APPROVED"
        or
        row[
            "student_status"
        ]
        !=
        "ACTIVE"
        or
        not row[
            "photo_filename"
        ]
    ):
        abort(404)

    filename = row[
        "photo_filename"
    ]

    if not re.fullmatch(
        r"student_[0-9a-f]{32}\.jpg",
        filename or "",
    ):
        abort(404)

    photo_path = (
        STUDENT_PHOTO_DIR
        / filename
    )

    if not photo_path.is_file():
        abort(404)

    response = send_file(
        photo_path,
        mimetype="image/jpeg",
        conditional=True,
    )

    response.headers[
        "Cache-Control"
    ] = (
        "private, no-store, max-age=0"
    )

    return response


@app.route(
    "/registration/status/<token>/qr.png"
)
def registration_status_qr(
    token
):

    if (
        not token
        or
        len(token) < 30
        or
        len(token) > 128
    ):
        abort(404)

    conn = get_db()

    student = conn.execute(
        """
        SELECT
            s.qr_token,
            s.qr_status,
            s.status AS student_status,
            r.status AS registration_status

        FROM student_registrations r

        JOIN students s
        ON s.registration_id = r.id

        WHERE r.public_status_token = ?
        """,
        (
            token,
        ),
    ).fetchone()

    conn.close()

    if not student:
        abort(404)

    if (
        student[
            "registration_status"
        ]
        !=
        "APPROVED"
        or
        student[
            "student_status"
        ]
        !=
        "ACTIVE"
        or
        student[
            "qr_status"
        ]
        !=
        "ACTIVE"
        or
        not student[
            "qr_token"
        ]
    ):
        abort(404)

    qr = qrcode.QRCode(
        version=None,

        error_correction=
            qrcode.constants
            .ERROR_CORRECT_M,

        box_size=10,
        border=4,
    )

    qr.add_data(
        "SSCQR:"
        +
        student[
            "qr_token"
        ]
    )

    qr.make(
        fit=True
    )

    image = qr.make_image(
        fill_color=
            "black",

        back_color=
            "white",
    )

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG",
    )

    buffer.seek(0)

    response = send_file(
        buffer,

        mimetype=
            "image/png",
    )

    response.headers[
        "Cache-Control"
    ] = (
        "no-store, "
        "private, "
        "max-age=0"
    )

    return response


# =========================================================
# APPROVALS
# =========================================================

@app.route("/approvals")
def pending_approvals():

    conn = get_db()

    pending = conn.execute(
        """
        SELECT *

        FROM student_registrations

        WHERE status = 'PENDING'

        ORDER BY id DESC
        """
    ).fetchall()

    recent = conn.execute(
        """
        SELECT *

        FROM student_registrations

        WHERE status IN
        (
            'APPROVED',
            'REJECTED'
        )

        ORDER BY
            COALESCE(
                reviewed_at,
                created_at
            )
            DESC

        LIMIT 8
        """
    ).fetchall()

    conn.close()

    pending_rows = []

    for row in pending:

        item = dict(row)

        item[
            "classes"
        ] = parse_selected_classes(
            item[
                "selected_classes"
            ]
        )

        pending_rows.append(
            item
        )

    recent_rows = []

    for row in recent:

        item = dict(row)

        item[
            "classes"
        ] = parse_selected_classes(
            item[
                "selected_classes"
            ]
        )

        recent_rows.append(
            item
        )

    return render_template(
        "approvals.html",

        pending=
            pending_rows,

        recent=
            recent_rows,
    )


@app.route(
    "/approvals/<int:registration_id>"
)
def approval_detail(
    registration_id
):

    conn = get_db()

    row = conn.execute(
        """
        SELECT *

        FROM student_registrations

        WHERE id = ?
        """,
        (
            registration_id,
        ),
    ).fetchone()

    conn.close()

    if not row:
        abort(404)

    registration = dict(row)

    registration[
        "classes"
    ] = parse_selected_classes(
        registration[
            "selected_classes"
        ]
    )

    return render_template(
        "approval_detail.html",

        registration=
            registration,
    )


@app.post(
    "/approvals/<int:registration_id>/approve"
)
def approve_registration(
    registration_id
):

    conn = get_db()

    try:

        conn.execute(
            "BEGIN IMMEDIATE"
        )

        row = conn.execute(
            """
            SELECT *

            FROM student_registrations

            WHERE id = ?
            """,
            (
                registration_id,
            ),
        ).fetchone()

        if not row:

            conn.rollback()
            abort(404)

        if (
            row["status"]
            != "PENDING"
        ):

            conn.rollback()

            flash(
                "This registration has already been reviewed.",
                "error",
            )

            return redirect(
                url_for(
                    "approval_detail",

                    registration_id=
                        registration_id,
                )
            )

        student_code = (
            make_student_code(
                conn
            )
        )

        qr_token = (
            make_qr_token(
                conn
            )
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        cursor = conn.execute(
            """
            INSERT INTO students
            (
                student_code,
                registration_id,
                full_name,
                phone,
                parent_name,
                parent_phone,
                school,
                grade,
                address,
                photo_filename,
                status,
                created_at,
                qr_token,
                qr_created_at,
                qr_status
            )

            VALUES
            (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                'ACTIVE',
                ?, ?, ?,
                'ACTIVE'
            )
            """,
            (
                student_code,
                row["id"],
                row["full_name"],
                row["phone"],
                row["parent_name"],
                row["parent_phone"],
                row["school"],
                row["grade"],
                row["address"],
                row["photo_filename"],
                now,
                qr_token,
                now,
            ),
        )

        student_id = (
            cursor.lastrowid
        )

        selected_classes = (
            parse_selected_classes(
                row[
                    "selected_classes"
                ]
            )
        )

        for class_item in selected_classes:

            conn.execute(
                """
                INSERT OR IGNORE
                INTO student_classes
                (
                    student_id,
                    class_code,
                    subject,
                    teacher,
                    created_at,
                    enrollment_status
                )

                VALUES
                (
                    ?, ?, ?, ?, ?,
                    'ACTIVE'
                )
                """,
                (
                    student_id,
                    class_item["id"],
                    class_item["subject"],
                    class_item["teacher"],
                    now,
                ),
            )

        conn.execute(
            """
            UPDATE student_registrations

            SET
                status = 'APPROVED',

                reviewed_at = ?,

                review_note = ?,

                rejection_reason = NULL,

                student_code = ?

            WHERE id = ?
            """,
            (
                now,
                f"Approved by {g.current_user['full_name']}",
                student_code,
                registration_id,
            ),
        )

        conn.commit()

        flash(
            (
                "Student approved successfully. "
                f"Student ID: {student_code}"
            ),
            "success",
        )

    except Exception:

        conn.rollback()
        raise

    finally:

        conn.close()

    return redirect(
        url_for(
            "approval_detail",

            registration_id=
                registration_id,
        )
    )


@app.post(
    "/approvals/<int:registration_id>/reject"
)
def reject_registration(
    registration_id
):

    reason = request.form.get(
        "rejection_reason",
        "",
    ).strip()

    if not reason:

        flash(
            "Please enter a rejection reason.",
            "error",
        )

        return redirect(
            url_for(
                "approval_detail",

                registration_id=
                    registration_id,
            )
        )

    conn = get_db()

    row = conn.execute(
        """
        SELECT *

        FROM student_registrations

        WHERE id = ?
        """,
        (
            registration_id,
        ),
    ).fetchone()

    if not row:

        conn.close()
        abort(404)

    if (
        row["status"]
        != "PENDING"
    ):

        conn.close()

        flash(
            "This registration has already been reviewed.",
            "error",
        )

        return redirect(
            url_for(
                "approval_detail",

                registration_id=
                    registration_id,
            )
        )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE student_registrations

        SET
            status = 'REJECTED',

            reviewed_at = ?,

            review_note = ?,

            rejection_reason = ?,

            student_code = NULL

        WHERE id = ?
        """,
        (
            now,
            f"Rejected by {g.current_user['full_name']}",
            reason,
            registration_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        "Registration rejected.",
        "success",
    )

    return redirect(
        url_for(
            "approval_detail",

            registration_id=
                registration_id,
        )
    )


# =========================================================
# STUDENT MANAGEMENT
# =========================================================

def get_student_class_history(
    conn,
    student_id,
):

    rows = conn.execute(
        """
        SELECT
            student_classes.id,
            student_classes.student_id,
            student_classes.class_code,
            student_classes.subject,
            student_classes.teacher,
            student_classes.created_at,
            student_classes.enrollment_status,
            student_classes.ended_at,

            classes.class_name,
            classes.grade AS class_grade,
            classes.day_of_week,
            classes.start_time,
            classes.end_time,
            classes.room,
            classes.monthly_fee,
            classes.status AS class_status,

            teachers.full_name AS current_teacher

        FROM student_classes

        LEFT JOIN classes
            ON classes.class_code =
                student_classes.class_code

        LEFT JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            student_classes.student_id = ?

        ORDER BY
            CASE
                WHEN (
                    student_classes.enrollment_status = 'ACTIVE'
                    OR student_classes.enrollment_status IS NULL
                )
                THEN 0
                ELSE 1
            END,
            classes.day_of_week,
            classes.start_time
        """,
        (
            student_id,
        ),
    ).fetchall()

    result = []

    for row in rows:

        item = dict(
            row
        )

        if (
            item.get(
                "day_of_week"
            )
            is not None
        ):

            item[
                "day_name"
            ] = DAY_NAMES[
                item[
                    "day_of_week"
                ]
            ]

        else:

            item[
                "day_name"
            ] = "-"

        item[
            "start_display"
        ] = (
            format_time(
                item[
                    "start_time"
                ]
            )
            if item.get(
                "start_time"
            )
            else "-"
        )

        item[
            "end_display"
        ] = (
            format_time(
                item[
                    "end_time"
                ]
            )
            if item.get(
                "end_time"
            )
            else "-"
        )

        result.append(
            item
        )

    return result


# =========================================================
# STUDENT LIST
# =========================================================

@app.route("/students")
def students():

    conn = get_db()

    rows = conn.execute(
        """
        SELECT
            students.*,

            (
                SELECT COUNT(*)

                FROM student_classes

                JOIN classes
                    ON classes.class_code =
                        student_classes.class_code

                WHERE
                    student_classes.student_id =
                        students.id

                AND (
                    student_classes.enrollment_status = 'ACTIVE'
                    OR student_classes.enrollment_status IS NULL
                )

                AND classes.status = 'ACTIVE'
            )
            AS active_class_count

        FROM students

        ORDER BY
            CASE
                WHEN students.status = 'ACTIVE'
                THEN 0
                ELSE 1
            END,
            students.id DESC
        """
    ).fetchall()

    result = []

    for row in rows:

        item = dict(
            row
        )

        item[
            "classes"
        ] = get_student_classes(
            conn,
            row[
                "id"
            ],
        )

        result.append(
            item
        )

    conn.close()

    return render_template(
        "students.html",

        students=
            result,
    )


# =========================================================
# STUDENT MANAGEMENT PROFILE
# =========================================================

@app.route(
    "/students/<int:student_id>"
)
def student_profile(
    student_id
):

    # =====================================================
    # PROFILE TAB
    # =====================================================

    allowed_tabs = {
        "overview",
        "classes",
        "attendance",
        "payments",
        "free-cards",
        "qr",
        "activity",
    }

    active_tab = request.args.get(
        "tab",
        "overview",
    ).strip().lower()

    if active_tab not in allowed_tabs:
        active_tab = "overview"

    # =====================================================
    # DATABASE
    # =====================================================

    conn = get_db()

    # =====================================================
    # STUDENT
    # =====================================================

    student_row = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student_row:
        conn.close()
        abort(404)

    student = dict(
        student_row
    )

    # =====================================================
    # CLASS HISTORY
    # =====================================================

    class_history = (
        get_student_class_history(
            conn,
            student_id,
        )
    )

    # =====================================================
    # AVAILABLE CLASSES
    # =====================================================

    available_classes = conn.execute(
        """
        SELECT
            classes.id,
            classes.class_code,
            classes.class_name,
            classes.grade,
            classes.subject,
            classes.day_of_week,
            classes.start_time,
            classes.end_time,
            classes.room,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM classes

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            classes.status = 'ACTIVE'

        AND teachers.status = 'ACTIVE'

        AND classes.grade = ?

        AND NOT EXISTS
        (
            SELECT 1

            FROM student_classes

            WHERE
                student_classes.student_id = ?

            AND student_classes.class_code =
                classes.class_code

            AND (
                student_classes.enrollment_status = 'ACTIVE'
                OR student_classes.enrollment_status IS NULL
            )
        )

        ORDER BY
            classes.subject,
            classes.day_of_week,
            classes.start_time
        """,
        (
            student["grade"],
            student_id,
        ),
    ).fetchall()

    available_result = []

    for row in available_classes:
        item = dict(row)

        if item.get("day_of_week") is not None:
            item["day_name"] = DAY_NAMES[
                item["day_of_week"]
            ]
        else:
            item["day_name"] = "-"

        item["start_display"] = (
            format_time(item["start_time"])
            if item.get("start_time")
            else "-"
        )

        item["end_display"] = (
            format_time(item["end_time"])
            if item.get("end_time")
            else "-"
        )

        available_result.append(item)

    # =====================================================
    # ACTIVE CLASS COUNT
    # =====================================================

    active_class_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM student_classes

        JOIN classes
            ON classes.class_code =
                student_classes.class_code

        WHERE
            student_classes.student_id = ?

        AND (
            student_classes.enrollment_status = 'ACTIVE'
            OR student_classes.enrollment_status IS NULL
        )

        AND classes.status = 'ACTIVE'
        """,
        (
            student_id,
        ),
    ).fetchone()[0]

    # =====================================================
    # ATTENDANCE HISTORY
    # =====================================================

    attendance_rows = conn.execute(
        """
        SELECT
            attendance.id,
            attendance.student_id,
            attendance.class_code,
            attendance.attendance_date,
            attendance.check_in_time,
            attendance.status,
            attendance.source,
            attendance.created_at,

            classes.class_name,

            teachers.full_name
                AS teacher_name

        FROM attendance

        LEFT JOIN classes
            ON classes.class_code =
                attendance.class_code

        LEFT JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            attendance.student_id = ?

        ORDER BY
            attendance.attendance_date DESC,
            attendance.check_in_time DESC,
            attendance.id DESC

        LIMIT 200
        """,
        (
            student_id,
        ),
    ).fetchall()

    attendance_records = [
        dict(row)
        for row in attendance_rows
    ]

    attendance_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM attendance

        WHERE student_id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()[0]

    # =====================================================
    # PAYMENT HISTORY
    # =====================================================

    payment_rows = conn.execute(
        """
        SELECT
            payments.id,
            payments.student_id,
            payments.class_code,
            payments.fee_month,
            payments.amount,
            payments.status,
            payments.payment_date,
            payments.created_at,
            payments.receipt_no,
            payments.payment_method,
            payments.notes,
            payments.recorded_by,
            payments.voided_at,
            payments.void_reason,

            classes.class_name,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM payments

        LEFT JOIN classes
            ON classes.class_code =
                payments.class_code

        LEFT JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            payments.student_id = ?

        ORDER BY
            payments.payment_date DESC,
            payments.id DESC

        LIMIT 200
        """,
        (
            student_id,
        ),
    ).fetchall()

    payment_records = [
        dict(row)
        for row in payment_rows
    ]

    # =====================================================
    # TOTAL CONFIRMED PAYMENTS
    # =====================================================

    total_paid = conn.execute(
        """
        SELECT COALESCE(
            SUM(amount),
            0
        )

        FROM payments

        WHERE
            student_id = ?

        AND status = 'CONFIRMED'
        """,
        (
            student_id,
        ),
    ).fetchone()[0]

    total_paid = float(
        total_paid
        or 0
    )

    # =====================================================
    # PAYMENT COUNTS
    # =====================================================

    confirmed_payment_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM payments

        WHERE
            student_id = ?

        AND status = 'CONFIRMED'
        """,
        (
            student_id,
        ),
    ).fetchone()[0]

    void_payment_count = conn.execute(
        """
        SELECT COUNT(*)

        FROM payments

        WHERE
            student_id = ?

        AND status = 'VOID'
        """,
        (
            student_id,
        ),
    ).fetchone()[0]

    # =====================================================
    # FREE CARD HISTORY
    # =====================================================

    free_card_sql = """
        SELECT
            free_cards.*,

            classes.class_name,
            classes.subject,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM free_cards

        LEFT JOIN classes
            ON classes.class_code =
                free_cards.class_code

        LEFT JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            free_cards.student_id = ?
    """

    free_card_params = [
        student_id,
    ]

    # Staff can only see their own Free Card requests.
    if (
        g.current_user
        and g.current_user["role"] == "STAFF"
    ):
        free_card_sql += """
            AND free_cards.requested_by_user_id = ?
        """
        free_card_params.append(
            g.current_user["id"]
        )

    free_card_sql += """
        ORDER BY
            free_cards.fee_month DESC,
            free_cards.id DESC
    """

    free_card_rows = conn.execute(
        free_card_sql,
        free_card_params,
    ).fetchall()

    free_card_records = [
        dict(row)
        for row in free_card_rows
    ]

    # =====================================================
    # APPROVED FREE CARD COUNT
    # =====================================================

    if (
        g.current_user
        and g.current_user["role"] == "STAFF"
    ):
        approved_free_cards = conn.execute(
            """
            SELECT COUNT(*)

            FROM free_cards

            WHERE
                student_id = ?

            AND status = 'APPROVED'

            AND requested_by_user_id = ?
            """,
            (
                student_id,
                g.current_user["id"],
            ),
        ).fetchone()[0]

        total_waived = conn.execute(
            """
            SELECT COALESCE(
                SUM(waived_amount),
                0
            )

            FROM free_cards

            WHERE
                student_id = ?

            AND status = 'APPROVED'

            AND requested_by_user_id = ?
            """,
            (
                student_id,
                g.current_user["id"],
            ),
        ).fetchone()[0]

    else:
        approved_free_cards = conn.execute(
            """
            SELECT COUNT(*)

            FROM free_cards

            WHERE
                student_id = ?

            AND status = 'APPROVED'
            """,
            (
                student_id,
            ),
        ).fetchone()[0]

        total_waived = conn.execute(
            """
            SELECT COALESCE(
                SUM(waived_amount),
                0
            )

            FROM free_cards

            WHERE
                student_id = ?

            AND status = 'APPROVED'
            """,
            (
                student_id,
            ),
        ).fetchone()[0]

    total_waived = float(
        total_waived
        or 0
    )

    # =====================================================
    # AUDIT / ACTIVITY HISTORY
    # OWNER ONLY
    # =====================================================

    activity_records = []

    if (
        g.current_user
        and g.current_user["role"] == "OWNER"
    ):
        activity_rows = conn.execute(
            """
            SELECT
                id,
                actor_user_id,
                actor_username,
                actor_full_name,
                actor_role,
                action,
                entity_type,
                entity_id,
                request_method,
                request_path,
                ip_address,
                status_code,
                created_at

            FROM audit_logs

            WHERE
                entity_type = 'STUDENT'

            AND entity_id = ?

            ORDER BY
                id DESC

            LIMIT 100
            """,
            (
                str(student_id),
            ),
        ).fetchall()

        activity_records = [
            dict(row)
            for row in activity_rows
        ]

    # =====================================================
    # CLOSE DATABASE
    # =====================================================

    conn.close()

    # =====================================================
    # RENDER PROFILE
    # =====================================================

    return render_template(
        "student_profile.html",

        student=student,
        active_tab=active_tab,
        classes=class_history,
        available_classes=available_result,
        active_class_count=active_class_count,
        attendance_records=attendance_records,
        attendance_count=attendance_count,
        payment_records=payment_records,
        total_paid=total_paid,
        confirmed_payment_count=confirmed_payment_count,
        void_payment_count=void_payment_count,
        free_card_records=free_card_records,
        approved_free_cards=approved_free_cards,
        total_waived=total_waived,
        activity_records=activity_records,
    )


# =========================================================
# EDIT STUDENT
# =========================================================

@app.route(
    "/students/<int:student_id>/edit",
    methods=[
        "GET",
        "POST",
    ],
)
def edit_student(
    student_id
):

    conn = get_db()

    student_row = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student_row:

        conn.close()

        abort(404)

    if request.method == "POST":

        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        phone = request.form.get(
            "phone",
            "",
        ).strip()

        parent_name = request.form.get(
            "parent_name",
            "",
        ).strip()

        parent_phone = request.form.get(
            "parent_phone",
            "",
        ).strip()

        school = request.form.get(
            "school",
            "",
        ).strip()

        grade = request.form.get(
            "grade",
            "",
        ).strip()

        address = request.form.get(
            "address",
            "",
        ).strip()

        if (
            not full_name
            or not phone
            or not parent_name
            or not parent_phone
            or not grade
        ):

            conn.close()

            flash(
                "Please fill all required fields.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_student",

                    student_id=
                        student_id,
                )
            )

        mismatch = conn.execute(
            """
            SELECT
                classes.class_code,
                classes.class_name,
                classes.grade

            FROM student_classes

            JOIN classes
                ON classes.class_code =
                    student_classes.class_code

            WHERE
                student_classes.student_id = ?

            AND (
                student_classes.enrollment_status = 'ACTIVE'
                OR student_classes.enrollment_status IS NULL
            )

            AND classes.status = 'ACTIVE'

            AND classes.grade != ?

            LIMIT 1
            """,
            (
                student_id,
                grade,
            ),
        ).fetchone()

        if mismatch:

            conn.close()

            flash(
                (
                    "Cannot change student grade while an "
                    "active class belongs to another grade. "
                    "Remove the old class first."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "edit_student",

                    student_id=
                        student_id,
                )
            )

        conn.execute(
            """
            UPDATE students

            SET
                full_name = ?,
                phone = ?,
                parent_name = ?,
                parent_phone = ?,
                school = ?,
                grade = ?,
                address = ?

            WHERE id = ?
            """,
            (
                full_name,
                phone,
                parent_name,
                parent_phone,
                school,
                grade,
                address,
                student_id,
            ),
        )

        conn.commit()

        conn.close()

        flash(
            "Student updated successfully.",
            "success",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    student = dict(
        student_row
    )

    conn.close()

    return render_template(
        "student_form.html",

        student=
            student,
    )


# =========================================================
# ACTIVATE / DEACTIVATE STUDENT
# =========================================================

@app.post(
    "/students/<int:student_id>/toggle"
)
def toggle_student(
    student_id
):

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student:

        conn.close()

        abort(404)

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    if (
        student[
            "status"
        ] == "ACTIVE"
    ):

        conn.execute(
            """
            UPDATE students

            SET
                status = 'INACTIVE',
                qr_status = 'REVOKED'

            WHERE id = ?
            """,
            (
                student_id,
            ),
        )

        message = (
            "Student deactivated. "
            "QR access has also been revoked."
        )

    else:

        new_token = (
            make_qr_token(
                conn
            )
        )

        conn.execute(
            """
            UPDATE students

            SET
                status = 'ACTIVE',
                qr_token = ?,
                qr_created_at = ?,
                qr_status = 'ACTIVE'

            WHERE id = ?
            """,
            (
                new_token,
                now,
                student_id,
            ),
        )

        message = (
            "Student activated successfully. "
            "A new QR token was generated."
        )

    conn.commit()

    conn.close()

    flash(
        message,
        "success",
    )

    return redirect(
        url_for(
            "student_profile",

            student_id=
                student_id,
        )
    )


# =========================================================
# ADD CLASS TO STUDENT
# =========================================================

@app.post(
    "/students/<int:student_id>/classes/add"
)
def add_student_class(
    student_id
):

    class_code = request.form.get(
        "class_code",
        "",
    ).strip()

    if not class_code:

        flash(
            "Please select a class.",
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student:

        conn.close()

        abort(404)

    if (
        student[
            "status"
        ] != "ACTIVE"
    ):

        conn.close()

        flash(
            "Activate the student before adding a class.",
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    class_row = conn.execute(
        """
        SELECT
            classes.*,

            teachers.full_name
                AS teacher_name,

            teachers.status
                AS teacher_status

        FROM classes

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            classes.class_code = ?
        """,
        (
            class_code,
        ),
    ).fetchone()

    if not class_row:

        conn.close()

        flash(
            "Class not found.",
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    if (
        class_row[
            "status"
        ] != "ACTIVE"

        or

        class_row[
            "teacher_status"
        ] != "ACTIVE"
    ):

        conn.close()

        flash(
            "Class or teacher is inactive.",
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    if (
        class_row[
            "grade"
        ] != student[
            "grade"
        ]
    ):

        conn.close()

        flash(
            (
                "This class does not match "
                "the student's grade."
            ),
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    existing = conn.execute(
        """
        SELECT *

        FROM student_classes

        WHERE
            student_id = ?

        AND class_code = ?
        """,
        (
            student_id,
            class_code,
        ),
    ).fetchone()

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    if existing:

        if (
            existing[
                "enrollment_status"
            ] == "ACTIVE"

            or

            existing[
                "enrollment_status"
            ] is None
        ):

            conn.close()

            flash(
                "Student is already enrolled in this class.",
                "error",
            )

            return redirect(
                url_for(
                    "student_profile",

                    student_id=
                        student_id,
                )
            )

        conn.execute(
            """
            UPDATE student_classes

            SET
                subject = ?,
                teacher = ?,
                created_at = ?,
                enrollment_status = 'ACTIVE',
                ended_at = NULL

            WHERE id = ?
            """,
            (
                class_row[
                    "subject"
                ],

                class_row[
                    "teacher_name"
                ],

                now,

                existing[
                    "id"
                ],
            ),
        )

    else:

        conn.execute(
            """
            INSERT INTO student_classes
            (
                student_id,
                class_code,
                subject,
                teacher,
                created_at,
                enrollment_status,
                ended_at
            )

            VALUES
            (
                ?, ?, ?, ?, ?,
                'ACTIVE',
                NULL
            )
            """,
            (
                student_id,
                class_code,

                class_row[
                    "subject"
                ],

                class_row[
                    "teacher_name"
                ],

                now,
            ),
        )

    conn.commit()

    conn.close()

    flash(
        (
            f"{class_row['class_name']} "
            "added to student successfully."
        ),
        "success",
    )

    return redirect(
        url_for(
            "student_profile",

            student_id=
                student_id,
        )
    )


# =========================================================
# REMOVE CLASS FROM STUDENT
# =========================================================

@app.post(
    "/students/<int:student_id>/classes/<class_code>/remove"
)
def remove_student_class(
    student_id,
    class_code,
):

    conn = get_db()

    enrollment = conn.execute(
        """
        SELECT *

        FROM student_classes

        WHERE
            student_id = ?

        AND class_code = ?
        """,
        (
            student_id,
            class_code,
        ),
    ).fetchone()

    if not enrollment:

        conn.close()

        abort(404)

    if (
        enrollment[
            "enrollment_status"
        ] == "INACTIVE"
    ):

        conn.close()

        flash(
            "This class is already inactive.",
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE student_classes

        SET
            enrollment_status = 'INACTIVE',
            ended_at = ?

        WHERE
            student_id = ?

        AND class_code = ?
        """,
        (
            now,
            student_id,
            class_code,
        ),
    )

    conn.commit()

    conn.close()

    flash(
        (
            f"Class {class_code} removed "
            "from the student."
        ),
        "success",
    )

    return redirect(
        url_for(
            "student_profile",

            student_id=
                student_id,
        )
    )


# =========================================================
# REVOKE STUDENT QR
# =========================================================

@app.post(
    "/students/<int:student_id>/qr/revoke"
)
def revoke_student_qr(
    student_id
):

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student:

        conn.close()

        abort(404)

    conn.execute(
        """
        UPDATE students

        SET qr_status = 'REVOKED'

        WHERE id = ?
        """,
        (
            student_id,
        ),
    )

    conn.commit()

    conn.close()

    flash(
        (
            "QR revoked successfully. "
            "The old QR can no longer mark attendance."
        ),
        "success",
    )

    return redirect(
        url_for(
            "student_profile",

            student_id=
                student_id,
        )
    )


# =========================================================
# REGENERATE STUDENT QR
# =========================================================

@app.post(
    "/students/<int:student_id>/qr/regenerate"
)
def regenerate_student_qr(
    student_id
):

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not student:

        conn.close()

        abort(404)

    if (
        student[
            "status"
        ] != "ACTIVE"
    ):

        conn.close()

        flash(
            (
                "The student is inactive. "
                "Activate the student before regenerating QR."
            ),
            "error",
        )

        return redirect(
            url_for(
                "student_profile",

                student_id=
                    student_id,
            )
        )

    token = make_qr_token(
        conn
    )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE students

        SET
            qr_token = ?,
            qr_created_at = ?,
            qr_status = 'ACTIVE'

        WHERE id = ?
        """,
        (
            token,
            now,
            student_id,
        ),
    )

    conn.commit()

    conn.close()

    flash(
        (
            "A new QR was generated successfully. "
            "The previous QR is now invalid."
        ),
        "success",
    )

    return redirect(
        url_for(
            "student_profile",

            student_id=
                student_id,
        )
    )


# =========================================================
# QR PASS
# =========================================================

@app.route(
    "/students/<int:student_id>/qr-pass"
)
def student_qr_pass(
    student_id
):

    conn = get_db()

    row = conn.execute(
        """
        SELECT *

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    if not row:

        conn.close()
        abort(404)

    student = dict(row)

    if not student[
        "qr_token"
    ]:

        token = make_qr_token(
            conn
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute(
            """
            UPDATE students

            SET
                qr_token = ?,
                qr_created_at = ?,
                qr_status = 'ACTIVE'

            WHERE id = ?
            """,
            (
                token,
                now,
                student_id,
            ),
        )

        conn.commit()

        student = dict(
            conn.execute(
                """
                SELECT *

                FROM students

                WHERE id = ?
                """,
                (
                    student_id,
                ),
            ).fetchone()
        )

    classes = get_student_classes(
        conn,
        student_id,
    )

    conn.close()

    return render_template(
        "qr_pass.html",

        student=
            student,

        classes=
            classes,
    )


@app.route(
    "/students/<int:student_id>/qr.png"
)
def qr_image(
    student_id
):

    conn = get_db()

    student = conn.execute(
        """
        SELECT
            qr_token,
            qr_status

        FROM students

        WHERE id = ?
        """,
        (
            student_id,
        ),
    ).fetchone()

    conn.close()

    if not student:
        abort(404)

    if (
        not student["qr_token"]
        or
        student["qr_status"] != "ACTIVE"
    ):
        abort(404)

    qr = qrcode.QRCode(
        version=None,

        error_correction=
            qrcode.constants
            .ERROR_CORRECT_M,

        box_size=10,
        border=4,
    )

    qr.add_data(
        "SSCQR:"
        +
        student[
            "qr_token"
        ]
    )

    qr.make(
        fit=True
    )

    image = qr.make_image(
        fill_color=
            "black",

        back_color=
            "white",
    )

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG",
    )

    buffer.seek(0)

    return send_file(
        buffer,

        mimetype=
            "image/png",
    )


# =========================================================
# SCANNER
# =========================================================

@app.route("/scan")
def scan_qr():

    return render_template(
        "scan.html"
    )


@app.post(
    "/api/qr/verify"
)
def verify_qr():

    data = (
        request.get_json(
            silent=True
        )
        or {}
    )

    raw = str(
        data.get(
            "qr_value",
            ""
        )
    ).strip()

    if not raw.startswith(
        "SSCQR:"
    ):

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Invalid Sri Sumana QR.",
            }
        ), 400

    token = raw[
        len("SSCQR:"):
    ]

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE qr_token = ?
        """,
        (
            token,
        ),
    ).fetchone()

    if not student:

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Unknown QR code.",
            }
        ), 404

    if (
        student["status"] != "ACTIVE"
        or
        student["qr_status"] != "ACTIVE"
    ):

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Student or QR is inactive.",
            }
        ), 403

    fee_month = datetime.now().strftime(
        "%Y-%m"
    )

    month_label = datetime.now().strftime(
        "%B %Y"
    )

    raw_classes = get_student_classes(
        conn,
        student["id"],
    )

    classes = []

    for row in raw_classes:

        item = dict(
            row
        )

        payment = (
            get_student_class_payment_status(
                conn,
                student["id"],
                item["class_code"],
                fee_month,
            )
        )

        item["class_name"] = (
            payment["class_name"]
        )

        item["payment"] = (
            payment
        )

        item["outstanding"] = (
            get_student_class_outstanding_summary(
                conn,
                student["id"],
                item["class_code"],
                fee_month,
            )
        )

        item["payment_status"] = (
            payment["status"]
        )

        item["monthly_fee"] = (
            payment["fee"]
        )

        item["paid"] = (
            payment["paid"]
        )

        item["balance"] = (
            payment["balance"]
        )

        item["payment_text"] = (
            payment_status_text(
                payment
            )
        )

        # Keep the existing scan.html class selector
        # compatible while also showing payment info.
        original_teacher = (
            item.get("teacher")
            or ""
        )

        item["teacher_name"] = (
            original_teacher
        )

        item["teacher"] = (
            original_teacher
            +
            " · "
            +
            item["payment_text"]
        )

        classes.append(
            item
        )

    conn.close()

    return jsonify(
        {
            "success":
                True,

            "student": {
                "id":
                    student["id"],

                "student_code":
                    student["student_code"],

                "full_name":
                    student["full_name"],

                "grade":
                    student["grade"],

                "school":
                    student["school"]
                    or "",
            },

            "classes":
                classes,

            "fee_month":
                fee_month,

            "month_label":
                month_label,

            "token":
                token,
        }
    )


# =========================================================
# ATTENDANCE
# =========================================================

@app.post(
    "/api/attendance/mark"
)
def mark_attendance():

    data = (
        request.get_json(
            silent=True
        )
        or {}
    )

    token = str(
        data.get(
            "token",
            ""
        )
    ).strip()

    class_code = str(
        data.get(
            "class_code",
            ""
        )
    ).strip()

    if (
        not token
        or not class_code
    ):

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "QR token and class are required.",
            }
        ), 400

    conn = get_db()

    student = conn.execute(
        """
        SELECT *

        FROM students

        WHERE qr_token = ?
        """,
        (
            token,
        ),
    ).fetchone()

    if not student:

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Student not found.",
            }
        ), 404

    if (
        student["status"] != "ACTIVE"
        or
        student["qr_status"] != "ACTIVE"
    ):

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Student or QR is inactive.",
            }
        ), 403

    enrollment = conn.execute(
        """
        SELECT
            student_classes.id

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            student_classes.student_id = ?

        AND student_classes.class_code = ?

        AND (
            student_classes.enrollment_status =
                'ACTIVE'

            OR

            student_classes.enrollment_status
                IS NULL
        )

        AND classes.status =
            'ACTIVE'
        """,
        (
            student["id"],
            class_code,
        ),
    ).fetchone()

    if not enrollment:

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    (
                        "Student is not enrolled "
                        "in this active class."
                    ),
            }
        ), 403

    now = datetime.now()

    attendance_date = (
        now.strftime(
            "%Y-%m-%d"
        )
    )

    check_in_time = (
        now.strftime(
            "%H:%M:%S"
        )
    )

    fee_month = now.strftime(
        "%Y-%m"
    )

    month_label = now.strftime(
        "%B %Y"
    )

    payment = (
        get_student_class_payment_status(
            conn,
            student["id"],
            class_code,
            fee_month,
        )
    )

    payment_text = (
        payment_status_text(
            payment
        )
    )

    outstanding_summary = (
        get_student_class_outstanding_summary(
            conn,
            student["id"],
            class_code,
            fee_month,
        )
    )

    existing = conn.execute(
        """
        SELECT *

        FROM attendance

        WHERE
            student_id = ?

        AND class_code = ?

        AND attendance_date = ?
        """,
        (
            student["id"],
            class_code,
            attendance_date,
        ),
    ).fetchone()

    if existing:

        previous_time = (
            existing[
                "check_in_time"
            ]
        )

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "duplicate":
                    True,

                "message":
                    (
                        "Attendance already marked today. "
                        f"Previous check-in: "
                        f"{previous_time}. "
                        f"Payment: {payment_text}"
                    ),

                "check_in_time":
                    previous_time,

                "payment":
                    payment,

                "fee_month":
                    fee_month,

                "month_label":
                    month_label,

                "outstanding":
                    outstanding_summary,
            }
        ), 409

    try:

        conn.execute(
            """
            INSERT INTO attendance
            (
                student_id,
                class_code,
                attendance_date,
                check_in_time,
                status,
                source,
                created_at
            )

            VALUES
            (
                ?, ?, ?, ?,
                'PRESENT',
                'QR',
                ?
            )
            """,
            (
                student["id"],
                class_code,
                attendance_date,
                check_in_time,

                now.strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            ),
        )

        conn.commit()

    except sqlite3.IntegrityError:

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "duplicate":
                    True,

                "message":
                    (
                        "Attendance already marked today. "
                        f"Payment: {payment_text}"
                    ),

                "payment":
                    payment,

                "fee_month":
                    fee_month,

                "month_label":
                    month_label,

                "outstanding":
                    outstanding_summary,
            }
        ), 409

    conn.close()

    return jsonify(
        {
            "success":
                True,

            "message":
                (
                    "Attendance marked successfully. "
                    f"Payment: {payment_text}"
                ),

            "attendance_date":
                attendance_date,

            "check_in_time":
                check_in_time,

            "class_code":
                class_code,

            "student": {
                "student_code":
                    student[
                        "student_code"
                    ],

                "full_name":
                    student[
                        "full_name"
                    ],
            },

            "payment":
                payment,

            "fee_month":
                fee_month,

            "month_label":
                month_label,

            "outstanding":
                outstanding_summary,
        }
    )


@app.route("/attendance")
def attendance_list():

    conn = get_db()

    rows = conn.execute(
        """
        SELECT
            attendance.*,

            students.student_code,

            students.full_name

        FROM attendance

        JOIN students

        ON students.id =
            attendance.student_id

        ORDER BY
            attendance.id DESC

        LIMIT 200
        """
    ).fetchall()

    conn.close()

    return render_template(
        "attendance.html",

        records=
            rows,
    )


# =========================================================
# TODAY ATTENDANCE
# =========================================================

@app.route("/attendance/today")
def attendance_today():

    now = datetime.now()

    today_date = now.strftime(
        "%Y-%m-%d"
    )

    today_day_index = (
        now.weekday()
    )

    today_day_name = (
        DAY_NAMES[
            today_day_index
        ]
    )

    conn = get_db()

    class_rows = conn.execute(
        """
        SELECT
            classes.id,
            classes.class_code,
            classes.class_name,
            classes.grade,
            classes.subject,
            classes.teacher_id,
            classes.day_of_week,
            classes.start_time,
            classes.end_time,
            classes.room,
            classes.status,

            teachers.full_name
                AS teacher_name,

            (
                SELECT COUNT(
                    DISTINCT
                    student_classes.student_id
                )

                FROM student_classes

                JOIN students
                    ON students.id =
                        student_classes.student_id

                WHERE
                    student_classes.class_code =
                        classes.class_code

                AND (
                    student_classes.enrollment_status =
                        'ACTIVE'

                    OR

                    student_classes.enrollment_status
                        IS NULL
                )

                AND students.status =
                    'ACTIVE'
            )
            AS enrolled_count,

            (
                SELECT COUNT(
                    DISTINCT
                    attendance.student_id
                )

                FROM attendance

                WHERE
                    attendance.class_code =
                        classes.class_code

                AND attendance.attendance_date = ?

                AND attendance.status =
                    'PRESENT'
            )
            AS present_count

        FROM classes

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            classes.status =
                'ACTIVE'

        AND classes.day_of_week = ?

        ORDER BY
            classes.start_time,
            classes.class_name
        """,
        (
            today_date,
            today_day_index,
        ),
    ).fetchall()

    classes = []

    total_expected = 0
    total_present = 0

    for row in class_rows:

        item = dict(
            row
        )

        item["start_display"] = (
            format_time(
                item["start_time"]
            )
        )

        item["end_display"] = (
            format_time(
                item["end_time"]
            )
        )

        enrolled_count = int(
            item["enrolled_count"]
            or 0
        )

        present_count = int(
            item["present_count"]
            or 0
        )

        not_marked = max(
            enrolled_count
            -
            present_count,
            0,
        )

        if enrolled_count > 0:

            attendance_rate = round(
                (
                    present_count
                    /
                    enrolled_count
                )
                *
                100,
                1,
            )

        else:

            attendance_rate = 0.0

        item["enrolled_count"] = (
            enrolled_count
        )

        item["present_count"] = (
            present_count
        )

        item["not_marked"] = (
            not_marked
        )

        item["attendance_rate"] = (
            attendance_rate
        )

        classes.append(
            item
        )

        total_expected += (
            enrolled_count
        )

        total_present += (
            present_count
        )

    total_not_marked = max(
        total_expected
        -
        total_present,
        0,
    )

    if total_expected > 0:

        overall_rate = round(
            (
                total_present
                /
                total_expected
            )
            *
            100,
                1,
        )

    else:

        overall_rate = 0.0

    attendance_rows = conn.execute(
        """
        SELECT
            attendance.id,
            attendance.student_id,
            attendance.class_code,
            attendance.attendance_date,
            attendance.check_in_time,
            attendance.status,
            attendance.source,

            students.student_code,
            students.full_name,

            classes.class_name,
            classes.subject

        FROM attendance

        JOIN students
            ON students.id =
                attendance.student_id

        LEFT JOIN classes
            ON classes.class_code =
                attendance.class_code

        WHERE
            attendance.attendance_date = ?

        ORDER BY
            attendance.check_in_time DESC,
            attendance.id DESC

        LIMIT 100
        """,
        (
            today_date,
        ),
    ).fetchall()

    recent_records = [
        dict(row)
        for row in attendance_rows
    ]

    conn.close()

    totals = {
        "classes":
            len(classes),

        "expected":
            total_expected,

        "present":
            total_present,

        "not_marked":
            total_not_marked,

        "rate":
            overall_rate,
    }

    return render_template(
        "attendance_today.html",

        classes=
            classes,

        recent_records=
            recent_records,

        today_date=
            today_date,

        today_day=
            today_day_name,

        totals=
            totals,
    )


# =========================================================
# CLASS ATTENDANCE
# =========================================================

@app.route("/attendance/classes")
def class_attendance():

    selected_date = (
        request.args.get(
            "date",
            "",
        ).strip()
    )

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    if not selected_date:
        selected_date = today

    try:
        selected_date_object = datetime.strptime(
            selected_date,
            "%Y-%m-%d",
        )

    except ValueError:
        selected_date = today

        selected_date_object = datetime.strptime(
            selected_date,
            "%Y-%m-%d",
        )

    selected_day_index = (
        selected_date_object.weekday()
    )

    conn = get_db()

    rows = conn.execute(
        """
        SELECT
            classes.id,
            classes.class_code,
            classes.class_name,
            classes.grade,
            classes.subject,
            classes.teacher_id,
            classes.day_of_week,
            classes.start_time,
            classes.end_time,
            classes.room,
            classes.status,

            teachers.full_name
                AS teacher_name,

            (
                SELECT COUNT(
                    DISTINCT
                    student_classes.student_id
                )

                FROM student_classes

                JOIN students
                    ON students.id =
                        student_classes.student_id

                WHERE
                    student_classes.class_code =
                        classes.class_code

                AND (
                    student_classes.enrollment_status =
                        'ACTIVE'

                    OR

                    student_classes.enrollment_status
                        IS NULL
                )

                AND students.status =
                    'ACTIVE'
            )
            AS enrolled_count,

            (
                SELECT COUNT(
                    DISTINCT
                    attendance.student_id
                )

                FROM attendance

                WHERE
                    attendance.class_code =
                        classes.class_code

                AND attendance.attendance_date = ?

                AND attendance.status =
                    'PRESENT'
            )
            AS present_count

        FROM classes

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            classes.status =
                'ACTIVE'

        ORDER BY
            CASE
                WHEN classes.day_of_week = ?
                THEN 0
                ELSE 1
            END,
            classes.day_of_week,
            classes.start_time,
            classes.class_name
        """,
        (
            selected_date,
            selected_day_index,
        ),
    ).fetchall()

    class_list = []

    total_students = 0
    total_present = 0
    scheduled_classes = 0

    for row in rows:

        item = dict(
            row
        )

        item["day_name"] = (
            DAY_NAMES[
                item["day_of_week"]
            ]
        )

        item["start_display"] = (
            format_time(
                item["start_time"]
            )
        )

        item["end_display"] = (
            format_time(
                item["end_time"]
            )
        )

        item["is_scheduled"] = (
            item["day_of_week"]
            ==
            selected_day_index
        )

        enrolled_count = int(
            item["enrolled_count"]
            or 0
        )

        present_count = int(
            item["present_count"]
            or 0
        )

        absent_count = max(
            enrolled_count
            -
            present_count,
            0,
        )

        if enrolled_count > 0:

            attendance_rate = round(
                (
                    present_count
                    /
                    enrolled_count
                )
                *
                100,
                1,
            )

        else:

            attendance_rate = 0.0

        item["enrolled_count"] = (
            enrolled_count
        )

        item["present_count"] = (
            present_count
        )

        item["absent_count"] = (
            absent_count
        )

        item["attendance_rate"] = (
            attendance_rate
        )

        if item["is_scheduled"]:

            scheduled_classes += 1

            total_students += (
                enrolled_count
            )

            total_present += (
                present_count
            )

        class_list.append(
            item
        )

    conn.close()

    total_absent = max(
        total_students
        -
        total_present,
        0,
    )

    if total_students > 0:

        overall_rate = round(
            (
                total_present
                /
                total_students
            )
            *
            100,
            1,
        )

    else:

        overall_rate = 0.0

    totals = {
        "scheduled_classes":
            scheduled_classes,

        "students":
            total_students,

        "present":
            total_present,

        "absent":
            total_absent,

        "rate":
            overall_rate,
    }

    return render_template(
        "attendance_classes.html",

        classes=
            class_list,

        selected_date=
            selected_date,

        today=
            today,

        selected_day=
            DAY_NAMES[
                selected_day_index
            ],

        totals=
            totals,
    )


# =========================================================
# TIMETABLE
# =========================================================

@app.route("/timetable")
def timetable():

    conn = get_db()

    rows = conn.execute(
        """
        SELECT
            classes.*,

            teachers.full_name
                AS teacher_name

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            classes.status = 'ACTIVE'

        ORDER BY
            classes.day_of_week,
            classes.start_time
        """
    ).fetchall()

    conn.close()

    schedule = {
        day: []
        for day in DAY_NAMES
    }

    now = datetime.now()

    for row in rows:

        item = dict(row)

        item[
            "start_display"
        ] = format_time(
            row["start_time"]
        )

        item[
            "end_display"
        ] = format_time(
            row["end_time"]
        )

        item[
            "status_label"
        ] = class_status(
            row,
            now,
        )

        schedule[
            DAY_NAMES[
                row[
                    "day_of_week"
                ]
            ]
        ].append(
            item
        )

    return render_template(
        "timetable.html",

        schedule=
            schedule,

        current_day=
            now.strftime(
                "%A"
            ),
    )


# =========================================================
# TEACHER MANAGEMENT LIST
# =========================================================

@app.route("/teachers")
def teachers():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    conn = get_db()

    rows = conn.execute(
        """
        SELECT *

        FROM teachers

        ORDER BY
            CASE
                WHEN status = 'ACTIVE'
                THEN 0
                ELSE 1
            END,

            full_name
        """
    ).fetchall()

    result = []

    for teacher in rows:

        item = dict(
            teacher
        )

        active_classes = conn.execute(
            """
            SELECT COUNT(*)

            FROM classes

            WHERE teacher_id = ?

            AND status = 'ACTIVE'
            """,
            (
                teacher["id"],
            ),
        ).fetchone()[0]

        all_classes = conn.execute(
            """
            SELECT COUNT(*)

            FROM classes

            WHERE teacher_id = ?
            """,
            (
                teacher["id"],
            ),
        ).fetchone()[0]

        student_count = conn.execute(
            """
            SELECT COUNT(
                DISTINCT
                student_classes.student_id
            )

            FROM student_classes

            JOIN classes

            ON classes.class_code =
                student_classes.class_code

            WHERE
                classes.teacher_id = ?

            AND (
                student_classes.enrollment_status = 'ACTIVE'

                OR student_classes.enrollment_status IS NULL
            )

            AND classes.status = 'ACTIVE'
            """,
            (
                teacher["id"],
            ),
        ).fetchone()[0]

        collected = conn.execute(
            """
            SELECT COALESCE(
                SUM(payments.amount),
                0
            )

            FROM payments

            JOIN classes

            ON classes.class_code =
                payments.class_code

            WHERE
                classes.teacher_id = ?

            AND payments.fee_month = ?

            AND payments.status = 'CONFIRMED'
            """,
            (
                teacher["id"],
                selected_month,
            ),
        ).fetchone()[0]

        teacher_income = conn.execute(
            """
            SELECT COALESCE(
                SUM(
                    payments.amount
                    *
                    classes.teacher_share_percent
                    /
                    100
                ),
                0
            )

            FROM payments

            JOIN classes

            ON classes.class_code =
                payments.class_code

            WHERE
                classes.teacher_id = ?

            AND payments.fee_month = ?

            AND payments.status = 'CONFIRMED'
            """,
            (
                teacher["id"],
                selected_month,
            ),
        ).fetchone()[0]

        item[
            "active_classes"
        ] = active_classes

        item[
            "all_classes"
        ] = all_classes

        item[
            "total_students"
        ] = student_count

        item[
            "collected"
        ] = float(
            collected
            or 0
        )

        item[
            "teacher_income"
        ] = float(
            teacher_income
            or 0
        )

        result.append(
            item
        )

    conn.close()

    return render_template(
        "teachers.html",

        teachers=
            result,

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m"
            ).strftime(
                "%B %Y"
            ),
    )


# =========================================================
# ADD TEACHER
# =========================================================

@app.route(
    "/teachers/new",
    methods=[
        "GET",
        "POST",
    ],
)
def new_teacher():

    if request.method == "POST":

        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        subject = request.form.get(
            "subject",
            "",
        ).strip()

        phone = request.form.get(
            "phone",
            "",
        ).strip()

        if not full_name:

            flash(
                "Teacher name is required.",
                "error",
            )

            return render_template(
                "teacher_form.html",

                mode=
                    "new",

                teacher=
                    None,
            )

        conn = get_db()

        teacher_code = (
            make_teacher_code(
                conn
            )
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute(
            """
            INSERT INTO teachers
            (
                teacher_code,
                full_name,
                subject,
                phone,
                status,
                created_at
            )

            VALUES
            (
                ?, ?, ?, ?,
                'ACTIVE',
                ?
            )
            """,
            (
                teacher_code,
                full_name,
                subject,
                phone,
                now,
            ),
        )

        conn.commit()
        conn.close()

        flash(
            (
                f"Teacher {teacher_code} "
                "added successfully."
            ),
            "success",
        )

        return redirect(
            url_for(
                "teachers"
            )
        )

    return render_template(
        "teacher_form.html",

        mode=
            "new",

        teacher=
            None,
    )


# =========================================================
# EDIT TEACHER
# =========================================================

@app.route(
    "/teachers/<int:teacher_id>/edit",
    methods=[
        "GET",
        "POST",
    ],
)
def edit_teacher(
    teacher_id
):

    conn = get_db()

    teacher_row = conn.execute(
        """
        SELECT *

        FROM teachers

        WHERE id = ?
        """,
        (
            teacher_id,
        ),
    ).fetchone()

    if not teacher_row:

        conn.close()
        abort(404)

    if request.method == "POST":

        full_name = request.form.get(
            "full_name",
            "",
        ).strip()

        subject = request.form.get(
            "subject",
            "",
        ).strip()

        phone = request.form.get(
            "phone",
            "",
        ).strip()

        if not full_name:

            conn.close()

            flash(
                "Teacher name is required.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_teacher",

                    teacher_id=
                        teacher_id,
                )
            )

        conn.execute(
            """
            UPDATE teachers

            SET
                full_name = ?,
                subject = ?,
                phone = ?

            WHERE id = ?
            """,
            (
                full_name,
                subject,
                phone,
                teacher_id,
            ),
        )

        class_rows = conn.execute(
            """
            SELECT class_code

            FROM classes

            WHERE teacher_id = ?
            """,
            (
                teacher_id,
            ),
        ).fetchall()

        for class_row in class_rows:

            conn.execute(
                """
                UPDATE student_classes

                SET teacher = ?

                WHERE class_code = ?
                """,
                (
                    full_name,

                    class_row[
                        "class_code"
                    ],
                ),
            )

        conn.commit()
        conn.close()

        flash(
            "Teacher updated successfully.",
            "success",
        )

        return redirect(
            url_for(
                "teachers"
            )
        )

    teacher = dict(
        teacher_row
    )

    conn.close()

    return render_template(
        "teacher_form.html",

        mode=
            "edit",

        teacher=
            teacher,
    )


# =========================================================
# ACTIVATE / DEACTIVATE TEACHER
# =========================================================

@app.post(
    "/teachers/<int:teacher_id>/toggle"
)
def toggle_teacher(
    teacher_id
):

    conn = get_db()

    teacher = conn.execute(
        """
        SELECT *

        FROM teachers

        WHERE id = ?
        """,
        (
            teacher_id,
        ),
    ).fetchone()

    if not teacher:

        conn.close()
        abort(404)

    current_status = (
        teacher[
            "status"
        ]
    )

    if current_status == "ACTIVE":

        active_class_count = conn.execute(
            """
            SELECT COUNT(*)

            FROM classes

            WHERE
                teacher_id = ?

            AND status = 'ACTIVE'
            """,
            (
                teacher_id,
            ),
        ).fetchone()[0]

        if active_class_count > 0:

            conn.close()

            flash(
                (
                    "This teacher still has "
                    f"{active_class_count} active class(es). "
                    "Deactivate or reassign those classes first."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "teachers"
                )
            )

        new_status = (
            "INACTIVE"
        )

    else:

        new_status = (
            "ACTIVE"
        )

    conn.execute(
        """
        UPDATE teachers

        SET status = ?

        WHERE id = ?
        """,
        (
            new_status,
            teacher_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        (
            f"{teacher['full_name']} "
            f"is now {new_status}."
        ),
        "success",
    )

    return redirect(
        url_for(
            "teachers"
        )
    )


# =========================================================
# TEACHER PROFILE
# =========================================================

@app.route(
    "/teachers/<int:teacher_id>"
)
def teacher_profile(
    teacher_id
):

    allowed_tabs = {
        "overview",
        "classes",
        "students",
        "income",
        "performance",
        "activity",
    }

    active_tab = request.args.get(
        "tab",
        "overview",
    ).strip().lower()

    if active_tab not in allowed_tabs:
        active_tab = "overview"

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    month_start, month_end = (
        month_bounds(
            selected_month
        )
    )

    conn = get_db()

    teacher = conn.execute(
        """
        SELECT *

        FROM teachers

        WHERE id = ?
        """,
        (
            teacher_id,
        ),
    ).fetchone()

    if not teacher:

        conn.close()
        abort(404)

    class_rows = conn.execute(
        """
        SELECT *

        FROM classes

        WHERE teacher_id = ?

        ORDER BY
            day_of_week,
            start_time
        """,
        (
            teacher_id,
        ),
    ).fetchall()

    class_stats = []

    total_expected = 0
    total_collected = 0
    total_outstanding = 0
    total_teacher_share = 0
    total_center_share = 0

    for class_row in class_rows:

        class_code = (
            class_row[
                "class_code"
            ]
        )

        start_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND created_at < ?

            AND (
                ended_at IS NULL

                OR ended_at >= ?
            )
            """,
            (
                class_code,
                month_start,
                month_start,
            ),
        ).fetchone()[0]

        new_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND created_at
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start,
                month_end,
            ),
        ).fetchone()[0]

        left_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND ended_at
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start,
                month_end,
            ),
        ).fetchone()[0]

        end_students = (
            count_students_at_month_end(
                conn,
                class_code,
                selected_month,
            )
        )

        growth = (
            end_students
            -
            start_students
        )

        monthly_fee = float(
            class_row[
                "monthly_fee"
            ]
            or 0
        )

        gross_expected = (
            end_students
            *
            monthly_fee
        )

        free_card_summary = get_class_free_card_summary(
            conn,
            class_code,
            selected_month,
        )

        waived_amount = min(
            free_card_summary[
                "waived_amount"
            ],
            gross_expected,
        )

        expected = max(
            gross_expected
            - waived_amount,
            0,
        )

        collected = conn.execute(
            """
            SELECT COALESCE(
                SUM(amount),
                0
            )

            FROM payments

            WHERE
                class_code = ?

            AND fee_month = ?

            AND status = 'CONFIRMED'
            """,
            (
                class_code,
                selected_month,
            ),
        ).fetchone()[0]

        collected = float(
            collected
            or 0
        )

        outstanding = max(
            expected
            -
            collected,
            0,
        )

        share_percent = float(
            class_row[
                "teacher_share_percent"
            ]
            or 0
        )

        teacher_share = (
            collected
            *
            share_percent
            /
            100
        )

        center_share = (
            collected
            -
            teacher_share
        )

        attendance_info = conn.execute(
            """
            SELECT
                COUNT(*)
                    AS total_records,

                COUNT(
                    DISTINCT attendance_date
                )
                    AS class_days

            FROM attendance

            WHERE
                class_code = ?

            AND attendance_date
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start[:10],
                month_end[:10],
            ),
        ).fetchone()

        class_days = (
            attendance_info[
                "class_days"
            ]
            or 0
        )

        total_records = (
            attendance_info[
                "total_records"
            ]
            or 0
        )

        average_attendance = (
            round(
                total_records
                /
                class_days,
                1,
            )

            if class_days > 0

            else 0
        )

        attendance_rate = 0

        if (
            class_days > 0
            and
            end_students > 0
        ):

            attendance_rate = round(
                (
                    total_records

                    /

                    (
                        class_days
                        *
                        end_students
                    )
                )
                *
                100,
                1,
            )

        attendance_rate = min(
            attendance_rate,
            100,
        )

        class_stats.append(
            {
                "class_code":
                    class_code,

                "class_name":
                    class_row[
                        "class_name"
                    ],

                "class_status":
                    class_row[
                        "status"
                    ],

                "monthly_fee":
                    monthly_fee,

                "teacher_share_percent":
                    share_percent,

                "day":
                    DAY_NAMES[
                        class_row[
                            "day_of_week"
                        ]
                    ],

                "start_display":
                    format_time(
                        class_row[
                            "start_time"
                        ]
                    ),

                "end_display":
                    format_time(
                        class_row[
                            "end_time"
                        ]
                    ),

                "start_students":
                    start_students,

                "new_students":
                    new_students,

                "left_students":
                    left_students,

                "end_students":
                    end_students,

                "growth":
                    growth,

                "gross_expected":
                    gross_expected,

                "free_card_count":
                    free_card_summary[
                        "free_card_count"
                    ],

                "waived_amount":
                    waived_amount,

                "expected":
                    expected,

                "collected":
                    collected,

                "outstanding":
                    outstanding,

                "teacher_share":
                    teacher_share,

                "center_share":
                    center_share,

                "class_days":
                    class_days,

                "average_attendance":
                    average_attendance,

                "attendance_rate":
                    attendance_rate,
            }
        )

        total_expected += (
            expected
        )

        total_collected += (
            collected
        )

        total_outstanding += (
            outstanding
        )

        total_teacher_share += (
            teacher_share
        )

        total_center_share += (
            center_share
        )


    teacher_start_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT
            student_classes.student_id
        )

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            classes.teacher_id = ?

        AND student_classes.created_at < ?

        AND (
            student_classes.ended_at IS NULL

            OR student_classes.ended_at >= ?
        )
        """,
        (
            teacher_id,
            month_start,
            month_start,
        ),
    ).fetchone()[0]


    teacher_end_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT
            student_classes.student_id
        )

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            classes.teacher_id = ?

        AND student_classes.created_at <= ?

        AND (
            student_classes.ended_at IS NULL

            OR student_classes.ended_at > ?
        )
        """,
        (
            teacher_id,
            month_end,
            month_end,
        ),
    ).fetchone()[0]


    teacher_new_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT
            student_classes.student_id
        )

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            classes.teacher_id = ?

        AND student_classes.created_at
            BETWEEN ? AND ?
        """,
        (
            teacher_id,
            month_start,
            month_end,
        ),
    ).fetchone()[0]


    teacher_left_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT
            student_classes.student_id
        )

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            classes.teacher_id = ?

        AND student_classes.ended_at
            BETWEEN ? AND ?
        """,
        (
            teacher_id,
            month_start,
            month_end,
        ),
    ).fetchone()[0]


    teacher_growth = (
        teacher_end_students
        -
        teacher_start_students
    )

    teacher_growth_percent = 0

    if teacher_start_students > 0:

        teacher_growth_percent = round(
            (
                teacher_growth
                /
                teacher_start_students
            )
            *
            100,
            1,
        )


    growth_history = []

    for offset in range(
        -5,
        1,
    ):

        history_month = (
            shift_month(
                selected_month,
                offset,
            )
        )

        _, history_end = (
            month_bounds(
                history_month
            )
        )

        history_count = conn.execute(
            """
            SELECT COUNT(
                DISTINCT
                student_classes.student_id
            )

            FROM student_classes

            JOIN classes

            ON classes.class_code =
                student_classes.class_code

            WHERE
                classes.teacher_id = ?

            AND student_classes.created_at <= ?

            AND (
                student_classes.ended_at IS NULL

                OR student_classes.ended_at > ?
            )
            """,
            (
                teacher_id,
                history_end,
                history_end,
            ),
        ).fetchone()[0]

        growth_history.append(
            {
                "month":
                    datetime.strptime(
                        history_month,
                        "%Y-%m"
                    ).strftime(
                        "%b"
                    ),

                "month_key":
                    history_month,

                "count":
                    history_count,
            }
        )

    max_growth_count = max(
        [
            item["count"]
            for item in growth_history
        ]
        or [0]
    )

    for item in growth_history:

        if max_growth_count > 0:

            item[
                "percent"
            ] = max(
                round(
                    item["count"]
                    /
                    max_growth_count
                    *
                    100
                ),
                3,
            )

        else:

            item[
                "percent"
            ] = 0

    # =====================================================
    # TEACHER STUDENTS FOR SELECTED MONTH
    # =====================================================

    student_rows = conn.execute(
        """
        SELECT
            students.id AS student_id,
            students.student_code,
            students.full_name,
            students.grade,
            students.school,
            students.phone,
            students.status AS student_status,
            GROUP_CONCAT(
                DISTINCT classes.class_code
            ) AS class_codes,
            GROUP_CONCAT(
                DISTINCT classes.class_name
            ) AS class_names

        FROM student_classes

        JOIN students
            ON students.id =
                student_classes.student_id

        JOIN classes
            ON classes.class_code =
                student_classes.class_code

        WHERE
            classes.teacher_id = ?

        AND student_classes.created_at <= ?

        AND (
            student_classes.ended_at IS NULL
            OR student_classes.ended_at > ?
        )

        GROUP BY
            students.id,
            students.student_code,
            students.full_name,
            students.grade,
            students.school,
            students.phone,
            students.status

        ORDER BY
            students.full_name
        """,
        (
            teacher_id,
            month_end,
            month_end,
        ),
    ).fetchall()

    teacher_students = [
        dict(row)
        for row in student_rows
    ]

    # =====================================================
    # PAYMENT / INCOME TRANSACTIONS
    # =====================================================

    payment_rows = conn.execute(
        """
        SELECT
            payments.id,
            payments.receipt_no,
            payments.student_id,
            payments.class_code,
            payments.fee_month,
            payments.amount,
            payments.payment_method,
            payments.payment_date,
            payments.status,
            students.student_code,
            students.full_name,
            classes.class_name,
            classes.teacher_share_percent

        FROM payments

        JOIN students
            ON students.id =
                payments.student_id

        JOIN classes
            ON classes.class_code =
                payments.class_code

        WHERE
            classes.teacher_id = ?

        AND payments.fee_month = ?

        ORDER BY
            payments.payment_date DESC,
            payments.id DESC
        """,
        (
            teacher_id,
            selected_month,
        ),
    ).fetchall()

    payment_records = []

    for row in payment_rows:
        item = dict(row)

        amount = float(
            item.get("amount")
            or 0
        )

        share_percent = float(
            item.get(
                "teacher_share_percent"
            )
            or 0
        )

        if item.get("status") == "CONFIRMED":
            item["teacher_income"] = (
                amount
                * share_percent
                / 100
            )
            item["center_income"] = (
                amount
                - item["teacher_income"]
            )
        else:
            item["teacher_income"] = 0.0
            item["center_income"] = 0.0

        payment_records.append(
            item
        )

    # =====================================================
    # TEACHER AUDIT ACTIVITY - OWNER ONLY
    # =====================================================

    activity_records = []

    if (
        g.current_user
        and g.current_user["role"] == "OWNER"
    ):
        activity_rows = conn.execute(
            """
            SELECT
                id,
                actor_username,
                actor_full_name,
                actor_role,
                action,
                entity_type,
                entity_id,
                status_code,
                created_at

            FROM audit_logs

            WHERE
                entity_type = 'TEACHER'

            AND entity_id = ?

            ORDER BY
                id DESC

            LIMIT 100
            """,
            (
                str(teacher_id),
            ),
        ).fetchall()

        activity_records = [
            dict(row)
            for row in activity_rows
        ]

    active_class_count = sum(
        1
        for item in class_stats
        if item["class_status"] == "ACTIVE"
    )

    total_attendance_records = sum(
        item["class_days"]
        * item["average_attendance"]
        for item in class_stats
    )

    average_attendance_rate = (
        round(
            sum(
                item["attendance_rate"]
                for item in class_stats
            )
            / len(class_stats),
            1,
        )
        if class_stats
        else 0
    )

    conn.close()

    return render_template(
        "teacher_profile.html",

        teacher=
            teacher,

        classes=
            class_stats,

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m"
            ).strftime(
                "%B %Y"
            ),

        previous_month=
            shift_month(
                selected_month,
                -1,
            ),

        next_month=
            shift_month(
                selected_month,
                1,
            ),

        total_expected=
            total_expected,

        total_collected=
            total_collected,

        total_outstanding=
            total_outstanding,

        total_teacher_share=
            total_teacher_share,

        total_center_share=
            total_center_share,

        start_students=
            teacher_start_students,

        end_students=
            teacher_end_students,

        new_students=
            teacher_new_students,

        left_students=
            teacher_left_students,

        growth=
            teacher_growth,

        growth_percent=
            teacher_growth_percent,

        growth_history=
            growth_history,

        active_tab=
            active_tab,

        teacher_students=
            teacher_students,

        payment_records=
            payment_records,

        activity_records=
            activity_records,

        active_class_count=
            active_class_count,

        total_attendance_records=
            total_attendance_records,

        average_attendance_rate=
            average_attendance_rate,
    )


# =========================================================
# FREE CARD MANAGEMENT
# =========================================================

def load_free_card_record(
    conn,
    free_card_id,
):
    return conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name AS student_name,
            students.grade,
            students.school,
            classes.class_name,
            classes.subject,
            classes.monthly_fee,
            teachers.full_name AS teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE free_cards.id = ?
        """,
        (
            free_card_id,
        ),
    ).fetchone()


@app.route("/free-cards/teacher-requests")
def teacher_free_card_approvals():
    """Owner-only queue for Free Card requests submitted by teachers."""
    conn = get_db()

    pending_rows = conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name AS student_name,
            students.grade,
            students.school,
            classes.class_name,
            classes.subject,
            classes.monthly_fee,
            teachers.full_name AS class_teacher_name,
            COALESCE(
                free_cards.requested_by_teacher_name,
                teachers.full_name
            ) AS request_teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE free_cards.status = 'PENDING'
        AND free_cards.request_source = 'TEACHER'
        ORDER BY free_cards.requested_at ASC, free_cards.id ASC
        """
    ).fetchall()

    recent_rows = conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name AS student_name,
            classes.class_name,
            COALESCE(
                free_cards.requested_by_teacher_name,
                teachers.full_name
            ) AS request_teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE free_cards.request_source = 'TEACHER'
        AND free_cards.status IN ('APPROVED', 'REJECTED')
        ORDER BY COALESCE(free_cards.reviewed_at, free_cards.requested_at) DESC
        LIMIT 10
        """
    ).fetchall()

    conn.close()

    return render_template(
        "teacher_free_card_approvals.html",
        pending_requests=pending_rows,
        recent_requests=recent_rows,
    )


@app.route("/free-cards")
def free_cards():
    selected_month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    if not valid_month(
        selected_month
    ):
        selected_month = datetime.now().strftime(
            "%Y-%m"
        )

    status_filter = request.args.get(
        "status",
        "ALL",
    ).strip().upper()

    allowed_statuses = {
        "ALL",
        "PENDING",
        "APPROVED",
        "REJECTED",
        "REVOKED",
    }

    if status_filter not in allowed_statuses:
        status_filter = "ALL"

    conn = get_db()

    where_parts = [
        "free_cards.fee_month = ?"
    ]
    params = [
        selected_month
    ]

    if g.current_user["role"] == "STAFF":
        where_parts.append(
            "free_cards.requested_by_user_id = ?"
        )
        params.append(
            g.current_user["id"]
        )

    if status_filter != "ALL":
        where_parts.append(
            "free_cards.status = ?"
        )
        params.append(
            status_filter
        )

    where_sql = " AND ".join(
        where_parts
    )

    rows = conn.execute(
        f"""
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name AS student_name,
            classes.class_name,
            classes.subject,
            classes.monthly_fee,
            teachers.full_name AS teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE {where_sql}
        ORDER BY free_cards.id DESC
        """,
        params,
    ).fetchall()

    stats_where = [
        "fee_month = ?"
    ]
    stats_params = [
        selected_month
    ]

    if g.current_user["role"] == "STAFF":
        stats_where.append(
            "requested_by_user_id = ?"
        )
        stats_params.append(
            g.current_user["id"]
        )

    stats_sql = " AND ".join(
        stats_where
    )

    stats = conn.execute(
        f"""
        SELECT
            SUM(
                CASE WHEN status = 'PENDING'
                THEN 1 ELSE 0 END
            ) AS pending_count,
            SUM(
                CASE WHEN status = 'APPROVED'
                THEN 1 ELSE 0 END
            ) AS approved_count,
            SUM(
                CASE WHEN status = 'REJECTED'
                THEN 1 ELSE 0 END
            ) AS rejected_count,
            SUM(
                CASE WHEN status = 'REVOKED'
                THEN 1 ELSE 0 END
            ) AS revoked_count,
            COALESCE(
                SUM(
                    CASE WHEN status = 'APPROVED'
                    THEN waived_amount ELSE 0 END
                ),
                0
            ) AS waived_amount
        FROM free_cards
        WHERE {stats_sql}
        """,
        stats_params,
    ).fetchone()

    conn.close()

    return render_template(
        "free_cards.html",
        free_cards=rows,
        selected_month=selected_month,
        month_label=datetime.strptime(
            selected_month,
            "%Y-%m",
        ).strftime("%B %Y"),
        status_filter=status_filter,
        stats=stats,
    )


@app.route(
    "/free-cards/new",
    methods=["GET", "POST"],
)
def new_free_card_request():
    conn = get_db()

    if request.method == "POST":
        student_id = request.form.get(
            "student_id",
            "",
        ).strip()

        class_code = request.form.get(
            "class_code",
            "",
        ).strip()

        fee_month = request.form.get(
            "fee_month",
            "",
        ).strip()

        reason = request.form.get(
            "reason",
            "",
        ).strip()

        if (
            not student_id.isdigit()
            or not class_code
            or not valid_month(fee_month)
            or len(reason) < 3
        ):
            conn.close()
            flash(
                "Student, class, month and reason are required.",
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        if len(reason) > 500:
            conn.close()
            flash(
                "Reason must be 500 characters or less.",
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        student = conn.execute(
            """
            SELECT *
            FROM students
            WHERE id = ?
            AND status = 'ACTIVE'
            """,
            (
                int(student_id),
            ),
        ).fetchone()

        if not student:
            conn.close()
            flash(
                "Active student not found.",
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        class_row = conn.execute(
            """
            SELECT *
            FROM classes
            WHERE class_code = ?
            """,
            (
                class_code,
            ),
        ).fetchone()

        if not class_row:
            conn.close()
            flash(
                "Class not found.",
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        month_start, month_end = month_bounds(
            fee_month
        )

        enrollment = conn.execute(
            """
            SELECT id
            FROM student_classes
            WHERE student_id = ?
            AND class_code = ?
            AND created_at <= ?
            AND (
                ended_at IS NULL
                OR ended_at > ?
            )
            """,
            (
                int(student_id),
                class_code,
                month_end,
                month_start,
            ),
        ).fetchone()

        if not enrollment:
            conn.close()
            flash(
                "Student is not enrolled in this class for the selected month.",
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        paid = conn.execute(
            """
            SELECT COALESCE(SUM(amount), 0)
            FROM payments
            WHERE student_id = ?
            AND class_code = ?
            AND fee_month = ?
            AND status = 'CONFIRMED'
            """,
            (
                int(student_id),
                class_code,
                fee_month,
            ),
        ).fetchone()[0]

        if float(paid or 0) > 0:
            conn.close()
            flash(
                (
                    "A confirmed payment already exists for this "
                    "student/class/month. Void or resolve that "
                    "payment before requesting a full Free Card."
                ),
                "error",
            )
            return redirect(
                url_for(
                    "new_free_card_request"
                )
            )

        duplicate = conn.execute(
            """
            SELECT id, status
            FROM free_cards
            WHERE student_id = ?
            AND class_code = ?
            AND fee_month = ?
            AND status IN ('PENDING', 'APPROVED')
            LIMIT 1
            """,
            (
                int(student_id),
                class_code,
                fee_month,
            ),
        ).fetchone()

        if duplicate:
            conn.close()
            flash(
                (
                    "A Pending or Approved Free Card already exists "
                    "for this student, class and month."
                ),
                "error",
            )
            return redirect(
                url_for(
                    "free_card_detail",
                    free_card_id=duplicate["id"],
                )
            )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        try:
            cursor = conn.execute(
                """
                INSERT INTO free_cards
                (
                    student_id,
                    class_code,
                    fee_month,
                    reason,
                    status,
                    waived_amount,
                    requested_by_user_id,
                    requested_by_name,
                    requested_at
                )
                VALUES
                (
                    ?, ?, ?, ?,
                    'PENDING',
                    0,
                    ?, ?, ?
                )
                """,
                (
                    int(student_id),
                    class_code,
                    fee_month,
                    reason,
                    g.current_user["id"],
                    g.current_user["full_name"],
                    now,
                ),
            )

            free_card_id = cursor.lastrowid
            conn.commit()

        except sqlite3.IntegrityError:
            conn.rollback()
            conn.close()
            flash(
                "A Free Card request already exists for this period.",
                "error",
            )
            return redirect(
                url_for(
                    "free_cards",
                    month=fee_month,
                )
            )

        conn.close()

        flash(
            "Free Card request submitted for Owner approval.",
            "success",
        )

        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    student_rows = conn.execute(
        """
        SELECT *
        FROM students
        WHERE status = 'ACTIVE'
        ORDER BY full_name
        """
    ).fetchall()

    students_data = []

    for student in student_rows:
        class_rows = conn.execute(
            """
            SELECT
                student_classes.class_code,
                classes.class_name,
                classes.monthly_fee,
                teachers.full_name AS teacher_name
            FROM student_classes
            JOIN classes
                ON classes.class_code = student_classes.class_code
            JOIN teachers
                ON teachers.id = classes.teacher_id
            WHERE student_classes.student_id = ?
            AND (
                student_classes.enrollment_status = 'ACTIVE'
                OR student_classes.enrollment_status IS NULL
            )
            AND classes.status = 'ACTIVE'
            ORDER BY classes.class_name
            """,
            (
                student["id"],
            ),
        ).fetchall()

        students_data.append(
            {
                "id": student["id"],
                "student_code": student["student_code"],
                "full_name": student["full_name"],
                "grade": student["grade"],
                "classes": [
                    dict(row)
                    for row in class_rows
                ],
            }
        )

    conn.close()

    return render_template(
        "free_card_form.html",
        students=students_data,
        default_month=datetime.now().strftime(
            "%Y-%m"
        ),
    )


@app.route(
    "/free-cards/<int:free_card_id>"
)
def free_card_detail(
    free_card_id,
):
    conn = get_db()
    record = load_free_card_record(
        conn,
        free_card_id,
    )

    if not record:
        conn.close()
        abort(404)

    if (
        g.current_user["role"] == "STAFF"
        and record["requested_by_user_id"]
        != g.current_user["id"]
    ):
        conn.close()
        abort(403)

    conn.close()

    return render_template(
        "free_card_detail.html",
        free_card=record,
    )


@app.post(
    "/free-cards/<int:free_card_id>/approve"
)
def approve_free_card(
    free_card_id,
):
    review_note = request.form.get(
        "review_note",
        "",
    ).strip()

    conn = get_db()
    record = load_free_card_record(
        conn,
        free_card_id,
    )

    if not record:
        conn.close()
        abort(404)

    if record["status"] != "PENDING":
        conn.close()
        flash(
            "Only Pending Free Card requests can be approved.",
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    paid = conn.execute(
        """
        SELECT COALESCE(SUM(amount), 0)
        FROM payments
        WHERE student_id = ?
        AND class_code = ?
        AND fee_month = ?
        AND status = 'CONFIRMED'
        """,
        (
            record["student_id"],
            record["class_code"],
            record["fee_month"],
        ),
    ).fetchone()[0]

    if float(paid or 0) > 0:
        conn.close()
        flash(
            (
                "Approval blocked because a confirmed payment exists "
                "for the same student/class/month. Resolve that "
                "payment first."
            ),
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    fee = float(
        record["monthly_fee"]
        or 0
    )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE free_cards
        SET
            status = 'APPROVED',
            waived_amount = ?,
            reviewed_by_user_id = ?,
            reviewed_by_name = ?,
            reviewed_at = ?,
            review_note = ?,
            rejected_reason = NULL
        WHERE id = ?
        AND status = 'PENDING'
        """,
        (
            fee,
            g.current_user["id"],
            g.current_user["full_name"],
            now,
            review_note,
            free_card_id,
        ),
    )

    conn.commit()
    conn.close()

    try:
        write_audit_log(
            action="FREE CARD APPROVED",
            entity_type="FREE_CARD",
            entity_id=str(free_card_id),
            metadata={
                "student_id": record["student_id"],
                "class_code": record["class_code"],
                "fee_month": record["fee_month"],
                "request_source": record["request_source"],
            },
        )
    except Exception:
        pass

    flash(
        (
            "Free Card approved. "
            f"Rs. {fee:,.2f} is waived for this fee month."
        ),
        "success",
    )

    return_to = request.form.get("return_to", "").strip()
    if is_safe_next_url(return_to):
        return redirect(return_to)

    return redirect(
        url_for(
            "free_card_detail",
            free_card_id=free_card_id,
        )
    )


@app.post(
    "/free-cards/<int:free_card_id>/reject"
)
def reject_free_card(
    free_card_id,
):
    reason = request.form.get(
        "rejected_reason",
        "",
    ).strip()

    if not reason:
        flash(
            "Rejection reason is required.",
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    conn = get_db()
    record = load_free_card_record(
        conn,
        free_card_id,
    )

    if not record:
        conn.close()
        abort(404)

    if record["status"] != "PENDING":
        conn.close()
        flash(
            "Only Pending Free Card requests can be rejected.",
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE free_cards
        SET
            status = 'REJECTED',
            reviewed_by_user_id = ?,
            reviewed_by_name = ?,
            reviewed_at = ?,
            rejected_reason = ?
        WHERE id = ?
        AND status = 'PENDING'
        """,
        (
            g.current_user["id"],
            g.current_user["full_name"],
            now,
            reason,
            free_card_id,
        ),
    )

    conn.commit()
    conn.close()

    try:
        write_audit_log(
            action="FREE CARD REJECTED",
            entity_type="FREE_CARD",
            entity_id=str(free_card_id),
            metadata={
                "student_id": record["student_id"],
                "class_code": record["class_code"],
                "fee_month": record["fee_month"],
                "request_source": record["request_source"],
                "reason": reason,
            },
        )
    except Exception:
        pass

    flash(
        "Free Card request rejected.",
        "success",
    )

    return_to = request.form.get("return_to", "").strip()
    if is_safe_next_url(return_to):
        return redirect(return_to)

    return redirect(
        url_for(
            "free_card_detail",
            free_card_id=free_card_id,
        )
    )


@app.post(
    "/free-cards/<int:free_card_id>/revoke"
)
def revoke_free_card(
    free_card_id,
):
    reason = request.form.get(
        "revoke_reason",
        "",
    ).strip()

    if not reason:
        flash(
            "Revoke reason is required.",
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    conn = get_db()
    record = load_free_card_record(
        conn,
        free_card_id,
    )

    if not record:
        conn.close()
        abort(404)

    if record["status"] != "APPROVED":
        conn.close()
        flash(
            "Only Approved Free Cards can be revoked.",
            "error",
        )
        return redirect(
            url_for(
                "free_card_detail",
                free_card_id=free_card_id,
            )
        )

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute(
        """
        UPDATE free_cards
        SET
            status = 'REVOKED',
            revoked_by_user_id = ?,
            revoked_by_name = ?,
            revoked_at = ?,
            revoke_reason = ?
        WHERE id = ?
        AND status = 'APPROVED'
        """,
        (
            g.current_user["id"],
            g.current_user["full_name"],
            now,
            reason,
            free_card_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        (
            "Free Card revoked. The fee becomes payable again "
            "for this month."
        ),
        "success",
    )

    return redirect(
        url_for(
            "free_card_detail",
            free_card_id=free_card_id,
        )
    )



# =========================================================
# FINANCE MODULE
# =========================================================


# =========================================================
# FINANCE - OVERVIEW
# =========================================================

@app.route(
    "/finance"
)
def finance_overview():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    conn = get_db()

    summary = get_payment_summary(
        conn,
        selected_month,
    )

    recent_rows = conn.execute(
        """
        SELECT
            payments.id,
            payments.receipt_no,
            payments.student_id,
            payments.class_code,
            payments.fee_month,
            payments.amount,
            payments.payment_method,
            payments.payment_date,
            payments.status,
            payments.created_at,

            students.student_code,
            students.full_name,

            classes.class_name,

            teachers.full_name
                AS teacher_name

        FROM payments

        JOIN students
            ON students.id =
                payments.student_id

        JOIN classes
            ON classes.class_code =
                payments.class_code

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            payments.fee_month = ?

        ORDER BY
            payments.id DESC

        LIMIT 8
        """,
        (
            selected_month,
        ),
    ).fetchall()

    recent_payments = [
        dict(row)
        for row in recent_rows
    ]

    transaction_stats_row = conn.execute(
        """
        SELECT
            COUNT(*) AS total_transactions,

            SUM(
                CASE
                    WHEN status = 'CONFIRMED'
                    THEN 1
                    ELSE 0
                END
            ) AS confirmed_transactions,

            SUM(
                CASE
                    WHEN status = 'VOID'
                    THEN 1
                    ELSE 0
                END
            ) AS void_transactions,

            COALESCE(
                SUM(
                    CASE
                        WHEN status = 'CONFIRMED'
                        THEN amount
                        ELSE 0
                    END
                ),
                0
            ) AS confirmed_amount

        FROM payments

        WHERE fee_month = ?
        """,
        (
            selected_month,
        ),
    ).fetchone()

    transaction_stats = {
        "total_transactions": int(
            transaction_stats_row[
                "total_transactions"
            ]
            or 0
        ),
        "confirmed_transactions": int(
            transaction_stats_row[
                "confirmed_transactions"
            ]
            or 0
        ),
        "void_transactions": int(
            transaction_stats_row[
                "void_transactions"
            ]
            or 0
        ),
        "confirmed_amount": float(
            transaction_stats_row[
                "confirmed_amount"
            ]
            or 0
        ),
    }

    approved_free_card_rows = conn.execute(
        """
        SELECT
            free_cards.id,
            free_cards.student_id,
            free_cards.class_code,
            free_cards.fee_month,
            free_cards.waived_amount,
            free_cards.status,

            students.student_code,
            students.full_name,

            classes.class_name,

            teachers.full_name
                AS teacher_name

        FROM free_cards

        JOIN students
            ON students.id =
                free_cards.student_id

        JOIN classes
            ON classes.class_code =
                free_cards.class_code

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            free_cards.fee_month = ?

        AND free_cards.status = 'APPROVED'

        ORDER BY
            free_cards.id DESC

        LIMIT 8
        """,
        (
            selected_month,
        ),
    ).fetchall()

    approved_free_cards = [
        dict(row)
        for row in approved_free_card_rows
    ]

    conn.close()

    return render_template(
        "finance_overview.html",

        active_tab=
            "overview",

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m",
            ).strftime(
                "%B %Y"
            ),

        previous_month=
            shift_month(
                selected_month,
                -1,
            ),

        next_month=
            shift_month(
                selected_month,
                1,
            ),

        summary=
            summary,

        recent_payments=
            recent_payments,

        approved_free_cards=
            approved_free_cards,

        transaction_stats=
            transaction_stats,
    )


# =========================================================
# FINANCE - OUTSTANDING
# =========================================================

@app.route(
    "/finance/outstanding"
)
def finance_outstanding():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    search = (
        request.args.get(
            "search",
            ""
        )
        .strip()
    )

    conn = get_db()

    month_start, month_end = (
        month_bounds(
            selected_month
        )
    )

    enrollment_rows = conn.execute(
        """
        SELECT
            student_classes.student_id,
            student_classes.class_code,
            student_classes.created_at,
            student_classes.ended_at,

            students.student_code,
            students.full_name,
            students.grade,
            students.phone,
            students.status
                AS student_status,

            classes.class_name,
            classes.subject,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM student_classes

        JOIN students
            ON students.id =
                student_classes.student_id

        JOIN classes
            ON classes.class_code =
                student_classes.class_code

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            student_classes.created_at <= ?

        AND (
            student_classes.ended_at IS NULL

            OR student_classes.ended_at > ?
        )

        ORDER BY
            students.full_name,
            classes.class_name
        """,
        (
            month_end,
            month_end,
        ),
    ).fetchall()

    outstanding_rows = []

    for enrollment in enrollment_rows:

        fee = float(
            enrollment[
                "monthly_fee"
            ]
            or 0
        )

        waiver = get_free_card_waiver(
            conn,
            enrollment[
                "student_id"
            ],
            enrollment[
                "class_code"
            ],
            selected_month,
            fee,
        )

        waiver = min(
            float(
                waiver
                or 0
            ),
            fee,
        )

        payable = max(
            fee - waiver,
            0,
        )

        paid = conn.execute(
            """
            SELECT COALESCE(
                SUM(amount),
                0
            )

            FROM payments

            WHERE
                student_id = ?

            AND class_code = ?

            AND fee_month = ?

            AND status = 'CONFIRMED'
            """,
            (
                enrollment[
                    "student_id"
                ],
                enrollment[
                    "class_code"
                ],
                selected_month,
            ),
        ).fetchone()[0]

        paid = float(
            paid
            or 0
        )

        balance = max(
            payable - paid,
            0,
        )

        if balance <= 0:
            continue

        fee_status = (
            "PARTIAL"
            if paid > 0
            else "UNPAID"
        )

        item = dict(
            enrollment
        )

        item[
            "monthly_fee"
        ] = fee

        item[
            "waived_amount"
        ] = waiver

        item[
            "payable"
        ] = payable

        item[
            "paid"
        ] = paid

        item[
            "balance"
        ] = balance

        item[
            "fee_status"
        ] = fee_status

        if search:

            search_text = (
                search.lower()
            )

            haystack = (
                str(
                    item[
                        "student_code"
                    ]
                    or ""
                )
                + " "
                + str(
                    item[
                        "full_name"
                    ]
                    or ""
                )
                + " "
                + str(
                    item[
                        "class_code"
                    ]
                    or ""
                )
                + " "
                + str(
                    item[
                        "class_name"
                    ]
                    or ""
                )
                + " "
                + str(
                    item[
                        "teacher_name"
                    ]
                    or ""
                )
            ).lower()

            if (
                search_text
                not in haystack
            ):
                continue

        outstanding_rows.append(
            item
        )

    outstanding_rows.sort(
        key=lambda item:
            item[
                "balance"
            ],
        reverse=True,
    )

    total_outstanding = sum(
        item[
            "balance"
        ]
        for item in outstanding_rows
    )

    partial_count = sum(
        1
        for item in outstanding_rows
        if item[
            "fee_status"
        ] == "PARTIAL"
    )

    unpaid_count = sum(
        1
        for item in outstanding_rows
        if item[
            "fee_status"
        ] == "UNPAID"
    )

    summary = {
        "accounts": len(
            outstanding_rows
        ),
        "partial_count": partial_count,
        "unpaid_count": unpaid_count,
        "outstanding": total_outstanding,
    }

    conn.close()

    return render_template(
        "finance_outstanding.html",

        active_tab=
            "outstanding",

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m",
            ).strftime(
                "%B %Y"
            ),

        previous_month=
            shift_month(
                selected_month,
                -1,
            ),

        next_month=
            shift_month(
                selected_month,
                1,
            ),

        search=
            search,

        outstanding_rows=
            outstanding_rows,

        summary=
            summary,
    )


# =========================================================
# FINANCE - PAYMENT HISTORY
# =========================================================

@app.route(
    "/finance/history"
)
def finance_history():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    search = (
        request.args.get(
            "search",
            ""
        )
        .strip()
    )

    status_filter = (
        request.args.get(
            "status",
            "ALL"
        )
        .strip()
        .upper()
    )

    allowed_statuses = {
        "ALL",
        "CONFIRMED",
        "VOID",
    }

    if (
        status_filter
        not in allowed_statuses
    ):
        status_filter = "ALL"

    conn = get_db()

    sql = """
        SELECT
            payments.*,

            students.student_code,
            students.full_name,
            students.grade,

            classes.class_name,
            classes.subject,
            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM payments

        JOIN students
            ON students.id =
                payments.student_id

        JOIN classes
            ON classes.class_code =
                payments.class_code

        JOIN teachers
            ON teachers.id =
                classes.teacher_id

        WHERE
            payments.fee_month = ?
    """

    params = [
        selected_month,
    ]

    if (
        status_filter
        != "ALL"
    ):

        sql += """
            AND payments.status = ?
        """

        params.append(
            status_filter
        )

    if search:

        sql += """
            AND (
                LOWER(
                    students.full_name
                ) LIKE ?

                OR LOWER(
                    students.student_code
                ) LIKE ?

                OR LOWER(
                    COALESCE(
                        payments.receipt_no,
                        ''
                    )
                ) LIKE ?

                OR LOWER(
                    classes.class_name
                ) LIKE ?

                OR LOWER(
                    payments.class_code
                ) LIKE ?
            )
        """

        search_value = (
            "%"
            + search.lower()
            + "%"
        )

        params.extend(
            [
                search_value,
                search_value,
                search_value,
                search_value,
                search_value,
            ]
        )

    sql += """
        ORDER BY
            payments.payment_date DESC,
            payments.id DESC

        LIMIT 500
    """

    rows = conn.execute(
        sql,
        params,
    ).fetchall()

    payment_history = [
        dict(row)
        for row in rows
    ]

    confirmed_amount = sum(
        float(
            row[
                "amount"
            ]
            or 0
        )
        for row in payment_history
        if row[
            "status"
        ] == "CONFIRMED"
    )

    void_amount = sum(
        float(
            row[
                "amount"
            ]
            or 0
        )
        for row in payment_history
        if row[
            "status"
        ] == "VOID"
    )

    confirmed_count = sum(
        1
        for row in payment_history
        if row[
            "status"
        ] == "CONFIRMED"
    )

    void_count = sum(
        1
        for row in payment_history
        if row[
            "status"
        ] == "VOID"
    )

    history_summary = {
        "transactions": len(
            payment_history
        ),
        "confirmed_count": confirmed_count,
        "void_count": void_count,
        "confirmed_amount": confirmed_amount,
        "void_amount": void_amount,
    }

    conn.close()

    return render_template(
        "finance_history.html",

        active_tab=
            "history",

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m",
            ).strftime(
                "%B %Y"
            ),

        previous_month=
            shift_month(
                selected_month,
                -1,
            ),

        next_month=
            shift_month(
                selected_month,
                1,
            ),

        search=
            search,

        status_filter=
            status_filter,

        payments=
            payment_history,

        summary=
            history_summary,
    )


# =========================================================
# PAYMENTS
# =========================================================

@app.route("/payments")
def payments():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    conn = get_db()

    summary = get_payment_summary(
        conn,
        selected_month,
    )

    rows = conn.execute(
        """
        SELECT
            payments.*,

            students.student_code,

            students.full_name,

            classes.class_name,

            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM payments

        JOIN students

        ON students.id =
            payments.student_id

        JOIN classes

        ON classes.class_code =
            payments.class_code

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            payments.fee_month = ?

        ORDER BY
            payments.id DESC
        """,
        (
            selected_month,
        ),
    ).fetchall()

    approved_free_cards = conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name,
            classes.class_name,
            teachers.full_name AS teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE free_cards.fee_month = ?
        AND free_cards.status = 'APPROVED'
        ORDER BY free_cards.id DESC
        """,
        (
            selected_month,
        ),
    ).fetchall()

    conn.close()

    return render_template(
        "payments.html",

        payments=
            rows,

        summary=
            summary,

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m"
            ).strftime(
                "%B %Y"
            ),

        approved_free_cards=
            approved_free_cards,
    )


@app.route(
    "/payments/new",
    methods=[
        "GET",
        "POST",
    ],
)
def new_payment():

    conn = get_db()

    if request.method == "POST":

        student_id = request.form.get(
            "student_id",
            "",
        ).strip()

        class_code = request.form.get(
            "class_code",
            "",
        ).strip()

        fee_month = request.form.get(
            "fee_month",
            "",
        ).strip()

        amount_text = request.form.get(
            "amount",
            "",
        ).strip()

        payment_method = request.form.get(
            "payment_method",
            "CASH",
        ).strip().upper()

        notes = request.form.get(
            "notes",
            "",
        ).strip()

        if (
            not student_id.isdigit()
            or
            not class_code
            or
            not valid_month(
                fee_month
            )
        ):

            conn.close()

            flash(
                (
                    "Please select student, "
                    "class and fee month."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        try:

            amount = float(
                amount_text
            )

        except ValueError:

            amount = 0

        if amount <= 0:

            conn.close()

            flash(
                (
                    "Payment amount must "
                    "be greater than zero."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        student = conn.execute(
            """
            SELECT *

            FROM students

            WHERE
                id = ?

            AND status = 'ACTIVE'
            """,
            (
                int(
                    student_id
                ),
            ),
        ).fetchone()

        if not student:

            conn.close()

            flash(
                "Student not found.",
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        enrollment = conn.execute(
            """
            SELECT
                student_classes.id

            FROM student_classes

            JOIN classes

            ON classes.class_code =
                student_classes.class_code

            WHERE
                student_classes.student_id = ?

            AND student_classes.class_code = ?

            AND (
                student_classes.enrollment_status = 'ACTIVE'

                OR student_classes.enrollment_status IS NULL
            )

            AND classes.status = 'ACTIVE'
            """,
            (
                int(
                    student_id
                ),
                class_code,
            ),
        ).fetchone()

        if not enrollment:

            conn.close()

            flash(
                (
                    "Student is not enrolled "
                    "in this active class."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        class_row = conn.execute(
            """
            SELECT *

            FROM classes

            WHERE
                class_code = ?

            AND status = 'ACTIVE'
            """,
            (
                class_code,
            ),
        ).fetchone()

        if not class_row:

            conn.close()

            flash(
                "Class not found or inactive.",
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        fee = float(
            class_row[
                "monthly_fee"
            ]
            or 0
        )

        approved_free_card = get_approved_free_card(
            conn,
            int(student_id),
            class_code,
            fee_month,
        )

        if approved_free_card:
            conn.close()

            flash(
                (
                    "Payment cannot be recorded because an Approved "
                    "Free Card covers this student/class/month."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        already_paid = conn.execute(
            """
            SELECT COALESCE(
                SUM(amount),
                0
            )

            FROM payments

            WHERE
                student_id = ?

            AND class_code = ?

            AND fee_month = ?

            AND status = 'CONFIRMED'
            """,
            (
                int(
                    student_id
                ),
                class_code,
                fee_month,
            ),
        ).fetchone()[0]

        already_paid = float(
            already_paid
            or 0
        )

        remaining = max(
            fee
            -
            already_paid,
            0,
        )

        if (
            fee > 0
            and
            amount > remaining
        ):

            conn.close()

            flash(
                (
                    "Payment exceeds balance. "
                    f"Remaining balance: "
                    f"Rs. {remaining:,.0f}"
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_payment"
                )
            )

        receipt_no = (
            make_receipt_no(
                conn
            )
        )

        now = datetime.now()

        cursor = conn.execute(
            """
            INSERT INTO payments
            (
                student_id,
                class_code,
                fee_month,
                amount,
                status,
                payment_date,
                created_at,
                receipt_no,
                payment_method,
                notes,
                recorded_by
            )

            VALUES
            (
                ?, ?, ?, ?,
                'CONFIRMED',
                ?, ?, ?, ?, ?, ?
            )
            """,
            (
                int(
                    student_id
                ),
                class_code,
                fee_month,
                amount,

                now.strftime(
                    "%Y-%m-%d"
                ),

                now.strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

                receipt_no,
                payment_method,
                notes,
                g.current_user["full_name"],
            ),
        )

        payment_id = (
            cursor.lastrowid
        )

        conn.commit()
        conn.close()

        return redirect(
            url_for(
                "payment_receipt",

                payment_id=
                    payment_id,
            )
        )


    student_rows = conn.execute(
        """
        SELECT *

        FROM students

        WHERE status = 'ACTIVE'

        ORDER BY full_name
        """
    ).fetchall()

    student_data = []

    for student in student_rows:

        class_rows = conn.execute(
            """
            SELECT
                student_classes.class_code,

                classes.class_name,

                classes.monthly_fee,

                teachers.full_name
                    AS teacher_name

            FROM student_classes

            JOIN classes

            ON classes.class_code =
                student_classes.class_code

            JOIN teachers

            ON teachers.id =
                classes.teacher_id

            WHERE
                student_classes.student_id = ?

            AND (
                student_classes.enrollment_status = 'ACTIVE'

                OR student_classes.enrollment_status IS NULL
            )

            AND classes.status = 'ACTIVE'

            ORDER BY
                classes.class_name
            """,
            (
                student["id"],
            ),
        ).fetchall()

        student_data.append(
            {
                "id":
                    student["id"],

                "student_code":
                    student["student_code"],

                "full_name":
                    student["full_name"],

                "classes": [
                    dict(row)
                    for row in class_rows
                ],
            }
        )

    conn.close()

    return render_template(
        "payment_form.html",

        students=
            student_data,

        default_month=
            datetime.now().strftime(
                "%Y-%m"
            ),
    )


@app.get(
    "/api/payment-balance"
)
def payment_balance():

    student_id = request.args.get(
        "student_id",
        "",
    )

    class_code = request.args.get(
        "class_code",
        "",
    )

    fee_month = request.args.get(
        "fee_month",
        "",
    )

    if (
        not student_id.isdigit()
        or
        not class_code
        or
        not valid_month(
            fee_month
        )
    ):

        return jsonify(
            {
                "success":
                    False,
            }
        ), 400

    conn = get_db()

    enrollment = conn.execute(
        """
        SELECT
            student_classes.id

        FROM student_classes

        JOIN classes

        ON classes.class_code =
            student_classes.class_code

        WHERE
            student_classes.student_id = ?

        AND student_classes.class_code = ?

        AND (
            student_classes.enrollment_status = 'ACTIVE'

            OR student_classes.enrollment_status IS NULL
        )
        """,
        (
            int(
                student_id
            ),
            class_code,
        ),
    ).fetchone()

    if not enrollment:

        conn.close()

        return jsonify(
            {
                "success":
                    False,

                "message":
                    "Student is not enrolled in this class.",
            }
        ), 403

    class_row = conn.execute(
        """
        SELECT monthly_fee

        FROM classes

        WHERE class_code = ?
        """,
        (
            class_code,
        ),
    ).fetchone()

    if not class_row:

        conn.close()

        return jsonify(
            {
                "success":
                    False,
            }
        ), 404

    fee = float(
        class_row[
            "monthly_fee"
        ]
        or 0
    )

    paid = conn.execute(
        """
        SELECT COALESCE(
            SUM(amount),
            0
        )

        FROM payments

        WHERE
            student_id = ?

        AND class_code = ?

        AND fee_month = ?

        AND status = 'CONFIRMED'
        """,
        (
            int(
                student_id
            ),
            class_code,
            fee_month,
        ),
    ).fetchone()[0]

    approved_free_card = get_approved_free_card(
        conn,
        int(student_id),
        class_code,
        fee_month,
    )

    if approved_free_card:
        waived = get_free_card_waiver(
            conn,
            int(student_id),
            class_code,
            fee_month,
            fee,
        )

        conn.close()

        return jsonify(
            {
                "success": True,
                "fee": fee,
                "paid": float(paid or 0),
                "balance": 0,
                "status": "FREE CARD",
                "free_card": True,
                "waived": waived,
            }
        )

    conn.close()

    paid = float(
        paid
        or 0
    )

    balance = max(
        fee
        -
        paid,
        0,
    )

    if paid <= 0:
        status = "UNPAID"

    elif paid < fee:
        status = "PARTIAL"

    else:
        status = "PAID"

    return jsonify(
        {
            "success":
                True,

            "fee":
                fee,

            "paid":
                paid,

            "balance":
                balance,

            "status":
                status,
        }
    )


@app.route(
    "/payments/<int:payment_id>"
)
def payment_receipt(
    payment_id
):

    conn = get_db()

    payment = conn.execute(
        """
        SELECT
            payments.*,

            students.student_code,

            students.full_name,

            students.phone,

            classes.class_name,

            classes.monthly_fee,

            teachers.full_name
                AS teacher_name

        FROM payments

        JOIN students

        ON students.id =
            payments.student_id

        JOIN classes

        ON classes.class_code =
            payments.class_code

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        WHERE
            payments.id = ?
        """,
        (
            payment_id,
        ),
    ).fetchone()

    if not payment:

        conn.close()
        abort(404)

    total_paid = conn.execute(
        """
        SELECT COALESCE(
            SUM(amount),
            0
        )

        FROM payments

        WHERE
            student_id = ?

        AND class_code = ?

        AND fee_month = ?

        AND status = 'CONFIRMED'
        """,
        (
            payment[
                "student_id"
            ],

            payment[
                "class_code"
            ],

            payment[
                "fee_month"
            ],
        ),
    ).fetchone()[0]

    conn.close()

    fee = float(
        payment[
            "monthly_fee"
        ]
        or 0
    )

    total_paid = float(
        total_paid
        or 0
    )

    balance = max(
        fee
        -
        total_paid,
        0,
    )

    if total_paid <= 0:
        fee_status = "UNPAID"

    elif total_paid < fee:
        fee_status = "PARTIAL"

    else:
        fee_status = "PAID"

    return render_template(
        "payment_receipt.html",

        payment=
            payment,

        total_paid=
            total_paid,

        balance=
            balance,

        fee_status=
            fee_status,
    )


@app.post(
    "/payments/<int:payment_id>/void"
)
def void_payment(
    payment_id
):

    reason = request.form.get(
        "void_reason",
        "",
    ).strip()

    if not reason:

        flash(
            "Void reason is required.",
            "error",
        )

        return redirect(
            url_for(
                "payment_receipt",

                payment_id=
                    payment_id,
            )
        )

    conn = get_db()

    payment = conn.execute(
        """
        SELECT *

        FROM payments

        WHERE id = ?
        """,
        (
            payment_id,
        ),
    ).fetchone()

    if not payment:

        conn.close()
        abort(404)

    if (
        payment[
            "status"
        ] != "CONFIRMED"
    ):

        conn.close()

        flash(
            "Only confirmed payments can be voided.",
            "error",
        )

        return redirect(
            url_for(
                "payment_receipt",

                payment_id=
                    payment_id,
            )
        )

    conn.execute(
        """
        UPDATE payments

        SET
            status = 'VOID',

            voided_at = ?,

            void_reason = ?

        WHERE id = ?
        """,
        (
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

            reason,

            payment_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        "Payment voided successfully.",
        "success",
    )

    return redirect(
        url_for(
            "payment_receipt",

            payment_id=
                payment_id,
        )
    )


# =========================================================
# CLASS MANAGEMENT
# =========================================================

@app.route("/classes")
def class_management():

    conn = get_db()

    rows = conn.execute(
        """
        SELECT
            classes.*,

            teachers.full_name
                AS teacher_name,

            (
                SELECT COUNT(
                    DISTINCT
                    student_classes.student_id
                )

                FROM student_classes

                WHERE
                    student_classes.class_code =
                        classes.class_code

                AND (
                    student_classes.enrollment_status = 'ACTIVE'

                    OR student_classes.enrollment_status IS NULL
                )
            )
            AS student_count

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        ORDER BY
            CASE
                WHEN classes.status = 'ACTIVE'
                THEN 0
                ELSE 1
            END,

            classes.day_of_week,

            classes.start_time
        """
    ).fetchall()

    class_list = []

    for row in rows:

        item = dict(row)

        item[
            "day_name"
        ] = DAY_NAMES[
            item[
                "day_of_week"
            ]
        ]

        item[
            "start_display"
        ] = format_time(
            item[
                "start_time"
            ]
        )

        item[
            "end_display"
        ] = format_time(
            item[
                "end_time"
            ]
        )

        class_list.append(
            item
        )

    conn.close()

    return render_template(
        "classes.html",

        classes=
            class_list,
    )


# =========================================================
# CLASS PROFILE
# =========================================================

@app.route(
    "/classes/<class_code>"
)
def class_profile(
    class_code
):

    allowed_tabs = {
        "overview",
        "students",
        "schedule",
        "attendance",
        "payments",
        "teacher",
    }

    active_tab = request.args.get(
        "tab",
        "overview",
    ).strip().lower()

    if active_tab not in allowed_tabs:
        active_tab = "overview"

    selected_month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    if not valid_month(selected_month):
        selected_month = datetime.now().strftime("%Y-%m")

    month_start, month_end = month_bounds(
        selected_month
    )

    conn = get_db()

    class_row = conn.execute(
        """
        SELECT
            classes.*,
            teachers.teacher_code,
            teachers.full_name AS teacher_name,
            teachers.subject AS teacher_subject,
            teachers.phone AS teacher_phone,
            teachers.status AS teacher_status
        FROM classes
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE classes.class_code = ?
        """,
        (
            class_code,
        ),
    ).fetchone()

    if not class_row:
        conn.close()
        abort(404)

    class_item = dict(class_row)
    class_item["day_name"] = DAY_NAMES[
        class_item["day_of_week"]
    ]
    class_item["start_display"] = format_time(
        class_item["start_time"]
    )
    class_item["end_display"] = format_time(
        class_item["end_time"]
    )

    enrollment_rows = conn.execute(
        """
        SELECT
            student_classes.id AS enrollment_id,
            student_classes.created_at AS enrolled_at,
            student_classes.enrollment_status,
            student_classes.ended_at,
            students.id AS student_id,
            students.student_code,
            students.full_name,
            students.phone,
            students.school,
            students.grade,
            students.status AS student_status
        FROM student_classes
        JOIN students
            ON students.id = student_classes.student_id
        WHERE student_classes.class_code = ?
        ORDER BY
            CASE
                WHEN (
                    student_classes.enrollment_status = 'ACTIVE'
                    OR student_classes.enrollment_status IS NULL
                )
                THEN 0
                ELSE 1
            END,
            students.full_name
        """,
        (
            class_code,
        ),
    ).fetchall()

    monthly_fee = float(
        class_item["monthly_fee"]
        or 0
    )

    students_data = []
    active_student_count = 0

    for row in enrollment_rows:
        student = dict(row)

        is_active = (
            student["enrollment_status"] == "ACTIVE"
            or student["enrollment_status"] is None
        )

        student["is_active"] = is_active

        if is_active:
            active_student_count += 1

        enrolled_at = student.get("enrolled_at") or ""
        ended_at = student.get("ended_at")

        in_selected_month = bool(
            enrolled_at
            and enrolled_at <= month_end
            and (
                ended_at is None
                or ended_at > month_end
            )
        )

        student["in_selected_month"] = in_selected_month

        paid = 0.0
        waiver = 0.0
        payable = 0.0
        balance = 0.0
        fee_status = "NOT DUE"

        if in_selected_month:
            waiver = get_free_card_waiver(
                conn,
                student["student_id"],
                class_code,
                selected_month,
                monthly_fee,
            )

            waiver = min(
                waiver,
                monthly_fee,
            )

            payable = max(
                monthly_fee - waiver,
                0,
            )

            paid = conn.execute(
                """
                SELECT COALESCE(SUM(amount), 0)
                FROM payments
                WHERE student_id = ?
                AND class_code = ?
                AND fee_month = ?
                AND status = 'CONFIRMED'
                """,
                (
                    student["student_id"],
                    class_code,
                    selected_month,
                ),
            ).fetchone()[0]

            paid = float(
                paid
                or 0
            )

            balance = max(
                payable - paid,
                0,
            )

            if waiver > 0 and payable <= 0:
                fee_status = "FREE CARD"
            elif paid <= 0:
                fee_status = "UNPAID"
            elif paid < payable:
                fee_status = "PARTIAL"
            else:
                fee_status = "PAID"

        student["waiver"] = waiver
        student["payable"] = payable
        student["paid"] = paid
        student["balance"] = balance
        student["fee_status"] = fee_status

        students_data.append(student)

    end_students = count_students_at_month_end(
        conn,
        class_code,
        selected_month,
    )

    gross_expected = (
        end_students
        * monthly_fee
    )

    free_card_summary = get_class_free_card_summary(
        conn,
        class_code,
        selected_month,
    )

    waived_amount = min(
        float(
            free_card_summary["waived_amount"]
            or 0
        ),
        gross_expected,
    )

    expected = max(
        gross_expected - waived_amount,
        0,
    )

    collected = conn.execute(
        """
        SELECT COALESCE(SUM(amount), 0)
        FROM payments
        WHERE class_code = ?
        AND fee_month = ?
        AND status = 'CONFIRMED'
        """,
        (
            class_code,
            selected_month,
        ),
    ).fetchone()[0]

    collected = float(
        collected
        or 0
    )

    outstanding = max(
        expected - collected,
        0,
    )

    collection_rate = (
        round(
            collected / expected * 100,
            1,
        )
        if expected > 0
        else 0
    )

    teacher_share_percent = float(
        class_item["teacher_share_percent"]
        or 0
    )

    teacher_income = (
        collected
        * teacher_share_percent
        / 100
    )

    center_income = (
        collected
        - teacher_income
    )

    attendance_rows = conn.execute(
        """
        SELECT
            attendance.*,
            students.student_code,
            students.full_name
        FROM attendance
        JOIN students
            ON students.id = attendance.student_id
        WHERE attendance.class_code = ?
        AND attendance.attendance_date BETWEEN ? AND ?
        ORDER BY
            attendance.attendance_date DESC,
            attendance.check_in_time DESC,
            attendance.id DESC
        LIMIT 500
        """,
        (
            class_code,
            month_start[:10],
            month_end[:10],
        ),
    ).fetchall()

    attendance_records = [
        dict(row)
        for row in attendance_rows
    ]

    attendance_summary = conn.execute(
        """
        SELECT
            COUNT(*) AS total_records,
            COUNT(DISTINCT attendance_date) AS class_days
        FROM attendance
        WHERE class_code = ?
        AND attendance_date BETWEEN ? AND ?
        """,
        (
            class_code,
            month_start[:10],
            month_end[:10],
        ),
    ).fetchone()

    attendance_total = int(
        attendance_summary["total_records"]
        or 0
    )

    class_days = int(
        attendance_summary["class_days"]
        or 0
    )

    average_attendance = (
        round(
            attendance_total / class_days,
            1,
        )
        if class_days > 0
        else 0
    )

    attendance_rate = 0

    if class_days > 0 and end_students > 0:
        attendance_rate = min(
            round(
                attendance_total
                / (class_days * end_students)
                * 100,
                1,
            ),
            100,
        )

    payment_rows = conn.execute(
        """
        SELECT
            payments.*,
            students.student_code,
            students.full_name
        FROM payments
        JOIN students
            ON students.id = payments.student_id
        WHERE payments.class_code = ?
        AND payments.fee_month = ?
        ORDER BY payments.id DESC
        LIMIT 500
        """,
        (
            class_code,
            selected_month,
        ),
    ).fetchall()

    payment_records = [
        dict(row)
        for row in payment_rows
    ]

    free_card_rows = conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        WHERE free_cards.class_code = ?
        AND free_cards.fee_month = ?
        ORDER BY free_cards.id DESC
        """,
        (
            class_code,
            selected_month,
        ),
    ).fetchall()

    free_card_records = [
        dict(row)
        for row in free_card_rows
    ]

    finance = {
        "gross_expected": gross_expected,
        "expected": expected,
        "collected": collected,
        "outstanding": outstanding,
        "collection_rate": collection_rate,
        "free_card_count": int(
            free_card_summary["free_card_count"]
            or 0
        ),
        "waived_amount": waived_amount,
        "teacher_income": teacher_income,
        "center_income": center_income,
    }

    attendance_stats = {
        "records": attendance_total,
        "class_days": class_days,
        "average": average_attendance,
        "rate": attendance_rate,
    }

    conn.close()

    return render_template(
        "class_profile.html",
        class_item=class_item,
        active_tab=active_tab,
        selected_month=selected_month,
        month_label=datetime.strptime(
            selected_month,
            "%Y-%m",
        ).strftime("%B %Y"),
        previous_month=shift_month(
            selected_month,
            -1,
        ),
        next_month=shift_month(
            selected_month,
            1,
        ),
        students=students_data,
        active_student_count=active_student_count,
        month_end_student_count=end_students,
        attendance_records=attendance_records,
        attendance_stats=attendance_stats,
        payment_records=payment_records,
        free_card_records=free_card_records,
        finance=finance,
    )


@app.route(
    "/classes/new",
    methods=[
        "GET",
        "POST",
    ],
)
def new_class():

    conn = get_db()

    teacher_rows = conn.execute(
        """
        SELECT *

        FROM teachers

        WHERE status = 'ACTIVE'

        ORDER BY full_name
        """
    ).fetchall()

    if request.method == "POST":

        class_name = request.form.get(
            "class_name",
            "",
        ).strip()

        grade = request.form.get(
            "grade",
            "",
        ).strip()

        subject = request.form.get(
            "subject",
            "",
        ).strip()

        teacher_id_text = request.form.get(
            "teacher_id",
            "",
        ).strip()

        day_text = request.form.get(
            "day_of_week",
            "",
        ).strip()

        start_time = request.form.get(
            "start_time",
            "",
        ).strip()

        end_time = request.form.get(
            "end_time",
            "",
        ).strip()

        room = request.form.get(
            "room",
            "",
        ).strip()

        fee_text = request.form.get(
            "monthly_fee",
            "",
        ).strip()

        share_text = request.form.get(
            "teacher_share_percent",
            "",
        ).strip()

        if (
            not grade
            or not subject
            or not teacher_id_text.isdigit()
            or not day_text.isdigit()
            or not start_time
            or not end_time
        ):

            conn.close()

            flash(
                "Please fill all required fields.",
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        teacher_id = int(
            teacher_id_text
        )

        day_of_week = int(
            day_text
        )

        if (
            day_of_week < 0
            or
            day_of_week > 6
        ):

            conn.close()

            flash(
                "Invalid day selected.",
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        if start_time >= end_time:

            conn.close()

            flash(
                "End time must be after start time.",
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        try:

            monthly_fee = float(
                fee_text
                or 0
            )

            teacher_share = float(
                share_text
                or 0
            )

        except ValueError:

            conn.close()

            flash(
                (
                    "Fee and teacher share "
                    "must be valid numbers."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        if (
            monthly_fee < 0
            or
            teacher_share < 0
            or
            teacher_share > 100
        ):

            conn.close()

            flash(
                (
                    "Please check fee and "
                    "teacher share values."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        teacher = conn.execute(
            """
            SELECT *

            FROM teachers

            WHERE
                id = ?

            AND status = 'ACTIVE'
            """,
            (
                teacher_id,
            ),
        ).fetchone()

        if not teacher:

            conn.close()

            flash(
                "Teacher not found.",
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        conflict = find_class_conflict(
            conn,
            teacher_id,
            day_of_week,
            start_time,
            end_time,
            room,
        )

        if conflict:

            conn.close()

            flash(
                conflict,
                "error",
            )

            return redirect(
                url_for(
                    "new_class"
                )
            )

        if not class_name:

            class_name = (
                f"{grade} {subject}"
            )

        class_code = (
            make_class_code(
                conn
            )
        )

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute(
            """
            INSERT INTO classes
            (
                class_code,
                class_name,
                grade,
                subject,
                teacher_id,
                day_of_week,
                start_time,
                end_time,
                room,
                monthly_fee,
                teacher_share_percent,
                status,
                created_at
            )

            VALUES
            (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?,
                'ACTIVE',
                ?
            )
            """,
            (
                class_code,
                class_name,
                grade,
                subject,
                teacher_id,
                day_of_week,
                start_time,
                end_time,
                room,
                monthly_fee,
                teacher_share,
                now,
            ),
        )

        conn.commit()
        conn.close()

        flash(
            (
                f"Class {class_code} "
                "created successfully."
            ),
            "success",
        )

        return redirect(
            url_for(
                "class_management"
            )
        )

    conn.close()

    return render_template(
        "class_form.html",

        mode=
            "new",

        class_item=
            None,

        teachers=
            teacher_rows,

        days=
            list(
                enumerate(
                    DAY_NAMES
                )
            ),
    )


@app.route(
    "/classes/<int:class_id>/edit",
    methods=[
        "GET",
        "POST",
    ],
)
def edit_class(
    class_id
):

    conn = get_db()

    class_row = conn.execute(
        """
        SELECT *

        FROM classes

        WHERE id = ?
        """,
        (
            class_id,
        ),
    ).fetchone()

    if not class_row:

        conn.close()
        abort(404)

    teacher_rows = conn.execute(
        """
        SELECT *

        FROM teachers

        WHERE status = 'ACTIVE'

        ORDER BY full_name
        """
    ).fetchall()

    if request.method == "POST":

        class_name = request.form.get(
            "class_name",
            "",
        ).strip()

        grade = request.form.get(
            "grade",
            "",
        ).strip()

        subject = request.form.get(
            "subject",
            "",
        ).strip()

        teacher_id_text = request.form.get(
            "teacher_id",
            "",
        ).strip()

        day_text = request.form.get(
            "day_of_week",
            "",
        ).strip()

        start_time = request.form.get(
            "start_time",
            "",
        ).strip()

        end_time = request.form.get(
            "end_time",
            "",
        ).strip()

        room = request.form.get(
            "room",
            "",
        ).strip()

        fee_text = request.form.get(
            "monthly_fee",
            "",
        ).strip()

        share_text = request.form.get(
            "teacher_share_percent",
            "",
        ).strip()

        if (
            not class_name
            or not grade
            or not subject
            or not teacher_id_text.isdigit()
            or not day_text.isdigit()
            or not start_time
            or not end_time
        ):

            conn.close()

            flash(
                "Please fill all required fields.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        teacher_id = int(
            teacher_id_text
        )

        day_of_week = int(
            day_text
        )

        if (
            day_of_week < 0
            or
            day_of_week > 6
        ):

            conn.close()

            flash(
                "Invalid day selected.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        try:

            monthly_fee = float(
                fee_text
                or 0
            )

            teacher_share = float(
                share_text
                or 0
            )

        except ValueError:

            conn.close()

            flash(
                "Invalid fee or teacher share.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        if (
            monthly_fee < 0
            or
            teacher_share < 0
            or
            teacher_share > 100
        ):

            conn.close()

            flash(
                (
                    "Please check monthly fee "
                    "and teacher share."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        if start_time >= end_time:

            conn.close()

            flash(
                "End time must be after start time.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        teacher = conn.execute(
            """
            SELECT *

            FROM teachers

            WHERE
                id = ?

            AND status = 'ACTIVE'
            """,
            (
                teacher_id,
            ),
        ).fetchone()

        if not teacher:

            conn.close()

            flash(
                "Teacher not found.",
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        conflict = find_class_conflict(
            conn,
            teacher_id,
            day_of_week,
            start_time,
            end_time,
            room,

            exclude_class_code=
                class_row[
                    "class_code"
                ],
        )

        if conflict:

            conn.close()

            flash(
                conflict,
                "error",
            )

            return redirect(
                url_for(
                    "edit_class",

                    class_id=
                        class_id,
                )
            )

        conn.execute(
            """
            UPDATE classes

            SET
                class_name = ?,
                grade = ?,
                subject = ?,
                teacher_id = ?,
                day_of_week = ?,
                start_time = ?,
                end_time = ?,
                room = ?,
                monthly_fee = ?,
                teacher_share_percent = ?

            WHERE id = ?
            """,
            (
                class_name,
                grade,
                subject,
                teacher_id,
                day_of_week,
                start_time,
                end_time,
                room,
                monthly_fee,
                teacher_share,
                class_id,
            ),
        )

        conn.execute(
            """
            UPDATE student_classes

            SET
                subject = ?,
                teacher = ?

            WHERE class_code = ?
            """,
            (
                subject,

                teacher[
                    "full_name"
                ],

                class_row[
                    "class_code"
                ],
            ),
        )

        conn.commit()
        conn.close()

        flash(
            "Class updated successfully.",
            "success",
        )

        return redirect(
            url_for(
                "class_management"
            )
        )

    class_item = dict(
        class_row
    )

    conn.close()

    return render_template(
        "class_form.html",

        mode=
            "edit",

        class_item=
            class_item,

        teachers=
            teacher_rows,

        days=
            list(
                enumerate(
                    DAY_NAMES
                )
            ),
    )


@app.post(
    "/classes/<int:class_id>/toggle"
)
def toggle_class(
    class_id
):

    conn = get_db()

    class_row = conn.execute(
        """
        SELECT *

        FROM classes

        WHERE id = ?
        """,
        (
            class_id,
        ),
    ).fetchone()

    if not class_row:

        conn.close()
        abort(404)

    new_status = (
        "INACTIVE"

        if class_row[
            "status"
        ] == "ACTIVE"

        else "ACTIVE"
    )

    if new_status == "ACTIVE":

        teacher = conn.execute(
            """
            SELECT *

            FROM teachers

            WHERE id = ?
            """,
            (
                class_row[
                    "teacher_id"
                ],
            ),
        ).fetchone()

        if (
            not teacher
            or
            teacher["status"] != "ACTIVE"
        ):

            conn.close()

            flash(
                (
                    "This class teacher is inactive. "
                    "Activate the teacher first."
                ),
                "error",
            )

            return redirect(
                url_for(
                    "class_management"
                )
            )

        conflict = find_class_conflict(
            conn,

            class_row[
                "teacher_id"
            ],

            class_row[
                "day_of_week"
            ],

            class_row[
                "start_time"
            ],

            class_row[
                "end_time"
            ],

            class_row[
                "room"
            ],

            exclude_class_code=
                class_row[
                    "class_code"
                ],
        )

        if conflict:

            conn.close()

            flash(
                conflict,
                "error",
            )

            return redirect(
                url_for(
                    "class_management"
                )
            )

    conn.execute(
        """
        UPDATE classes

        SET status = ?

        WHERE id = ?
        """,
        (
            new_status,
            class_id,
        ),
    )

    conn.commit()
    conn.close()

    flash(
        (
            "Class status changed to "
            f"{new_status}."
        ),
        "success",
    )

    return redirect(
        url_for(
            "class_management"
        )
    )


# =========================================================
# OWNER REPORT BUILDER
# =========================================================

def build_monthly_report(
    conn,
    month_key,
):

    month_start, month_end = (
        month_bounds(
            month_key
        )
    )

    class_rows = conn.execute(
        """
        SELECT
            classes.*,

            teachers.teacher_code,

            teachers.full_name
                AS teacher_name,

            teachers.subject
                AS teacher_subject

        FROM classes

        JOIN teachers

        ON teachers.id =
            classes.teacher_id

        ORDER BY
            teachers.full_name,
            classes.grade,
            classes.subject
        """
    ).fetchall()

    class_reports = []

    teacher_map = {}

    total_expected = 0
    total_collected = 0
    total_outstanding = 0
    total_waived_amount = 0
    total_free_card_count = 0

    total_teacher_income = 0
    total_center_income = 0

    total_attendance_records = 0
    total_possible_attendance = 0
    total_class_days = 0

    for class_row in class_rows:

        class_code = (
            class_row[
                "class_code"
            ]
        )

        start_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND created_at < ?

            AND (
                ended_at IS NULL
                OR ended_at >= ?
            )
            """,
            (
                class_code,
                month_start,
                month_start,
            ),
        ).fetchone()[0]

        new_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND created_at
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start,
                month_end,
            ),
        ).fetchone()[0]

        left_students = conn.execute(
            """
            SELECT COUNT(
                DISTINCT student_id
            )

            FROM student_classes

            WHERE
                class_code = ?

            AND ended_at
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start,
                month_end,
            ),
        ).fetchone()[0]

        end_students = (
            count_students_at_month_end(
                conn,
                class_code,
                month_key,
            )
        )

        growth = (
            end_students
            -
            start_students
        )

        growth_percent = 0

        if start_students > 0:

            growth_percent = round(
                (
                    growth
                    /
                    start_students
                )
                *
                100,
                1,
            )

        monthly_fee = float(
            class_row[
                "monthly_fee"
            ]
            or 0
        )

        gross_expected = (
            end_students
            *
            monthly_fee
        )

        free_card_summary = get_class_free_card_summary(
            conn,
            class_code,
            month_key,
        )

        waived_amount = min(
            free_card_summary[
                "waived_amount"
            ],
            gross_expected,
        )

        expected = max(
            gross_expected
            - waived_amount,
            0,
        )

        collected = conn.execute(
            """
            SELECT COALESCE(
                SUM(amount),
                0
            )

            FROM payments

            WHERE
                class_code = ?

            AND fee_month = ?

            AND status = 'CONFIRMED'
            """,
            (
                class_code,
                month_key,
            ),
        ).fetchone()[0]

        collected = float(
            collected
            or 0
        )

        outstanding = max(
            expected
            -
            collected,
            0,
        )

        collection_rate = (
            round(
                collected
                /
                expected
                *
                100,
                1,
            )

            if expected > 0

            else 0
        )

        teacher_share_percent = float(
            class_row[
                "teacher_share_percent"
            ]
            or 0
        )

        teacher_income = (
            collected
            *
            teacher_share_percent
            /
            100
        )

        center_income = (
            collected
            -
            teacher_income
        )

        attendance_info = conn.execute(
            """
            SELECT
                COUNT(*)
                    AS attendance_records,

                COUNT(
                    DISTINCT attendance_date
                )
                    AS class_days

            FROM attendance

            WHERE
                class_code = ?

            AND attendance_date
                BETWEEN ? AND ?
            """,
            (
                class_code,
                month_start[:10],
                month_end[:10],
            ),
        ).fetchone()

        attendance_records = (
            attendance_info[
                "attendance_records"
            ]
            or 0
        )

        class_days = (
            attendance_info[
                "class_days"
            ]
            or 0
        )

        possible_attendance = (
            class_days
            *
            end_students
        )

        attendance_rate = (
            round(
                attendance_records
                /
                possible_attendance
                *
                100,
                1,
            )

            if possible_attendance > 0

            else 0
        )

        attendance_rate = min(
            attendance_rate,
            100,
        )

        class_reports.append(
            {
                "class_code":
                    class_code,

                "class_name":
                    class_row[
                        "class_name"
                    ],

                "grade":
                    class_row[
                        "grade"
                    ],

                "subject":
                    class_row[
                        "subject"
                    ],

                "status":
                    class_row[
                        "status"
                    ],

                "teacher_id":
                    class_row[
                        "teacher_id"
                    ],

                "teacher_name":
                    class_row[
                        "teacher_name"
                    ],

                "teacher_code":
                    class_row[
                        "teacher_code"
                    ],

                "monthly_fee":
                    monthly_fee,

                "teacher_share_percent":
                    teacher_share_percent,

                "start_students":
                    start_students,

                "new_students":
                    new_students,

                "left_students":
                    left_students,

                "end_students":
                    end_students,

                "growth":
                    growth,

                "growth_percent":
                    growth_percent,

                "gross_expected":
                    gross_expected,

                "free_card_count":
                    free_card_summary[
                        "free_card_count"
                    ],

                "waived_amount":
                    waived_amount,

                "expected":
                    expected,

                "collected":
                    collected,

                "outstanding":
                    outstanding,

                "collection_rate":
                    collection_rate,

                "teacher_income":
                    teacher_income,

                "center_income":
                    center_income,

                "attendance_records":
                    attendance_records,

                "class_days":
                    class_days,

                "attendance_rate":
                    attendance_rate,
            }
        )

        total_expected += expected
        total_collected += collected
        total_outstanding += outstanding
        total_waived_amount += waived_amount
        total_free_card_count += free_card_summary[
            "free_card_count"
        ]

        total_teacher_income += (
            teacher_income
        )

        total_center_income += (
            center_income
        )

        total_attendance_records += (
            attendance_records
        )

        total_possible_attendance += (
            possible_attendance
        )

        total_class_days += (
            class_days
        )

        teacher_id = (
            class_row[
                "teacher_id"
            ]
        )

        if teacher_id not in teacher_map:

            teacher_map[
                teacher_id
            ] = {
                "id":
                    teacher_id,

                "teacher_code":
                    class_row[
                        "teacher_code"
                    ],

                "teacher_name":
                    class_row[
                        "teacher_name"
                    ],

                "subject":
                    class_row[
                        "teacher_subject"
                    ],

                "classes":
                    0,

                "expected":
                    0,

                "free_card_count":
                    0,

                "waived_amount":
                    0,

                "collected":
                    0,

                "outstanding":
                    0,

                "teacher_income":
                    0,

                "center_income":
                    0,

                "students":
                    set(),
            }

        teacher_item = (
            teacher_map[
                teacher_id
            ]
        )

        teacher_item[
            "classes"
        ] += 1

        teacher_item[
            "expected"
        ] += expected

        teacher_item[
            "free_card_count"
        ] += free_card_summary[
            "free_card_count"
        ]

        teacher_item[
            "waived_amount"
        ] += waived_amount

        teacher_item[
            "collected"
        ] += collected

        teacher_item[
            "outstanding"
        ] += outstanding

        teacher_item[
            "teacher_income"
        ] += teacher_income

        teacher_item[
            "center_income"
        ] += center_income

        teacher_students = conn.execute(
            """
            SELECT DISTINCT student_id

            FROM student_classes

            WHERE
                class_code = ?

            AND created_at <= ?

            AND (
                ended_at IS NULL
                OR ended_at > ?
            )
            """,
            (
                class_code,
                month_end,
                month_end,
            ),
        ).fetchall()

        for student in teacher_students:

            teacher_item[
                "students"
            ].add(
                student[
                    "student_id"
                ]
            )


    start_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT student_id
        )

        FROM student_classes

        WHERE
            created_at < ?

        AND (
            ended_at IS NULL
            OR ended_at >= ?
        )
        """,
        (
            month_start,
            month_start,
        ),
    ).fetchone()[0]


    end_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT student_id
        )

        FROM student_classes

        WHERE
            created_at <= ?

        AND (
            ended_at IS NULL
            OR ended_at > ?
        )
        """,
        (
            month_end,
            month_end,
        ),
    ).fetchone()[0]


    new_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT student_id
        )

        FROM student_classes

        WHERE
            created_at
                BETWEEN ? AND ?
        """,
        (
            month_start,
            month_end,
        ),
    ).fetchone()[0]


    left_students = conn.execute(
        """
        SELECT COUNT(
            DISTINCT student_id
        )

        FROM student_classes

        WHERE
            ended_at
                BETWEEN ? AND ?
        """,
        (
            month_start,
            month_end,
        ),
    ).fetchone()[0]


    student_growth = (
        end_students
        -
        start_students
    )

    student_growth_percent = 0

    if start_students > 0:

        student_growth_percent = round(
            (
                student_growth
                /
                start_students
            )
            *
            100,
            1,
        )


    attendance_rate = (
        round(
            total_attendance_records
            /
            total_possible_attendance
            *
            100,
            1,
        )

        if total_possible_attendance > 0

        else 0
    )

    attendance_rate = min(
        attendance_rate,
        100,
    )


    collection_rate = (
        round(
            total_collected
            /
            total_expected
            *
            100,
            1,
        )

        if total_expected > 0

        else 0
    )


    teacher_reports = []

    for teacher_item in teacher_map.values():

        teacher_item[
            "student_count"
        ] = len(
            teacher_item[
                "students"
            ]
        )

        del teacher_item[
            "students"
        ]

        teacher_reports.append(
            teacher_item
        )

    teacher_reports.sort(
        key=lambda item:
            item[
                "collected"
            ],

        reverse=True,
    )


    free_card_rows = conn.execute(
        """
        SELECT
            free_cards.*,
            students.student_code,
            students.full_name AS student_name,
            classes.class_name,
            teachers.full_name AS teacher_name
        FROM free_cards
        JOIN students
            ON students.id = free_cards.student_id
        JOIN classes
            ON classes.class_code = free_cards.class_code
        JOIN teachers
            ON teachers.id = classes.teacher_id
        WHERE free_cards.fee_month = ?
        AND free_cards.status = 'APPROVED'
        ORDER BY teachers.full_name, classes.class_name, students.full_name
        """,
        (
            month_key,
        ),
    ).fetchall()


    return {
        "financial": {
            "expected":
                total_expected,

            "collected":
                total_collected,

            "outstanding":
                total_outstanding,

            "teacher_income":
                total_teacher_income,

            "center_income":
                total_center_income,

            "collection_rate":
                collection_rate,

            "free_card_count":
                total_free_card_count,

            "waived_amount":
                total_waived_amount,
        },

        "students": {
            "start":
                start_students,

            "new":
                new_students,

            "left":
                left_students,

            "end":
                end_students,

            "growth":
                student_growth,

            "growth_percent":
                student_growth_percent,
        },

        "attendance": {
            "records":
                total_attendance_records,

            "class_days":
                total_class_days,

            "rate":
                attendance_rate,
        },

        "classes":
            class_reports,

        "teachers":
            teacher_reports,

        "free_cards": [
            dict(row)
            for row in free_card_rows
        ],
    }


# =========================================================
# REPORTS
# =========================================================

@app.route("/reports")
def reports():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    conn = get_db()

    report = build_monthly_report(
        conn,
        selected_month,
    )

    conn.close()

    return render_template(
        "reports.html",

        report=
            report,

        selected_month=
            selected_month,

        month_label=
            datetime.strptime(
                selected_month,
                "%Y-%m"
            ).strftime(
                "%B %Y"
            ),

        previous_month=
            shift_month(
                selected_month,
                -1,
            ),

        next_month=
            shift_month(
                selected_month,
                1,
            ),
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route(
    "/reports/export/excel"
)
def export_reports_excel():

    selected_month = (
        request.args.get(
            "month"
        )

        or

        datetime.now().strftime(
            "%Y-%m"
        )
    )

    if not valid_month(
        selected_month
    ):

        selected_month = (
            datetime.now().strftime(
                "%Y-%m"
            )
        )

    conn = get_db()

    report = build_monthly_report(
        conn,
        selected_month,
    )

    conn.close()

    month_label = (
        datetime.strptime(
            selected_month,
            "%Y-%m"
        ).strftime(
            "%B %Y"
        )
    )

    workbook = Workbook()

    # =====================================================
    # SUMMARY
    # =====================================================

    summary_sheet = (
        workbook.active
    )

    summary_sheet.title = (
        "Summary"
    )

    summary_sheet[
        "A1"
    ] = (
        "Sri Sumana Education Center"
    )

    summary_sheet[
        "A2"
    ] = (
        "Owner Monthly Report"
    )

    summary_sheet[
        "A3"
    ] = month_label

    summary_sheet[
        "A1"
    ].font = Font(
        bold=True,
        size=16,
    )

    summary_sheet[
        "A2"
    ].font = Font(
        bold=True,
        size=13,
    )

    summary_sheet.append(
        []
    )

    summary_sheet.append(
        [
            "Financial Summary",
            "Value",
        ]
    )

    summary_sheet.append(
        [
            "Expected Revenue",
            report[
                "financial"
            ][
                "expected"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Monthly Collection",
            report[
                "financial"
            ][
                "collected"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Outstanding",
            report[
                "financial"
            ][
                "outstanding"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Teacher Income",
            report[
                "financial"
            ][
                "teacher_income"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Center Income",
            report[
                "financial"
            ][
                "center_income"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Collection Rate %",
            report[
                "financial"
            ][
                "collection_rate"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Approved Free Cards",
            report[
                "financial"
            ][
                "free_card_count"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Waived Amount",
            report[
                "financial"
            ][
                "waived_amount"
            ],
        ]
    )

    summary_sheet.append(
        []
    )

    summary_sheet.append(
        [
            "Student Growth",
            "Value",
        ]
    )

    summary_sheet.append(
        [
            "Start of Month",
            report[
                "students"
            ][
                "start"
            ],
        ]
    )

    summary_sheet.append(
        [
            "New Students",
            report[
                "students"
            ][
                "new"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Students Left",
            report[
                "students"
            ][
                "left"
            ],
        ]
    )

    summary_sheet.append(
        [
            "End of Month",
            report[
                "students"
            ][
                "end"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Net Growth",
            report[
                "students"
            ][
                "growth"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Growth %",
            report[
                "students"
            ][
                "growth_percent"
            ],
        ]
    )

    summary_sheet.append(
        []
    )

    summary_sheet.append(
        [
            "Attendance",
            "Value",
        ]
    )

    summary_sheet.append(
        [
            "Attendance Records",
            report[
                "attendance"
            ][
                "records"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Class Days",
            report[
                "attendance"
            ][
                "class_days"
            ],
        ]
    )

    summary_sheet.append(
        [
            "Attendance Rate %",
            report[
                "attendance"
            ][
                "rate"
            ],
        ]
    )


    # =====================================================
    # TEACHERS SHEET
    # =====================================================

    teacher_sheet = (
        workbook.create_sheet(
            "Teachers"
        )
    )

    teacher_sheet.append(
        [
            "Teacher Code",
            "Teacher",
            "Subject",
            "Classes",
            "Students",
            "Expected",
            "Free Cards",
            "Waived",
            "Collected",
            "Outstanding",
            "Teacher Income",
            "Center Income",
        ]
    )

    for teacher in report[
        "teachers"
    ]:

        teacher_sheet.append(
            [
                teacher[
                    "teacher_code"
                ],

                teacher[
                    "teacher_name"
                ],

                teacher[
                    "subject"
                ],

                teacher[
                    "classes"
                ],

                teacher[
                    "student_count"
                ],

                teacher[
                    "expected"
                ],

                teacher[
                    "free_card_count"
                ],

                teacher[
                    "waived_amount"
                ],

                teacher[
                    "collected"
                ],

                teacher[
                    "outstanding"
                ],

                teacher[
                    "teacher_income"
                ],

                teacher[
                    "center_income"
                ],
            ]
        )


    # =====================================================
    # CLASSES SHEET
    # =====================================================

    class_sheet = (
        workbook.create_sheet(
            "Classes"
        )
    )

    class_sheet.append(
        [
            "Class Code",
            "Class",
            "Teacher",
            "Status",
            "Students",
            "Growth",
            "Monthly Fee",
            "Expected",
            "Free Cards",
            "Waived",
            "Collected",
            "Outstanding",
            "Collection %",
            "Teacher Income",
            "Center Income",
            "Attendance %",
            "Attendance Records",
        ]
    )

    for class_item in report[
        "classes"
    ]:

        class_sheet.append(
            [
                class_item[
                    "class_code"
                ],

                class_item[
                    "class_name"
                ],

                class_item[
                    "teacher_name"
                ],

                class_item[
                    "status"
                ],

                class_item[
                    "end_students"
                ],

                class_item[
                    "growth"
                ],

                class_item[
                    "monthly_fee"
                ],

                class_item[
                    "expected"
                ],

                class_item[
                    "free_card_count"
                ],

                class_item[
                    "waived_amount"
                ],

                class_item[
                    "collected"
                ],

                class_item[
                    "outstanding"
                ],

                class_item[
                    "collection_rate"
                ],

                class_item[
                    "teacher_income"
                ],

                class_item[
                    "center_income"
                ],

                class_item[
                    "attendance_rate"
                ],

                class_item[
                    "attendance_records"
                ],
            ]
        )


    # =====================================================
    # FREE CARDS SHEET
    # =====================================================

    free_card_sheet = workbook.create_sheet(
        "Free Cards"
    )

    free_card_sheet.append(
        [
            "Student ID",
            "Student",
            "Class",
            "Teacher",
            "Fee Month",
            "Waived Amount",
            "Requested By",
            "Requested At",
            "Approved By",
            "Approved At",
            "Reason",
        ]
    )

    for free_card in report[
        "free_cards"
    ]:
        free_card_sheet.append(
            [
                free_card[
                    "student_code"
                ],
                free_card[
                    "student_name"
                ],
                free_card[
                    "class_name"
                ],
                free_card[
                    "teacher_name"
                ],
                free_card[
                    "fee_month"
                ],
                free_card[
                    "waived_amount"
                ],
                free_card[
                    "requested_by_name"
                ],
                free_card[
                    "requested_at"
                ],
                free_card[
                    "reviewed_by_name"
                ],
                free_card[
                    "reviewed_at"
                ],
                free_card[
                    "reason"
                ],
            ]
        )


    # =====================================================
    # EXCEL FORMATTING
    # =====================================================

    for sheet in workbook.worksheets:

        for cell in sheet[1]:

            cell.font = Font(
                bold=True
            )

        for column_cells in sheet.columns:

            max_length = 0

            column_letter = (
                column_cells[
                    0
                ].column_letter
            )

            for cell in column_cells:

                cell.alignment = Alignment(
                    vertical=
                        "center"
                )

                value = (
                    ""
                    if cell.value is None
                    else str(
                        cell.value
                    )
                )

                max_length = max(
                    max_length,
                    len(value),
                )

            sheet.column_dimensions[
                column_letter
            ].width = min(
                max(
                    max_length + 3,
                    12,
                ),
                35,
            )


    for row_number in range(
        6,
        13,
    ):

        summary_sheet[
            f"B{row_number}"
        ].number_format = (
            '#,##0.00'
        )


    for row_number in range(
        2,
        teacher_sheet.max_row + 1,
    ):

        for column_letter in [
            "F",
            "H",
            "I",
            "J",
            "K",
            "L",
        ]:

            teacher_sheet[
                f"{column_letter}{row_number}"
            ].number_format = (
                '#,##0.00'
            )


    for row_number in range(
        2,
        class_sheet.max_row + 1,
    ):

        for column_letter in [
            "G",
            "H",
            "J",
            "K",
            "L",
            "N",
            "O",
        ]:

            class_sheet[
                f"{column_letter}{row_number}"
            ].number_format = (
                '#,##0.00'
            )


    for row_number in range(
        2,
        free_card_sheet.max_row + 1,
    ):
        free_card_sheet[
            f"F{row_number}"
        ].number_format = '#,##0.00'


    output = io.BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    filename = (
        "Sri-Sumana-Report-"
        f"{selected_month}.xlsx"
    )

    return send_file(
        output,

        as_attachment=True,

        download_name=
            filename,

        mimetype=(
            "application/"
            "vnd.openxmlformats-"
            "officedocument."
            "spreadsheetml.sheet"
        ),
    )


# =========================================================
# SECURITY CENTER
# OWNER ONLY
# =========================================================

@app.route("/security")
def security_center():

    conn = get_db()

    now_epoch = int(time.time())
    fifteen_minutes_ago = now_epoch - 900

    owner_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'OWNER'
        AND status = 'ACTIVE'
        """
    ).fetchone()[0]

    active_staff_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'STAFF'
        AND status = 'ACTIVE'
        """
    ).fetchone()[0]

    inactive_staff_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM app_users
        WHERE role = 'STAFF'
        AND status != 'ACTIVE'
        """
    ).fetchone()[0]

    login_failure_events_15m = conn.execute(
        """
        SELECT COUNT(*)
        FROM rate_limit_events
        WHERE scope IN (
            'login_failure_ip',
            'login_failure_user'
        )
        AND created_at_epoch >= ?
        """,
        (fifteen_minutes_ago,),
    ).fetchone()[0]

    login_ip_events_15m = conn.execute(
        """
        SELECT COUNT(*)
        FROM rate_limit_events
        WHERE scope = 'login_failure_ip'
        AND created_at_epoch >= ?
        """,
        (fifteen_minutes_ago,),
    ).fetchone()[0]

    login_user_events_15m = conn.execute(
        """
        SELECT COUNT(*)
        FROM rate_limit_events
        WHERE scope = 'login_failure_user'
        AND created_at_epoch >= ?
        """,
        (fifteen_minutes_ago,),
    ).fetchone()[0]

    security_rows = conn.execute(
        """
        SELECT
            id,
            actor_username,
            actor_full_name,
            actor_role,
            action,
            entity_type,
            entity_id,
            request_method,
            request_path,
            ip_address,
            status_code,
            created_at
        FROM audit_logs
        WHERE
            entity_type IN (
                'AUTH',
                'ACCOUNT',
                'STAFF_ACCESS',
                'SYSTEM'
            )
            OR action LIKE '%LOGIN%'
            OR action LIKE '%LOGOUT%'
            OR action LIKE '%PASSWORD%'
            OR action LIKE '%STAFF%'
        ORDER BY id DESC
        LIMIT 10
        """
    ).fetchall()

    recent_security_activity = [
        dict(row)
        for row in security_rows
    ]

    conn.close()

    backup_items = []

    if BACKUP_DIR.exists():
        backup_paths = sorted(
            BACKUP_DIR.glob("sri_sumana_*.db"),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        )

        for backup_path in backup_paths[:10]:
            stat = backup_path.stat()

            backup_items.append(
                {
                    "name": backup_path.name,
                    "size_mb": round(
                        stat.st_size / (1024 * 1024),
                        2,
                    ),
                    "modified_at": datetime.fromtimestamp(
                        stat.st_mtime
                    ).strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    latest_backup = (
        backup_items[0]
        if backup_items
        else None
    )

    database_exists = DB_PATH.exists()
    database_size_mb = 0

    if database_exists:
        database_size_mb = round(
            DB_PATH.stat().st_size / (1024 * 1024),
            2,
        )

    session_lifetime = app.config.get(
        "PERMANENT_SESSION_LIFETIME"
    )

    session_hours = 0

    if session_lifetime:
        session_hours = round(
            session_lifetime.total_seconds() / 3600,
            1,
        )

    security_status = {
        "https_enabled": bool(HTTPS_ENABLED),
        "request_secure": bool(request.is_secure),
        "trust_proxy": bool(TRUST_PROXY),
        "cookie_secure": bool(
            app.config.get("SESSION_COOKIE_SECURE")
        ),
        "cookie_httponly": bool(
            app.config.get("SESSION_COOKIE_HTTPONLY")
        ),
        "cookie_samesite": app.config.get(
            "SESSION_COOKIE_SAMESITE",
            "Not Set",
        ),
        "session_hours": session_hours,
        "csrf_same_origin": True,
        "rate_limiting": True,
        "audit_logging": True,
        "security_headers": True,
        "database_exists": database_exists,
        "backup_retention": BACKUP_RETENTION_COUNT,
        "database_size_mb": database_size_mb,
    }

    account_stats = {
        "owners": int(owner_count or 0),
        "active_staff": int(active_staff_count or 0),
        "inactive_staff": int(inactive_staff_count or 0),
    }

    login_stats = {
        "events_15m": int(login_failure_events_15m or 0),
        "ip_events_15m": int(login_ip_events_15m or 0),
        "user_events_15m": int(login_user_events_15m or 0),
        "ip_limit": 20,
        "user_limit": 8,
        "window_minutes": 15,
    }

    return render_template(
        "security.html",
        security_status=security_status,
        account_stats=account_stats,
        login_stats=login_stats,
        backup_items=backup_items,
        latest_backup=latest_backup,
        recent_security_activity=recent_security_activity,
    )


@app.post("/security/backup")
def security_create_backup():

    try:
        backup_path = create_database_backup(
            force=True
        )

        if not backup_path:
            flash(
                "Database backup could not be created.",
                "error",
            )
        else:
            flash(
                (
                    "Database backup created successfully: "
                    + backup_path.name
                ),
                "success",
            )

    except Exception as error:
        print(
            "SECURITY BACKUP ERROR:",
            error,
        )

        flash(
            "Database backup failed.",
            "error",
        )

    return redirect(
        url_for("security_center")
    )


# =========================================================
# AUDIT LOGS
# OWNER ONLY
# =========================================================

@app.route(
    "/audit-logs"
)
def audit_logs():

    search = request.args.get(
        "search",
        "",
    ).strip()

    action_filter = request.args.get(
        "action",
        "ALL",
    ).strip()

    entity_filter = request.args.get(
        "entity",
        "ALL",
    ).strip()

    role_filter = request.args.get(
        "role",
        "ALL",
    ).strip().upper()

    date_filter = request.args.get(
        "date",
        "",
    ).strip()

    if role_filter not in {
        "ALL",
        "OWNER",
        "STAFF",
    }:
        role_filter = "ALL"

    conn = get_db()

    sql = """
        SELECT *
        FROM audit_logs
        WHERE 1 = 1
    """

    params = []

    if action_filter != "ALL":
        sql += """
            AND action = ?
        """
        params.append(action_filter)

    if entity_filter != "ALL":
        sql += """
            AND entity_type = ?
        """
        params.append(entity_filter)

    if role_filter != "ALL":
        sql += """
            AND actor_role = ?
        """
        params.append(role_filter)

    if date_filter:
        sql += """
            AND substr(created_at, 1, 10) = ?
        """
        params.append(date_filter)

    if search:
        search_value = f"%{search}%"

        sql += """
            AND
            (
                actor_username LIKE ?
                OR actor_full_name LIKE ?
                OR action LIKE ?
                OR entity_type LIKE ?
                OR entity_id LIKE ?
                OR metadata_json LIKE ?
            )
        """

        params.extend(
            [
                search_value,
                search_value,
                search_value,
                search_value,
                search_value,
                search_value,
            ]
        )

    sql += """
        ORDER BY id DESC
        LIMIT 500
    """

    rows = conn.execute(
        sql,
        params,
    ).fetchall()

    action_rows = conn.execute(
        """
        SELECT DISTINCT action
        FROM audit_logs
        WHERE action IS NOT NULL
        ORDER BY action
        """
    ).fetchall()

    entity_rows = conn.execute(
        """
        SELECT DISTINCT entity_type
        FROM audit_logs
        WHERE entity_type IS NOT NULL
        ORDER BY entity_type
        """
    ).fetchall()

    total_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        """
    ).fetchone()[0]

    today_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        WHERE substr(created_at, 1, 10) = ?
        """,
        (
            datetime.now().strftime(
                "%Y-%m-%d"
            ),
        ),
    ).fetchone()[0]

    owner_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        WHERE actor_role = 'OWNER'
        """
    ).fetchone()[0]

    staff_count = conn.execute(
        """
        SELECT COUNT(*)
        FROM audit_logs
        WHERE actor_role = 'STAFF'
        """
    ).fetchone()[0]

    conn.close()

    logs = []

    for row in rows:
        item = dict(row)
        metadata = {}

        if item.get("metadata_json"):
            try:
                metadata = json.loads(
                    item["metadata_json"]
                )
            except (
                json.JSONDecodeError,
                TypeError,
            ):
                metadata = {}

        item["metadata"] = metadata
        logs.append(item)

    stats = {
        "total": total_count,
        "today": today_count,
        "owner": owner_count,
        "staff": staff_count,
    }

    return render_template(
        "audit_logs.html",
        logs=logs,
        stats=stats,
        search=search,
        action_filter=action_filter,
        entity_filter=entity_filter,
        role_filter=role_filter,
        date_filter=date_filter,
        action_options=[
            row["action"]
            for row in action_rows
        ],
        entity_options=[
            row["entity_type"]
            for row in entity_rows
        ],
    )


# =========================================================
# ROUTE ARCHITECTURE SELF-CHECK
# =========================================================

def validate_route_architecture():
    expected_routes = {
        "/register": "public_register_student",
        "/admin/students/new": "register_student",
        "/dashboard": "dashboard",
        "/teacher/dashboard": "teacher_dashboard",
    }

    actual_routes = {
        rule.rule: rule.endpoint
        for rule in app.url_map.iter_rules()
    }

    errors = []

    for path, endpoint in expected_routes.items():
        actual_endpoint = actual_routes.get(path)

        if actual_endpoint != endpoint:
            errors.append(
                f"{path} must map to {endpoint!r}, "
                f"but maps to {actual_endpoint!r}."
            )

    if "public_register_student" not in PUBLIC_ENDPOINTS:
        errors.append(
            "Public student registration endpoint is not marked public."
        )

    if "register_student" in PUBLIC_ENDPOINTS:
        errors.append(
            "Admin manual registration endpoint must not be public."
        )

    if "register_student" not in STAFF_ENDPOINTS:
        errors.append(
            "Staff must retain access to admin manual registration."
        )

    if errors:
        raise RuntimeError(
            "Route architecture safety check failed:\n- "
            + "\n- ".join(errors)
        )


# =========================================================
# RUN APP
# =========================================================

if __name__ == "__main__":

    init_db()
    validate_route_architecture()

    try:
        backup_path = create_database_backup()

        if backup_path:
            print(
                "Database backup ready:",
                backup_path,
            )
    except Exception as error:
        print(
            "BACKUP ERROR:",
            error,
        )

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False,
    )
