from pathlib import Path
import sqlite3
import shutil
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "sri_sumana_demo.db"
BACKUP_DIR = BASE_DIR / "backups"

TEST_START = "2026-08-01 00:00:00"

if not DB_PATH.exists():
    raise SystemExit(f"Database not found: {DB_PATH}")

BACKUP_DIR.mkdir(exist_ok=True)
stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_path = BACKUP_DIR / f"sri_sumana_demo_before_outstanding_test_{stamp}.db"
shutil.copy2(DB_PATH, backup_path)

print(f"\nBackup created:\n{backup_path}\n")

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
conn.execute("PRAGMA foreign_keys = ON")

try:
    student_code = input("Enter Student ID / Code (example: S001): ").strip()
    class_code = input("Enter Class Code (example: C003): ").strip()

    student = conn.execute(
        """
        SELECT id, student_code, full_name, status
        FROM students
        WHERE student_code = ?
        """,
        (student_code,),
    ).fetchone()

    if not student:
        raise SystemExit("Student not found.")

    enrollment = conn.execute(
        """
        SELECT
            sc.id,
            sc.student_id,
            sc.class_code,
            sc.created_at,
            sc.enrollment_status,
            sc.ended_at,
            c.class_name,
            c.monthly_fee
        FROM student_classes sc
        JOIN classes c
            ON c.class_code = sc.class_code
        WHERE sc.student_id = ?
          AND sc.class_code = ?
        """,
        (student["id"], class_code),
    ).fetchone()

    if not enrollment:
        raise SystemExit("This student is not enrolled in that class.")

    print("\nSelected:")
    print(f"Student: {student['student_code']} - {student['full_name']}")
    print(f"Class:   {enrollment['class_code']} - {enrollment['class_name']}")
    print(f"Current enrollment date: {enrollment['created_at']}")
    print(f"Monthly fee: Rs. {float(enrollment['monthly_fee'] or 0):,.2f}")

    payments = conn.execute(
        """
        SELECT fee_month, SUM(amount) AS total_paid
        FROM payments
        WHERE student_id = ?
          AND class_code = ?
          AND fee_month BETWEEN '2026-08' AND '2026-10'
          AND status = 'CONFIRMED'
        GROUP BY fee_month
        ORDER BY fee_month
        """,
        (student["id"], class_code),
    ).fetchall()

    if payments:
        print("\nWARNING: Confirmed payments already exist in this test range:")
        for row in payments:
            print(f"  {row['fee_month']}: Rs. {float(row['total_paid'] or 0):,.2f}")
        print("\nThose paid months will NOT appear as unpaid/outstanding.")

    confirm = input(
        f"\nChange only this enrollment start date to {TEST_START}? (yes/no): "
    ).strip().lower()

    if confirm not in {"yes", "y"}:
        print("Cancelled. No database changes made.")
        raise SystemExit(0)

    conn.execute(
        """
        UPDATE student_classes
        SET created_at = ?,
            enrollment_status = COALESCE(enrollment_status, 'ACTIVE'),
            ended_at = NULL
        WHERE id = ?
        """,
        (TEST_START, enrollment["id"]),
    )

    conn.commit()

    updated = conn.execute(
        """
        SELECT created_at, enrollment_status, ended_at
        FROM student_classes
        WHERE id = ?
        """,
        (enrollment["id"],),
    ).fetchone()

    print("\nTEST DATA READY")
    print(f"Student: {student['student_code']} - {student['full_name']}")
    print(f"Class: {enrollment['class_code']} - {enrollment['class_name']}")
    print(f"Enrollment start: {updated['created_at']}")
    print(f"Enrollment status: {updated['enrollment_status']}")
    print("\nNow restart Flask and scan this student's QR.")
    print("Expected: August/September/October unpaid months appear unless already paid/free-card.")

finally:
    conn.close()
