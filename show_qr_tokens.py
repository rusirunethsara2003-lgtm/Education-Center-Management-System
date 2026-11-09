import sqlite3

DB_PATH = "sri_sumana_demo.db"

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row

rows = conn.execute(
    """
    SELECT
        student_code,
        full_name,
        qr_token
    FROM students
    WHERE qr_status = 'ACTIVE'
    ORDER BY student_code
    """
).fetchall()

conn.close()

if not rows:
    print("No active QR tokens found.")
else:
    print()
    print("ACTIVE STUDENT QR VALUES")
    print("=" * 70)

    for row in rows:
        print(
            row["student_code"],
            "|",
            row["full_name"],
            "|",
            "SSCQR:" + row["qr_token"],
        )