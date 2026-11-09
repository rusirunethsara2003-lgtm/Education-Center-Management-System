from pathlib import Path
from datetime import datetime
import shutil
import sys

ROOT = Path(__file__).resolve().parent
# If this updater is copied to the project root it works directly.
# If run from elsewhere, pass the project folder as first argument.
if len(sys.argv) > 1:
    PROJECT = Path(sys.argv[1]).expanduser().resolve()
else:
    PROJECT = Path.cwd().resolve()

APP = PROJECT / "app.py"
SCAN = PROJECT / "templates" / "scan.html"

if not APP.is_file():
    raise SystemExit(f"app.py not found: {APP}")
if not SCAN.is_file():
    raise SystemExit(f"scan.html not found: {SCAN}")

stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(APP, APP.with_name(f"app.before_outstanding_{stamp}.py"))
shutil.copy2(SCAN, SCAN.with_name(f"scan.before_outstanding_{stamp}.html"))

app = APP.read_text(encoding="utf-8")
scan = SCAN.read_text(encoding="utf-8")

# =========================================================
# APP.PY
# =========================================================
helper = r'''

def get_student_class_outstanding_summary(
    conn,
    student_id,
    class_code,
    through_month=None,
):
    """Return unpaid/partial fee months for one student and class.

    Uses the existing enrollment/payment/free-card rules. Paid months and
    approved Free Card months are excluded from the outstanding total.
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
'''

if "def get_student_class_outstanding_summary(" not in app:
    marker = "def payment_status_text(\n    payment,\n):"
    pos = app.find(marker)
    if pos < 0:
        raise SystemExit("Could not locate payment_status_text() in app.py")
    app = app[:pos] + helper + "\n" + app[pos:]

verify_old = '''        item["payment"] = (\n            payment\n        )\n\n        item["payment_status"] = ('''
verify_new = '''        item["payment"] = (\n            payment\n        )\n\n        item["outstanding"] = (\n            get_student_class_outstanding_summary(\n                conn,\n                student["id"],\n                item["class_code"],\n                fee_month,\n            )\n        )\n\n        item["payment_status"] = ('''
if 'item["outstanding"] = (' not in app:
    if verify_old not in app:
        raise SystemExit("Could not locate QR class payment block in app.py")
    app = app.replace(verify_old, verify_new, 1)

mark_old = '''    payment_text = (\n        payment_status_text(\n            payment\n        )\n    )\n\n    existing = conn.execute('''
mark_new = '''    payment_text = (\n        payment_status_text(\n            payment\n        )\n    )\n\n    outstanding_summary = (\n        get_student_class_outstanding_summary(\n            conn,\n            student["id"],\n            class_code,\n            fee_month,\n        )\n    )\n\n    existing = conn.execute('''
if "outstanding_summary = (" not in app:
    if mark_old not in app:
        raise SystemExit("Could not locate mark_attendance payment block in app.py")
    app = app.replace(mark_old, mark_new, 1)

start = app.find("def mark_attendance():")
end = app.find('@app.route("/attendance")', start)
if start < 0 or end < 0:
    raise SystemExit("Could not isolate mark_attendance() in app.py")
block = app[start:end]

if '"outstanding":\n                    outstanding_summary' not in block:
    old = '''                "month_label":\n                    month_label,\n'''
    new = '''                "month_label":\n                    month_label,\n\n                "outstanding":\n                    outstanding_summary,\n'''
    if block.count(old) < 2:
        raise SystemExit("Could not locate duplicate attendance responses")
    block = block.replace(old, new, 2)

if '"outstanding":\n                outstanding_summary' not in block:
    old = '''            "month_label":\n                month_label,\n'''
    new = '''            "month_label":\n                month_label,\n\n            "outstanding":\n                outstanding_summary,\n'''
    if old not in block:
        raise SystemExit("Could not locate success attendance response")
    block = block.replace(old, new, 1)

app = app[:start] + block + app[end:]

# =========================================================
# SCAN.HTML
# =========================================================
css_marker = '''    .payment-message.free-card {\n        background: var(--blue-soft);\n        color: var(--blue);\n    }\n'''
css_add = r'''

    /* =====================================================
       OUTSTANDING / ARREARS
    ====================================================== */

    .outstanding-panel {
        display: none;
        margin: 0 14px 14px;
        overflow: hidden;
        border: 1px solid #fecaca;
        border-radius: 10px;
        background: #fff7f7;
    }

    .outstanding-panel.show {
        display: block;
    }

    .outstanding-head {
        padding: 11px 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        border-bottom: 1px solid #fecaca;
        background: #fff1f1;
    }

    .outstanding-head strong {
        color: var(--red);
        font-size: 9px;
        font-weight: 850;
    }

    .outstanding-count {
        min-height: 24px;
        padding: 0 8px;
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        background: #ffffff;
        color: var(--red);
        font-size: 7px;
        font-weight: 800;
    }

    .outstanding-list {
        padding: 6px 12px;
    }

    .outstanding-row {
        padding: 9px 0;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        border-bottom: 1px solid #fee2e2;
    }

    .outstanding-row:last-child {
        border-bottom: 0;
    }

    .outstanding-row-copy {
        min-width: 0;
    }

    .outstanding-row-copy strong {
        display: block;
        color: var(--text);
        font-size: 8px;
    }

    .outstanding-row-copy span {
        display: block;
        margin-top: 3px;
        color: var(--muted);
        font-size: 7px;
    }

    .outstanding-row-amount {
        flex: 0 0 auto;
        color: var(--red);
        font-size: 9px;
        font-weight: 850;
    }

    .outstanding-total {
        padding: 11px 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        border-top: 1px solid #fecaca;
        background: #ffffff;
    }

    .outstanding-total span {
        color: var(--muted);
        font-size: 8px;
        font-weight: 700;
    }

    .outstanding-total strong {
        color: var(--red);
        font-size: 13px;
        font-weight: 900;
    }
'''

if ".outstanding-panel {" not in scan:
    if css_marker not in scan:
        raise SystemExit("Could not locate payment message CSS in scan.html")
    scan = scan.replace(css_marker, css_marker + css_add, 1)

html_marker = '''                <div\n                    id="paymentMessage"\n                    class="payment-message unpaid"\n                >\n                    Select a class to view payment status.\n                </div>\n'''
html_add = r'''

                <div
                    id="outstandingPanel"
                    class="outstanding-panel"
                >

                    <div class="outstanding-head">
                        <strong>
                            ⚠ Outstanding Class Fees
                        </strong>

                        <span
                            id="outstandingCount"
                            class="outstanding-count"
                        >
                            0 months
                        </span>
                    </div>

                    <div
                        id="outstandingList"
                        class="outstanding-list"
                    ></div>

                    <div class="outstanding-total">
                        <span>
                            Total Outstanding
                        </span>

                        <strong id="outstandingTotal">
                            Rs. 0
                        </strong>
                    </div>

                </div>
'''

if 'id="outstandingPanel"' not in scan:
    if html_marker not in scan:
        raise SystemExit("Could not locate paymentMessage HTML in scan.html")
    scan = scan.replace(html_marker, html_marker + html_add, 1)

# Insert outstanding renderer before clearPayment()
js_marker = '''    function clearPayment() {\n'''
js_add = r'''
    function renderOutstanding(
        summary
    ) {

        const panel =
            document.getElementById(
                "outstandingPanel"
            );

        const list =
            document.getElementById(
                "outstandingList"
            );

        const count =
            document.getElementById(
                "outstandingCount"
            );

        const total =
            document.getElementById(
                "outstandingTotal"
            );

        const months =
            Array.isArray(
                summary?.months
            )
            ? summary.months
            : [];

        const totalOutstanding =
            Number(
                summary?.total_outstanding
                || 0
            );

        list.innerHTML = "";

        if (
            months.length === 0
            ||
            totalOutstanding <= 0
        ) {
            panel.classList.remove(
                "show"
            );

            count.textContent =
                "0 months";

            total.textContent =
                money(0);

            return;
        }

        months.forEach(
            function(item) {

                const row =
                    document.createElement(
                        "div"
                    );

                row.className =
                    "outstanding-row";

                const copy =
                    document.createElement(
                        "div"
                    );

                copy.className =
                    "outstanding-row-copy";

                const month =
                    document.createElement(
                        "strong"
                    );

                month.textContent =
                    item.month_label
                    ||
                    item.fee_month
                    ||
                    "Fee month";

                const meta =
                    document.createElement(
                        "span"
                    );

                meta.textContent =
                    (
                        String(
                            item.status
                            ||
                            "UNPAID"
                        ).toUpperCase()
                        +
                        " · Paid "
                        +
                        money(
                            item.paid
                            || 0
                        )
                    );

                const amount =
                    document.createElement(
                        "div"
                    );

                amount.className =
                    "outstanding-row-amount";

                amount.textContent =
                    money(
                        item.balance
                        || 0
                    );

                copy.appendChild(month);
                copy.appendChild(meta);
                row.appendChild(copy);
                row.appendChild(amount);
                list.appendChild(row);

            }
        );

        count.textContent =
            (
                months.length
                +
                (
                    months.length === 1
                    ? " month"
                    : " months"
                )
            );

        total.textContent =
            money(
                totalOutstanding
            );

        panel.classList.add(
            "show"
        );

    }


    function clearOutstanding() {

        renderOutstanding(
            {
                months: [],
                total_outstanding: 0
            }
        );

    }


'''

if "function renderOutstanding(" not in scan:
    if js_marker not in scan:
        raise SystemExit("Could not locate clearPayment() in scan.html")
    scan = scan.replace(js_marker, js_add + js_marker, 1)

# renderPayment: render the new summary and show Add Payment for any old due month.
button_old = '''        const addPaymentButton =\n            document.getElementById(\n                "addPaymentButton"\n            );\n\n\n        if (\n            status === "UNPAID"\n            ||\n            status === "PARTIAL"\n        ) {'''
button_new = '''        const outstanding =\n            classItem?.outstanding\n            ||\n            {};\n\n\n        renderOutstanding(\n            outstanding\n        );\n\n\n        const hasOutstanding =\n            Number(\n                outstanding.total_outstanding\n                || 0\n            ) > 0;\n\n\n        const addPaymentButton =\n            document.getElementById(\n                "addPaymentButton"\n            );\n\n\n        if (\n            status === "UNPAID"\n            ||\n            status === "PARTIAL"\n            ||\n            hasOutstanding\n        ) {'''
if "const hasOutstanding =" not in scan:
    if button_old not in scan:
        raise SystemExit("Could not locate Add Payment visibility logic in scan.html")
    scan = scan.replace(button_old, button_new, 1)

clear_old = '''        document.getElementById(\n            "addPaymentButton"\n        ).classList.remove(\n            "show"\n        );\n\n    }\n'''
clear_new = '''        document.getElementById(\n            "addPaymentButton"\n        ).classList.remove(\n            "show"\n        );\n\n\n        clearOutstanding();\n\n    }\n'''
# Limit replacement to clearPayment() area.
cp_start = scan.find("    function clearPayment() {")
cp_end = scan.find("    /* =====================================================\n       QR VERIFY", cp_start)
if cp_start < 0 or cp_end < 0:
    raise SystemExit("Could not isolate clearPayment() in scan.html")
cp_block = scan[cp_start:cp_end]
if "clearOutstanding();" not in cp_block:
    if clear_old not in cp_block:
        raise SystemExit("Could not patch clearPayment() in scan.html")
    cp_block = cp_block.replace(clear_old, clear_new, 1)
scan = scan[:cp_start] + cp_block + scan[cp_end:]

# Result payment button should also remain available when old balances exist.
call_old = '''            renderResultPayment(\n                payment\n            );'''
call_new = '''            if (\n                classItem\n                &&\n                data.outstanding\n            ) {\n                classItem.outstanding =\n                    data.outstanding;\n            }\n\n\n            renderResultPayment(\n                payment,\n                classItem\n            );'''
if "renderResultPayment(\n                payment,\n                classItem" not in scan:
    if call_old not in scan:
        raise SystemExit("Could not locate renderResultPayment call")
    scan = scan.replace(call_old, call_new, 1)

sig_old = '''    function renderResultPayment(\n        payment\n    ) {'''
sig_new = '''    function renderResultPayment(\n        payment,\n        classItem\n    ) {'''
if sig_new not in scan:
    if sig_old not in scan:
        raise SystemExit("Could not locate renderResultPayment function")
    scan = scan.replace(sig_old, sig_new, 1)

result_vis_old = '''        if (\n            status === "UNPAID"\n            ||\n            status === "PARTIAL"\n        ) {\n\n            resultAddPaymentButton.classList.add('''
result_vis_new = '''        const hasOutstanding =\n            Number(\n                classItem?.outstanding?.total_outstanding\n                || 0\n            ) > 0;\n\n\n        if (\n            status === "UNPAID"\n            ||\n            status === "PARTIAL"\n            ||\n            hasOutstanding\n        ) {\n\n            resultAddPaymentButton.classList.add('''
# This exact block appears once in renderResultPayment, after prior patch adds another hasOutstanding earlier.
rp_start = scan.find("    function renderResultPayment(")
rp_end = scan.find("    /* =====================================================\n       ADD PAYMENT", rp_start)
rp_block = scan[rp_start:rp_end]
if "classItem?.outstanding?.total_outstanding" not in rp_block:
    if result_vis_old not in rp_block:
        raise SystemExit("Could not patch result Add Payment visibility")
    rp_block = rp_block.replace(result_vis_old, result_vis_new, 1)
scan = scan[:rp_start] + rp_block + scan[rp_end:]

# openPaymentForm: pay oldest outstanding month first.
op_start = scan.find("    function openPaymentForm() {")
op_end = scan.find("    /* =====================================================\n       SCAN NEXT", op_start)
if op_start < 0 or op_end < 0:
    raise SystemExit("Could not isolate openPaymentForm()")
op_block = scan[op_start:op_end]

old_logic = '''        if (\n            status !== "UNPAID"\n            &&\n            status !== "PARTIAL"\n        ) {\n\n            return;\n\n        }\n\n\n        const balance =\n            Number(\n                payment.balance\n                ??\n                classItem.balance\n                ??\n                0\n            );'''
new_logic = '''        const outstanding =\n            classItem.outstanding\n            ||\n            {};\n\n\n        const dueMonths =\n            Array.isArray(\n                outstanding.months\n            )\n            ? outstanding.months\n            : [];\n\n\n        const oldestDue =\n            dueMonths.length > 0\n            ? dueMonths[0]\n            : null;\n\n\n        if (\n            !oldestDue\n            &&\n            status !== "UNPAID"\n            &&\n            status !== "PARTIAL"\n        ) {\n\n            return;\n\n        }\n\n\n        const targetMonth =\n            oldestDue?.fee_month\n            ||\n            verifiedData.fee_month\n            ||\n            "";\n\n\n        const balance =\n            Number(\n                oldestDue?.balance\n                ??\n                payment.balance\n                ??\n                classItem.balance\n                ??\n                0\n            );'''
if "const oldestDue =" not in op_block:
    if old_logic not in op_block:
        raise SystemExit("Could not patch old outstanding payment target logic")
    op_block = op_block.replace(old_logic, new_logic, 1)

month_old = '''        params.set(\n            "fee_month",\n            String(\n                verifiedData.fee_month\n                ||\n                ""\n            )\n        );'''
month_new = '''        params.set(\n            "fee_month",\n            String(\n                targetMonth\n            )\n        );'''
if "String(\n                targetMonth" not in op_block:
    if month_old not in op_block:
        raise SystemExit("Could not patch fee_month parameter")
    op_block = op_block.replace(month_old, month_new, 1)
scan = scan[:op_start] + op_block + scan[op_end:]

APP.write_text(app, encoding="utf-8")
SCAN.write_text(scan, encoding="utf-8")

print("Outstanding fee update applied successfully.")
print("Backups created:")
print(" -", APP.with_name(f"app.before_outstanding_{stamp}.py"))
print(" -", SCAN.with_name(f"scan.before_outstanding_{stamp}.html"))
print("Database was not changed.")
