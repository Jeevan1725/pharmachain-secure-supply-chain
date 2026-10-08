"""
PharmaChain - SSE End Semester Report Generator (Styled)
Builds a professionally formatted Word document.
Run: python scripts/build_report.py
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOC_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "PharmaChain_SSE_Report.docx")
)

# ---- Theme colors ----
PRIMARY = RGBColor(0x1F, 0x3A, 0x5F)       # deep navy
ACCENT = RGBColor(0x2E, 0x74, 0xB5)        # medium blue
CODE_FG = RGBColor(0x1E, 0x1E, 0x1E)       # near-black
CODE_BG = "F2F2F2"                          # light grey shading
HDR_BG = "1F3A5F"                           # header fill (navy)
HDR_FG = RGBColor(0xFF, 0xFF, 0xFF)        # header text (white)
ZEBRA_BG = "F7FAFD"                         # zebra row


# ---------- helpers ----------

def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def style_table(table, header=True, zebra=True):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = 1  # center
            if header and r_idx == 0:
                shade(cell, HDR_BG)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = HDR_FG
                        run.font.size = Pt(10.5)
            elif zebra and r_idx % 2 == 0:
                shade(cell, ZEBRA_BG)


def add_code_block(doc, code_text):
    """Insert a monospace, shaded code paragraph."""
    for line in code_text.split("\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = CODE_FG
        # shading
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), CODE_BG)
        pPr.append(shd)
    # small gap after code block
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def style_heading(paragraph, level=1):
    for run in paragraph.runs:
        run.font.name = "Calibri"
        if level == 1:
            run.font.size = Pt(20)
            run.font.color.rgb = PRIMARY
            run.font.bold = True
        elif level == 2:
            run.font.size = Pt(15)
            run.font.color.rgb = ACCENT
            run.font.bold = True
        elif level == 3:
            run.font.size = Pt(12.5)
            run.font.color.rgb = PRIMARY
            run.font.bold = True


def h1(doc, text):
    p = doc.add_heading(text, level=1)
    style_heading(p, 1)
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    return p


def h2(doc, text):
    p = doc.add_heading(text, level=2)
    style_heading(p, 2)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    return p


def h3(doc, text):
    p = doc.add_heading(text, level=3)
    style_heading(p, 3)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    return p


def add_page_numbers(doc):
    """Add 'Page X of Y' in footer."""
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def field(instruction):
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), instruction)
        return fld

    p.add_run("Page ")
    p._p.append(field("PAGE"))
    p.add_run(" of ")
    p._p.append(field("NUMPAGES"))


def set_base_font(doc):
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(6)


# ---------- document scaffolding ----------

def build_title_page(doc):
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(80)
    r = t.add_run("PharmaChain")
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = PRIMARY

    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run("Secure Pharmaceutical Supply Chain Tracking System")
    r.font.size = Pt(16)
    r.font.color.rgb = ACCENT

    doc.add_paragraph()
    doc.add_paragraph()

    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run("Course: 24CYS401 - Secure Software Engineering\n").bold = True
    info.add_run("Examination: End Semester Laboratory Examination\n").bold = True
    info.add_run("Domain: Pharmaceutical Supply Chain\n").bold = True
    info.add_run("Team: Jeeva\n").bold = True

    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    tag = doc.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tag.add_run('"Trusted at every step - from factory to patient."')
    r.italic = True
    r.font.size = Pt(12)
    r.font.color.rgb = PRIMARY

    doc.add_page_break()


def build_toc_and_thread(doc):
    h1(doc, "Table of Contents")
    doc.add_paragraph(
        "(In Word: right-click below and choose 'Update Field' to populate. "
        "Or use References -> Table of Contents.)"
    ).italic = True
    doc.add_paragraph()
    doc.add_page_break()

    h1(doc, "Global Traceability Thread")
    p = doc.add_paragraph()
    r = p.add_run("SR-03")
    r.bold = True
    r.font.color.rgb = ACCENT
    p.add_run(
        " - Every ownership transfer must be digitally signed, "
        "hash-chained, and audit-logged."
    )
    doc.add_paragraph(
        "This requirement threads through every phase of this report "
        "(requirements -> use case -> DFD -> threat -> attack tree -> "
        "user story -> sprint -> implementation -> test -> deployment)."
    )
    doc.add_page_break()


# ---------- Phase 1 ----------

def add_phase_1(doc):
    h1(doc, "Phase 1 - Agile Process and Development Approach")

    h2(doc, "1.1 Selected Approach: Scrum-XP")
    doc.add_paragraph(
        "Chosen: Scrum with embedded XP (Extreme Programming) practices."
    )

    h3(doc, "Justification")
    for b in [
        "Five independent organizations (Manufacturer, Distributor, Warehouse, Retailer, Customer) operate under FDA DSCSA constraints.",
        "Scrum sprints allow phased cross-organization onboarding (1-2 orgs per sprint).",
        "XP practices (TDD, pair programming, CI, code review) enforce rigor for a security-critical, auditable system.",
        "Regulatory changes (e.g., new FDA serialization rule) enter the backlog as new stories without replanning.",
        "Continuous delivery via CI/CD ships signed, tested Docker images every sprint.",
    ]:
        doc.add_paragraph(b, style="List Bullet")

    h3(doc, "Sprint Configuration")
    for b in [
        "Sprint length: 2 weeks",
        "Roles: Product Owner, Scrum Master, Dev Team (4), Security Champion",
        "Ceremonies: Planning, Daily Scrum, Review, Retrospective, Refinement",
        "XP practices: TDD, Pair Programming, CI, Small Releases, Simple Design",
    ]:
        doc.add_paragraph(b, style="List Bullet")

    h2(doc, "1.2 Agile Manifesto Mapping")
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "#", "Principle", "Application to PharmaChain"
    for r in [
        ("1", "Individuals & interactions over processes/tools",
         "Cross-org daily scrums jointly resolve shipment discrepancies."),
        ("2", "Working software over comprehensive documentation",
         "Each sprint ships a running module deployable in Minikube."),
        ("3", "Customer collaboration over contract negotiation",
         "Retailer + patient feedback in Sprint Review shapes usability."),
        ("4", "Responding to change over following a plan",
         "New FDA rule becomes a new backlog story - no replan."),
        ("5", "Continuous delivery of valuable software",
         "CI/CD delivers signed Docker images on every merge to main."),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = val
    style_table(t)

    h2(doc, "1.3 Refactoring Evidence")

    h3(doc, "Refactoring 1: Ownership Transfer (Monolithic to Layered)")
    doc.add_paragraph("Before (v1 - insecure):")
    add_code_block(doc,
        "@app.post('/transfer')\n"
        "def transfer(product_id, new_owner):\n"
        "    conn = sqlite3.connect('pharmachain.db')\n"
        "    cur = conn.cursor()\n"
        "    cur.execute(f\"SELECT * FROM products WHERE id='{product_id}'\")  # SQLi\n"
        "    product = cur.fetchone()\n"
        "    if product:\n"
        "        cur.execute(f\"UPDATE products SET owner='{new_owner}' WHERE id='{product_id}'\")\n"
        "        conn.commit()  # no auth, no audit, no signature\n"
        "    return {'status': 'ok'}"
    )
    doc.add_paragraph("After (v2 - secure, layered):")
    add_code_block(doc,
        "class OwnershipService:\n"
        "    def transfer(self, product_id, new_owner, actor, signature):\n"
        "        self.authz.require(actor, 'TRANSFER_OWNERSHIP', product_id)\n"
        "        self.signer.verify(actor, product_id, new_owner, signature)\n"
        "        product = self.repo.get(product_id)\n"
        "        with self.repo.transaction():\n"
        "            product.owner = new_owner\n"
        "            self.repo.save(product)\n"
        "            tx_hash = self.ledger.append({...})\n"
        "        self.audit.log(actor, 'TRANSFER', product_id, tx_hash)\n"
        "        return TransferResult(True, tx_hash)"
    )
    p = doc.add_paragraph()
    p.add_run("Improvements: ").bold = True
    p.add_run(
        "Eliminated SQL injection, added RBAC + ABAC, digital signature "
        "(non-repudiation), immutable ledger, audit log, and dependency "
        "injection for testability."
    )

    h3(doc, "Refactoring 2: Serial Uniqueness Validation (DRY)")
    doc.add_paragraph(
        "Before: duplicated uniqueness check in 4 services "
        "(product_registration, shipment_service, warehouse_receipt, delivery)."
    )
    doc.add_paragraph(
        "After: single SerialValidator class with parameterized query, "
        "called by every service. Single point of change, consistent errors, testable."
    )

    h2(doc, "1.4 Limitations of Agile for a Security-Critical System")
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = "Table Grid"
    h = t2.rows[0].cells
    h[0].text, h[1].text = "Limitation / Risk", "Mitigation"
    for r in [
        ("Under-specified security requirements in fast sprints",
         "Add abuse stories + Definition of Done requires STRIDE review"),
        ("Informal cross-org communication may cause API contract drift",
         "Freeze OpenAPI contracts at sprint start; CI validates schema"),
        ("Regulatory compliance demands documentation Agile de-emphasizes",
         "Auto-generated Compliance Traceability Matrix"),
        ("Velocity pressure may tempt skipping security testing",
         "SAST + dependency scan mandatory for Definition of Done"),
        ("Cross-org governance ambiguity - who owns the backlog?",
         "Steering Committee (one rep per org) as PO delegate"),
    ]:
        cells = t2.add_row().cells
        cells[0].text, cells[1].text = r[0], r[1]
    style_table(t2)

    doc.add_page_break()


# ---------- Phase 2 ----------

def add_phase_2(doc):
    h1(doc, "Phase 2 - Requirements Engineering")

    h2(doc, "2.1 Stakeholders and User Types")
    doc.add_paragraph(
        "PharmaChain involves 5 primary actors (one per organization in the "
        "supply chain) plus 3 justified supporting roles required for "
        "regulatory compliance and security."
    )

    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"
    hdr[1].text = "Stakeholder"
    hdr[2].text = "Type"
    hdr[3].text = "Goal"
    hdr[4].text = "Key Concern"
    rows = [
        ("1", "Manufacturer", "Primary Actor",
         "Create and serialize drug products",
         "Anti-counterfeit, batch traceability"),
        ("2", "Distributor", "Primary Actor",
         "Move products between orgs, transfer ownership",
         "Non-repudiation of transfer"),
        ("3", "Warehouse", "Primary Actor",
         "Receive, store, dispatch shipments",
         "Tamper-proof receipts"),
        ("4", "Retailer", "Primary Actor",
         "Receive products, sell to customers",
         "Authenticity verification"),
        ("5", "Customer / Patient", "Primary Actor",
         "Verify drug authenticity",
         "Trust, privacy"),
        ("6", "Auditor / Regulator (FDA)", "Secondary Actor",
         "Inspect immutable history",
         "Compliance, tamper evidence"),
        ("7", "System Administrator", "Supporting Actor",
         "Onboard orgs, manage keys",
         "Least privilege, key lifecycle"),
        ("8", "Security Officer", "Supporting Actor",
         "Monitor anomalies and threats",
         "Real-time alerts, SIEM"),
    ]
    for r in rows:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = val
    style_table(t)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Justification for extra actors: ").bold = True
    p.add_run(
        "PharmaChain is regulated (FDA DSCSA) and cross-organizational. "
        "An Auditor is required for regulatory inspection. An Admin is needed "
        "to onboard/offboard organizations. A Security Officer is required to "
        "detect fraud, key abuse, and anomalous transfers."
    )

    h2(doc, "2.2 Functional Requirements (FR)")
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"
    hdr[1].text = "Requirement"
    hdr[2].text = "Actor"
    hdr[3].text = "Priority"
    frs = [
        ("FR-01", "Register a new product with a globally unique serial number", "Manufacturer", "High"),
        ("FR-02", "Create a shipment linking source org, destination org, product batch", "Manufacturer / Distributor", "High"),
        ("FR-03", "Confirm warehouse receipt with seal/hash verification", "Warehouse", "High"),
        ("FR-04", "Transfer ownership between organizations with signature", "Distributor / Retailer", "High"),
        ("FR-05", "Confirm delivery to retailer / customer", "Retailer", "High"),
        ("FR-06", "Query product status and full history", "Any authenticated actor", "High"),
        ("FR-07", "Verify product authenticity via QR / serial lookup", "Customer", "High"),
        ("FR-08", "View immutable audit trail (read-only)", "Auditor", "High"),
        ("FR-09", "Onboard / offboard an organization with PKI certificate", "Admin", "High"),
        ("FR-10", "Rotate or revoke signing keys", "Admin", "Medium"),
        ("FR-11", "Raise and resolve a dispute between organizations", "Distributor / Retailer", "Medium"),
        ("FR-12", "Trigger product recall (trace backward to source)", "Manufacturer / Auditor", "Medium"),
        ("FR-13", "Receive real-time shipment status notifications", "All logistics actors", "Medium"),
        ("FR-14", "Detect anomalous ownership transfers (rule/ML based)", "Security Officer", "Medium"),
    ]
    for r in frs:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = val
    style_table(t)

    h2(doc, "2.3 Non-Functional Requirements (NFR)")
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"
    hdr[1].text = "Category"
    hdr[2].text = "Requirement"
    hdr[3].text = "Target"
    nfrs = [
        ("NFR-01", "Availability", "System uptime", "99.9% (monthly)"),
        ("NFR-02", "Performance", "Product history lookup latency", "p95 < 1s"),
        ("NFR-03", "Scalability", "Concurrent shipments per day", "10,000+"),
        ("NFR-04", "Scalability", "Organizations onboarded", "100+"),
        ("NFR-05", "Interoperability", "Standards compliance", "GS1 EPCIS 2.0, FDA DSCSA"),
        ("NFR-06", "Privacy", "Customer PII handling", "GDPR + HIPAA-aligned"),
        ("NFR-07", "Maintainability", "Test coverage on critical modules", ">= 80%"),
        ("NFR-08", "Disaster Recovery", "RPO / RTO", "RPO <= 5 min, RTO <= 1 hour"),
        ("NFR-09", "Auditability", "Log retention", "7 years (pharma regulation)"),
        ("NFR-10", "Usability", "Common operation time", "<= 3 clicks for status check"),
    ]
    for r in nfrs:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = val
    style_table(t)

    h2(doc, "2.4 Security Requirements (SR)")
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"
    hdr[1].text = "Category"
    hdr[2].text = "Requirement"
    hdr[3].text = "CIA/Auth"
    srs = [
        ("SR-01", "Authentication", "Multi-factor authentication per user", "Auth"),
        ("SR-02", "Authentication", "mTLS between organizations (PKI)", "Auth"),
        ("SR-03", "Integrity", "Every ownership transfer must be digitally signed, hash-chained, and audit-logged", "I, NR"),
        ("SR-04", "Integrity", "Product records SHA-256 hashed on every change", "I"),
        ("SR-05", "Integrity", "Anti-replay via nonce + timestamp on every signed write", "I"),
        ("SR-06", "Authorization", "Role-Based Access Control (RBAC) per role", "Authz"),
        ("SR-07", "Authorization", "Attribute-Based Access Control (ABAC) per organization", "Authz"),
        ("SR-08", "Confidentiality", "TLS 1.3 in transit, AES-256 at rest", "C"),
        ("SR-09", "Confidentiality", "Field-level encryption for customer PII", "C"),
        ("SR-10", "Audit", "Append-only immutable audit log of every state change", "I, NR"),
        ("SR-11", "Audit", "Logs forwarded to SIEM with tamper-evident hashing", "I"),
        ("SR-12", "Non-Repudiation", "Digital signatures on all state-changing actions", "NR"),
        ("SR-13", "Key Management", "HSM-backed keys, 90-day rotation", "C, I"),
        ("SR-14", "Anomaly Detection", "Alert on duplicate serial, odd-hour transfer, or unusual pattern", "I"),
        ("SR-15", "Availability", "DDoS protection at API Gateway", "A"),
        ("SR-16", "Trust Boundaries", "Cross-org calls only via authenticated, signed APIs", "Auth, I"),
    ]
    for r in srs:
        cells = t.add_row().cells
        for i, val in enumerate(r):
            cells[i].text = val
    style_table(t)

    h2(doc, "2.5 Prioritization (MoSCoW)")
    t = doc.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Priority"
    hdr[1].text = "Requirements"
    moscow = [
        ("Must Have", "FR-01..FR-05, FR-08, FR-09, SR-01..SR-12, NFR-01, NFR-06, NFR-09"),
        ("Should Have", "FR-06, FR-07, FR-10..FR-12, SR-13..SR-16, NFR-02..NFR-05, NFR-07"),
        ("Could Have", "FR-13, FR-14, NFR-08, NFR-10"),
        ("Won't Have (this release)", "Public blockchain anchoring, cross-chain interop"),
    ]
    for r in moscow:
        cells = t.add_row().cells
        cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "2.6 CIA / Auth / Authz / Audit Mapping")
    t = doc.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Category"
    hdr[1].text = "Requirements"
    mapping = [
        ("Confidentiality (C)", "SR-08, SR-09, SR-13 (keys), NFR-06"),
        ("Integrity (I)", "SR-03, SR-04, SR-05, SR-10, SR-11, SR-14"),
        ("Availability (A)", "NFR-01, SR-15"),
        ("Authentication (Auth)", "SR-01, SR-02, SR-16"),
        ("Authorization (Authz)", "SR-06, SR-07"),
        ("Audit (Aud)", "SR-10, SR-11, NFR-09"),
        ("Non-Repudiation (NR)", "SR-03, SR-12"),
    ]
    for r in mapping:
        cells = t.add_row().cells
        cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "2.7 Traceability Thread (Carried Forward)")
    p = doc.add_paragraph()
    r = p.add_run("SR-03")
    r.bold = True
    r.font.color.rgb = ACCENT
    p.add_run(
        " = Every ownership transfer must be digitally signed, "
        "hash-chained, and audit-logged."
    )
    doc.add_paragraph(
        "Mapped to FR-04 (transfer ownership), NFR-09 (log retention), "
        "SR-12 (non-repudiation). Will thread through Use Case, DFD, Threat "
        "Model, Attack Tree, User Story, Code, Test, and Deployment phases."
    )

    h2(doc, "2.8 Assumptions and Constraints")
    h3(doc, "Assumptions")
    for b in [
        "Each organization has PKI-issued certificates from a shared CA.",
        "Organizations are known/trusted (permissioned network, not public).",
        "Customers have smartphones with QR scanning capability.",
    ]:
        doc.add_paragraph(b, style="List Bullet")

    h3(doc, "Constraints")
    for b in [
        "Must comply with FDA DSCSA and GS1 EPCIS 2.0.",
        "Must use HSM for signing key storage (no soft keys in production).",
        "Regulatory log retention: 7 years minimum.",
    ]:
        doc.add_paragraph(b, style="List Bullet")

    doc.add_page_break()

# ---------- main ----------

def main():
    doc = Document()
    set_base_font(doc)
    # margins
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    add_page_numbers(doc)

    build_title_page(doc)
    build_toc_and_thread(doc)
    add_phase_1(doc)
    add_phase_2(doc)
    # Future phases appended here.

    doc.save(DOC_PATH)
    print(f"[OK] Styled report written to: {DOC_PATH}")


if __name__ == "__main__":
    main()
