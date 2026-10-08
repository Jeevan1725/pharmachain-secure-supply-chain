"""
PharmaChain - SSE End Semester Report Generator (Styled)
Builds a professionally formatted Word document.
Run: python scripts/build_report.py
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOC_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "PharmaChain_SSE_Report.docx")
)

PRIMARY = RGBColor(0x1F, 0x3A, 0x5F)
ACCENT = RGBColor(0x2E, 0x74, 0xB5)
CODE_FG = RGBColor(0x1E, 0x1E, 0x1E)
CODE_BG = "F2F2F2"
HDR_BG = "1F3A5F"
HDR_FG = RGBColor(0xFF, 0xFF, 0xFF)
ZEBRA_BG = "F7FAFD"


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
            cell.vertical_alignment = 1
            if header and r_idx == 0:
                shade(cell, HDR_BG)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = HDR_FG
                        run.font.size = Pt(10.5)
            elif zebra and r_idx % 2 == 0:
                shade(cell, ZEBRA_BG)


def add_figure(doc, filename, caption):
    img_path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "diagrams", filename)
    )
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=Inches(6.0))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cr = cap.add_run(caption)
        cr.italic = True
        cr.font.size = Pt(10)
        cr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        doc.add_paragraph()
    else:
        warn = doc.add_paragraph()
        warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
        wr = warn.add_run(f"[Missing diagram: diagrams/{filename}]")
        wr.italic = True
        wr.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)


def add_code_block(doc, code_text):
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
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), CODE_BG)
        pPr.append(shd)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def style_heading(paragraph, level=1):
    for run in paragraph.runs:
        run.font.name = "Calibri"
        if level == 1:
            run.font.size = Pt(20); run.font.color.rgb = PRIMARY; run.font.bold = True
        elif level == 2:
            run.font.size = Pt(15); run.font.color.rgb = ACCENT; run.font.bold = True
        elif level == 3:
            run.font.size = Pt(12.5); run.font.color.rgb = PRIMARY; run.font.bold = True


def h1(doc, text):
    p = doc.add_heading(text, level=1); style_heading(p, 1)
    p.paragraph_format.space_before = Pt(18); p.paragraph_format.space_after = Pt(8)
    return p


def h2(doc, text):
    p = doc.add_heading(text, level=2); style_heading(p, 2)
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(6)
    return p


def h3(doc, text):
    p = doc.add_heading(text, level=3); style_heading(p, 3)
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(4)
    return p


def add_page_numbers(doc):
    section = doc.sections[0]
    p = section.footer.paragraphs[0]
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


def build_title_page(doc):
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(80)
    r = t.add_run("PharmaChain"); r.font.size = Pt(40); r.font.bold = True; r.font.color.rgb = PRIMARY
    s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run("Secure Pharmaceutical Supply Chain Tracking System"); r.font.size = Pt(16); r.font.color.rgb = ACCENT
    doc.add_paragraph(); doc.add_paragraph()
    info = doc.add_paragraph(); info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run("Course: 24CYS401 - Secure Software Engineering\n").bold = True
    info.add_run("Examination: End Semester Laboratory Examination\n").bold = True
    info.add_run("Domain: Pharmaceutical Supply Chain\n").bold = True
    info.add_run("Team: Jeeva\n").bold = True
    doc.add_paragraph(); doc.add_paragraph(); doc.add_paragraph()
    tag = doc.add_paragraph(); tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tag.add_run('"Trusted at every step - from factory to patient."'); r.italic = True; r.font.size = Pt(12); r.font.color.rgb = PRIMARY
    doc.add_page_break()


def build_toc_and_thread(doc):
    h1(doc, "Table of Contents")
    doc.add_paragraph("(In Word: right-click below and choose 'Update Field' to populate. Or use References -> Table of Contents.)").italic = True
    doc.add_paragraph(); doc.add_page_break()
    h1(doc, "Global Traceability Thread")
    p = doc.add_paragraph()
    r = p.add_run("SR-03"); r.bold = True; r.font.color.rgb = ACCENT
    p.add_run(" - Every ownership transfer must be digitally signed, hash-chained, and audit-logged.")
    doc.add_paragraph("This requirement threads through every phase of this report (requirements -> use case -> DFD -> threat -> attack tree -> user story -> sprint -> implementation -> test -> deployment).")
    doc.add_page_break()


# ---------- Phase 1 ----------

def add_phase_1(doc):
    h1(doc, "Phase 1 - Agile Process and Development Approach")
    h2(doc, "1.1 Selected Approach: Scrum-XP")
    doc.add_paragraph("Chosen: Scrum with embedded XP (Extreme Programming) practices.")
    h3(doc, "Justification")
    for b in [
        "Five independent organizations (Manufacturer, Distributor, Warehouse, Retailer, Customer) operate under FDA DSCSA constraints.",
        "Scrum sprints allow phased cross-organization onboarding (1-2 orgs per sprint).",
        "XP practices (TDD, pair programming, CI, code review) enforce rigor for a security-critical, auditable system.",
        "Regulatory changes (e.g., new FDA serialization rule) enter the backlog as new stories without replanning.",
        "Continuous delivery via CI/CD ships signed, tested Docker images every sprint.",
    ]: doc.add_paragraph(b, style="List Bullet")
    h3(doc, "Sprint Configuration")
    for b in [
        "Sprint length: 2 weeks",
        "Roles: Product Owner, Scrum Master, Dev Team (4), Security Champion",
        "Ceremonies: Planning, Daily Scrum, Review, Retrospective, Refinement",
        "XP practices: TDD, Pair Programming, CI, Small Releases, Simple Design",
    ]: doc.add_paragraph(b, style="List Bullet")
    h2(doc, "1.2 Agile Manifesto Mapping")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "#", "Principle", "Application to PharmaChain"
    for r in [
        ("1", "Individuals & interactions over processes/tools", "Cross-org daily scrums jointly resolve shipment discrepancies."),
        ("2", "Working software over comprehensive documentation", "Each sprint ships a running module deployable in Minikube."),
        ("3", "Customer collaboration over contract negotiation", "Retailer + patient feedback in Sprint Review shapes usability."),
        ("4", "Responding to change over following a plan", "New FDA rule becomes a new backlog story - no replan."),
        ("5", "Continuous delivery of valuable software", "CI/CD delivers signed Docker images on every merge to main."),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
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
    p.add_run("Eliminated SQL injection, added RBAC + ABAC, digital signature (non-repudiation), immutable ledger, audit log, and dependency injection for testability.")
    h3(doc, "Refactoring 2: Serial Uniqueness Validation (DRY)")
    doc.add_paragraph("Before: duplicated uniqueness check in 4 services (product_registration, shipment_service, warehouse_receipt, delivery).")
    doc.add_paragraph("After: single SerialValidator class with parameterized query, called by every service. Single point of change, consistent errors, testable.")
    h2(doc, "1.4 Limitations of Agile for a Security-Critical System")
    t2 = doc.add_table(rows=1, cols=2); t2.style = "Table Grid"
    h = t2.rows[0].cells; h[0].text, h[1].text = "Limitation / Risk", "Mitigation"
    for r in [
        ("Under-specified security requirements in fast sprints", "Add abuse stories + Definition of Done requires STRIDE review"),
        ("Informal cross-org communication may cause API contract drift", "Freeze OpenAPI contracts at sprint start; CI validates schema"),
        ("Regulatory compliance demands documentation Agile de-emphasizes", "Auto-generated Compliance Traceability Matrix"),
        ("Velocity pressure may tempt skipping security testing", "SAST + dependency scan mandatory for Definition of Done"),
        ("Cross-org governance ambiguity - who owns the backlog?", "Steering Committee (one rep per org) as PO delegate"),
    ]:
        cells = t2.add_row().cells
        cells[0].text, cells[1].text = r[0], r[1]
    style_table(t2)
    doc.add_page_break()


# ---------- Phase 2 ----------

def add_phase_2(doc):
    h1(doc, "Phase 2 - Requirements Engineering")
    h2(doc, "2.1 Stakeholders and User Types")
    doc.add_paragraph("PharmaChain involves 5 primary actors (one per organization in the supply chain) plus 3 justified supporting roles required for regulatory compliance and security.")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Stakeholder"; hdr[2].text = "Type"; hdr[3].text = "Goal"; hdr[4].text = "Key Concern"
    rows = [
        ("1", "Manufacturer", "Primary Actor", "Create and serialize drug products", "Anti-counterfeit, batch traceability"),
        ("2", "Distributor", "Primary Actor", "Move products between orgs, transfer ownership", "Non-repudiation of transfer"),
        ("3", "Warehouse", "Primary Actor", "Receive, store, dispatch shipments", "Tamper-proof receipts"),
        ("4", "Retailer", "Primary Actor", "Receive products, sell to customers", "Authenticity verification"),
        ("5", "Customer / Patient", "Primary Actor", "Verify drug authenticity", "Trust, privacy"),
        ("6", "Auditor / Regulator (FDA)", "Secondary Actor", "Inspect immutable history", "Compliance, tamper evidence"),
        ("7", "System Administrator", "Supporting Actor", "Onboard orgs, manage keys", "Least privilege, key lifecycle"),
        ("8", "Security Officer", "Supporting Actor", "Monitor anomalies and threats", "Real-time alerts, SIEM"),
    ]
    for r in rows:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Justification for extra actors: ").bold = True
    p.add_run("PharmaChain is regulated (FDA DSCSA) and cross-organizational. An Auditor is required for regulatory inspection. An Admin is needed to onboard/offboard organizations. A Security Officer is required to detect fraud, key abuse, and anomalous transfers.")

    h2(doc, "2.2 Functional Requirements (FR)")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Requirement"; hdr[2].text = "Actor"; hdr[3].text = "Priority"
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
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "2.3 Non-Functional Requirements (NFR)")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Category"; hdr[2].text = "Requirement"; hdr[3].text = "Target"
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
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "2.4 Security Requirements (SR)")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Category"; hdr[2].text = "Requirement"; hdr[3].text = "CIA/Auth"
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
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "2.5 Prioritization (MoSCoW)")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Priority"; hdr[1].text = "Requirements"
    for r in [
        ("Must Have", "FR-01..FR-05, FR-08, FR-09, SR-01..SR-12, NFR-01, NFR-06, NFR-09"),
        ("Should Have", "FR-06, FR-07, FR-10..FR-12, SR-13..SR-16, NFR-02..NFR-05, NFR-07"),
        ("Could Have", "FR-13, FR-14, NFR-08, NFR-10"),
        ("Won't Have (this release)", "Public blockchain anchoring, cross-chain interop"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "2.6 CIA / Auth / Authz / Audit Mapping")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Category"; hdr[1].text = "Requirements"
    for r in [
        ("Confidentiality (C)", "SR-08, SR-09, SR-13 (keys), NFR-06"),
        ("Integrity (I)", "SR-03, SR-04, SR-05, SR-10, SR-11, SR-14"),
        ("Availability (A)", "NFR-01, SR-15"),
        ("Authentication (Auth)", "SR-01, SR-02, SR-16"),
        ("Authorization (Authz)", "SR-06, SR-07"),
        ("Audit (Aud)", "SR-10, SR-11, NFR-09"),
        ("Non-Repudiation (NR)", "SR-03, SR-12"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "2.7 Traceability Thread (Carried Forward)")
    p = doc.add_paragraph()
    r = p.add_run("SR-03"); r.bold = True; r.font.color.rgb = ACCENT
    p.add_run(" = Every ownership transfer must be digitally signed, hash-chained, and audit-logged.")
    doc.add_paragraph("Mapped to FR-04 (transfer ownership), NFR-09 (log retention), SR-12 (non-repudiation). Will thread through Use Case, DFD, Threat Model, Attack Tree, User Story, Code, Test, and Deployment phases.")

    h2(doc, "2.8 Assumptions and Constraints")
    h3(doc, "Assumptions")
    for b in [
        "Each organization has PKI-issued certificates from a shared CA.",
        "Organizations are known/trusted (permissioned network, not public).",
        "Customers have smartphones with QR scanning capability.",
    ]: doc.add_paragraph(b, style="List Bullet")
    h3(doc, "Constraints")
    for b in [
        "Must comply with FDA DSCSA and GS1 EPCIS 2.0.",
        "Must use HSM for signing key storage (no soft keys in production).",
        "Regulatory log retention: 7 years minimum.",
    ]: doc.add_paragraph(b, style="List Bullet")
    doc.add_page_break()


# ---------- Phase 3 ----------

def add_phase_3(doc):
    h1(doc, "Phase 3 - Requirements Analysis and UML")
    h2(doc, "3.1 Use Case Diagram")
    doc.add_paragraph("The diagram shows 8 actors (5 primary supply-chain roles + 3 supporting roles) and 10 major use cases with include/extend relationships that enforce security invariants.")
    add_figure(doc, "03-usecase.png", "Figure 3.1 - PharmaChain Use Case Diagram")

    h3(doc, "3.1.1 Actors and Their Use Cases")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Actor"; hdr[1].text = "Use Cases"
    for r in [
        ("Manufacturer", "UC-01 Register Product, UC-02 Create Shipment, UC-10 Trigger Recall"),
        ("Distributor", "UC-02 Create Shipment, UC-04 Transfer Ownership, Raise Dispute"),
        ("Warehouse", "UC-03 Confirm Warehouse Receipt"),
        ("Retailer", "UC-04 Transfer Ownership, UC-05 Confirm Delivery"),
        ("Customer", "UC-06 Verify Product Authenticity"),
        ("Auditor", "UC-07 View Audit Trail, UC-10 Trigger Recall"),
        ("Admin", "UC-08 Onboard Organization, UC-09 Manage Keys"),
        ("Security Officer", "Monitor Anomalies (via audit + ledger alerts)"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h3(doc, "3.1.2 Include / Extend Relationships")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Source Use Case"; hdr[1].text = "Relationship"; hdr[2].text = "Target Use Case"
    for r in [
        ("UC-02 Create Shipment", "<<include>>", "Sign Transaction"),
        ("UC-04 Transfer Ownership", "<<include>>", "Sign Transaction"),
        ("UC-04 Transfer Ownership", "<<include>>", "Write Audit Log"),
        ("UC-06 Verify Authenticity", "<<include>>", "Query Ledger"),
        ("UC-07 View Audit Trail", "<<include>>", "Query Ledger"),
        ("Raise Dispute", "<<extend>>", "UC-04 Transfer Ownership"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "3.2 Use Case Specification 1 - Transfer Ownership")
    t = doc.add_table(rows=0, cols=2); t.style = "Table Grid"
    ucs1 = [
        ("Use Case ID", "UC-04"),
        ("Name", "Transfer Ownership"),
        ("Actor", "Distributor (current owner)"),
        ("Related FR / SR", "FR-04, SR-03, SR-05, SR-06, SR-07, SR-10, SR-12"),
        ("Precondition", "Product exists; actor authenticated via mTLS + MFA; actor holds current ownership; HSM-backed private key available."),
        ("Main Flow",
         "1. Actor selects product.\n2. Actor enters target organization.\n3. System verifies RBAC (role) + ABAC (organization).\n4. Actor signs transfer payload with private key.\n5. System verifies signature + nonce + timestamp.\n6. System writes append-only audit log entry.\n7. System updates ownership on ledger (hash-chained).\n8. System returns confirmation with tx_hash."),
        ("Alternative Flow 3a", "Actor not authorized -> reject with 403, log attempt, alert Security Officer."),
        ("Alternative Flow 5a", "Signature invalid or nonce reused -> abort transfer, raise alert, log."),
        ("Exception Flow 7a", "Ledger write fails -> rollback DB transaction, retry with exponential backoff."),
        ("Postcondition", "Ownership transferred atomically; immutable ledger entry with tx_hash; audit record created; notification sent to both orgs."),
        ("Security Requirements", "SR-03 (signature + hash chain), SR-05 (nonce), SR-06/07 (RBAC+ABAC), SR-10 (audit), SR-12 (non-repudiation)."),
    ]
    for k, v in ucs1:
        row = t.add_row().cells
        row[0].text = k; row[1].text = v
        if row[0].paragraphs[0].runs: row[0].paragraphs[0].runs[0].bold = True
    style_table(t, header=False)

    h2(doc, "3.3 Use Case Specification 2 - Confirm Warehouse Receipt")
    t = doc.add_table(rows=0, cols=2); t.style = "Table Grid"
    ucs2 = [
        ("Use Case ID", "UC-03"),
        ("Name", "Confirm Warehouse Receipt"),
        ("Actor", "Warehouse operator"),
        ("Related FR / SR", "FR-03, SR-02, SR-04, SR-10, SR-14"),
        ("Precondition", "Shipment exists in state IN_TRANSIT; shipment ID matches incoming goods; seal/hash available for verification."),
        ("Main Flow",
         "1. Operator scans shipment QR code.\n2. System validates QR signature (PKI).\n3. System retrieves shipment record from ledger.\n4. Operator enters/verifies physical seal hash.\n5. System compares seal hash to expected value.\n6. Operator signs receipt with HSM-backed key.\n7. System updates status to RECEIVED, writes audit + ledger.\n8. Notifications sent to Manufacturer and Distributor."),
        ("Alternative Flow 5a", "Seal hash mismatch -> quarantine shipment, alert Security Officer, audit log."),
        ("Alternative Flow 2a", "QR signature invalid -> reject, do not proceed, log attempt."),
        ("Exception Flow 7a", "Ledger write fails -> rollback, retry with exponential backoff."),
        ("Postcondition", "Shipment marked RECEIVED; receipt record created; audit and ledger entries written; origin orgs notified."),
        ("Security Requirements", "SR-02 (mTLS), SR-04 (SHA-256 hash), SR-10 (audit), SR-14 (anomaly)."),
    ]
    for k, v in ucs2:
        row = t.add_row().cells
        row[0].text = k; row[1].text = v
        if row[0].paragraphs[0].runs: row[0].paragraphs[0].runs[0].bold = True
    style_table(t, header=False)

    h2(doc, "3.4 Scenario-Based Analysis Model")
    doc.add_paragraph("Scenario: Receive Shipment and Confirm Receipt (adapted from the exam template scenario 'Attend Examination and Submit Answers', retargeted to PharmaChain).")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Step"; hdr[1].text = "Actor Action"; hdr[2].text = "System Response"
    for r in [
        ("1", "Warehouse scans shipment QR", "Validates QR signature against PKI"),
        ("2", "-", "Retrieves shipment record from ledger"),
        ("3", "Verifies physical seal hash", "Compares seal hash to expected value"),
        ("4", "Clicks 'Confirm Receipt'", "Prompts for digital signature"),
        ("5", "Signs with HSM key", "Verifies signature + RBAC + ABAC"),
        ("6", "-", "Writes audit entry (append-only) + ledger hash"),
        ("7", "-", "Updates status to RECEIVED"),
        ("8", "-", "Sends notification to Manufacturer + Distributor"),
        ("Alt 3a", "Seal mismatch detected", "Quarantine + alert + audit log"),
        ("Alt 5a", "Invalid signature", "Reject, retry prompt, log attempt"),
        ("Exc 2a", "Ledger write fails", "Rollback DB, retry, escalate"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "3.5 Consistency With Phase 2 Requirements")
    doc.add_paragraph("Every use case maps to at least one functional requirement (FR) and carries forward the relevant security requirements (SR).")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Use Case"; hdr[1].text = "FR(s)"; hdr[2].text = "SR(s)"
    for r in [
        ("UC-01 Register Product", "FR-01", "SR-04, SR-10"),
        ("UC-02 Create Shipment", "FR-02", "SR-03, SR-05, SR-10"),
        ("UC-03 Confirm Receipt", "FR-03", "SR-04, SR-10, SR-14"),
        ("UC-04 Transfer Ownership", "FR-04", "SR-03, SR-05, SR-06, SR-07, SR-10, SR-12"),
        ("UC-05 Confirm Delivery", "FR-05", "SR-10, SR-12"),
        ("UC-06 Verify Authenticity", "FR-07", "SR-04"),
        ("UC-07 View Audit Trail", "FR-08", "SR-10, SR-11"),
        ("UC-08 Onboard Organization", "FR-09", "SR-01, SR-02, SR-13"),
        ("UC-09 Manage Keys", "FR-10", "SR-13"),
        ("UC-10 Trigger Recall", "FR-12", "SR-10"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    doc.add_page_break()


# ---------- Phase 4 ----------

def add_phase_4(doc):
    h1(doc, "Phase 4 - Data and Information Flow Modeling")
    h2(doc, "4.1 ER Diagram")
    doc.add_paragraph("The data model supports multi-organization ownership, signed transfers, and a hash-linked audit log. Primary keys are UUIDs to avoid enumeration attacks; foreign keys enforce referential integrity.")
    add_figure(doc, "04-er.png", "Figure 4.1 - PharmaChain Entity-Relationship Diagram")
    h3(doc, "4.1.1 Entities and Their Roles")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Entity"; hdr[1].text = "Purpose"; hdr[2].text = "Key Attributes"
    for r in [
        ("Organization", "Onboarded supply chain participant", "org_id (PK), name, type, public_key, status"),
        ("User", "Employee of an organization", "user_id (PK), org_id (FK), role, mfa_secret"),
        ("Product", "Serialized drug unit/batch", "product_id (PK), serial_no (UNIQUE), manufacturer_id (FK)"),
        ("Shipment", "Movement of goods between orgs", "shipment_id (PK), product_id (FK), source_org, dest_org"),
        ("OwnershipRecord", "Signed transfer of ownership", "record_id (PK), from_org, to_org, signature, tx_hash"),
        ("WarehouseReceipt", "Confirmation of receipt at warehouse", "receipt_id (PK), shipment_id (FK), seal_hash"),
        ("Delivery", "Final delivery to retailer", "delivery_id (PK), shipment_id (FK), customer_id (FK)"),
        ("AuditLog", "Append-only immutable log", "log_id (PK), actor_id (FK), hash_prev, hash_curr, signature"),
        ("KeyStore", "Public keys & lifecycle", "key_id (PK), org_id (FK), public_key, revoked, valid_from/to"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "4.2 Level-0 DFD (Context Diagram)")
    doc.add_paragraph("The context diagram shows PharmaChain as a single process and the 7 external entities that interact with it. Every exchange crosses at least one trust boundary.")
    add_figure(doc, "04-dfd-l0.png", "Figure 4.2 - PharmaChain Level-0 DFD (Context)")
    h2(doc, "4.3 Level-1 DFD with Trust Boundaries")
    doc.add_paragraph("The Level-1 DFD decomposes the system into 6 processes (P1-P6), 6 data stores (D1-D6), and 3 trust boundaries that enforce cross-organization security policies.")
    add_figure(doc, "04-dfd-l1.png", "Figure 4.3 - PharmaChain Level-1 DFD with Trust Boundaries")
    h3(doc, "4.3.1 DFD Elements")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Type"; hdr[1].text = "Element"; hdr[2].text = "Responsibility"
    for r in [
        ("External Entity", "Manufacturer / Distributor / Warehouse / Retailer / Customer / Auditor", "Originate or consume data"),
        ("Process P1", "Register Product", "Validate serial, hash, write to ledger"),
        ("Process P2", "Create Shipment", "Link product, source, destination"),
        ("Process P3", "Confirm Receipt", "Verify seal hash, write receipt"),
        ("Process P4", "Transfer Ownership", "RBAC + signature + nonce + ledger append"),
        ("Process P5", "Confirm Delivery", "Write final delivery record"),
        ("Process P6", "Audit & Monitor", "Read-only ledger queries, anomaly detection"),
        ("Data Store D1", "Product DB", "Serialized product master data"),
        ("Data Store D2", "Shipment DB", "Shipment and route metadata"),
        ("Data Store D3", "Receipt DB", "Warehouse receipt records"),
        ("Data Store D4", "Ownership DB", "Ownership transfer history"),
        ("Data Store D5", "Immutable Ledger", "Hash-chained audit log"),
        ("Data Store D6", "Delivery DB", "Final delivery records"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h3(doc, "4.3.2 Trust Boundaries")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Boundary"; hdr[1].text = "Between"; hdr[2].text = "Enforcement"
    for r in [
        ("TB1", "External organizations <-> API Gateway", "mTLS + OAuth2 + rate limiting"),
        ("TB2", "API Gateway <-> Application Zone", "JWT validation + RBAC + ABAC"),
        ("TB3", "Application Zone <-> Ledger Zone", "Signed writes + HSM-backed keys"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "4.4 Consistency With Phase 2 & 3")
    doc.add_paragraph("Every entity maps to a use case from Phase 3 and to requirements from Phase 2. The traceability thread (SR-03) is enforced in DFD process P4 and the OwnershipRecord + AuditLog entities.")
    doc.add_page_break()


# ---------- Phase 5 ----------

def add_phase_5(doc):
    h1(doc, "Phase 5 - Software Architecture and Design Engineering")
    h2(doc, "5.1 Proposed Architecture")
    doc.add_paragraph("Layered microservices architecture with a permissioned blockchain (Hyperledger Fabric) for the immutable ledger. This style fits PharmaChain because each organization integrates independently, and the ledger provides cross-organization trust without a central authority.")
    add_figure(doc, "05-architecture.png", "Figure 5.1 - PharmaChain Layered Architecture")
    h2(doc, "5.2 Architecture Justification")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Choice"; hdr[1].text = "Justification"
    for r in [
        ("Microservices", "Each organization integrates independently; services scale separately"),
        ("API Gateway", "Single secure entry point with mTLS + OAuth2 + rate limiting"),
        ("Event-Driven (Kafka)", "Real-time shipment notifications and anomaly detection"),
        ("Permissioned blockchain", "Immutable, tamper-evident audit across organizations"),
        ("CQRS (read/write split)", "Audit reads are isolated from transactional writes"),
        ("Service Mesh (sidecar)", "Enforces mTLS between microservices"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "5.3 Components and Interfaces")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Component"; hdr[1].text = "Responsibility"; hdr[2].text = "Interface"
    for r in [
        ("API Gateway", "Auth, rate limit, routing", "REST/HTTPS"),
        ("Auth Service", "MFA, PKI, JWT issuance", "gRPC"),
        ("Product Service", "Product CRUD, serial validation", "REST"),
        ("Shipment Service", "Shipment lifecycle, notifications", "REST + Kafka"),
        ("Ownership Service", "Transfers, signatures, ledger writes", "REST + Fabric SDK"),
        ("Warehouse Service", "Receipts, seal verification", "REST"),
        ("Delivery Service", "Final delivery confirmation", "REST"),
        ("Audit Service", "Read-only ledger queries", "REST"),
        ("Notification Service", "Email / SMS alerts", "Kafka consumer"),
        ("Anomaly Detection", "ML rules on transfers", "Kafka consumer"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "5.4 Design Patterns Applied")
    for b in [
        "API Gateway - single secure entry point",
        "Saga - distributed ownership transfer across services",
        "CQRS - separate read (audit) and write (transaction) paths",
        "Circuit Breaker - resilience between organization services",
        "Sidecar (Service Mesh) - mTLS between microservices",
    ]: doc.add_paragraph(b, style="List Bullet")
    h2(doc, "5.5 Mapping Requirements to Components")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Concern"; hdr[1].text = "Components"
    for r in [
        ("Authentication", "Auth Service + API Gateway"),
        ("Authorization", "API Gateway + Auth Service (RBAC + ABAC)"),
        ("Ownership transfer (SR-03)", "Ownership Service + Ledger Service + Saga"),
        ("Audit / Logging", "Audit Service + Ledger + ELK stack"),
        ("Anomaly detection", "Anomaly Detection + SIEM"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    doc.add_page_break()


# ---------- Phase 6 ----------

def add_phase_6(doc):
    h1(doc, "Phase 6 - User Interface Design")
    doc.add_paragraph("The UI follows a clean enterprise design language: navy sidebar, breadcrumb navigation, KPI cards with gradient accents, and hash-verified data tables. All 4 screens are interactive HTML mockups built with a shared stylesheet (ui/style.css).")
    h2(doc, "6.1 Login Screen")
    add_figure(doc, "01-login.png", "Figure 6.1 - Login Screen (split layout with brand panel)")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Aspect"; hdr[1].text = "Detail"
    for r in [
        ("User", "Any organization employee"),
        ("Goal", "Authenticate securely and enter workspace"),
        ("Navigation", "Login -> Dashboard"),
        ("Inputs", "Organization, email, password, MFA code"),
        ("Feedback", "Inline validation, inline errors, disabled button until valid"),
        ("Error Handling", "Generic error messages (no user enumeration), lockout after 5 fails"),
        ("Security", "TLS 1.3, mTLS option, MFA required, rate limiting"),
        ("Golden Rules", "Consistency, Error prevention, Visibility"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "6.2 Product Registration Screen (Manufacturer)")
    add_figure(doc, "02-product-registration.png", "Figure 6.2 - Product Registration Screen")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Aspect"; hdr[1].text = "Detail"
    for r in [
        ("User", "Manufacturer operator"),
        ("Goal", "Register new product with unique serial"),
        ("Navigation", "Dashboard -> Products -> Register New"),
        ("Inputs", "Product name, batch, serial (auto), expiry, quantity, storage"),
        ("Feedback", "KPI cards show today's stats; inline uniqueness check on serial"),
        ("Error Handling", "Duplicate serial -> inline warning; HSM signature failure -> retry + alert"),
        ("Security", "Auto-generated serial (no tampering), signed submission, RBAC enforced"),
        ("Golden Rules", "Feedback, Consistency, Visibility"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "6.3 Shipment Tracking Screen (Distributor)")
    add_figure(doc, "03-shipment-tracking.png", "Figure 6.3 - Shipment Tracking Screen")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Aspect"; hdr[1].text = "Detail"
    for r in [
        ("User", "Distributor / Warehouse / Retailer"),
        ("Goal", "Track shipment, verify chain of custody, act on it"),
        ("Navigation", "Dashboard -> Shipments -> SHP-xxxx"),
        ("Inputs", "Shipment ID or QR scan"),
        ("Feedback", "Live timeline, temperature chip, progress bar, hash verification"),
        ("Error Handling", "Unknown ID -> 'not found, verify'; hash mismatch -> red alert + audit log"),
        ("Security", "Read-only for non-owners; actions require signature; nonce check"),
        ("Golden Rules", "Navigation, User control, Consistency"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "6.4 Audit Trail Screen (Auditor)")
    add_figure(doc, "04-audit-trail.png", "Figure 6.4 - Audit Trail Screen")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Aspect"; hdr[1].text = "Detail"
    for r in [
        ("User", "Auditor / Regulator (FDA)"),
        ("Goal", "Inspect immutable history and verify hash chain"),
        ("Navigation", "Dashboard -> Audit Trail"),
        ("Inputs", "Filters: product serial, date range, event type"),
        ("Feedback", "Chain integrity KPI, per-row Valid/Invalid badges, glowing verify dots"),
        ("Error Handling", "Hash mismatch -> red 'Invalid' badge + alert to Security Officer"),
        ("Security", "Read-only view, signed export to PDF, 7-year retention"),
        ("Golden Rules", "Consistency, Visibility, Error prevention"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "6.5 Golden Rules Applied Across All Screens")
    for b in [
        "Consistency - unified sidebar, colors, typography, spacing",
        "User control - breadcrumb navigation, cancel/back everywhere",
        "Feedback - inline errors, loading states, success toasts",
        "Error prevention - required fields, disabled submit until valid, uniqueness check",
        "Navigation - persistent sidebar, breadcrumbs, single-click access",
        "Visibility - KPI cards, status badges, chain-of-custody timeline",
    ]: doc.add_paragraph(b, style="List Bullet")
    doc.add_page_break()


# ---------- Phase 7 ----------

def add_phase_7(doc):
    h1(doc, "Phase 7 - Threat Modeling and Security Analysis")
    h2(doc, "7.1 Assets and CIA Classification")
    doc.add_paragraph("Nine critical assets were identified across the PharmaChain system. Each is classified by Confidentiality (C), Integrity (I), and Availability (A) requirements.")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Asset"; hdr[2].text = "C"; hdr[3].text = "I"; hdr[4].text = "A"
    for r in [
        ("1", "Product records (serials, batch)", "Yes", "Yes", "Yes"),
        ("2", "Ownership transfer records", "Yes", "Yes", "Yes"),
        ("3", "Shipment data", "Yes", "Yes", "Yes"),
        ("4", "Customer PII (delivery addresses)", "Yes", "Yes", "No"),
        ("5", "Digital signing keys", "Yes", "Yes", "Yes"),
        ("6", "Audit log / ledger", "No", "Yes", "Yes"),
        ("7", "Authentication credentials", "Yes", "Yes", "No"),
        ("8", "Pricing & contract data", "Yes", "Yes", "No"),
        ("9", "Ledger integrity (hash chain)", "No", "Yes", "Yes"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "7.2 STRIDE Threat Analysis")
    doc.add_paragraph("STRIDE was applied to each element of the Level-1 DFD. Eleven threats were identified, along with mitigations and residual risk.")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "DFD Element"; hdr[2].text = "Threat"; hdr[3].text = "STRIDE"; hdr[4].text = "Impact / Mitigation"
    for r in [
        ("T1", "API Gateway (TB1)", "Credential stuffing", "S / Account takeover / MFA + rate limit + CAPTCHA"),
        ("T2", "Product DB (D1)", "Serial tampering", "T / Fake products / SHA-256 + signature"),
        ("T3", "Ownership Service (P4)", "Repudiation of transfer", "R / Disputed ownership / Digital signature + ledger"),
        ("T4", "Audit Log (D5)", "Information disclosure", "I / Privacy leak / Encryption at rest + RBAC"),
        ("T5", "Shipment API (P2)", "Denial of Service", "D / Supply chain halt / Rate limit + WAF + autoscale"),
        ("T6", "Admin console (P6)", "Privilege escalation", "E / Full compromise / Least privilege + MFA"),
        ("T7", "Kafka topic", "Message injection", "T / Fake shipment / mTLS + schema validation + signing"),
        ("T8", "Ledger write (D5)", "Replay attack", "T / Duplicate transfers / Nonce + timestamp + idempotency"),
        ("T9", "Customer portal", "PII scraping", "I / GDPR breach / Rate limit + auth + field encryption"),
        ("T10", "Service mesh", "MITM between microservices", "S/T / Data theft / mTLS between every call"),
        ("T11", "QR code on product", "Forgery of QR", "S / Fake receipt / Signed QR (PKI) verified server-side"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "7.3 Information Flow Analysis")
    doc.add_paragraph("Three sensitive assets were traced from source to destination, with the controls enforced at every hop.")
    h3(doc, "Asset 1 - Ownership Record")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Flow Step"; hdr[1].text = "Control"
    for r in [
        ("Source: Distributor app", "Signed with HSM-backed private key"),
        ("In transit: Distributor -> API Gateway", "TLS 1.3 + mTLS certificate"),
        ("API Gateway -> Ownership Service", "JWT validation, RBAC + ABAC"),
        ("Ownership Service -> Ledger", "Hash-chained write, signature stored"),
        ("Ledger -> Audit Service", "Read-only path, no write access"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h3(doc, "Asset 2 - Customer PII")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Flow Step"; hdr[1].text = "Control"
    for r in [
        ("Customer -> Portal", "TLS 1.3, token-based session"),
        ("Portal -> Delivery Service", "Field-level encryption (AES-256)"),
        ("Delivery Service -> PostgreSQL", "Encrypted at rest, column-level encryption"),
        ("Delivery Service -> Notification", "Masked PII, role-limited access"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h3(doc, "Asset 3 - Signing Keys")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Flow Step"; hdr[1].text = "Control"
    for r in [
        ("Generation", "Inside HSM, never exported"),
        ("Storage", "HSM with FIPS 140-2 Level 3"),
        ("Usage", "Short-lived session tokens only"),
        ("Rotation", "Every 90 days, dual-control"),
        ("Revocation", "Immediate, propagated to KeyStore"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "7.4 Vulnerability Analysis")
    doc.add_paragraph("Six high-impact vulnerabilities were identified from the threat analysis. Each maps to a threat and has a specific mitigation.")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Element"; hdr[2].text = "Threat"; hdr[3].text = "Impact"; hdr[4].text = "Mitigation"
    for r in [
        ("V1", "Login API", "T1 (credential stuffing)", "Account takeover", "MFA + lockout + CAPTCHA + IP throttle"),
        ("V2", "QR verification", "T11 (QR forgery)", "Fake receipt accepted", "Server-side signature verification"),
        ("V3", "Ledger SDK", "T8 (replay)", "Duplicate transfers", "Nonce + timestamp + idempotency key"),
        ("V4", "Kafka topic", "T7 (injection)", "Fake shipments", "Schema validation + message signing + mTLS"),
        ("V5", "Admin API", "T6 (IDOR)", "Unauthorized access", "ABAC check per org, UUIDv4 IDs"),
        ("V6", "Blob storage", "T4 (misconfig)", "Data leak", "Private buckets + signed URLs + audit logs"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "7.5 Residual Risk Summary")
    doc.add_paragraph("After applying the mitigations above, the highest residual risks are: (1) insider threat at the Ledger Service - mitigated by dual-control and segregation of duties; (2) quantum-computing attack on RSA keys - deferred to future work with post-quantum migration roadmap.")
    doc.add_page_break()


# ---------- Phase 8 ----------

def add_phase_8(doc):
    h1(doc, "Phase 8 - Attack Tree and Security Architecture Refinement")
    h2(doc, "8.1 Attack Tree")
    doc.add_paragraph("Critical attacker goal: fraudulently transfer ownership of another organization's product. The attack tree below decomposes this into sub-goals and enumerates preventive and detective controls.")
    add_figure(doc, "08-attack-tree.png", "Figure 8.1 - Attack Tree: Fraudulent Ownership Transfer")

    h2(doc, "8.2 Attack Paths and Controls")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Path"; hdr[1].text = "Sub-goal"; hdr[2].text = "Preventive Control"; hdr[3].text = "Detective Control"
    for r in [
        ("A1", "Steal Distributor's private key", "HSM-backed keys (no export), MFA", "Anomaly alert on unusual login IP/time"),
        ("A1a", "Compromise HSM", "FIPS 140-2 L3, tamper-responsive", "HSM audit logs + physical monitoring"),
        ("A1c", "Phish employee", "Email filtering, awareness training", "SIEM correlation of suspicious events"),
        ("A2", "Exploit API authorization bug", "ABAC per org, UUIDv4 IDs", "API access log review"),
        ("A2a", "Guess product ID", "UUIDv4 (122 bits entropy)", "Rate limit on 404s"),
        ("A3", "Replay old signed transfer", "Nonce + timestamp + idempotency", "Ledger duplicate detection"),
        ("A4", "Insider at Ledger Service", "Dual-control, segregation of duties", "Peer review + immutable audit log"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "8.3 Security Architecture Refinement")
    doc.add_paragraph("The following controls were added to the architecture in Phase 5 as a direct result of this attack tree analysis:")
    for b in [
        "HSM-backed signing service (was: soft keys)",
        "Nonce + timestamp validation at API Gateway (was: no replay defense)",
        "Dual-approval workflow for high-value transfers (was: single signer)",
        "SIEM with anomaly rules (was: manual log review)",
        "Signed QR codes verified server-side (was: trust QR contents)",
        "Automatic key rotation every 90 days (was: manual ad-hoc)",
    ]: doc.add_paragraph(b, style="List Bullet")

    h2(doc, "8.4 Highest-Risk Issues After Refinement")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Rank"; hdr[1].text = "Risk"; hdr[2].text = "Control Status"
    for r in [
        ("1", "Private key theft -> fraudulent transfer", "Controlled: HSM + MFA + rotation + anomaly alerts"),
        ("2", "Replay of signed transfer", "Controlled: nonce + timestamp + idempotency"),
        ("3", "Insider at Ledger Service", "Controlled: dual-control + segregation + immutable audit"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "8.5 Traceability Thread (SR-03)")
    doc.add_paragraph("SR-03 (every ownership transfer must be digitally signed, hash-chained, and audit-logged) maps directly to attack path A3 (Replay) which is mitigated by nonce + timestamp + hash chain. This closes the thread from requirement -> attack -> control.")
    doc.add_page_break()


# ---------- Phase 9 ----------

def add_phase_9(doc):
    h1(doc, "Phase 9 - Product Backlog and Jira/Scrum")
    doc.add_paragraph("The PharmaChain backlog was managed in a live Jira Cloud Scrum project (site: j33v4n.atlassian.net, project key: SCMSEC). The backlog, sprint boards, and burndown charts below are direct screenshots from Jira.")

    h2(doc, "9.0 Live Jira Backlog (Evidence)")
    add_figure(doc, "01-backlog.png", "Figure 9.1 - Jira backlog showing 5 epics, 12 stories, 2 sprints")
    doc.add_paragraph()
    doc.add_paragraph("The Timeline view below shows story start/due dates and the dependency arrows between blocked and blocking stories.")
    add_figure(doc, "05-timeline.png", "Figure 9.2 - Jira Timeline view with dependencies")

    h2(doc, "9.1 Epics")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells; hdr[0].text = "Epic"; hdr[1].text = "Description"
    for r in [
        ("E1 (SCMSEC-1)", "Identity & Access - PKI, MFA, RBAC, ABAC"),
        ("E2 (SCMSEC-2)", "Product Lifecycle - registration, serialization, lookup"),
        ("E3 (SCMSEC-3)", "Shipment & Logistics - shipments, receipts, deliveries"),
        ("E4 (SCMSEC-4)", "Ownership & Ledger - signed transfers, hash chain"),
        ("E5 (SCMSEC-5)", "Audit & Compliance - immutable log, anomaly detection, reporting"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)

    h2(doc, "9.2 User Stories")
    doc.add_paragraph("12 user stories are defined across the 5 epics. Each follows the As a ... I want ... so that ... format with acceptance criteria.")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Epic"; hdr[2].text = "User Story"; hdr[3].text = "Priority"; hdr[4].text = "Acceptance Criteria"
    for r in [
        ("SCMSEC-6",  "E1", "As a Manufacturer, I want to register with PKI so that my identity is trusted", "High", "Certificate issued; mTLS works"),
        ("SCMSEC-7",  "E1", "As an Admin, I want to onboard an org so that it can join the network", "High", "Org has keys and RBAC roles"),
        ("SCMSEC-8",  "E1", "As a User, I want MFA so that my account is protected", "High", "TOTP required; recovery codes"),
        ("SCMSEC-9",  "E2", "As a Manufacturer, I want to register a product with a unique serial so it is traceable", "High", "Duplicate serial rejected; ledger entry"),
        ("SCMSEC-10", "E2", "As a Customer, I want to verify product authenticity so I trust my purchase", "High", "Signed query returns valid/invalid"),
        ("SCMSEC-11", "E3", "As a Distributor, I want to create a shipment so goods move", "High", "Shipment ID + signed record"),
        ("SCMSEC-12", "E3", "As a Warehouse, I want to confirm receipt so chain of custody continues", "High", "Seal hash verified; audit logged"),
        ("SCMSEC-13", "E3", "As a Retailer, I want to confirm delivery so the sale is complete", "High", "Delivery record + notification"),
        ("SCMSEC-14", "E4", "As a Distributor, I want to transfer ownership so legal title moves", "High", "Signature verified; ledger updated"),
        ("SCMSEC-15", "E5", "As an Auditor, I want to view immutable audit logs so I can investigate", "High", "Read-only; hash verified"),
        ("SCMSEC-16", "E5", "As a Security Officer, I want anomaly alerts so fraud is detected", "Medium", "Alert on duplicate/odd transfer"),
        ("SCMSEC-17", "E3", "As a Distributor, I want real-time status so I can plan logistics", "Medium", "Status < 1s; push notifications"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "9.3 Jira Scrum Project Setup")
    for b in [
        "Jira Cloud site: j33v4n.atlassian.net",
        "Project key: SCMSEC",
        "Board type: Scrum",
        "Epics: SCMSEC-1 .. SCMSEC-5",
        "Stories: SCMSEC-6 .. SCMSEC-17",
        "Workflow: TO DO -> IN PROGRESS -> TESTING -> DONE",
    ]: doc.add_paragraph(b, style="List Bullet")

    h2(doc, "9.4 Sprint Plan")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Sprint"; hdr[1].text = "Sprint Goal"; hdr[2].text = "Stories"; hdr[3].text = "Duration"
    for r in [
        ("Sprint 1 - Secure Foundation", "Secure Foundation - Identity, Product, Audit", "SCMSEC-6, 7, 8, 9, 15", "2 weeks"),
        ("Sprint 2 - Logistics", "Logistics and Ownership - Shipments, Receipts, Transfers, Anomaly", "SCMSEC-10..14, 16, 17", "2 weeks"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    doc.add_page_break()


# ---------- Phase 10 ----------

def add_phase_10(doc):
    h1(doc, "Phase 10 - Sprint Execution and Scrum Metrics")
    h2(doc, "10.1 Sprint Board at Start (Jira Evidence)")
    doc.add_paragraph("Sprint 1 - Secure Foundation was started in Jira. The board below is a live screenshot from the Jira Cloud site.")
    add_figure(doc, "02-sprint1-board.png", "Figure 10.1 - Sprint 1 board at sprint start (TO DO / IN PROGRESS / DONE)")
    h2(doc, "10.2 Board With Progress (Day 3)")
    doc.add_paragraph("After three days of execution, several stories have been moved through IN PROGRESS, TESTING, and DONE.")
    add_figure(doc, "03-board-progress.png", "Figure 10.2 - Sprint 1 board with in-flight and completed stories")
    h2(doc, "10.3 Burndown Chart")
    add_figure(doc, "04-burndown.png", "Figure 10.3 - Sprint burndown chart from Jira")
    h2(doc, "10.4 Sprint Board Snapshot (tabular)")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "TO DO"; hdr[1].text = "IN PROGRESS"; hdr[2].text = "TESTING"; hdr[3].text = "DONE"
    for r in [
        ("SCMSEC-17", "SCMSEC-9", "SCMSEC-15", "SCMSEC-6, SCMSEC-7, SCMSEC-8"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "10.5 Daily Scrum Entry (Day 3)")
    t = doc.add_table(rows=1, cols=4); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Member"; hdr[1].text = "Yesterday"; hdr[2].text = "Today"; hdr[3].text = "Blockers"
    for r in [
        ("Dev A", "Finished SCMSEC-8", "Start SCMSEC-9", "None"),
        ("Dev B", "SCMSEC-14 signature logic", "Unit tests for SCMSEC-14", "HSM access pending"),
        ("QA", "Test plan for SCMSEC-6", "Execute SCMSEC-6 tests", "None"),
        ("Security", "Threat review SCMSEC-14", "STRIDE recheck", "None"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "10.6 Sprint Burndown Data")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Day"; hdr[1].text = "Ideal SP"; hdr[2].text = "Actual SP"
    for r in [
        ("0", "20", "20"), ("2", "16", "17"), ("4", "12", "13"),
        ("6", "8", "9"), ("8", "4", "5"), ("10", "0", "2 (carry-over SCMSEC-17)"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "10.7 Scrum Metrics")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Metric"; hdr[1].text = "Value"
    for r in [
        ("Velocity (Sprint 1)", "18 story points"),
        ("Velocity (Sprint 2)", "18 story points"),
        ("Average velocity", "18 story points"),
        ("Defects found", "4 (1 High, 2 Medium, 1 Low)"),
        ("Defects fixed", "3"),
        ("Defects carried over", "1 (Low)"),
        ("Stories completed", "11 of 12"),
        ("Carry-over", "SCMSEC-17 (real-time status)"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h2(doc, "10.8 Defect Log")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Severity"; hdr[2].text = "Description"; hdr[3].text = "Fix"; hdr[4].text = "Status"
    for r in [
        ("DEF-001", "High", "Replay in ownership transfer API", "Nonce + timestamp validation", "Fixed"),
        ("DEF-002", "Medium", "Duplicate serial race condition", "DB unique constraint + transaction", "Fixed"),
        ("DEF-003", "Medium", "Anomaly alert fired on legitimate night shift", "Refined rule (shift-aware)", "Fixed"),
        ("DEF-004", "Low", "Timestamp formatting inconsistency", "Deferred to Sprint 3", "Open"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    h2(doc, "10.9 Sprint Review Outcome")
    doc.add_paragraph("Stakeholders (Manufacturer, Distributor, Auditor) accepted 11 stories. The team demonstrated end-to-end flow: product registration -> shipment -> warehouse receipt -> ownership transfer -> audit trail.")
    doc.add_paragraph("Retailer requested better anomaly alert UI for non-technical operators. This was added to the Sprint 3 backlog.")
    h2(doc, "10.10 Sprint Retrospective")
    t = doc.add_table(rows=1, cols=2); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "What Went Well"; hdr[1].text = "What To Improve"
    for r in [
        ("Cross-org daily scrums surfaced blockers early", "Add HSM access earlier in the sprint"),
        ("Security review as part of Definition of Done caught replay bug", "Pair-program security-critical stories (SCMSEC-14)"),
        ("Velocity stable at 18 SP across both sprints", "Refine carry-over estimation"),
    ]:
        cells = t.add_row().cells; cells[0].text, cells[1].text = r[0], r[1]
    style_table(t)
    h3(doc, "Two Improvement Actions for Sprint 3")
    for b in [
        "Provision HSM development keys before Sprint 3 planning (owner: DevOps).",
        "Adopt pair programming on all signature-related stories (owner: Dev Team).",
    ]: doc.add_paragraph(b, style="List Bullet")
    doc.add_page_break()


# ---------- Phase 11 ----------

def add_phase_11(doc):
    h1(doc, "Phase 11 - Secure Development and Build Environment")

    h2(doc, "11.1 Repository and Branch Strategy")
    doc.add_paragraph("The project uses Git with GitFlow-style branching. Every commit is signed and all changes to main require a pull request with two approving reviews.")
    for b in [
        "Remote: GitHub (private repo)",
        "Branch model: GitFlow (main, develop, feature/*, release/*, hotfix/*)",
        "Protected branches: main, release/*",
        "Required reviews: 2",
        "Signed commits: yes",
        "CODEOWNERS: security team owns security/ and ledger/",
    ]: doc.add_paragraph(b, style="List Bullet")

    h2(doc, "11.2 Secure Build Controls")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Control"; hdr[2].text = "Implementation"
    for r in [
        ("1", "Least privilege", "GitHub teams per role; no admin access by default; branch permissions"),
        ("2", "Secret management", "Secrets in GitHub Actions Secrets + HashiCorp Vault; never in source"),
        ("3", "Dependency control", "requirements.txt with pinned versions; Dependabot + pip-audit in CI"),
        ("4", "Code review", "Mandatory 2-reviewer PRs; CODEOWNERS for security-critical paths"),
        ("5", "Protected branches", "main cannot be force-pushed; requires passing CI status checks"),
        ("6", "Reproducible builds", "Pinned base image (python:3.12-slim) + pinned pip hashes"),
        ("7", "Artifact integrity", "Docker images signed with Sigstore Cosign; SHA-256 digests stored"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "11.3 No Hard-Coded Secrets")
    doc.add_paragraph("Secrets are never committed. Verified by grep across source:")
    add_code_block(doc,
        "# PASS: only env var lookups\n"
        "DB_PASSWORD = os.environ['DB_PASSWORD']\n"
        "HSM_TOKEN   = os.environ['HSM_TOKEN']\n"
        "\n"
        "# FAIL (would break CI):\n"
        "# DB_PASSWORD = 'admin123'   <- rejected by secret scanner"
    )

    h2(doc, "11.4 Static Analysis (Bandit)")
    doc.add_paragraph("Bandit is run on every commit. The current scan finds zero issues:")
    add_figure(doc, "12-bandit.png", "Figure 11.1 - Bandit SAST scan: 0 high/medium/low issues")
    doc.add_paragraph("CI fails the pipeline if any HIGH severity issue is introduced.")
    doc.add_page_break()


# ---------- Phase 12 ----------

def add_phase_12(doc):
    h1(doc, "Phase 12 - Secure Coding and Refactoring")

    h2(doc, "12.1 Module Implemented")
    doc.add_paragraph("We implemented the ownership-transfer module and the authentication module. Both went through a weak v1 -> secure v2 refactoring.")

    h2(doc, "12.2 Refactoring Evidence - Ownership Transfer")
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
        "        if not self.authz.can(actor, 'TRANSFER_OWNERSHIP', product_id):\n"
        "            self.audit.log('TRANSFER_DENIED', {...})\n"
        "            raise UnauthorizedError()\n"
        "        if not self.signer.verify(actor, product_id, new_owner, signature):\n"
        "            raise SignatureError()\n"
        "        with self.repo.transaction():\n"
        "            product.owner = new_owner\n"
        "            self.repo.save(product)\n"
        "            tx_hash = self.ledger.append({...})\n"
        "        self.audit.log('TRANSFER', {'tx_hash': tx_hash, ...})\n"
        "        return TransferResult(True, tx_hash)"
    )
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Aspect"; hdr[1].text = "Before"; hdr[2].text = "After"
    for r in [
        ("SQL injection", "Vulnerable (f-string)", "Parameterized query"),
        ("Authentication", "None", "mTLS + MFA"),
        ("Authorization", "None", "RBAC + ABAC"),
        ("Non-repudiation", "None", "HSM-backed signature"),
        ("Audit trail", "None", "Immutable ledger + audit log"),
        ("Testability", "Low", "High (dependency injection)"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "12.3 Refactoring Evidence - Authentication")
    doc.add_paragraph("v1 (insecure):")
    add_code_block(doc,
        "def login(email, password):\n"
        "    user = db.get_user(email)\n"
        "    if user.password == password:   # plaintext compare\n"
        "        return True\n"
        "    return False  # leaks timing / no MFA / no rate-limit"
    )
    doc.add_paragraph("v2 (secure):")
    add_code_block(doc,
        "def login(self, email, password, mfa_code):\n"
        "    self._validate_email(email)           # input validation\n"
        "    self._validate_mfa_code(mfa_code)\n"
        "    if self._is_locked(email):            # rate limit\n"
        "        return AuthResult(False, reason='Account locked')\n"
        "    user = self.user_repo.find_by_email(email)\n"
        "    if user is None or not self._verify_password(password, user.password_hash):\n"
        "        self._record_failure(email)\n"
        "        return AuthResult(False, reason='Invalid credentials')  # generic\n"
        "    if not self.mfa_service.verify(user.mfa_secret, mfa_code):\n"
        "        return AuthResult(False, reason='Invalid credentials')\n"
        "    return AuthResult(True, user_id=user.user_id)"
    )

    h2(doc, "12.4 Security Controls Demonstrated")
    for b in [
        "Input validation - email format, MFA code format, serial regex",
        "Password hashing - bcrypt with 12-round cost; passwords truncated to 72 bytes",
        "Authorization - RBAC + ABAC check before any state change",
        "Error handling - generic messages prevent user enumeration",
        "Sensitive data - no plaintext passwords, no PII in logs",
        "Rate limiting - account lockout after 5 failures in 15 min",
        "Constant-time comparison - hmac.compare_digest for secret comparisons",
    ]: doc.add_paragraph(b, style="List Bullet")

    h2(doc, "12.5 Test Coverage")
    doc.add_paragraph("11 unit tests covering both modules, with 81% line coverage.")
    add_figure(doc, "13-pytest.png", "Figure 12.1 - Pytest + coverage for secure code")
    doc.add_page_break()


# ---------- Phase 13 ----------

def add_phase_13(doc):
    h1(doc, "Phase 13 - Containerized Development: Docker and Kubernetes")

    h2(doc, "13.1 Dockerfile")
    doc.add_paragraph("The image is based on python:3.12-slim, runs as a non-root user, exposes only port 8080, and contains no secrets.")
    add_figure(doc, "13-dockerfile.png", "Figure 13.1 - Dockerfile with security best practices")

    h3(doc, "Container Security Practices Applied")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Practice"; hdr[2].text = "Implementation"
    for r in [
        ("1", "Minimal base image", "python:3.12-slim (not full Debian)"),
        ("2", "Non-root execution", "useradd appuser + USER appuser"),
        ("3", "Controlled ports", "Only EXPOSE 8080"),
        ("4", "No secrets in image", "Env vars injected at runtime via Secrets"),
        ("5", "Version control", "Pinned base image tag"),
        ("6", "Removed unnecessary packages", "Slim base, minimal deps"),
        ("7", "Health check", "HEALTHCHECK directive"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "13.2 .dockerignore")
    doc.add_paragraph("The .dockerignore excludes .git, .venv, tests, docs, secrets, and all evidence files to keep the image minimal and leak-free.")

    h2(doc, "13.3 Kubernetes Deployment")
    add_figure(doc, "17-k8s-deployment.png", "Figure 13.2 - Kubernetes Deployment manifest")

    h3(doc, "Kubernetes Security Controls")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "#"; hdr[1].text = "Control"; hdr[2].text = "Implementation"
    for r in [
        ("1", "Namespace isolation", "Dedicated namespace: scm"),
        ("2", "Non-root security context", "runAsNonRoot, runAsUser 1000, drop ALL capabilities"),
        ("3", "Read-only root filesystem", "readOnlyRootFilesystem: true"),
        ("4", "Resource limits", "CPU 500m, memory 512Mi limits"),
        ("5", "Network policy", "Ingress only from api-gateway; egress to postgres + fabric-peer only"),
        ("6", "Secrets", "K8s Secrets + Vault; never in manifests"),
        ("7", "Health probes", "liveness + readiness on /health"),
        ("8", "Replica count", "2 replicas for availability"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "13.4 Manifest Validation")
    doc.add_paragraph("All 6 Kubernetes manifests validated successfully (kind + metadata + schema). Client-side YAML validation was used since no local cluster was available.")
    add_figure(doc, "15-k8s-dryrun.png", "Figure 13.3 - K8s manifests validated (namespace, configmap, secret, deployment, service, networkpolicy)")
    doc.add_page_break()


# ---------- Phase 14 ----------

def add_phase_14(doc):
    h1(doc, "Phase 14 - CI/CD and Security Testing")

    h2(doc, "14.1 CI/CD Pipeline")
    doc.add_paragraph("GitHub Actions pipeline (.github/workflows/ci.yml) runs on every push and PR. It performs the following stages in order:")
    t = doc.add_table(rows=1, cols=3); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "Stage"; hdr[1].text = "Action"; hdr[2].text = "Fails on"
    for r in [
        ("1. Checkout", "actions/checkout@v4", "-"),
        ("2. Setup Python", "actions/setup-python@v5 (3.12)", "-"),
        ("3. Install deps", "pip install -e . + test tools", "Install error"),
        ("4. Unit tests", "pytest --cov=src", "Any failure"),
        ("5. SAST", "bandit -r src/", "High severity issue"),
        ("6. Dependency scan", "pip-audit", "Known CVE"),
        ("7. Build container", "docker build", "Build failure"),
        ("8. Artifact integrity", "cosign sign (informational)", "-"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)
    add_figure(doc, "16-cicd-yaml.png", "Figure 14.1 - GitHub Actions CI/CD pipeline")

    h2(doc, "14.2 Test Results")
    doc.add_paragraph("The full test suite includes 11 unit tests, 2 integration tests, and 4 fuzz tests, all passing.")
    add_figure(doc, "14-pytest-full.png", "Figure 14.2 - Full test suite: 17 passed")

    h2(doc, "14.3 Dependency Vulnerability Scan")
    doc.add_paragraph("pip-audit initially found 3 vulnerabilities in python-jose and ecdsa. Both libraries were unused, so they were removed. The re-scan is clean.")
    add_figure(doc, "18-pip-audit-before.png", "Figure 14.3 - pip-audit BEFORE: 3 vulnerabilities found")
    add_figure(doc, "18-pip-audit-fixed.png", "Figure 14.4 - pip-audit AFTER: no known vulnerabilities")

    h2(doc, "14.4 Fuzzing")
    doc.add_paragraph("Hypothesis fuzzes the serial validator with 500 random strings per run. The validator never crashes - invalid input always raises ValidationError, valid serials pass.")

    h2(doc, "14.5 Defect Log")
    add_figure(doc, "20-defect-log.png", "Figure 14.5 - Defect log with severity, fix, and retest status")
    t = doc.add_table(rows=1, cols=5); t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "ID"; hdr[1].text = "Severity"; hdr[2].text = "Description"; hdr[3].text = "Fix"; hdr[4].text = "Status"
    for r in [
        ("DEF-001", "Medium", "Empty shipment_id crashes API with 500", "Input validation", "Fixed"),
        ("DEF-002", "High", "Vulnerable deps: python-jose, ecdsa", "Removed unused packages", "Fixed"),
        ("DEF-003", "Low", "datetime.utcnow() deprecation warning", "Deferred to Sprint 3", "Open"),
    ]:
        cells = t.add_row().cells
        for i, val in enumerate(r): cells[i].text = val
    style_table(t)

    h2(doc, "14.6 Traceability Thread (SR-03)")
    doc.add_paragraph("SR-03 (signed + hash-chained + audit-logged ownership transfer) is enforced by OwnershipService and verified by two integration tests: test_integration_transfer_writes_to_both_ledger_and_audit and test_integration_chain_grows_across_multiple_transfers.")
    doc.add_page_break()


# ---------- main ----------

def main():
    doc = Document()
    set_base_font(doc)
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
    add_phase_3(doc)
    add_phase_4(doc)
    add_phase_5(doc)
    add_phase_6(doc)
    add_phase_7(doc)
    add_phase_8(doc)
    add_phase_9(doc)
    add_phase_10(doc)
    add_phase_11(doc)
    add_phase_12(doc)
    add_phase_13(doc)
    add_phase_14(doc)
    # Future phases appended here.

    doc.save(DOC_PATH)
    print(f"[OK] Styled report written to: {DOC_PATH}")


if __name__ == "__main__":
    main()
