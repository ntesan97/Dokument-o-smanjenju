# ── Imports ───────────────────────────────────────────────────────────────────
import streamlit as st
import tempfile
import os
from pathlib import Path
from datetime import datetime, date
import pandas as pd
from lxml import etree

# =============================================================================
# SELLER DETAILS — fill in your company's permanent data here
# =============================================================================
SELLER = {
    "pib":          "110014338",
    "name":         "SERVIER doo",
    "street":       "Milutina Milankovića 11a",
    "city":         "Novi Beograd",
    "post_code":    "11070",
    "country":      "RS",
    "mb":           "21285293",
    "email":        "fakture@servier.rs",
    "bank_account": "325-950050031087-338",
}

# =============================================================================
# CONVERSION LOGIC
# =============================================================================
NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cec": "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "xsd": "http://www.w3.org/2001/XMLSchema",
    "sbt": "http://mfin.gov.rs/srbdt/srbdtext",
    "":    "urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2",
}


def _fmt_date(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s[:10]


def _str(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def _dec(val, decimals: int = 2) -> str:
    try:
        return f"{float(val):.{decimals}f}"
    except (TypeError, ValueError):
        return "0.00"


def _add(parent, tag: str, text: str, **attribs):
    parts = tag.split(":")
    qname = etree.QName(NS[parts[0]], parts[1]) if len(parts) == 2 else tag
    el = etree.SubElement(parent, qname, **attribs)
    el.text = text
    return el


def _sub(parent, tag: str):
    parts = tag.split(":")
    qname = etree.QName(NS[parts[0]], parts[1]) if len(parts) == 2 else tag
    return etree.SubElement(parent, qname)


def _read_kv(df: pd.DataFrame) -> dict:
    kv = {}
    for _, row in df.iterrows():
        k = _str(row.iloc[0])
        v = row.iloc[1] if len(row) > 1 else None
        if k and not pd.isna(v) and v is not None:
            kv[k] = v
    return kv


def _read_lines(df: pd.DataFrame) -> list:
    """Read credit note lines — header is row index 1."""
    headers = [_str(df.iloc[1, c]) for c in range(df.shape[1])]
    lines = []
    for i in range(2, df.shape[0]):
        row = df.iloc[i]
        if _str(row.iloc[0]) == "":
            continue
        lines.append({headers[c]: row.iloc[c] for c in range(len(headers))})
    return lines


def _vat_rate_from_group(group_str: str) -> float:
    """Extract VAT % from posting group name like 'DOBRA-10%' or 'DOBRA-20%'."""
    import re
    m = re.search(r"(\d+)%", str(group_str))
    if m:
        return float(m.group(1))
    return 10.0  # fallback


def build_credit_note_xml(xlsx_path: str, orig_invoice_no: str, orig_invoice_date: str) -> bytes:
    xl = pd.ExcelFile(xlsx_path)

    gen_df  = pd.read_excel(xl, sheet_name="General",                         header=None)
    lines_df= pd.read_excel(xl, sheet_name="Edit - Posted Sales Credit Mem",  header=None)
    tot_df  = pd.read_excel(xl, sheet_name="Edit - Posted Sales Credit Mem1", header=None)
    inv2_df = pd.read_excel(xl, sheet_name="Invoicing",                        header=None)
    reg_df  = pd.read_excel(xl, sheet_name="Registration Numbers",             header=None)

    gen   = _read_kv(gen_df)
    tot   = _read_kv(tot_df)
    inv2  = _read_kv(inv2_df)
    reg   = _read_kv(reg_df)
    lines = _read_lines(lines_df)

    # Header fields
    doc_no     = _str(gen.get("No.", ""))
    issue_date = date.today().strftime("%Y-%m-%d")  # always today per schema requirement
    vat_date   = _fmt_date(gen.get("VAT Date", gen.get("Posting Date")))
    ext_doc_no = _str(gen.get("External Document No.", ""))

    buyer_name   = _str(gen.get("Sell-to Customer Name",  inv2.get("Bill-to Name", "")))
    # Use real PIB from Registration Numbers sheet, not internal BC customer code
    buyer_pib    = str(int(float(reg.get("VAT Registration No.", 0) or 0))).zfill(9)
    buyer_mb     = str(int(float(reg.get("Registration No.", 0) or 0))).zfill(8)
    buyer_street = _str(gen.get("Sell-to Address",        inv2.get("Bill-to Address", "")))
    buyer_city   = _str(gen.get("Sell-to City",           inv2.get("Bill-to City", "")))
    buyer_zip    = _str(gen.get("Sell-to Post Code",      inv2.get("Bill-to Post Code", "")))

    total_excl_vat = float(tot.get("Total Excl. VAT (RSD)", 0) or 0)
    total_vat      = float(tot.get("Total VAT (RSD)", 0) or 0)
    total_incl_vat = float(tot.get("Total Incl. VAT (RSD)", 0) or 0)

    # VAT groups from line posting groups
    vat_groups = {}
    line_ext_total = 0.0
    for ln in lines:
        line_amt = float(ln.get("Line Amount Excl. VAT", 0) or 0)
        line_ext_total += line_amt
        vat_rate = _vat_rate_from_group(ln.get("VAT Prod. Posting Group", ""))
        vg = vat_groups.setdefault(vat_rate, {"taxable": 0.0, "tax": 0.0})
        vg["taxable"] += line_amt
        vg["tax"]     += line_amt * (vat_rate / 100)

    if not vat_groups:
        vat_groups[10.0] = {"taxable": total_excl_vat, "tax": total_vat}

    # ── Build XML ─────────────────────────────────────────────────────────────
    nsmap = {
        None:  NS[""],
        "cbc": NS["cbc"],
        "cac": NS["cac"],
        "cec": NS["cec"],
        "xsi": NS["xsi"],
        "xsd": NS["xsd"],
        "sbt": NS["sbt"],
    }
    root = etree.Element(etree.QName(NS[""], "CreditNote"), nsmap=nsmap)

    _add(root, "cbc:CustomizationID",
         "urn:cen.eu:en16931:2017#compliant#urn:mfin.gov.rs:srbdt:2022")
    _add(root, "cbc:ID", doc_no)
    _add(root, "cbc:IssueDate", issue_date)
    _add(root, "cbc:CreditNoteTypeCode", "381")
    if ext_doc_no:
        _add(root, "cbc:Note", ext_doc_no)
    _add(root, "cbc:DocumentCurrencyCode", "RSD")

    # BillingReference — original invoice this credit note corrects
    br  = _sub(root, "cac:BillingReference")
    idr = _sub(br,   "cac:InvoiceDocumentReference")
    _add(idr, "cbc:ID", orig_invoice_no)
    if orig_invoice_date:
        _add(idr, "cbc:IssueDate", orig_invoice_date)

    # Supplier
    sup_party = _sub(_sub(root, "cac:AccountingSupplierParty"), "cac:Party")
    _add(sup_party, "cbc:EndpointID", SELLER["pib"]).set("schemeID", "9948")
    _add(_sub(sup_party, "cac:PartyName"), "cbc:Name", SELLER["name"])
    pa = _sub(sup_party, "cac:PostalAddress")
    _add(pa, "cbc:StreetName", SELLER["street"])
    _add(pa, "cbc:CityName",   SELLER["city"])
    _add(pa, "cbc:PostalZone", SELLER["post_code"])
    _add(_sub(pa, "cac:Country"), "cbc:IdentificationCode", SELLER["country"])
    pts = _sub(sup_party, "cac:PartyTaxScheme")
    _add(pts, "cbc:CompanyID", f"RS{SELLER['pib']}")
    _add(_sub(pts, "cac:TaxScheme"), "cbc:ID", "VAT")
    ple = _sub(sup_party, "cac:PartyLegalEntity")
    _add(ple, "cbc:RegistrationName", SELLER["name"])
    _add(ple, "cbc:CompanyID", SELLER["mb"])
    _add(_sub(sup_party, "cac:Contact"), "cbc:ElectronicMail", SELLER["email"])

    # Customer
    cust_party = _sub(_sub(root, "cac:AccountingCustomerParty"), "cac:Party")
    _add(cust_party, "cbc:EndpointID", buyer_pib).set("schemeID", "9948")
    _add(_sub(cust_party, "cac:PartyName"), "cbc:Name", buyer_name)
    cpa = _sub(cust_party, "cac:PostalAddress")
    _add(cpa, "cbc:StreetName", buyer_street)
    _add(cpa, "cbc:CityName",   buyer_city)
    if buyer_zip:
        _add(cpa, "cbc:PostalZone", buyer_zip)
    _add(_sub(cpa, "cac:Country"), "cbc:IdentificationCode", "RS")
    cpts = _sub(cust_party, "cac:PartyTaxScheme")
    _add(cpts, "cbc:CompanyID", f"RS{buyer_pib}")
    _add(_sub(cpts, "cac:TaxScheme"), "cbc:ID", "VAT")
    cust_ple = _sub(cust_party, "cac:PartyLegalEntity")
    _add(cust_ple, "cbc:RegistrationName", buyer_name)
    _add(cust_ple, "cbc:CompanyID", buyer_mb)

    # Delivery & payment
    _add(_sub(root, "cac:Delivery"), "cbc:ActualDeliveryDate", vat_date)
    pm = _sub(root, "cac:PaymentMeans")
    _add(pm, "cbc:PaymentMeansCode", "30")
    _add(_sub(pm, "cac:PayeeFinancialAccount"), "cbc:ID", SELLER["bank_account"])

    # TaxTotal
    tt = _sub(root, "cac:TaxTotal")
    _add(tt, "cbc:TaxAmount", _dec(total_vat)).set("currencyID", "RSD")
    for rate, grp in sorted(vat_groups.items()):
        tst = _sub(tt, "cac:TaxSubtotal")
        _add(tst, "cbc:TaxableAmount", _dec(grp["taxable"])).set("currencyID", "RSD")
        _add(tst, "cbc:TaxAmount",     _dec(grp["tax"])).set("currencyID", "RSD")
        tc2 = _sub(tst, "cac:TaxCategory")
        _add(tc2, "cbc:ID", "S")
        _add(tc2, "cbc:Percent", str(int(rate)))
        _add(_sub(tc2, "cac:TaxScheme"), "cbc:ID", "VAT")

    # LegalMonetaryTotal
    lmt = _sub(root, "cac:LegalMonetaryTotal")
    _add(lmt, "cbc:LineExtensionAmount",  _dec(line_ext_total)).set("currencyID", "RSD")
    _add(lmt, "cbc:TaxExclusiveAmount",   _dec(total_excl_vat)).set("currencyID", "RSD")
    _add(lmt, "cbc:TaxInclusiveAmount",   _dec(total_incl_vat)).set("currencyID", "RSD")
    _add(lmt, "cbc:AllowanceTotalAmount", "0.00").set("currencyID", "RSD")
    _add(lmt, "cbc:PrepaidAmount",        "0.00").set("currencyID", "RSD")
    _add(lmt, "cbc:PayableRoundingAmount","0.00").set("currencyID", "RSD")
    _add(lmt, "cbc:PayableAmount",        _dec(total_incl_vat)).set("currencyID", "RSD")

    # CreditNote lines
    for idx, ln in enumerate(lines, start=1):
        qty      = _str(ln.get("Quantity", "1"))
        uom      = _str(ln.get("Unit of Measure Code", "XKI")) or "XKI"
        desc     = _str(ln.get("Description", ""))
        item_no  = _str(ln.get("No.", ""))
        unit_price = float(ln.get("Unit Price Excl. VAT", 0) or 0)
        line_amt   = float(ln.get("Line Amount Excl. VAT", 0) or 0)
        vat_rate   = _vat_rate_from_group(ln.get("VAT Prod. Posting Group", ""))

        cnl = _sub(root, "cac:CreditNoteLine")
        _add(cnl, "cbc:ID", str(idx))
        _add(cnl, "cbc:CreditedQuantity", qty).set("unitCode", uom)
        _add(cnl, "cbc:LineExtensionAmount", _dec(line_amt)).set("currencyID", "RSD")

        item = _sub(cnl, "cac:Item")
        _add(item, "cbc:Name", desc)
        _add(_sub(item, "cac:SellersItemIdentification"), "cbc:ID", item_no)
        ctc = _sub(item, "cac:ClassifiedTaxCategory")
        _add(ctc, "cbc:ID", "S")
        _add(ctc, "cbc:Percent", str(int(vat_rate)))
        _add(_sub(ctc, "cac:TaxScheme"), "cbc:ID", "VAT")

        _add(_sub(cnl, "cac:Price"), "cbc:PriceAmount", _dec(unit_price)).set("currencyID", "RSD")

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")


# =============================================================================
# STREAMLIT UI
# =============================================================================
st.set_page_config(
    page_title="Credit Memo → UBL XML",
    page_icon="📋",
    layout="centered",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background-color: #0f0f0f; color: #e8e8e8; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { max-width: 640px; padding-top: 4rem; padding-bottom: 4rem; }

.app-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.1rem; font-weight: 600;
    letter-spacing: 0.15em; text-transform: uppercase;
    color: #f5a623; margin-bottom: 0.25rem;
}
.app-subtitle {
    font-size: 0.85rem; color: #555;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.05em; margin-bottom: 3rem;
}
.section-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem; letter-spacing: 0.12em;
    text-transform: uppercase; color: #444;
    margin-top: 2rem; margin-bottom: 0.4rem;
}

[data-testid="stFileUploader"] {
    background: #1a1a1a; border: 1.5px dashed #2e2e2e;
    border-radius: 4px; padding: 1rem; transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #f5a623; }
[data-testid="stFileUploader"] label {
    color: #888 !important;
    font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem;
}

/* Text inputs */
[data-testid="stTextInput"] input {
    background: #1a1a1a !important;
    border: 1px solid #2e2e2e !important;
    border-radius: 3px !important;
    color: #e8e8e8 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: #f5a623 !important;
    box-shadow: 0 0 0 1px #f5a62340 !important;
}
[data-testid="stTextInput"] label {
    color: #666 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.75rem !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}

.stButton > button {
    background: #f5a623; color: #0f0f0f;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
    font-size: 0.85rem; letter-spacing: 0.1em; text-transform: uppercase;
    border: none; border-radius: 3px; padding: 0.65rem 2rem;
    width: 100%; margin-top: 1.5rem; transition: background 0.15s, transform 0.1s;
}
.stButton > button:hover { background: #ffbb44; transform: translateY(-1px); }
.stButton > button:active { transform: translateY(0); }

.stDownloadButton > button {
    background: transparent; color: #f5a623;
    font-family: 'IBM Plex Mono', monospace; font-weight: 600;
    font-size: 0.85rem; letter-spacing: 0.1em; text-transform: uppercase;
    border: 1.5px solid #f5a623; border-radius: 3px;
    padding: 0.65rem 2rem; width: 100%; margin-top: 0.5rem;
    transition: all 0.15s;
}
.stDownloadButton > button:hover { background: #f5a623; color: #0f0f0f; }

.info-row {
    display: flex; justify-content: space-between;
    font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #444;
    margin-top: 3rem; padding-top: 1rem; border-top: 1px solid #1e1e1e;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="app-title">Credit Memo Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">xlsx → ubl creditnote xml · serbian e-faktura</div>', unsafe_allow_html=True)

# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop your Credit Memo Excel file here or click to browse",
    type=["xlsx"],
    label_visibility="visible",
)

# ── Original invoice reference (required by schema) ───────────────────────────
st.markdown('<div class="section-label">Original Invoice Reference (faktura na koju se odnosi ovo knjižno odobrenje)</div>', unsafe_allow_html=True)

col1, col2 = st.columns([3, 2])
with col1:
    orig_invoice_no = st.text_input(
        "Broj originalne fakture (npr. SI25/26-0097)",
        placeholder="e.g. SI25/26-0097",
    )
with col2:
    orig_invoice_date_val = st.date_input(
        "Invoice Date",
        value=date.today(),
    )
    orig_invoice_date = orig_invoice_date_val.strftime("%Y-%m-%d")

# ── Convert ───────────────────────────────────────────────────────────────────
if uploaded:
    st.markdown(f"**`{uploaded.name}`** — ready to convert")

    if st.button("Convert to XML"):
        if not orig_invoice_no.strip():
            st.error("Please enter the original invoice number before converting.")
        else:
            with st.spinner("Processing..."):
                tmp_path = None
                try:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        tmp.write(uploaded.read())
                        tmp_path = tmp.name

                    xml_bytes = build_credit_note_xml(
                        tmp_path,
                        orig_invoice_no.strip(),
                        orig_invoice_date.strip(),
                    )
                    os.unlink(tmp_path)

                    output_name = Path(uploaded.name).stem + ".xml"
                    st.success(f"✓ Converted successfully — {len(xml_bytes):,} bytes")
                    st.download_button(
                        label="⬇  Download XML",
                        data=xml_bytes,
                        file_name=output_name,
                        mime="application/xml",
                    )

                except Exception as e:
                    if tmp_path:
                        try:
                            os.unlink(tmp_path)
                        except Exception:
                            pass
                    st.error(f"Conversion failed:\n\n{e}")

st.markdown("""
<div class="info-row">
    <span>Reads: General · Credit Memo Lines · Totals · Invoicing</span>
    <span>Schema: EN 16931 / mfin.gov.rs 2022</span>
</div>
""", unsafe_allow_html=True)
