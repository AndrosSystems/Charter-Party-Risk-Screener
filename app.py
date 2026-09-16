
import io
import json
import re

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="Charter Party Risk Screener", page_icon="⚓", layout="wide")

SHORT_DISCLAIMER = """
This tool provides **preliminary commercial and contractual decision support only**. It is not legal advice,
does not determine the legal effect or enforceability of any clause, and must not be relied upon as the sole
basis for negotiation, performance, claims, termination, sanctions, war-risk, cargo, payment or other material
decisions. Qualified maritime legal and commercial review is required before reliance.
"""

DISCLAIMER = """
**Important legal & use disclaimer**

This application is provided solely as a preliminary commercial and contractual decision-support and
issue-spotting tool. It does **not** constitute legal advice, maritime legal advice, professional consultancy,
contractual interpretation, or a determination of the validity, enforceability, effect, suitability or risk of
any charter party, clause, fixture or transaction.

Results are generated from automated text screening, user-provided information and, where enabled,
AI-assisted analysis. Outputs may be incomplete, inaccurate, outdated, affected by extraction errors, or
misleading where contractual context is missing. The significance of any wording depends on the complete
charter party, recap, incorporated forms and riders, amendments, bills of lading, governing law, factual
circumstances, trade, vessel, cargo, counterparties, sanctions exposure, regulatory regime and negotiations
between the parties.

A finding marked **Found**, **Possible** or **Not found** is a text-screening result only. It is not a legal
conclusion. In particular, **Not found** does not establish that an issue is absent from the agreement or its
legal effect, and a detected term does not establish that the relevant risk has been adequately allocated.

Regulatory and sanctions-related outputs are not live compliance checks, sanctions screening, counterparty
verification or regulatory advice. Laws, sanctions regimes, BIMCO clauses, regulatory requirements and market
practice can change. Users must verify current requirements and authoritative source material before acting.

AI-generated comments are hypotheses, review prompts and summaries only. They may contain errors or omit
material issues. Users remain responsible for reviewing the complete contract and for all commercial,
operational and legal decisions made in reliance on or following use of this application. Material decisions —
including fixing, signing, performance, withdrawal, termination, refusal of orders, deviation, sanctions
compliance, cargo handling, payment, claims or dispute action — should be reviewed by appropriately qualified
maritime legal and commercial professionals and handled under the user's approved organisational procedures.

No representation or warranty is made regarding the accuracy, completeness, reliability, suitability or
availability of the information or outputs produced by this application. **To the maximum extent permitted by
applicable law**, the developer and contributors accept no liability for losses, damages, claims, costs or
consequences arising from use of, inability to use, or reliance upon the application or its outputs. Nothing in
this disclaimer excludes or limits liability where such exclusion or limitation is prohibited by applicable law.

Users must not upload confidential, privileged, commercially sensitive or personal information to third-party
AI services unless they are authorised to do so and have assessed the applicable confidentiality, privacy,
data-processing and professional-obligation requirements. This MVP contains no application database for
contract storage, but hosting and AI providers may process or retain information under their own terms and
privacy policies.

Use of this application does not create a lawyer-client, consultant-client, fiduciary or other professional
relationship. The application should be used as a structured review aid, not as a substitute for professional
judgment.
"""

ISSUES = [
    {"area":"Governing law & dispute resolution","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":5,
     "patterns":[r"\bgoverning law\b",r"\blaw and arbitration\b",r"\barbitration\b",r"\bjurisdiction\b",r"\benglish law\b",r"\blondon arbitration\b"],
     "why":"Determines the legal framework, forum, procedure and often the interpretation of the entire contract."},
    {"area":"Payment of freight / hire","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":5,
     "patterns":[r"\bfreight\b",r"\bhire\b",r"\bpayment\b",r"\badvance payment\b",r"\bremittance\b",r"\bdue date\b"],
     "why":"Payment timing, deductions, banking mechanics and default consequences are core commercial terms."},
    {"area":"Withdrawal / suspension / termination for non-payment","types":["Time","Bareboat","Other / Unknown"],"priority":5,
     "patterns":[r"\bwithdraw(al)?\b",r"\bsuspend\b",r"\banti[- ]technicality\b",r"\bgrace period\b",r"\bnon[- ]payment\b",r"\bterminate\b"],
     "why":"Non-payment remedies can create major operational and legal exposure."},
    {"area":"Laytime, NOR & demurrage","types":["Voyage","Other / Unknown"],"priority":5,
     "patterns":[r"\blaytime\b",r"\bdemurrage\b",r"\bnotice of readiness\b",r"\bNOR\b",r"\btime to count\b",r"\bdespatch\b"],
     "why":"Defines time counting and delay allocation at load/discharge ports."},
    {"area":"Off-hire","types":["Time","Other / Unknown"],"priority":5,
     "patterns":[r"\boff[- ]hire\b",r"\boffhire\b",r"\bhire shall cease\b",r"\bloss of time\b"],
     "why":"Allocates time-related financial risk when vessel service is interrupted or impaired."},
    {"area":"Safe port / safe berth / trading limits","types":["Voyage","Time","Other / Unknown"],"priority":5,
     "patterns":[r"\bsafe port\b",r"\bsafe berth\b",r"\btrading limit",r"\btrading area\b",r"\bport.*safe\b",r"\bberth.*safe\b"],
     "why":"Affects employment orders, port safety allocation and exposure to unsafe destinations."},
    {"area":"War risks / piracy / dangerous areas","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":5,
     "patterns":[r"\bwar risk",r"\bCONWARTIME\b",r"\bVOYWAR\b",r"\bpira(cy|tes)\b",r"\bhostilit",r"\bwar zone\b",r"\bblockade\b"],
     "why":"Allocates rights, costs and responsibilities where vessel, cargo or crew may be exposed to war-related risks."},
    {"area":"Sanctions","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":5,
     "patterns":[r"\bsanction",r"\brestricted party\b",r"\bdesignated person\b",r"\bOFAC\b",r"\bEU sanctions\b",r"\bprohibited trade\b"],
     "why":"Sanctions can affect counterparties, cargo, trade, banks, ports and performance rights."},
    {"area":"Cargo description / exclusions / dangerous cargo","types":["Voyage","Time","Other / Unknown"],"priority":5,
     "patterns":[r"\bcargo\b",r"\bdangerous cargo\b",r"\bhazardous cargo\b",r"\bcargo exclusion",r"\bIMDG\b"],
     "why":"Cargo obligations and exclusions are central to operational, safety and liability allocation."},
    {"area":"Bills of lading / letters of indemnity","types":["Voyage","Time","Other / Unknown"],"priority":5,
     "patterns":[r"\bbill[s]? of lading\b",r"\bB\/L\b",r"\bletter of indemnity\b",r"\bLOI\b",r"\bwithout presentation\b"],
     "why":"Document issuance and delivery against LOIs can create substantial cargo and P&I exposure."},
    {"area":"Owners' lien / security","types":["Voyage","Time","Other / Unknown"],"priority":4,
     "patterns":[r"\blien\b",r"\bowners'? lien\b",r"\bcargo lien\b",r"\bsub[- ]freight\b",r"\bsub[- ]hire\b"],
     "why":"Security rights can materially affect recovery of unpaid sums."},
    {"area":"Speed & consumption / performance warranty","types":["Time","Other / Unknown"],"priority":4,
     "patterns":[r"\bspeed\b",r"\bconsumption\b",r"\bperformance warranty\b",r"\bweather condition",r"\bbeaufort\b",r"\babout\b.*\bknots\b"],
     "why":"Performance wording can generate significant underperformance and bunker-consumption claims."},
    {"area":"Bunkers / fuel quality / delivery-redelivery quantities","types":["Time","Bareboat","Other / Unknown"],"priority":4,
     "patterns":[r"\bbunker",r"\bfuel oil\b",r"\bVLSFO\b",r"\bMGO\b",r"\bfuel specification\b",r"\bredelivery.*fuel\b"],
     "why":"Fuel quality, quantity, price and responsibility affect both operations and claims."},
    {"area":"Delivery / redelivery / final voyage","types":["Time","Bareboat","Other / Unknown"],"priority":4,
     "patterns":[r"\bdelivery\b",r"\bredelivery\b",r"\bfinal voyage\b",r"\blast voyage\b",r"\bdelivery range\b",r"\bredelivery range\b"],
     "why":"Timing, location and condition at delivery/redelivery are major exposure points."},
    {"area":"Maintenance / repairs / drydocking","types":["Time","Bareboat","Other / Unknown"],"priority":4,
     "patterns":[r"\bmaintenance\b",r"\brepair",r"\bdry[- ]?dock",r"\bclass\b",r"\bseaworthy\b",r"\bmaintain the vessel\b"],
     "why":"Defines technical responsibility, downtime allocation and preservation of class/condition."},
    {"area":"Employment, orders & indemnities","types":["Time","Other / Unknown"],"priority":4,
     "patterns":[r"\bemployment\b",r"\border[s]?\b",r"\bindemnif",r"\bindemnity\b",r"\bmaster.*orders\b",r"\bcharterers'? orders\b"],
     "why":"Operational control and indemnity wording affects the consequences of charterers' employment orders."},
    {"area":"Stevedore damage","types":["Voyage","Time","Other / Unknown"],"priority":3,
     "patterns":[r"\bstevedore\b",r"\bstevedoring\b",r"\bdamage.*stevedore\b"],
     "why":"Responsibility, notice and repair timing for stevedore damage can create recurring disputes."},
    {"area":"Exceptions / force majeure","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":4,
     "patterns":[r"\bforce majeure\b",r"\bexceptions\b",r"\bexcepted peril",r"\bact of god\b",r"\bprevent performance\b"],
     "why":"Excuse-of-performance wording varies materially and should be read in the context of the entire charter."},
    {"area":"Deviation / liberty / route","types":["Voyage","Time","Other / Unknown"],"priority":3,
     "patterns":[r"\bdeviation\b",r"\bliberty\b",r"\broute\b",r"\bdeviate\b",r"\breasonable deviation\b"],
     "why":"Route and liberty wording can affect performance, cargo exposure and time/cost allocation."},
    {"area":"Ice / weather / port closure","types":["Voyage","Time","Other / Unknown"],"priority":3,
     "patterns":[r"\bice\b",r"\bweather\b",r"\bport closure\b",r"\bweather routing\b",r"\bheavy weather\b"],
     "why":"Weather and ice provisions may affect orders, delays, extra costs and route decisions."},
    {"area":"Cyber security","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":3,
     "patterns":[r"\bcyber\b",r"\binformation security\b",r"\bsecurity incident\b",r"\bmalware\b",r"\bransomware\b"],
     "why":"Cyber incidents can disrupt performance, data, navigation and contractual communications."},
    {"area":"EU ETS / emission allowances","types":["Voyage","Time","Other / Unknown"],"priority":4,
     "patterns":[r"\bEU ETS\b",r"\bemission allowance",r"\bemissions trading\b",r"\bEUA\b",r"\ballowances\b.*\bemission"],
     "why":"Where applicable, allocation of emissions costs, data and allowance-transfer responsibilities should be explicit."},
    {"area":"FuelEU Maritime","types":["Voyage","Time","Other / Unknown"],"priority":4,
     "patterns":[r"\bFuelEU\b",r"\bGHG intensity\b",r"\bcompliance balance\b",r"\bFuelEU penalty\b"],
     "why":"For relevant EU trades, FuelEU obligations can create cost, fuel-choice, pooling/banking and data-allocation issues."},
    {"area":"CII / carbon-intensity operations","types":["Time","Other / Unknown"],"priority":3,
     "patterns":[r"\bCII\b",r"\bcarbon intensity indicator\b",r"\bCII rating\b",r"\battained CII\b"],
     "why":"Operational carbon-intensity obligations can affect vessel employment, speed, instructions and performance allocation."},
    {"area":"Electronic bills of lading","types":["Voyage","Time","Other / Unknown"],"priority":2,
     "patterns":[r"\belectronic bill",r"\beBL\b",r"\bpaperless trading\b",r"\belectronic.*bill[s]? of lading\b"],
     "why":"If electronic trade documents are contemplated, system, liability and cost allocation should be understood."},
    {"area":"Notices / communications","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":3,
     "patterns":[r"\bnotice\b",r"\bnotification\b",r"\bin writing\b",r"\bemail\b",r"\bdeemed received\b"],
     "why":"Many contractual rights depend on valid and timely notices."},
    {"area":"Confidentiality / data","types":["Voyage","Time","Bareboat","Other / Unknown"],"priority":2,
     "patterns":[r"\bconfidential",r"\bdata protection\b",r"\bpersonal data\b",r"\bnon[- ]disclosure\b"],
     "why":"Commercial and operational information may require confidentiality and data-handling rules."},
]

def extract_text_from_pdf(file_bytes):
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((p.extract_text() or "") for p in reader.pages)

def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs)

def load_contract(uploaded):
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith(".pdf"):
        return extract_text_from_pdf(raw)
    if name.endswith(".docx"):
        return extract_text_from_docx(raw)
    if name.endswith(".txt") or name.endswith(".md"):
        return raw.decode("utf-8", errors="ignore")
    raise ValueError("Unsupported file type.")

def clean_text(text):
    text = text.replace("\x00", " ")
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def find_matches(text, patterns, context=260, max_hits=4):
    hits = []
    for pat in patterns:
        try:
            rx = re.compile(pat, flags=re.I | re.S)
        except Exception:
            continue
        for m in rx.finditer(text):
            start = max(0, m.start() - context)
            end = min(len(text), m.end() + context)
            snippet = re.sub(r"\s+", " ", text[start:end].strip())
            hits.append((m.start(), snippet))
    hits.sort(key=lambda x: x[0])
    out = []
    for _, s in hits:
        if not any(s[:120] == old[:120] for old in out):
            out.append(s)
        if len(out) >= max_hits:
            break
    return out

def detection_state(hits):
    if len(hits) >= 2:
        return "Found"
    if len(hits) == 1:
        return "Possible"
    return "Not found"

def priority_label(base_priority, state):
    factor = {"Not found":1.0, "Possible":0.75, "Found":0.35}[state]
    score = base_priority * factor
    return "High" if score >= 4.0 else ("Medium" if score >= 2.4 else "Low")

def screen_contract(text, contract_type):
    rows = []
    for issue in ISSUES:
        applicable = contract_type in issue["types"]
        hits = find_matches(text, issue["patterns"]) if applicable else []
        state = detection_state(hits) if applicable else "Not applicable"
        rows.append({
            "Risk / Clause Area": issue["area"],
            "Applicable": "Yes" if applicable else "No",
            "Detection": state,
            "Review Priority": priority_label(issue["priority"], state) if applicable else "—",
            "Why It Matters": issue["why"],
            "Evidence Snippet": hits[0] if hits else "",
            "Additional Hits": max(0, len(hits)-1),
            "User Review": "Unreviewed" if applicable else "Not applicable",
            "User Notes": "",
        })
    return pd.DataFrame(rows)

def get_api_key():
    try:
        if "OPENROUTER_API_KEY" in st.secrets:
            return st.secrets["OPENROUTER_API_KEY"]
    except Exception:
        pass
    return st.session_state.get("manual_api_key","")

def call_openrouter(prompt, model="openrouter/free", max_tokens=2200):
    key = get_api_key()
    if not key:
        raise ValueError("No OpenRouter API key configured.")
    r = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={"Authorization":f"Bearer {key}","Content-Type":"application/json","X-Title":"Charter Party Risk Screener"},
        json={"model":model,"messages":[{"role":"user","content":prompt}],"temperature":0.15,"max_tokens":max_tokens},
        timeout=120,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]

def build_markdown_report(meta, results, ai_text=None):
    lines = [
        "# Charter Party Risk Screening Report","",
        f"**Contract type:** {meta['contract_type']}",
        f"**Perspective:** {meta['perspective']}",
        f"**Document label:** {meta['document_label']}","",
        "## Screening summary","",
    ]
    for _, r in results.iterrows():
        if r["Applicable"] == "No":
            continue
        lines += [
            f"### {r['Risk / Clause Area']}",
            f"- Detection: **{r['Detection']}**",
            f"- Review priority: **{r['Review Priority']}**",
            f"- User review: **{r['User Review']}**",
        ]
        if str(r.get("User Notes","")).strip():
            lines.append(f"- User notes: {r['User Notes']}")
        if str(r.get("Evidence Snippet","")).strip():
            lines.append(f"- Evidence snippet: {r['Evidence Snippet']}")
        lines.append("")
    if ai_text:
        lines += ["## AI-assisted issue-spotting brief","",ai_text,""]
    lines += ["---","",DISCLAIMER.replace("**","")]
    return "\n".join(lines)

st.title("⚓ Charter Party Risk Screener")
st.subheader("Commercial & Contractual Decision Support")
st.caption("Preliminary clause coverage screening • human review • AI-assisted issue spotting • exportable risk register")

st.warning(SHORT_DISCLAIMER)
with st.expander("Full legal & use disclaimer — read before use", expanded=False):
    st.markdown(DISCLAIMER)

usage_ack = st.checkbox(
    "I understand that this application provides preliminary decision support only and is not legal advice.",
    value=False,
)

with st.sidebar:
    st.header("Analysis settings")
    contract_type = st.selectbox("Charter / contract type",["Voyage","Time","Bareboat","Other / Unknown"])
    perspective = st.selectbox("Review perspective",["Neutral / Both parties","Owner / Disponent Owner","Charterer"])
    document_label = st.text_input("Document label", value="Charter Party Review")
    st.divider()
    st.header("AI")
    st.text_input("OpenRouter API key (optional if stored in Streamlit Secrets)",type="password",key="manual_api_key")
    ai_model = st.text_input("OpenRouter model", value="openrouter/free")
    st.caption("The deterministic screener works without AI.")

tab1, tab2, tab3 = st.tabs(["1 · Screen Contract","2 · Review Register","3 · Method & References"])

with tab1:
    st.header("Load charter-party text")
    st.write("Upload a PDF, DOCX or TXT file, or paste the contract text. The app screens only the text you provide.")
    uploaded = st.file_uploader("Upload contract",type=["pdf","docx","txt","md"])
    pasted = st.text_area("Or paste charter-party / recap / rider text",height=240,placeholder="Paste text here...")

    contract_text = ""
    if uploaded is not None:
        try:
            contract_text = load_contract(uploaded)
            st.success(f"Loaded {uploaded.name}")
        except Exception as e:
            st.error(f"Could not read the file: {e}")
    elif pasted.strip():
        contract_text = pasted

    contract_text = clean_text(contract_text)

    if contract_text:
        c1,c2,c3 = st.columns(3)
        c1.metric("Characters",f"{len(contract_text):,}")
        c2.metric("Approx. words",f"{len(contract_text.split()):,}")
        c3.metric("Contract type",contract_type)

        if st.button("Run deterministic screening",type="primary", disabled=not usage_ack):
            st.session_state["contract_text"] = contract_text
            st.session_state["screen_results"] = screen_contract(contract_text, contract_type)
            st.session_state["screen_contract_type"] = contract_type
            st.success("Screening complete. Open the Review Register tab.")
    else:
        st.info("Upload or paste a contract to begin.")

with tab2:
    st.header("Review register")
    results = st.session_state.get("screen_results")
    contract_text_saved = st.session_state.get("contract_text","")

    if results is None:
        st.info("Run a screening in Tab 1 first.")
    else:
        if st.session_state.get("screen_contract_type") != contract_type:
            st.warning("You changed the contract type after screening. Re-run the screening for correct applicability.")

        summary = results[results["Applicable"]=="Yes"]["Review Priority"].value_counts()
        a,b,c = st.columns(3)
        a.metric("High-priority review areas",int(summary.get("High",0)))
        b.metric("Medium-priority review areas",int(summary.get("Medium",0)))
        c.metric("Low-priority review areas",int(summary.get("Low",0)))

        review_options = ["Unreviewed","Confirmed adequate","Needs commercial review","Needs legal review","Confirmed issue","Not applicable"]

        edited = st.data_editor(
            results,use_container_width=True,hide_index=True,
            column_config={
                "User Review":st.column_config.SelectboxColumn("User Review",options=review_options,required=True),
                "User Notes":st.column_config.TextColumn("User Notes"),
                "Evidence Snippet":st.column_config.TextColumn("Evidence Snippet",width="large"),
                "Why It Matters":st.column_config.TextColumn("Why It Matters",width="medium"),
            },
            disabled=["Risk / Clause Area","Applicable","Detection","Review Priority","Why It Matters","Evidence Snippet","Additional Hits"],
            key="risk_register_editor",
        )
        st.session_state["screen_results"] = edited.copy()

        st.caption("'Not found' means the keyword screener did not find a clear match. It does not prove the legal issue is absent.")

        st.download_button("Download risk register CSV",edited.to_csv(index=False).encode("utf-8"),
                           file_name="charter_party_risk_register.csv",mime="text/csv")

        st.subheader("AI-assisted review")
        include_full_text = st.checkbox(
            "Send contract text to the AI service",
            value=False,
            help="If selected, the text is transmitted to the configured OpenRouter model. Do not use confidential, privileged or commercially sensitive text unless you are authorised to do so."
        )
        ai_transfer_ack = True
        if include_full_text:
            st.warning(
                "The selected contract text will be sent to a third-party AI provider. Provider processing and retention are governed by that provider's terms and privacy policy."
            )
            ai_transfer_ack = st.checkbox(
                "I am authorised to send this contract text to the configured AI provider.",
                value=False,
                key="ai_transfer_ack",
            )

        if st.button("Generate AI issue-spotting brief", disabled=(not usage_ack or not ai_transfer_ack)):
            compact_register = edited[["Risk / Clause Area","Applicable","Detection","Review Priority","Evidence Snippet","User Review","User Notes"]].to_dict(orient="records")
            source_text = contract_text_saved[:50000] if include_full_text else ""

            prompt = f"""
You are a maritime charter-party contract review assistant supporting qualified commercial and legal professionals.

CONTRACT TYPE: {contract_type}
REVIEW PERSPECTIVE: {perspective}

DETERMINISTIC SCREENING REGISTER:
{json.dumps(compact_register, indent=2)}

CONTRACT TEXT PROVIDED TO MODEL:
{source_text if source_text else "[Not provided — rely only on the screening register and snippets]"}

Produce:
1. Executive commercial review
2. Highest-priority clauses/issues to review
3. Owner-side exposure
4. Charterer-side exposure
5. Ambiguous or missing allocations
6. Questions to resolve before fixing / signing
7. Operational handover points for chartering, operations, claims and technical teams
8. Regulatory / compliance clauses that may be relevant
9. Limitations

Rules:
- Preliminary issue spotting only; not legal advice.
- Do not state that a clause is valid, invalid, enforceable or unenforceable.
- Do not invent wording not present in the supplied evidence.
- Distinguish text actually detected, absence of a detected term, and a review question.
- Do not treat 'Not found' as proof that an issue is legally unaddressed.
- Do not reproduce or reconstruct copyrighted standard charter-party wording.
- You may refer to public industry topics such as war risks, sanctions, force majeure, cyber security,
  EU ETS, FuelEU Maritime, CII and electronic bills of lading, but do not claim that a BIMCO clause
  is incorporated unless the supplied text says so.
- If governing law or incorporated form is unclear, identify that uncertainty.
- If perspective is Owner or Charterer, explain exposure from that perspective while noting the counterparty's concern.
- Keep the brief practical for chartering / operations teams and suitable for escalation to maritime counsel.
"""
            try:
                with st.spinner("Generating AI review..."):
                    ai_text = call_openrouter(prompt,model=ai_model,max_tokens=2500)
                st.session_state["ai_review_text"] = ai_text
                st.markdown(ai_text)
                st.warning("AI-generated issue spotting only — qualified maritime legal/commercial review is required before reliance.")
            except Exception as e:
                st.error(f"AI request failed: {e}")

        ai_text = st.session_state.get("ai_review_text")
        if ai_text:
            report = build_markdown_report(
                {"contract_type":contract_type,"perspective":perspective,"document_label":document_label},
                edited,ai_text
            )
            st.download_button("Download full screening report (Markdown)",report.encode("utf-8"),
                               file_name="charter_party_screening_report.md",mime="text/markdown")

with tab3:
    st.header("How the screener works")
    st.markdown("""
### Deterministic layer
The app searches the text you provide for issue-specific terms and expressions and creates a structured review register.

The result is **clause coverage screening**, not legal interpretation.

- **Found** — multiple relevant term matches were located.
- **Possible** — one relevant match was located.
- **Not found** — no clear keyword/pattern match was located.
- **Not applicable** — the issue is outside the selected MVP charter-type profile.

The **Review Priority** is a triage priority only. It is not a probability of loss, legal risk score or claim value.

### Human review layer
The user can mark each area as:
- Confirmed adequate
- Needs commercial review
- Needs legal review
- Confirmed issue
- Not applicable

### AI layer
The AI receives the structured register and, only if the user explicitly opts in, a bounded portion of the contract text.

### Confidentiality and third-party AI processing
The deterministic screening can be used without sending the full contract text to the AI service.
If the user explicitly enables full-text AI review, the selected text is transmitted to the configured
third-party AI provider. Do not transmit confidential, privileged, commercially sensitive or personal
information unless you are authorised to do so and have assessed the relevant provider terms, privacy
requirements and professional obligations.

### Regulatory scope
Sanctions, emissions and regulatory flags are review prompts only. This MVP does not perform live sanctions
screening, counterparty verification or real-time legal/regulatory monitoring.
""")

    st.header("Public reference framework")
    st.markdown("""
This MVP uses public industry sources to define topics worth screening. It does not contain or reproduce copyrighted charter-party forms.

- BIMCO GENCON 2022 overview
- BIMCO War Risks Clauses: VOYWAR 2025 and CONWARTIME 2025
- BIMCO Sanctions Clause for Voyage Charter Parties 2020
- BIMCO Force Majeure Clause 2022
- BIMCO Cyber Security Clause 2019
- BIMCO ETS clause framework for time/voyage charter parties
- BIMCO FuelEU Maritime Clause for Time Charter Parties 2024
- BIMCO CII Operations Clause for Time Charter Parties 2022
- BIMCO Electronic Bills of Lading Clause 2014

Always consult current official text and explanatory material where parties intend to incorporate a standard clause.
""")
    st.warning(DISCLAIMER)
