# Charter Party Risk Screener

AI-assisted commercial and contractual decision-support tool for preliminary charter-party review. Upload or paste voyage, time or bareboat charter-party text to screen major risk areas, identify clause coverage, compare Owner/Charterer perspectives, add human validation and generate an optional AI-assisted issue-spotting brief.

## Features

- PDF, DOCX, TXT or pasted text
- Voyage, Time, Bareboat and Other/Unknown modes
- structured clause / risk review register
- Owner / Charterer / Neutral review perspective
- human confirmation and notes
- optional OpenRouter AI issue-spotting brief
- explicit opt-in before sending contract text to the AI provider
- CSV and Markdown export
- visible legal/use disclaimer and separate Terms of Use

## Important legal and use notice

This software is a **preliminary decision-support and issue-spotting tool only**. It is not legal advice, does not determine contractual validity or enforceability, and must not be relied upon as the sole basis for material commercial, operational or legal decisions.

Automated findings such as **Found**, **Possible** and **Not found** are text-screening results rather than legal conclusions. AI-generated outputs may be inaccurate, incomplete or affected by missing contractual context.

Regulatory and sanctions-related flags are not live sanctions screening, counterparty verification or regulatory advice.

Before using the application, review the full [Terms of Use / Legal & Use Notice](TERMS_OF_USE.md).

## Confidentiality and AI processing

The deterministic screener can operate without sending the full contract text to the AI service.

If the user explicitly enables full-text AI review, the selected text is transmitted to the configured third-party AI provider. Do not send confidential, privileged, commercially sensitive or personal information unless you are authorised to do so and have assessed the relevant provider terms, privacy obligations and internal policies.

This MVP contains no application database for persistent contract storage; hosting and AI providers may nevertheless process or retain information under their own terms and privacy policies.

## Standard forms and copyrighted material

The app does not reproduce BIMCO or other proprietary standard forms. It screens user-provided text and uses public industry topics to organise review areas.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## AI setup

Create `.streamlit/secrets.toml`:

```toml
OPENROUTER_API_KEY = "your_key_here"
```

Never commit this file to GitHub.

The default model field is:

```text
openrouter/free
```

Free-model availability and rate limits can change.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository named `Charter-Party-Risk-Screener`.
2. Upload `app.py`, `requirements.txt`, `README.md`, `TERMS_OF_USE.md`, `.gitignore`, and the optional demo file.
3. Create a Streamlit Community Cloud app from the repository.
4. Main file path: `app.py`.
5. Add the OpenRouter key under Streamlit **Secrets**.
6. Deploy.

## Recommended GitHub repository description

> AI-assisted charter party risk screening and decision-support tool for voyage, time and bareboat contracts, with structured clause review, Owner/Charterer perspectives and human validation.

## Suggested next versions

- clause-by-clause section extraction
- comparison of two charter-party versions / rider drafts
- configurable company risk policies
- owner / charterer review profiles
- sanctions and regulatory live-data connectors
- claims and operations handover workflow
- amendment / redline change detection
