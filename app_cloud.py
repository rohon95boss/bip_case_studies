import streamlit as st
from openai import OpenAI
import os
import json
import csv
from pptx import Presentation
from pathlib import Path

# ==============================================================
# 🔑 API Key Setup (Secrets Only: local + cloud)
# ==============================================================

if "OPENAI_API_KEY" not in st.secrets:
    st.error("❌ No API key found. Please add it to .streamlit/secrets.toml (local) "
             "or in Streamlit Cloud Settings → Secrets.")
    st.stop()

api_key = st.secrets["OPENAI_API_KEY"]
client = OpenAI(api_key=api_key)

# ==============================================================
# 📂 Helper Functions
# ==============================================================

def extract_text_from_ppt(ppt_file):
    """Extract all text from a PPTX file and return as string."""
    prs = Presentation(ppt_file)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_runs.append(shape.text)
    return "\n".join(text_runs)


def save_raw_and_csv(text, output_dir, base_name):
    """Save extracted text into raw.json and parsed.csv."""
    raw_path = output_dir / "raw.json"
    with open(raw_path, "w") as f:
        json.dump({"text": text}, f, indent=2)

    csv_path = output_dir / "parsed.csv"
    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["line_number", "content"])
        for i, line in enumerate(text.splitlines(), 1):
            writer.writerow([i, line])

    return raw_path, csv_path


def generate_case_study_metadata(text):
    """Call OpenAI to extract structured fields from text."""
    prompt = f"""
    You are a consulting analyst. Analyze the following case study text:

    {text}

    Extract the following fields in JSON:
    - case_study_name: short clear title
    - category: a consulting subcategory (e.g., Data Migration, Regulatory Compliance, Program Management)
    - function: the office/function impacted (e.g., COO Office, Compliance Office, PMO Office)
    - challenge: 3–4 sentences describing the challenge
    - solution: 3–4 sentences describing the solution (use BIP as the firm, anonymize the client)
    - results: 3–4 sentences describing the results
    - business_functions: list of up to 5 short (max 2 words) categories
    - tags: 3 buzzword-like tags (no #, just plain words, max 3 words each)
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4,
    )

    try:
        content = response.choices[0].message.content
        data = json.loads(content)
    except Exception:
        data = {
            "case_study_name": "Unknown Case Study",
            "category": "Uncategorized",
            "function": "General Office",
            "challenge": "Challenge not available.",
            "solution": "Solution not available.",
            "results": "Results not available.",
            "business_functions": ["Consulting"],
            "tags": ["General", "Business", "CaseStudy"],
        }
    return data


def populate_template(template_path, output_path, metadata):
    """Fill in the PPTX template with extracted metadata."""
    prs = Presentation(template_path)
    slide = prs.slides[0]  # assuming everything is on first slide

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text.strip()

        if "Insert name of case study" in text:
            shape.text = f"Case Study – {metadata['case_study_name']}"

        elif "(Insert Category)" in text:
            shape.text = metadata["category"]

        elif "(Insert Function)" in text:
            shape.text = metadata["function"]

        elif "(Insert Challenge Here)" in text:
            shape.text = metadata["challenge"]

        elif "(Insert Solution Here)" in text:
            shape.text = metadata["solution"]

        elif "(Insert Results Here)" in text:
            shape.text = metadata["results"]

        elif "tagone" in text.lower():
            shape.text = metadata["tags"][0]

        elif "tagtwo" in text.lower():
            shape.text = metadata["tags"][1]

        elif "tagthree" in text.lower():
            shape.text = metadata["tags"][2]

        elif "BusinessFunction1" in text:
            for i, func in enumerate(metadata["business_functions"], 1):
                if f"BusinessFunction{i}" in text:
                    shape.text = func

    prs.save(output_path)


# ==============================================================
# 🎨 Streamlit UI
# ==============================================================

st.title("📊 BIP Case Study Automation (Cloud Deployment)")

uploaded_files = st.file_uploader(
    "Upload up to 20 PPTX files",
    type=["pptx"],
    accept_multiple_files=True
)

TEMPLATE_PATH = "BIP_MCG_Case Study_Insert Case Study Name.pptx"

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.write(f"Processing: {uploaded_file.name}")

        # Create output folder
        base_name = Path(uploaded_file.name).stem
        output_dir = Path("data") / base_name
        output_dir.mkdir(parents=True, exist_ok=True)

        # Save original PPT
        original_path = output_dir / f"Original - {uploaded_file.name}"
        with open(original_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Extract raw text
        text = extract_text_from_ppt(original_path)

        # Save raw.json + parsed.csv
        raw_path, csv_path = save_raw_and_csv(text, output_dir, base_name)

        # Generate metadata via OpenAI
        metadata = generate_case_study_metadata(text)

        # Save new PPT
        new_ppt_path = output_dir / f"BIP_MCG_Case Study_{metadata['case_study_name']}.pptx"
        populate_template(TEMPLATE_PATH, new_ppt_path, metadata)

        st.success(f"✅ Case study saved: {new_ppt_path}")

