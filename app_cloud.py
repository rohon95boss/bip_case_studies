import streamlit as st
import os
import json
import pandas as pd
import tempfile
import zipfile
import io
from pptx import Presentation
from openai import OpenAI
from dotenv import load_dotenv

# -----------------------------
# Setup
# -----------------------------
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

TEMPLATE = "templates/BIP_MCG_Case Study_Insert Case Study Name.pptx"

# -----------------------------
# Helpers
# -----------------------------
def extract_text_from_ppt(ppt_file):
    """Extract all text from a PPTX file."""
    prs = Presentation(ppt_file)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_runs.append(shape.text)
    return text_runs


def analyze_case(texts):
    """Call OpenAI to parse case study into structured JSON."""
    joined_text = " ".join(texts)[:8000]
    prompt = f"""
    You are a consultant. Read this case study text and return structured JSON.
    Rules:
    - Never include real client names: replace with "the client".
    - Always refer to delivering company as "BIP".
    - Case Study Name:
      • If the PPT text has a clear subject, use it directly.
      • If not, infer a professional name based on the content.
      • Keep it short: max 3–4 words, no dashes/colons.
    - Category: must be a consulting subcategory (e.g., Regulatory Compliance, Data Migration, Risk & Controls).
    - Function: which office(s) are impacted (e.g., COO Office, PMO Office).
    - Challenge, Solution, Results: 3–4 sentences each, concise.
    - Business Categories: 5 unique items, ≤2 words each.
    - Hashtags: 3 unique buzzwords, 2–3 words each, WITHOUT the # symbol
      (the # is handled in the template separately).

    Case Study Text:
    {joined_text}

    Return JSON with keys:
    case_study_name, category, function, challenge, solution, results, business_categories, hashtags
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"}
    )
    return response.choices[0].message.content


def replace_in_shape(shape, placeholder, value):
    """Replace placeholder text in runs, preserving formatting."""
    if not shape.has_text_frame or not placeholder:
        return False
    replaced = False
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if placeholder in r.text:
                r.text = r.text.replace(placeholder, value)
                replaced = True
    return replaced


def create_case_ppt(analysis_json, folder, orig_name):
    """Generate outputs and return paths."""
    data = json.loads(analysis_json)

    # Case study name cleanup
    name = data.get("case_study_name", "").strip()
    name = name.replace("(", "").replace(")", "")
    data["case_study_name"] = " ".join(name.split()[:6]) or "Case Study"

    # Trim challenge/solution/results
    for field in ["challenge", "solution", "results"]:
        text_val = data.get(field, "")
        sentences = text_val.split(". ")
        if len(sentences) > 4:
            text_val = ". ".join(sentences[:4])
        data[field] = text_val.strip()

    # Hashtags (no # added)
    hashtags = list(dict.fromkeys(data.get("hashtags") or []))[:3]
    cleaned_tags = []
    for tag in hashtags:
        tag = tag.replace("#", "").strip()
        tag = " ".join(tag.split()[:3])
        if tag:
            cleaned_tags.append(tag)
    while len(cleaned_tags) < 3:
        cleaned_tags.append("")
    hashtags = cleaned_tags

    # Business categories
    bcs = data.get("business_categories") or []
    seen = set()
    clean_bcs = []
    for x in bcs:
        item = " ".join(x.split()[:2]).strip()
        if item and item.lower() not in seen:
            seen.add(item.lower())
            clean_bcs.append(item)
    while len(clean_bcs) < 5:
        clean_bcs.append("")
    clean_bcs = clean_bcs[:5]

    # Load template
    prs = Presentation(TEMPLATE)

    # Replace placeholders
    for slide in prs.slides:
        for shape in slide.shapes:
            replace_in_shape(shape, "Insert name of case study", data["case_study_name"])
            replace_in_shape(shape, "Insert Category", data.get("category", ""))
            replace_in_shape(shape, "Insert Function", data.get("function", ""))
            replace_in_shape(shape, "Insert Challenge Here", data.get("challenge", ""))
            replace_in_shape(shape, "Insert Solution Here", data.get("solution", ""))
            replace_in_shape(shape, "Insert Results Here", data.get("results", ""))
            replace_in_shape(shape, "TagOne", hashtags[0])
            replace_in_shape(shape, "TagTwo", hashtags[1])
            replace_in_shape(shape, "TagThree", hashtags[2])
            replace_in_shape(shape, "Business Category 1", clean_bcs[0])
            replace_in_shape(shape, "Business Category 2", clean_bcs[1])
            replace_in_shape(shape, "Business Category 3", clean_bcs[2])
            replace_in_shape(shape, "Business Category 4", clean_bcs[3])
            replace_in_shape(shape, "Business Category 5", clean_bcs[4])

    # Save outputs
    final_ppt = os.path.join(folder, f"BIP_MCG_Case Study_{data['case_study_name']}.pptx")
    prs.save(final_ppt)

    raw_json = os.path.join(folder, "raw.json")
    with open(raw_json, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    parsed_csv = os.path.join(folder, "parsed.csv")
    pd.DataFrame({"content": data.values()}).to_csv(parsed_csv, index=False)

    original_ppt = os.path.join(folder, f"Original - {orig_name}")
    return [final_ppt, raw_json, parsed_csv, original_ppt]


def make_zip(file_paths):
    """Bundle output files into a zip archive."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for file in file_paths:
            if os.path.exists(file):
                zf.write(file, arcname=os.path.basename(file))
    zip_buffer.seek(0)
    return zip_buffer

# -----------------------------
# Streamlit App
# -----------------------------
st.title("📊 BIP Case Study Processor (Cloud Version)")

uploaded_files = st.file_uploader(
    "Upload up to 20 PPT files",
    accept_multiple_files=True,
    type=["pptx"]
)

if uploaded_files:
    for f in uploaded_files[:20]:
        case_name = os.path.splitext(f.name)[0]
        st.write(f"Processing {case_name}...")

        with tempfile.TemporaryDirectory() as tmpdir:
            # Save original
            orig_path = os.path.join(tmpdir, f"Original - {f.name}")
            with open(orig_path, "wb") as orig_f:
                orig_f.write(f.getbuffer())

            # Extract text
            texts = extract_text_from_ppt(f)

            # Analyze
            analysis_json = analyze_case(texts)

            # Generate outputs
            outputs = create_case_ppt(analysis_json, tmpdir, f.name)

            # Bundle into zip
            zip_file = make_zip(outputs)

            # Download button
            st.download_button(
                label=f"📥 Download {case_name} Outputs",
                data=zip_file,
                file_name=f"{case_name}_outputs.zip",
                mime="application/zip"
            )

        st.success(f"✅ {case_name} processed and ready for download!")

