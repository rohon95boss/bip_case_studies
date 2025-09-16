import streamlit as st

st.title("🔍 Secrets Debugger")
keys = list(st.secrets.keys())
st.write("Secrets keys detected:", keys)

if "OPENAI_API_KEY" in st.secrets:
    st.success("✅ OPENAI_API_KEY is available!")
else:
    st.error("❌ OPENAI_API_KEY not found.")

if "TEST_KEY" in st.secrets:
    st.info("ℹ️ TEST_KEY also found (good sign).")
else:
    st.warning("⚠️ TEST_KEY not found.")
