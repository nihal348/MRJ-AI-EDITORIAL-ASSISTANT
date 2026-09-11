# ============================================================
# 4. FREE LLM AI ANALYSIS LAYER (Groq API)
# ============================================================
def run_ai_analysis(text, literature):
    api_key = os.getenv("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY")
    if not api_key:
        return {"status": "NOT_RUN", "message": "API Key missing. Please set GROQ_API_KEY in Secrets."}

    # Clean whitespace from the API key to avoid header errors
    api_key = str(api_key).strip().strip('"').strip("'")

    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    prompt = f"""
    Analyze this academic manuscript as an editorial pre-screening assistant.
    Check for:
    1. Structural completeness (Abstract, Methods, Results, Funding, Conflicts).
    2. Internal consistency (Sample sizes, causal statements in observational studies).
    3. Ethics statements.

    MANUSCRIPT EXTRACT:
    {text[:25000]}

    Return JSON strictly in this structure:
    {{
        "overall_status": "GREEN/AMBER/RED",
        "findings": [
            {{"category": "category_name", "severity": "BLOCKING/EDITOR_REVIEW/SUGGESTION", "finding": "description"}}
        ]
    }}
    """
    
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": "You are a journal editor assistant. Always return valid JSON only."},
            {"role": "user", "content": prompt}
        ],
        "response_format": {"type": "json_object"},
        "temperature": 0.1
    }

    try:
        r = requests.post(url, headers=headers, json=payload, timeout=30)
        r.raise_for_status()
        return {"status": "OK", "result": r.json()["choices"][0]["message"]["content"]}
    except requests.exceptions.HTTPError as err:
        return {"status": "ERROR", "message": f"HTTP Error {r.status_code}: {r.text}"}
    except Exception as e:
        return {"status": "ERROR", "message": str(e)}
