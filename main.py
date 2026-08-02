# main.py
# Legal Guide AI - FastAPI Backend
# Run with: uvicorn main:app --reload

import os
import re
import io
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from dotenv import load_dotenv
import google.generativeai as genai
import pdfplumber
from docx import Document as DocxDocument

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

app = FastAPI(title="Legal Guide AI Backend")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SYSTEM_PROMPT = """You are Legal Guide AI, an assistant that helps ordinary
people in India understand their legal rights and procedures in plain English.
Rules:
1. Explain in simple, everyday language (avoid legal jargon).
2. Structure answers with: Relevant Rights, Guidance, Next Steps.
3. Include helpline numbers for urgent situations (e.g. Cyber Crime: 1930, Women: 181).
4. End with a disclaimer that this is not formal legal advice.
"""

def ask_gemini(prompt: str) -> str:
    key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not key:
        raise HTTPException(status_code=500, detail="Google Gemini API key is missing. Set GEMINI_API_KEY in .env")
    
    genai.configure(api_key=key)
    models_to_try = ["gemini-3.6-flash", "gemini-flash-latest", "gemini-3.5-flash", "gemini-2.0-flash"]
    
    last_err = ""
    for m in models_to_try:
        try:
            model = genai.GenerativeModel(m)
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            last_err = str(e)
            if "404" in last_err or "not found" in last_err or "no longer available" in last_err:
                continue
            elif "429" in last_err or "Quota" in last_err:
                raise HTTPException(status_code=429, detail=f"Gemini API Quota Exceeded (429): {last_err}")
            else:
                raise HTTPException(status_code=500, detail=f"Gemini API Error: {last_err}")
                
    raise HTTPException(status_code=500, detail=f"Model Error: {last_err}")


class ChatRequest(BaseModel):
    message: str

class RightsRequest(BaseModel):
    situation: str

class SimplifyRequest(BaseModel):
    legal_text: str

class ScamCheckRequest(BaseModel):
    content: str


@app.get("/")
def home():
    return {"status": "Legal Guide AI Backend is running"}

@app.post("/chat")
def chat(req: ChatRequest):
    prompt = f"{SYSTEM_PROMPT}\n\nUser Question: {req.message}"
    return {"response": ask_gemini(prompt)}

@app.post("/rights-finder")
def rights_finder(req: RightsRequest):
    prompt = f"{SYSTEM_PROMPT}\n\nSituation: {req.situation}"
    return {"response": ask_gemini(prompt)}

@app.post("/simplify")
def simplify(req: SimplifyRequest):
    prompt = f"Simplify into plain English:\n{req.legal_text}"
    return {"simplified": ask_gemini(prompt)}

@app.post("/scam-check")
def scam_check(req: ScamCheckRequest):
    flags = []
    if re.search(r"http[s]?://[^\s]+", req.content):
        flags.append("Contains link")
    if re.search(r"(otp|urgent|account blocked|lottery)", req.content, re.IGNORECASE):
        flags.append("Scam keywords")
        
    prompt = f"Analyze for scam/fraud:\n{req.content}"
    return {"flags": flags, "analysis": ask_gemini(prompt)}

@app.post("/analyze-document")
async def analyze_document(file: UploadFile = File(...)):
    content = await file.read()
    text = ""
    if file.filename.endswith(".pdf"):
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    elif file.filename.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(content))
        text = "\n".join([p.text for p in doc.paragraphs])
        
    if not text.strip():
        raise HTTPException(status_code=400, detail="Could not extract text from document")
        
    prompt = f"Analyze key clauses and risks:\n{text[:6000]}"
    return {"analysis": ask_gemini(prompt)}
