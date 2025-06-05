import argparse, json, os, re, sys
from datetime import datetime, timezone
from uuid import uuid4

try:
    from docx import Document           # pip install python-docx
except ImportError:
    sys.exit("Missing dependency: python-docx (pip install python-docx)")

# --------------------------------------------------------------------------- #
# CONSTANTS                                                                   #
# --------------------------------------------------------------------------- #
LER_RS_CONTEXT = (
    "http://schema.hropenstandards.org/4.5RC/recruiting/json/ler-rs/"
    "ProvisionalVerifiableCredentialResumeType.json"
)
LER_RS_SCHEMA = {"id": LER_RS_CONTEXT, "type": "JsonSchemaValidator2018"}

SECTION_MAP = {
    # common résumé headings → canonical LER-RS property
    "experience": "positionHistory",
    "professional experience": "positionHistory",
    "work experience": "positionHistory",
    "work history": "positionHistory",
    "employment history": "positionHistory",
    "employment": "positionHistory",
    "education": "educationHistory",
    "academic background": "educationHistory",
    "skills": "competency",
    "competencies": "competency",
    "awards": "achievements",
    "honors": "achievements",
}

ROLE_RE = re.compile(
    r"\b(manager|director|engineer|developer|analyst|lead|officer|consultant|specialist)\b",
    re.I,
)

# --------------------------------------------------------------------------- #
# DOCX → SECTIONS                                                             #
# --------------------------------------------------------------------------- #
def looks_like_heading(text: str) -> bool:
    return (
        text.isupper()
        and 2 <= len(text.split()) <= 5
        and sum(c.isalpha() for c in text) >= 4
    )

def parse_docx(path: str):
    doc = Document(path)
    name, current = None, None
    out: dict[str, list[str]] = {}
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        style = (p.style.name if p.style else "").lower()
        if style == "title" or style.startswith("heading 1"):
            if name is None:
                name = txt
            continue
        if style.startswith("heading") and not style.startswith("heading 1") or looks_like_heading(txt):
            current = txt
            out[current] = []
            continue
        if current:
            out[current].append(txt)
    return name, out

# --------------------------------------------------------------------------- #
# HELPERS                                                                     #
# --------------------------------------------------------------------------- #
def canonical(head: str):
    return SECTION_MAP.get(head.lower().strip())

def new_uuid():
    return f"urn:uuid:{uuid4()}"

def to_position(text: str):
    parts = re.split(r"\s[–-]\s", text, 1)
    if len(parts) == 2:
        left, right = parts
        jt, org = (left, right) if ROLE_RE.search(left) else (right, left)
    else:
        jt, org = text, None
    return {
        "id": new_uuid(),
        "type": "Position",
        "jobTitle": jt or None,
        "organization": org or None,
        "description": text or None,
    }

def to_edu(text: str):
    return {"id": new_uuid(), "type": "Education", "description": text or None}

def to_comp(text: str):
    return {"id": new_uuid(), "type": "Competency", "name": text or None}

def to_ach(text: str):
    return {"id": new_uuid(), "type": "Achievement", "description": text or None}

MAKE_OBJ = {
    "positionHistory": to_position,
    "educationHistory": to_edu,
    "competency":       to_comp,
    "achievements":     to_ach,
}

# --------------------------------------------------------------------------- #
# BUILD VC                                                                    #
# --------------------------------------------------------------------------- #
def build_vc(name, sections, args):
    # ------------ personal --------------------------------------------------
    given = args.given_name or (name.split()[0] if name else None)
    family = args.family_name or (name.split()[-1] if name and len(name.split()) > 1 else None)
    personal = {
        "givenName": given,
        "familyName": family,
        "email": args.email or None,
        "phone": args.phone or None,
    }
    # always include keys even if null
    personal = {k: personal.get(k, None) for k in ["givenName", "familyName", "email", "phone"]}

    subject = {
        "id": args.subject_id or new_uuid(),
        "type": "ResumeSubject",
        "personalData": personal,
    }

    # sections
    for head, paras in sections.items():
        key = canonical(head)
        if not key:
            continue
        objs = [MAKE_OBJ[key](txt) for txt in paras if txt]
        if objs:
            subject.setdefault(key, []).extend(objs)

    # guarantee at least ONE section with one object
    if not any(key in subject for key in ("positionHistory", "educationHistory", "competency", "achievements")):
        subject["positionHistory"] = [{
            "id": new_uuid(),
            "type": "Position",
            "jobTitle": None,
            "organization": None,
            "description": None,
        }]

    issuer = {
        "type": "Organization",
        "id":  args.issuer_id  or new_uuid(),
        "name": args.issuer_name or "Self-issued résumé builder",
    }

    vc = {
        "@context": ["https://www.w3.org/2018/credentials/v1", LER_RS_CONTEXT],
        "type": ["VerifiableCredential", "ProvisionalVerifiableCredentialResume"],
        "id": new_uuid(),
        "credentialSchema": LER_RS_SCHEMA,
        "issuer": issuer,
        "issuanceDate": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "credentialSubject": subject,
        "proof": {
            "type": "JsonWebSignature2020",
            "created": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
            "proofPurpose": "assertionMethod",
            "verificationMethod": new_uuid(),
            "jws": "",                 # sign later
        },
    }
    if name:
        vc["name"] = f"{name} – Resume"
    return vc

# --------------------------------------------------------------------------- #
# CLI                                                                         #
# --------------------------------------------------------------------------- #
def choose_docx():
    docs = [f for f in os.listdir() if f.lower().endswith(".docx")]
    if not docs:
        sys.exit("No .docx files found in current directory.")
    print("Select résumé to convert:")
    for i, f in enumerate(docs, 1):
        print(f"{i}. {f}")
    sel = input("Enter number: ").strip()
    if not sel.isdigit() or not (1 <= int(sel) <= len(docs)):
        sys.exit("Invalid choice.")
    return docs[int(sel) - 1]

def default_out(path):
    return f"{os.path.splitext(os.path.basename(path))[0]}_ler_rs.json"

def replace_nulls(obj):
    """Recursively replace None with 'Unknown' for required strings,
       and prune empty dicts/lists."""
    if obj is None:
        return "Unknown"
    if isinstance(obj, list):
        return [replace_nulls(x) for x in obj] or ["Unknown"]
    if isinstance(obj, dict):
        return {k: replace_nulls(v) for k, v in obj.items()}
    return obj

def main():
    ap = argparse.ArgumentParser(
        description="Convert a .docx résumé to an LER-RS Provisional VC"
    )
    ap.add_argument("input",  nargs="?", help="Input .docx résumé")
    ap.add_argument("output", nargs="?", help="Output JSON file (default: <input>_ler_rs.json)")

    # Identifiers
    ap.add_argument("--subject-id")
    ap.add_argument("--issuer-id")
    ap.add_argument("--issuer-name")

    # Personal data overrides
    ap.add_argument("--given-name")
    ap.add_argument("--family-name")
    ap.add_argument("--email")
    ap.add_argument("--phone")
    ap.add_argument("--country")
    ap.add_argument("--region")

    args = ap.parse_args()

    docx_path = args.input or choose_docx()
    base_name = os.path.splitext(docx_path)[0]
    out_path = args.output or f"{base_name}_ler_rs.json"

    name, sections = parse_docx(docx_path)
    vc = build_vc(name, sections, args)

    # Replace nulls and handle proof
    vc = replace_nulls(vc)
    if "proof" in vc and "jws" in vc["proof"]:
        vc["proof"]["jws"] = vc["proof"]["jws"] or "dummy"

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(vc, fh, indent=2, ensure_ascii=False)

    print(f"LER-RS résumé written → {out_path}")

if __name__ == "__main__":
    main()
