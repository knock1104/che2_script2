# -*- coding: utf-8 -*-
"""
Streamlit 예배 자료 업로드 + Word 저장 + GitHub 임시저장/제출 (+ 성경 JSON 연동)

✅ 성경 본문 로드 (권별 파일)
- 기본: bible_books_json/{book_name}.json
- 폴백: bible_books_json/books/{book_name}.json
포맷:
{
  "1": {"1":"...", "2":"..."},
  "2": {"1":"...", ...}
}

✅ 자료 추가 UI는 "그대로" 유지
- + 자료 추가 → 자료 카드 생성(기존처럼)
- 자료 카드 내부만 "구역(섹션) 추가" + 섹션 안에 성경/이미지/파일 아이템을 넣는 구조로 변경
- 구역 순서 조정 가능(▲▼)

✅ 버튼 텍스트 깨짐(줄바꿈/글자 쪼개짐) 해결
- CSS: button white-space: nowrap
- use_container_width=True
- 버튼 라벨 짧게 구성
"""

# ---------------------------
# 페이지 설정
# ---------------------------
import streamlit as st
st.set_page_config(page_title="설교 자료 업로드", page_icon="🙏", layout="wide")

# ---------------------------
# 표준/서드파티 import
# ---------------------------
import io, os, re, json, uuid, base64, tempfile, requests, hashlib, mimetypes
from copy import deepcopy
from typing import List, Dict, Any, Optional
from datetime import date, datetime, timezone

# python-docx / PIL
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
except Exception:
    st.warning("python-docx가 설치되지 않았습니다. 터미널에서: pip install python-docx")
    Document = None

try:
    from PIL import Image
except Exception:
    st.warning("Pillow가 설치되지 않았습니다. 터미널에서: pip install pillow")
    Image = None

# ---------------------------
# 스타일 (✅ 버튼 글자 줄바꿈 방지 포함)
# ---------------------------
st.markdown(
    """
    <style>
    .small-note { color:#666; font-size:0.9rem; }
    .section-title { font-weight:700; font-size:1.1rem; margin-top:0.5rem; }
    .landing-card {
        padding: 16px; border: 1px solid #e5e7eb; border-radius: 12px; background: #fff;
        box-shadow: 0 4px 12px rgba(0,0,0,0.04);
    }

    /* ✅ 버튼 글자 줄바꿈 방지 + 높이 약간 키움 */
    div.stButton > button {
        white-space: nowrap !important;
        padding: 0.45rem 0.8rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------------------
# 세션 상태 초기화
# ---------------------------
if "materials" not in st.session_state:
    st.session_state.materials: List[Dict[str, Any]] = []
if "preview_idx" not in st.session_state:
    st.session_state.preview_idx = 0
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "position" not in st.session_state:
    st.session_state.position = ""
if "can_edit" not in st.session_state:
    st.session_state.can_edit = False
if "worship_date" not in st.session_state:
    st.session_state.worship_date = date.today()
if "submission_id" not in st.session_state:
    st.session_state.submission_id = None

BASE_SERVICES = ["1부", "2부", "3부", "오후예배"]
if "services_options" not in st.session_state:
    st.session_state.services_options = BASE_SERVICES.copy()
if "services_selected" not in st.session_state:
    st.session_state.services_selected: List[str] = []

# ---------------------------
# 성경 JSON 설정 (권별 파일)
# ---------------------------
# ✅ 너가 말한 경로: bible_books_json/창세기.json
BIBLE_BOOKS_DIR = st.secrets.get("BIBLE_BOOKS_DIR", "bible_books_json")

# 성경 책 목록(표준어)
BOOK_NAMES = [
    # OT
    "창세기","출애굽기","레위기","민수기","신명기",
    "여호수아","사사기","룻기","사무엘상","사무엘하",
    "열왕기상","열왕기하","역대상","역대하","에스라",
    "느헤미야","에스더","욥기","시편","잠언",
    "전도서","아가","이사야","예레미야","예레미야애가",
    "에스겔","다니엘","호세아","요엘","아모스",
    "오바댜","요나","미가","나훔","하박국",
    "스바냐","학개","스가랴","말라기",
    # NT
    "마태복음","마가복음","누가복음","요한복음","사도행전",
    "로마서","고린도전서","고린도후서","갈라디아서","에베소서",
    "빌립보서","골로새서","데살로니가전서","데살로니가후서","디모데전서",
    "디모데후서","디도서","빌레몬서","히브리서","야고보서",
    "베드로전서","베드로후서","요한1서","요한2서","요한3서",
    "유다서","요한계시록"
]

CHAPTER_COUNT = {
    "창세기":50,"출애굽기":40,"레위기":27,"민수기":36,"신명기":34,"여호수아":24,"사사기":21,"룻기":4,"사무엘상":31,"사무엘하":24,
    "열왕기상":22,"열왕기하":25,"역대상":29,"역대하":36,"에스라":10,"느헤미야":13,"에스더":10,"욥기":42,"시편":150,"잠언":31,
    "전도서":12,"아가":8,"이사야":66,"예레미야":52,"예레미야애가":5,"에스겔":48,"다니엘":12,"호세아":14,"요엘":3,"아모스":9,
    "오바댜":1,"요나":4,"미가":7,"나훔":3,"하박국":3,"스바냐":3,"학개":2,"스가랴":14,"말라기":4,
    "마태복음":28,"마가복음":16,"누가복음":24,"요한복음":21,"사도행전":28,"로마서":16,"고린도전서":16,"고린도후서":13,"갈라디아서":6,"에베소서":6,
    "빌립보서":4,"골로새서":4,"데살로니가전서":5,"데살로니가후서":3,"디모데전서":6,"디모데후서":4,"디도서":3,"빌레몬서":1,"히브리서":13,"야고보서":5,
    "베드로전서":5,"베드로후서":3,"요한1서":5,"요한2서":1,"요한3서":1,"유다서":1,"요한계시록":22
}

# ---------------------------
# 랜딩 (권한/접근)
# ---------------------------
def render_landing():
    st.title("Ch2 설교 자료 업로더")
    st.markdown(
        "<div class='landing-card'>"
        "<b>역할을 선택하고 입장하세요.</b><br>"
        "교역자는 작성/수정이 가능하며, 미디어부는 확인만 가능합니다.<br>"
        "<span class='small-note'>[테스트 안내] 현재는 모든 액세스 코드가 <b>0001</b>이면 입장 가능합니다.</span>"
        "</div>",
        unsafe_allow_html=True
    )
    st.write("")
    with st.form("landing_form"):
        role = st.radio("역할 선택", ["교역자", "미디어부"], horizontal=True)
        c1, c2 = st.columns(2)
        with c1:
            user_name = st.text_input("이름")
        with c2:
            position = st.selectbox(
                "직분 선택",
                ['원로목사', "담임목사", "부목사", '강도사', "전도사", "미디어부"],
                index=2
            )
        access_code = st.text_input("개인 액세스 코드", type="password", placeholder="예) 0001")
        submitted = st.form_submit_button("입장")
    if submitted:
        if access_code == "0001":
            st.session_state.authenticated = True
            st.session_state.role = role
            st.session_state.user_name = user_name.strip()
            st.session_state.position = position
            st.session_state.can_edit = (role == "교역자")
            st.success("입장되었습니다.")
            st.rerun()
        else:
            st.error("액세스 코드가 올바르지 않습니다. (테스트: 0001)")

if not st.session_state.authenticated:
    render_landing()
    st.stop()

# ---------------------------
# 상단 사용자/권한 표시
# ---------------------------
can_edit = st.session_state.get("can_edit", False)
role_badge = "🟢 편집 가능" if can_edit else "🔒 읽기 전용(확인만)"
st.markdown(
    f"**접속자:** {st.session_state.user_name or '이름 미입력'} "
    f"({st.session_state.position or '직분 미선택'}) · "
    f"{st.session_state.role} · {role_badge}"
)

# ---------------------------
# GitHub 설정 (없으면 기능 일부 비활성)
# ---------------------------
def _secrets_get(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

GITHUB_OWNER = _secrets_get("GITHUB_OWNER")
GITHUB_REPO  = _secrets_get("GITHUB_REPO")
GITHUB_TOKEN = _secrets_get("GITHUB_TOKEN")
GITHUB_BRANCH = _secrets_get("GITHUB_BRANCH", "main")

GITHUB_ENABLED = bool(GITHUB_OWNER and GITHUB_REPO and GITHUB_TOKEN)

def _gh_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
    }

def _gh_api_base():
    return f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"

def gh_put_bytes(path: str, content_bytes: bytes, message: str):
    if not GITHUB_ENABLED:
        raise RuntimeError("GitHub 설정이 없습니다. secrets.toml에 GITHUB_OWNER/GITHUB_REPO/GITHUB_TOKEN을 넣어주세요.")
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    get = requests.get(url, headers=_gh_headers())
    sha = get.json().get("sha") if get.status_code == 200 else None
    b64 = base64.b64encode(content_bytes).decode("utf-8")
    payload = {"message": message, "content": b64, "branch": GITHUB_BRANCH}
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=_gh_headers(), json=payload)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub 업로드 실패: {r.status_code} {r.text}")
    return r.json()

def gh_get_bytes(path: str) -> bytes:
    if not GITHUB_ENABLED:
        raise RuntimeError("GitHub 설정이 없습니다. secrets.toml에 GITHUB_OWNER/GITHUB_REPO/GITHUB_TOKEN을 넣어주세요.")
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        raise FileNotFoundError(f"GitHub 파일 없음: {path}")
    content = r.json()["content"]
    return base64.b64decode(content)

def gh_list_dir(path: str):
    if not GITHUB_ENABLED:
        return []
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        return []
    return r.json()

# ---------------------------
# 자료 유틸
# ---------------------------
def add_material():
    st.session_state.materials.append({
        "id": str(uuid.uuid4()),
        "kind": "자료 만들기",      # 자료 카드 유형(기존 UI 유지용)
        "description": "",
        # ✅ 섹션 기반
        "sections": [{
            "id": str(uuid.uuid4()),
            "title": "구역 1",
            "storyboard": "",
            "items": []
        }]
    })

def remove_material(mid: str):
    st.session_state.materials = [m for m in st.session_state.materials if m["id"] != mid]

def move_material(mid: str, direction: str):
    mats = st.session_state.materials
    idx = next((i for i, m in enumerate(mats) if m["id"] == mid), None)
    if idx is None:
        return
    if direction == "up" and idx > 0:
        mats[idx-1], mats[idx] = mats[idx], mats[idx-1]
        st.session_state.materials = mats
        st.rerun()
    elif direction == "down" and idx < len(mats)-1:
        mats[idx+1], mats[idx] = mats[idx], mats[idx+1]
        st.session_state.materials = mats
        st.rerun()

def add_rich_text(paragraph, text: str):
    if not text:
        return
    pattern = r'(\*\*.*?\*\*|==.*?==)'
    parts = re.split(pattern, text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("==") and part.endswith("=="):
            run = paragraph.add_run(part[2:-2])
            try:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            except Exception:
                pass
        else:
            paragraph.add_run(part)

# ---------------------------
# ✅ 섹션/아이템 유틸
# ---------------------------
def ensure_material_sections(mat: dict):
    if "sections" not in mat or not isinstance(mat["sections"], list):
        mat["sections"] = []
    if len(mat["sections"]) == 0:
        mat["sections"].append({
            "id": str(uuid.uuid4()),
            "title": "구역 1",
            "storyboard": "",
            "items": []
        })

def add_section(mat: dict):
    ensure_material_sections(mat)
    n = len(mat["sections"]) + 1
    mat["sections"].append({
        "id": str(uuid.uuid4()),
        "title": f"구역 {n}",
        "storyboard": "",
        "items": []
    })

def remove_section(mat: dict, sid: str):
    mat["sections"] = [s for s in mat.get("sections", []) if s["id"] != sid]
    if len(mat["sections"]) == 0:
        ensure_material_sections(mat)

def move_section(mat: dict, sid: str, direction: str):
    secs = mat.get("sections", [])
    idx = next((i for i, s in enumerate(secs) if s["id"] == sid), None)
    if idx is None:
        return
    if direction == "up" and idx > 0:
        secs[idx-1], secs[idx] = secs[idx], secs[idx-1]
    elif direction == "down" and idx < len(secs)-1:
        secs[idx+1], secs[idx] = secs[idx], secs[idx+1]
    mat["sections"] = secs

def add_item_to_section(sec: dict, t: str):
    sec.setdefault("items", [])
    base = {"id": str(uuid.uuid4()), "type": t}
    if t == "성경":
        base.update({
            "book": BOOK_NAMES[0],
            "chap": 1,
            "v_from": 1,
            "v_to": 1,
            "verse_text": ""
        })
    elif t == "이미지":
        base.update({"files": []})
    elif t == "기타파일":
        base.update({"file": None})
    sec["items"].append(base)

def remove_item(sec: dict, iid: str):
    sec["items"] = [x for x in sec.get("items", []) if x["id"] != iid]

# ---------------------------
# 성경 로더 (권별 JSON)
# ---------------------------
@st.cache_data(show_spinner=False, ttl=60 * 30)
def load_chapter_verses_from_repo(book_name: str, chap: int) -> List[Dict[str, Any]]:
    """
    GitHub에서 로드:
      - bible_books_json/{book}.json
      - 폴백: bible_books_json/books/{book}.json
    포맷:
      {"1":{"1":"...","2":"..."}, "2":{...}}
    반환:
      [{"verse":1,"text":"..."}, ...]
    """
    candidates = [
        f"{BIBLE_BOOKS_DIR}/{book_name}.json",
        f"{BIBLE_BOOKS_DIR}/books/{book_name}.json",
    ]

    last_err = None
    raw = None

    for path in candidates:
        try:
            raw = gh_get_bytes(path).decode("utf-8")
            last_err = None
            break
        except Exception as e:
            last_err = e
            continue

    if raw is None:
        raise FileNotFoundError(f"GitHub 파일 없음: {candidates[0]} (폴백도 실패: {candidates[1]})")

    book_data = json.loads(raw)
    chapter_dict = book_data.get(str(int(chap)), {}) or {}

    verses = []
    for verse_k, text in chapter_dict.items():
        try:
            vn = int(str(verse_k))
        except Exception:
            continue
        verses.append({"verse": vn, "text": (text or "").strip()})
    verses.sort(key=lambda x: x["verse"])
    return verses

# ---------------------------
# 성경 위젯 (아이템 단위)
# ---------------------------
def render_bible_item(x: dict, disabled: bool, prefix: str):
    """
    x: 섹션 내부 아이템(type="성경")
    prefix: streamlit key 충돌 방지용 접두사
    """
    c1, c2, c3 = st.columns([1.5, 0.8, 1.4])

    with c1:
        x["book"] = st.selectbox(
            "책",
            options=BOOK_NAMES,
            index=BOOK_NAMES.index(x.get("book", BOOK_NAMES[0])) if x.get("book") in BOOK_NAMES else 0,
            key=f"{prefix}_book",
            disabled=disabled
        )

    max_chap = CHAPTER_COUNT.get(x["book"], 1)
    with c2:
        x["chap"] = st.number_input(
            "장",
            min_value=1,
            max_value=max_chap,
            step=1,
            value=int(x.get("chap", 1)),
            key=f"{prefix}_chap",
            disabled=disabled
        )

    try:
        verses = load_chapter_verses_from_repo(x["book"], int(x["chap"]))
        max_verse = len(verses) if verses else 1
    except Exception as e:
        st.error(f"성경 본문 로드 실패: {e}")
        verses = []
        max_verse = 1

    with c3:
        v1, v2 = st.columns(2)
        with v1:
            x["v_from"] = st.number_input(
                "절(시작)",
                min_value=1,
                max_value=max_verse,
                value=int(x.get("v_from", 1)),
                key=f"{prefix}_vfrom",
                disabled=disabled
            )
        with v2:
            x["v_to"] = st.number_input(
                "절(끝)",
                min_value=int(x["v_from"]),
                max_value=max_verse,
                value=int(x.get("v_to", x["v_from"])),
                key=f"{prefix}_vto",
                disabled=disabled
            )

    preview = ""
    if verses:
        lines = []
        for v in verses:
            if int(x["v_from"]) <= v["verse"] <= int(x["v_to"]):
                lines.append(f"{x['book']} {int(x['chap'])}:{v['verse']} {v['text']}")
        preview = "\n".join(lines)

    st.text_area("미리보기", value=preview, height=150, disabled=True, key=f"{prefix}_preview")

    # ✅ 버튼 텍스트가 깨지지 않게(짧은 라벨 + use_container_width)
    if st.button("📥 추가", key=f"{prefix}_insert", disabled=disabled, use_container_width=True):
        prev = x.get("verse_text", "") or ""
        block = preview.strip()
        if block:
            x["verse_text"] = (prev + ("\n" if prev else "") + block).strip()
            st.session_state[f"{prefix}_text"] = x["verse_text"]
            st.success("본문에 추가했습니다.")
            st.rerun()
        else:
            st.warning("추가할 본문이 없습니다.")

    # 본문 편집 박스(큰 높이)
    text_key = f"{prefix}_text"
    if text_key not in st.session_state:
        st.session_state[text_key] = x.get("verse_text", "")
    x["verse_text"] = st.text_area(
        "본문 내용(편집 가능)",
        key=text_key,
        height=220,
        disabled=disabled
    )

# ---------------------------
# 파일 업로드 보조(메타데이터화)
# ---------------------------
def sanitize_filename(name: str) -> str:
    name = os.path.basename(name or "upload.bin")
    return name.replace("/", "_").replace("\\", "_").strip()

def _sha1(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()

def upload_streamlit_file_to_github(uploaded_file, dest_dir: str, msg_prefix: str = "[file]") -> dict:
    if uploaded_file is None:
        return {}
    data = uploaded_file.getvalue()
    sha1 = _sha1(data)
    orig_name = getattr(uploaded_file, "name", "upload.bin")
    safe_name = sanitize_filename(orig_name)
    dest_path = f"{dest_dir}/{sha1[:10]}_{safe_name}"
    gh_put_bytes(dest_path, data, message=f"{msg_prefix} upload {safe_name}")
    return {
        "name": orig_name,
        "path": dest_path,
        "size": len(data),
        "content_type": getattr(uploaded_file, "type", mimetypes.guess_type(orig_name)[0]),
        "sha1": sha1,
    }

def detach_sections_files_for_github(materials: List[Dict[str, Any]], files_dir: str, msg_prefix: str) -> List[Dict[str, Any]]:
    """
    섹션 내부의 이미지/기타파일 업로드를 dict 메타로 변환해서 저장 가능하게 만듦
    """
    out = []
    for m in materials:
        m2 = deepcopy(m)
        ensure_material_sections(m2)
        for sec in m2["sections"]:
            sec.setdefault("items", [])
            for x in sec["items"]:
                if x.get("type") == "이미지":
                    metas = []
                    for f in x.get("files") or []:
                        if hasattr(f, "getvalue"):
                            metas.append(upload_streamlit_file_to_github(f, files_dir, msg_prefix))
                        elif isinstance(f, dict) and "path" in f:
                            metas.append(f)
                    x["files"] = metas

                elif x.get("type") == "기타파일":
                    f = x.get("file")
                    if hasattr(f, "getvalue"):
                        x["file"] = upload_streamlit_file_to_github(f, files_dir, msg_prefix)
                    elif isinstance(f, dict) and "path" in f:
                        pass
                    else:
                        x["file"] = None
        out.append(m2)
    return out

# ---------------------------
# build_docx (섹션 구조 반영)
# ---------------------------
def build_docx(worship_date: date, services: List[str], materials: List[Dict[str, Any]],
               user_name: str, position: str, role: str) -> bytes:
    if Document is None:
        raise RuntimeError("python-docx가 설치되지 않았습니다. 'pip install python-docx' 실행 후 다시 시도해주세요.")
    doc = Document()

    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style.font.size = Pt(11)

    title = doc.add_paragraph()
    run = title.add_run("설교 자료")
    run.bold = True
    run.font.size = Pt(20)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    meta.add_run(f"날짜: {worship_date.strftime('%Y-%m-%d')}\n").bold = True
    meta.add_run("예배 구분: " + (", ".join(services) if services else "(미선택)") + "\n").bold = True
    if user_name or position or role:
        meta.add_run(f"작성자/권한: {user_name or '(미입력)'} ({position or '직분 미선택'}) - {role or '권한 미지정'}").bold = True

    doc.add_paragraph("")
    doc.add_heading("자료 (스토리보드)", level=1)

    if not materials:
        doc.add_paragraph("(추가된 자료가 없습니다)")
    else:
        for midx, m in enumerate(materials, start=1):
            ensure_material_sections(m)
            doc.add_heading(f"{midx}. {m.get('kind','자료')}", level=2)

            if (m.get("description") or "").strip():
                p = doc.add_paragraph()
                p.add_run("자료 설명: ").bold = True
                add_rich_text(p, m.get("description",""))

            for sidx, sec in enumerate(m["sections"], start=1):
                doc.add_heading(f"구역 {sidx}. {sec.get('title','')}", level=3)

                if (sec.get("storyboard") or "").strip():
                    p = doc.add_paragraph()
                    p.add_run("스토리보드: ").bold = True
                    add_rich_text(p, sec.get("storyboard",""))

                for x in sec.get("items", []):
                    t = x.get("type","")
                    doc.add_paragraph(f"- [{t}]")

                    if t == "성경":
                        vt = (x.get("verse_text") or "").strip()
                        if vt:
                            for line in vt.splitlines():
                                p = doc.add_paragraph()
                                add_rich_text(p, line)
                        else:
                            doc.add_paragraph("(성경 본문 없음)")

                    elif t == "이미지":
                        files = x.get("files") or []
                        if files:
                            for f in files:
                                try:
                                    if isinstance(f, dict) and "path" in f:
                                        img_bytes = gh_get_bytes(f["path"])
                                        _, ext = os.path.splitext(f.get("name") or f["path"])
                                        with tempfile.NamedTemporaryFile(delete=False, suffix=ext or ".img") as tmp:
                                            tmp.write(img_bytes)
                                            tmp.flush()
                                            doc.add_picture(tmp.name, width=Inches(5))
                                    elif hasattr(f, "getvalue"):
                                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(getattr(f, "name", ""))[1]) as tmp:
                                            tmp.write(f.getvalue())
                                            tmp.flush()
                                            doc.add_picture(tmp.name, width=Inches(5))
                                except Exception:
                                    doc.add_paragraph("(이미지 삽입 실패)")
                        else:
                            doc.add_paragraph("(이미지 없음)")

                    elif t == "기타파일":
                        f = x.get("file")
                        if isinstance(f, dict) and "name" in f:
                            doc.add_paragraph(f"첨부 파일: {f['name']} (문서에 직접 삽입되지 않습니다)")
                        elif f is not None and hasattr(f, "getvalue"):
                            doc.add_paragraph(f"첨부 파일: {getattr(f,'name','파일')} (문서에 직접 삽입되지 않습니다)")
                        else:
                            doc.add_paragraph("(첨부 파일 없음)")

                doc.add_paragraph("")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()

# ---------------------------
# 제출 직렬화/역직렬화 + 경로
# ---------------------------
def serialize_submission():
    return {
        "worship_date": str(st.session_state.get("worship_date")),
        "services": st.session_state.get("services_selected", []),
        "materials": st.session_state.get("materials", []),
        "user_name": st.session_state.get("user_name"),
        "position": st.session_state.get("position"),
        "role": st.session_state.get("role"),
        "saved_at": datetime.now(timezone.utc).isoformat()
    }

def load_into_session(payload: dict):
    st.session_state.worship_date = date.fromisoformat(payload.get("worship_date"))
    st.session_state.services_selected = payload.get("services", [])
    st.session_state.materials = payload.get("materials", [])

def gh_paths(user_name: str, worship_date: date, submission_id: str = None):
    base = _secrets_get("GITHUB_BASE_DIR", "worship_submissions")
    d = worship_date.strftime("%Y-%m-%d")
    safe_user = (user_name or "unknown").strip().replace("/", "_")
    sub_id = "draft" if submission_id is None else submission_id
    folder = f"{base}/{d}/{safe_user}/{sub_id}"
    return {
        "folder": folder,
        "files_dir": f"{folder}/files",
        "json": f"{folder}/submission.json",
        "docx": f"{folder}/submission.docx",
    }

# ---------------------------
# ① 날짜/예배 선택
# ---------------------------
st.markdown("<div class='section-title'>① 날짜/예배 선택</div>", unsafe_allow_html=True)
worship_date = st.date_input("예배 날짜", value=st.session_state.worship_date, format="YYYY-MM-DD", disabled=not can_edit)
st.session_state.worship_date = worship_date

c1, c2 = st.columns([2, 1])
with c1:
    st.session_state.services_selected = st.multiselect(
        "예배 구분 선택",
        options=st.session_state.services_options,
        default=st.session_state.services_selected,
        help="해당 날짜에 해당되는 예배를 모두 선택하세요.",
        disabled=not can_edit
    )
with c2:
    new_service = st.text_input("직접 입력", placeholder="예: 청년예배 / 새벽기도", disabled=not can_edit)
    add_new = st.button("추가", disabled=not can_edit)
    if add_new and new_service.strip():
        if new_service not in st.session_state.services_options:
            st.session_state.services_options.append(new_service.strip())
        if new_service not in st.session_state.services_selected:
            st.session_state.services_selected.append(new_service.strip())
        st.rerun()

services = st.session_state.services_selected
st.divider()

# ---------------------------
# ② 자료 추가 (UI는 그대로 유지)
# ---------------------------
st.markdown("<div class='section-title'>② 자료 추가 (성경/이미지/기타/설교 전문)</div>", unsafe_allow_html=True)
st.caption("• 설명(스토리보드)에서 **굵게**, ==형광펜== 으로 강조하면 Word에 그대로 반영됩니다.")

add_btn = st.button("+ 자료 추가", disabled=not can_edit)
if add_btn and can_edit:
    add_material()

to_remove: List[str] = []
for i, item in enumerate(st.session_state.materials):
    with st.container(border=True):
        top_cols = st.columns([1.2, 0.2, 0.2, 0.2])
        with top_cols[0]:
            item["kind"] = st.selectbox(
                "자료 유형",
                ["자료 만들기", "설교 전문"],
                index=["자료 만들기", "설교 전문"].index(item.get("kind", "자료 만들기")),
                key=f"kind_{item['id']}",
                disabled=not can_edit
            )
        with top_cols[1]:
            st.write("")
            st.button("▲", key=f"up_{item['id']}", disabled=(not can_edit or i == 0),
                      on_click=move_material, args=(item["id"], "up"))
        with top_cols[2]:
            st.write("")
            st.button("▼", key=f"down_{item['id']}", disabled=(not can_edit or i == len(st.session_state.materials)-1),
                      on_click=move_material, args=(item["id"], "down"))
        with top_cols[3]:
            st.write("")
            if st.button("삭제", key=f"del_{item['id']}", disabled=not can_edit):
                to_remove.append(item["id"])

        # ✅ 내부만 섹션 기반으로 변경
        ensure_material_sections(item)

        if item["kind"] == "자료 만들기":
            item["description"] = st.text_area(
                "자료 설명(전체)",
                value=item.get("description", ""),
                key=f"mat_desc_{item['id']}",
                height=90,
                disabled=not can_edit
            )

            st.divider()

            if st.button("➕ 구역 추가", key=f"add_section_{item['id']}", disabled=not can_edit, use_container_width=True):
                add_section(item)
                st.rerun()

            for si, sec in enumerate(item["sections"]):
                with st.container(border=True):
                    head = st.columns([1.6, 0.2, 0.2, 0.2])
                    with head[0]:
                        sec["title"] = st.text_input(
                            "구역 제목",
                            value=sec.get("title", ""),
                            key=f"sec_title_{item['id']}_{sec['id']}",
                            disabled=not can_edit
                        )
                    with head[1]:
                        st.write("")
                        st.button("▲", key=f"sec_up_{item['id']}_{sec['id']}",
                                  disabled=(not can_edit or si == 0),
                                  on_click=move_section, args=(item, sec["id"], "up"))
                    with head[2]:
                        st.write("")
                        st.button("▼", key=f"sec_dn_{item['id']}_{sec['id']}",
                                  disabled=(not can_edit or si == len(item["sections"]) - 1),
                                  on_click=move_section, args=(item, sec["id"], "down"))
                    with head[3]:
                        st.write("")
                        if st.button("삭제", key=f"sec_del_{item['id']}_{sec['id']}", disabled=not can_edit):
                            remove_section(item, sec["id"])
                            st.rerun()

                    sec["storyboard"] = st.text_area(
                        "구역 스토리보드/설명",
                        value=sec.get("storyboard", ""),
                        key=f"sec_story_{item['id']}_{sec['id']}",
                        height=140,
                        placeholder="이 구역에서 무엇을 보여줄지 / 타이밍 / 강조 포인트 등을 적으세요.",
                        disabled=not can_edit
                    )

                    # ✅ 버튼 폭/줄바꿈 깨짐 해결: use_container_width + 짧은 라벨 + columns 비율 확대
                    add_cols = st.columns([1, 1, 1], gap="small")
                    with add_cols[0]:
                        if st.button("📖 성경", key=f"add_bible_{item['id']}_{sec['id']}",
                                     disabled=not can_edit, use_container_width=True):
                            add_item_to_section(sec, "성경")
                            st.rerun()
                    with add_cols[1]:
                        if st.button("🖼️ 이미지", key=f"add_img_{item['id']}_{sec['id']}",
                                     disabled=not can_edit, use_container_width=True):
                            add_item_to_section(sec, "이미지")
                            st.rerun()
                    with add_cols[2]:
                        if st.button("📎 파일", key=f"add_file_{item['id']}_{sec['id']}",
                                     disabled=not can_edit, use_container_width=True):
                            add_item_to_section(sec, "기타파일")
                            st.rerun()

                    for xi, x in enumerate(sec.get("items", [])):
                        with st.container(border=True):
                            h = st.columns([1.4, 0.2])
                            with h[0]:
                                st.markdown(f"**아이템 {xi+1} · {x['type']}**")
                            with h[1]:
                                if st.button("삭제", key=f"del_item_{item['id']}_{sec['id']}_{x['id']}", disabled=not can_edit):
                                    remove_item(sec, x["id"])
                                    st.rerun()

                            if x["type"] == "성경":
                                render_bible_item(
                                    x,
                                    disabled=not can_edit,
                                    prefix=f"b_{item['id']}_{sec['id']}_{x['id']}"
                                )

                            elif x["type"] == "이미지":
                                up_key = f"img_{item['id']}_{sec['id']}_{x['id']}"
                                uploads = st.file_uploader(
                                    "이미지 업로드 (여러 장 가능)",
                                    type=["png", "jpg", "jpeg"],
                                    accept_multiple_files=True,
                                    key=up_key,
                                    disabled=not can_edit
                                )
                                existing = x.get("files") or []
                                if uploads and len(uploads) > 0:
                                    x["files"] = existing + uploads
                                else:
                                    x["files"] = existing

                                if x["files"]:
                                    names = []
                                    for f in x["files"]:
                                        names.append(f.get("name") if isinstance(f, dict) else getattr(f, "name", "(file)"))
                                    st.caption("업로드됨: " + ", ".join([n for n in names if n]))

                            elif x["type"] == "기타파일":
                                file_key = f"file_{item['id']}_{sec['id']}_{x['id']}"
                                up = st.file_uploader(
                                    "파일 업로드 (PDF/DOCX/HWP 등)",
                                    type=None,
                                    accept_multiple_files=False,
                                    key=file_key,
                                    disabled=not can_edit
                                )
                                if up is not None:
                                    x["file"] = up
                                else:
                                    x["file"] = x.get("file", None)

                                if x.get("file"):
                                    f = x["file"]
                                    nm = f.get("name") if isinstance(f, dict) else getattr(f, "name", "(file)")
                                    st.caption(f"첨부: {nm}")

        elif item["kind"] == "설교 전문":
            # 설교 전문은 기존처럼 텍스트 중심 + 구역형으로 확장하고 싶으면 섹션 구조 재사용 가능
            # 일단 단일 텍스트 유지(요청 없어서 보수적으로)
            item.setdefault("full_text", "")
            item["full_text"] = st.text_area(
                "설교 전문 입력 (줄바꿈 유지 / **굵게**, ==형광펜== 지원)",
                value=item.get("full_text", ""),
                key=f"full_{item['id']}",
                height=340,
                disabled=not can_edit
            )

            item["description"] = st.text_area(
                "설명(스토리보드)",
                value=item.get("description", ""),
                key=f"desc_{item['id']}",
                height=120,
                placeholder="노출 타이밍, 강조 부분 등. **굵게**, ==형광펜== 으로 강조 가능합니다.",
                disabled=not can_edit
            )

if to_remove and can_edit:
    for rid in to_remove:
        remove_material(rid)

st.divider()

# ---------------------------
# ③ Word 파일 생성(로컬 다운로드)
# ---------------------------
st.markdown("<div class='section-title'>③ Word 저장 (로컬 미리 받기)</div>", unsafe_allow_html=True)
col1, col2 = st.columns([1, 2])
with col1:
    do_save = st.button("📄 업로드 하기 (Word 저장)", type="primary", disabled=not can_edit)

if do_save and can_edit:
    try:
        docx_bytes = build_docx(
            worship_date=worship_date,
            services=services,
            materials=st.session_state.materials,
            user_name=st.session_state.user_name,
            position=st.session_state.position,
            role=st.session_state.role
        )
        filename = f"설교자료_{worship_date.strftime('%Y%m%d')}_{'-'.join(services) if services else '미지정'}.docx"
        st.success("Word 파일이 생성되었습니다. 아래 버튼으로 다운로드하세요.")
        st.download_button(
            "⬇️ Word 파일 다운로드",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"문서 생성 중 오류가 발생했습니다: {e}")

st.divider()

# ---------------------------
# ④ 저장/불러오기/제출 (GitHub)
# ---------------------------
st.markdown("#### 저장/제출")
if not GITHUB_ENABLED:
    st.info("GitHub 저장/제출 기능을 사용하려면 secrets.toml에 GITHUB_OWNER/GITHUB_REPO/GITHUB_TOKEN을 설정해야 합니다.")

b1, b2, b3, _ = st.columns([1, 1, 1, 3])
with b1:
    save_draft = st.button("💾 임시 저장", disabled=(not can_edit or not GITHUB_ENABLED))
with b2:
    load_draft = st.button("↩️ 불러오기", disabled=(not can_edit or not GITHUB_ENABLED))
with b3:
    submit_now = st.button("✅ 제출", disabled=(not can_edit or not GITHUB_ENABLED))

if save_draft and can_edit and GITHUB_ENABLED:
    try:
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        mats_detached = detach_sections_files_for_github(
            st.session_state.materials, p["files_dir"], msg_prefix="[draft-files]"
        )
        data = serialize_submission()
        data["materials"] = mats_detached
        gh_put_bytes(
            p["json"],
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            message=f"[draft] {st.session_state.user_name} {worship_date} 저장"
        )
        st.success("임시 저장되었습니다. (GitHub)")
    except Exception as e:
        st.error(f"임시 저장 실패: {e}")

if load_draft and can_edit and GITHUB_ENABLED:
    try:
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        draft_bytes = gh_get_bytes(p["json"])
        payload = json.loads(draft_bytes.decode("utf-8"))
        load_into_session(payload)
        st.success("임시 저장본을 불러왔습니다.")
        st.rerun()
    except Exception as e:
        st.error(f"불러오기 실패 또는 저장본 없음: {e}")

if submit_now and can_edit and GITHUB_ENABLED:
    try:
        sub_id = st.session_state.submission_id or datetime.now().strftime("%H%M%S") + "-" + uuid.uuid4().hex[:6]
        st.session_state.submission_id = sub_id
        p = gh_paths(st.session_state.user_name, worship_date, submission_id=sub_id)

        mats_detached = detach_sections_files_for_github(
            st.session_state.materials, p["files_dir"], msg_prefix="[submit-files]"
        )

        docx_bytes = build_docx(
            worship_date=worship_date,
            services=st.session_state.services_selected,
            materials=st.session_state.materials,
            user_name=st.session_state.user_name,
            position=st.session_state.position,
            role=st.session_state.role
        )

        data = serialize_submission()
        data["status"] = "submitted"
        data["submission_id"] = sub_id
        data["materials"] = mats_detached

        gh_put_bytes(
            p["json"],
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            message=f"[submit] {st.session_state.user_name} {worship_date} 제출"
        )
        gh_put_bytes(
            p["docx"],
            docx_bytes,
            message=f"[submit-docx] {st.session_state.user_name} {worship_date} DOCX"
        )
        st.success("제출 완료! 미디어부 화면에서 확인 가능합니다.")
    except Exception as e:
        st.error(f"제출 실패: {e}")

st.divider()

# ---------------------------
# ⑤ 미디어부 제출함 (검토/다운로드)
# ---------------------------
if st.session_state.role == "미디어부":
    st.markdown("### 📬 제출함(미디어부) — 날짜별/제출자별 목록")
    if not GITHUB_ENABLED:
        st.info("GitHub 설정이 없어 제출함을 불러올 수 없습니다.")
    else:
        base = _secrets_get("GITHUB_BASE_DIR", "worship_submissions")
        days = gh_list_dir(base)
        if not days:
            st.info("아직 제출된 자료가 없습니다.")
        else:
            day_names = sorted([d["name"] for d in days if d.get("type") == "dir"], reverse=True)
            sel_day = st.selectbox("날짜 선택", options=day_names)
            if sel_day:
                day_dir = f"{base}/{sel_day}"
                users = gh_list_dir(day_dir) or []
                for u in users:
                    if u.get("type") != "dir":
                        continue
                    with st.expander(f"👤 {u['name']} — {sel_day} 제출물들"):
                        subs = gh_list_dir(u["path"]) or []
                        for s in subs:
                            if s.get("type") != "dir":
                                continue
                            files = gh_list_dir(s["path"]) or []
                            json_item = next((f for f in files if f.get("name") == "submission.json"), None)
                            docx_item = next((f for f in files if f.get("name") == "submission.docx"), None)

                            c1, c2, c3 = st.columns([2, 1, 2])
                            with c1:
                                st.markdown(f"**제출 ID:** {s['name']}")
                            with c2:
                                if json_item:
                                    try:
                                        payload = json.loads(gh_get_bytes(json_item["path"]).decode("utf-8"))
                                        info = (
                                            f"- 예배: {', '.join(payload.get('services', [])) or '(미지정)'}\n"
                                            f"- 자료개수: {len(payload.get('materials', []))}\n"
                                            f"- 제출시각(UTC): {payload.get('saved_at','')}\n"
                                        )
                                        st.caption(info)
                                    except Exception:
                                        st.caption("메타 로드 실패")
                                else:
                                    st.caption("메타 없음")
                            with c3:
                                if docx_item:
                                    try:
                                        docx_bytes = gh_get_bytes(docx_item["path"])
                                        st.download_button(
                                            "⬇️ Word 다운로드",
                                            data=docx_bytes,
                                            file_name=f"설교자료_{sel_day}_{u['name']}_{s['name']}.docx",
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                            key=f"dl_{sel_day}_{u['name']}_{s['name']}"
                                        )
                                    except Exception as e:
                                        st.error(f"다운로드 오류: {e}")
                                else:
                                    st.caption("DOCX 없음")

# ---------------------------
# 풋터
# ---------------------------
st.markdown(
    """
    <hr/>
    <div class='small-note'>
    ⚙️ 이미지 외의 기타 파일은 Word에 직접 삽입되지 않으며, 파일명과 설명이 기록됩니다.<br>
    ✍️ 강조법: **굵게**, ==형광펜== (Word 변환 시 자동 적용)<br>
    🔗 성경 본문은 bible_books_json/{책이름}.json(또는 books 폴더)에서 로드됩니다.
    </div>
    """,
    unsafe_allow_html=True
)
