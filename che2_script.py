# -*- coding: utf-8 -*-
"""
Streamlit 예배 자료 업로드 + Word 저장 + GitHub 임시저장/제출 (+ 성경 JSON 연동)

✅ 이번 버전 변경점
- "자료 추가" 방식 → "구역(Section)" 중심 구조
- 구역을 추가/삭제/위·아래 이동(순서 조정) 가능
- 각 구역 안에 "성경/이미지/기타 파일" 자료를 여러 개 추가 가능
- 구역마다 "스토리보드/설명"을 작성 가능 (전문 올리기처럼 구역 단위 작성)

✅ 성경 로드
- 기본: bible_books_json/{book_name}.json
- 폴백: bible_books_json/books/{book_name}.json
- 책 파일 포맷(권별):
  {"1":{"1":"...","2":"..."}, "2":{...}}   (장→절→본문)

✅ GitHub 업로드/저장
- 임시저장/제출 시 파일(이미지/첨부)을 GitHub에 업로드하고 메타데이터로 치환
- 제출 시 DOCX도 함께 업로드

⚠️ secrets.toml 필요 (Streamlit Cloud 또는 .streamlit/secrets.toml)
GITHUB_TOKEN="..."
GITHUB_OWNER="knock1104"
GITHUB_REPO="che2_script2"
GITHUB_BRANCH="main"                # optional
GITHUB_BASE_DIR="worship_submissions" # optional
BIBLE_BOOKS_DIR="bible_books_json"   # optional: 기본값
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
# 스타일
# ---------------------------
st.markdown(
    """
    <style>
    .small-note { color:#666; font-size:0.9rem; }
    .section-title { font-weight:800; font-size:1.12rem; margin-top:0.5rem; }
    .landing-card {
        padding: 16px; border: 1px solid #e5e7eb; border-radius: 12px; background: #fff;
        box-shadow: 0 4px 12px rgba(0,0,0,0.04);
    }
    .chip { display:inline-block; padding:2px 8px; border:1px solid #e5e7eb; border-radius:999px; font-size:0.85rem; color:#444; background:#fafafa; }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------------------
# 세션 상태 초기화
# ---------------------------
if "sections" not in st.session_state:
    st.session_state.sections: List[Dict[str, Any]] = []
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
# 성경 JSON 설정
# ---------------------------
BIBLE_BOOKS_DIR = st.secrets.get("BIBLE_BOOKS_DIR", "bible_books_json")

CHAPTER_COUNT = {
    "창세기":50,"출애굽기":40,"레위기":27,"민수기":36,"신명기":34,"여호수아":24,"사사기":21,"룻기":4,"사무엘상":31,"사무엘하":24,
    "열왕기상":22,"열왕기하":25,"역대상":29,"역대하":36,"에스라":10,"느헤미야":13,"에스더":10,"욥기":42,"시편":150,"잠언":31,
    "전도서":12,"아가":8,"이사야":66,"예레미야":52,"예레미야애가":5,"에스겔":48,"다니엘":12,"호세아":14,"요엘":3,"아모스":9,
    "오바댜":1,"요나":4,"미가":7,"나훔":3,"하박국":3,"스바냐":3,"학개":2,"스가랴":14,"말라기":4,
    "마태복음":28,"마가복음":16,"누가복음":24,"요한복음":21,"사도행전":28,"로마서":16,"고린도전서":16,"고린도후서":13,"갈라디아서":6,"에베소서":6,
    "빌립보서":4,"골로새서":4,"데살로니가전서":5,"데살로니가후서":3,"디모데전서":6,"디모데후서":4,"디도서":3,"빌레몬서":1,"히브리서":13,"야고보서":5,
    "베드로전서":5,"베드로후서":3,"요한1서":5,"요한2서":1,"요한3서":1,"유다서":1,"요한계시록":22
}
BOOK_NAMES = list(CHAPTER_COUNT.keys())

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
# GitHub 유틸
# ---------------------------
def _gh_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github+json",
    }

def _gh_api_base():
    owner = st.secrets["GITHUB_OWNER"]
    repo = st.secrets["GITHUB_REPO"]
    return f"https://api.github.com/repos/{owner}/{repo}"

def gh_put_bytes(path: str, content_bytes: bytes, message: str):
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    get = requests.get(url, headers=_gh_headers())
    sha = get.json().get("sha") if get.status_code == 200 else None
    b64 = base64.b64encode(content_bytes).decode("utf-8")
    payload = {
        "message": message,
        "content": b64,
        "branch": st.secrets.get("GITHUB_BRANCH", "main"),
    }
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=_gh_headers(), json=payload)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub 업로드 실패: {r.status_code} {r.text}")
    return r.json()

def gh_get_bytes(path: str) -> bytes:
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        raise FileNotFoundError(f"GitHub 파일 없음: {path}")
    content = r.json()["content"]
    return base64.b64decode(content)

def gh_list_dir(path: str):
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        return []
    return r.json()

# ---------------------------
# 유틸: 텍스트 강조(**, == ==)
# ---------------------------
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
# 섹션/아이템 데이터 구조
# ---------------------------
def _new_bible_item() -> Dict[str, Any]:
    return {
        "id": str(uuid.uuid4()),
        "type": "성경",
        "verse_text": "",
        # picker state (저장용)
        "book": BOOK_NAMES[0],
        "chap": 1,
        "v_from": 1,
        "v_to": 1,
    }

def _new_image_item() -> Dict[str, Any]:
    return {
        "id": str(uuid.uuid4()),
        "type": "이미지",
        "files": [],  # UploadedFile or metadata dict
    }

def _new_file_item() -> Dict[str, Any]:
    return {
        "id": str(uuid.uuid4()),
        "type": "기타파일",
        "file": None,  # UploadedFile or metadata dict
    }

def add_section():
    st.session_state.sections.append({
        "id": str(uuid.uuid4()),
        "title": "",
        "storyboard": "",
        "items": []
    })

def remove_section(sec_id: str):
    st.session_state.sections = [s for s in st.session_state.sections if s["id"] != sec_id]

def move_section(sec_id: str, direction: str):
    secs = st.session_state.sections
    idx = next((i for i, s in enumerate(secs) if s["id"] == sec_id), None)
    if idx is None:
        return
    if direction == "up" and idx > 0:
        secs[idx-1], secs[idx] = secs[idx], secs[idx-1]
        st.session_state.sections = secs
        st.rerun()
    if direction == "down" and idx < len(secs) - 1:
        secs[idx+1], secs[idx] = secs[idx], secs[idx+1]
        st.session_state.sections = secs
        st.rerun()

def add_item(sec: Dict[str, Any], item_type: str):
    if item_type == "성경":
        sec["items"].append(_new_bible_item())
    elif item_type == "이미지":
        sec["items"].append(_new_image_item())
    elif item_type == "기타파일":
        sec["items"].append(_new_file_item())

def remove_item(sec: Dict[str, Any], item_id: str):
    sec["items"] = [it for it in sec.get("items", []) if it["id"] != item_id]

def move_item(sec: Dict[str, Any], item_id: str, direction: str):
    items = sec.get("items", [])
    idx = next((i for i, it in enumerate(items) if it["id"] == item_id), None)
    if idx is None:
        return
    if direction == "up" and idx > 0:
        items[idx-1], items[idx] = items[idx], items[idx-1]
        sec["items"] = items
        st.rerun()
    if direction == "down" and idx < len(items) - 1:
        items[idx+1], items[idx] = items[idx], items[idx+1]
        sec["items"] = items
        st.rerun()

# ---------------------------
# 성경 로더 (권별 파일)
# ---------------------------
@st.cache_data(show_spinner=False, ttl=60*30)
def load_chapter_verses_from_github(book_name: str, chap: int) -> List[Dict[str, Any]]:
    """
    1) {BIBLE_BOOKS_DIR}/{book_name}.json
    2) 폴백: {BIBLE_BOOKS_DIR}/books/{book_name}.json
    포맷: {"1":{"1":"...","2":"..."}, "2":{...}}
    반환: [{"verse":1,"text":"..."}, ...]
    """
    last_err = None
    for path in [
        f"{BIBLE_BOOKS_DIR}/{book_name}.json",
        f"{BIBLE_BOOKS_DIR}/books/{book_name}.json",
    ]:
        try:
            raw = gh_get_bytes(path).decode("utf-8")
            book_data = json.loads(raw)
            chapter_dict = book_data.get(str(chap), {}) or {}
            verses = []
            for verse_k, text in chapter_dict.items():
                try:
                    vn = int(verse_k)
                except Exception:
                    continue
                verses.append({"verse": vn, "text": (text or "").strip()})
            verses.sort(key=lambda x: x["verse"])
            return verses
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"GitHub 파일 없음: {BIBLE_BOOKS_DIR}/{book_name}.json (폴백도 실패: {last_err})")

def render_bible_item(sec: Dict[str, Any], item: Dict[str, Any], disabled: bool):
    st.markdown("**📖 성경 구절**")
    c1, c2, c3 = st.columns([1.4, 0.8, 1.2])

    with c1:
        book_name = st.selectbox(
            "책",
            options=BOOK_NAMES,
            index=BOOK_NAMES.index(item.get("book", BOOK_NAMES[0])) if item.get("book") in BOOK_NAMES else 0,
            key=f"b_book_{sec['id']}_{item['id']}",
            disabled=disabled,
        )
    item["book"] = book_name
    max_chap = CHAPTER_COUNT[book_name]

    with c2:
        chap = st.number_input(
            "장", min_value=1, max_value=max_chap, step=1,
            value=int(item.get("chap", 1)),
            key=f"b_chap_{sec['id']}_{item['id']}",
            disabled=disabled
        )
    item["chap"] = int(chap)

    try:
        verses = load_chapter_verses_from_github(book_name, int(chap))
        max_verse = len(verses) or 1
    except Exception as e:
        st.error(f"성경 본문 로드 실패: {e}")
        verses = []
        max_verse = 1

    with c3:
        v1, v2 = st.columns(2)
        with v1:
            v_from = st.number_input(
                "절(시작)", min_value=1, max_value=max_verse,
                value=int(item.get("v_from", 1)),
                key=f"b_vfrom_{sec['id']}_{item['id']}",
                disabled=disabled
            )
        with v2:
            v_to = st.number_input(
                "절(끝)", min_value=int(v_from), max_value=max_verse,
                value=int(item.get("v_to", v_from)),
                key=f"b_vto_{sec['id']}_{item['id']}",
                disabled=disabled
            )
    item["v_from"], item["v_to"] = int(v_from), int(v_to)

    preview = ""
    if verses:
        lines = []
        for v in verses:
            if v_from <= v["verse"] <= v_to:
                lines.append(f"{book_name} {int(chap)}:{v['verse']} {v['text']}")
        preview = "\n".join(lines)

    st.text_area("미리보기", value=preview, height=120, disabled=True, key=f"b_prev_{sec['id']}_{item['id']}")

    if st.button("📥 구절 추가(누적)", key=f"b_add_{sec['id']}_{item['id']}", disabled=disabled):
        prev = item.get("verse_text", "") or ""
        new_block = preview.strip()
        if new_block:
            item["verse_text"] = (prev + ("\n" if prev else "") + new_block).strip()
            st.session_state[f"b_txt_{sec['id']}_{item['id']}"] = item["verse_text"]
            st.success("본문에 구절을 추가했습니다.")
            st.rerun()
        else:
            st.warning("추가할 본문이 없습니다. 책/장/절을 확인해 주세요.")

    # 편집 가능한 본문 내용
    text_key = f"b_txt_{sec['id']}_{item['id']}"
    if text_key not in st.session_state:
        st.session_state[text_key] = item.get("verse_text", "")
    txt = st.text_area("본문 내용(편집 가능)", key=text_key, height=140, disabled=disabled)
    item["verse_text"] = txt

def render_image_item(sec: Dict[str, Any], item: Dict[str, Any], disabled: bool):
    st.markdown("**🖼️ 이미지**")
    existing = item.get("files") or []
    if existing:
        with st.expander("기존 이미지(메타) 보기", expanded=False):
            names = []
            for f in existing:
                if isinstance(f, dict):
                    names.append(f.get("name") or os.path.basename(f.get("path", "")))
                elif hasattr(f, "name"):
                    names.append(f.name)
            st.write(", ".join(names) if names else "(목록 없음)")

    uploads = st.file_uploader(
        "이미지 업로드 (PNG/JPG) — 여러 장 선택 가능",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key=f"img_up_{sec['id']}_{item['id']}",
        disabled=disabled
    )
    if uploads and len(uploads) > 0:
        item["files"] = existing + uploads
    else:
        item["files"] = existing

def render_file_item(sec: Dict[str, Any], item: Dict[str, Any], disabled: bool):
    st.markdown("**📎 기타 파일**")
    existing = item.get("file")
    if existing:
        if isinstance(existing, dict):
            st.caption(f"기존 첨부: {existing.get('name','(이름 없음)')}")
        elif hasattr(existing, "name"):
            st.caption(f"기존 첨부: {existing.name}")

    up = st.file_uploader(
        "첨부 업로드 (예: PDF/PPTX/XLSX/HWP/DOCX 등)",
        type=None,
        accept_multiple_files=False,
        key=f"file_up_{sec['id']}_{item['id']}",
        disabled=disabled
    )
    if up is not None:
        item["file"] = up
    else:
        item["file"] = existing

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

def detach_section_files_for_github(sections: List[Dict[str, Any]], files_dir: str, msg_prefix: str) -> List[Dict[str, Any]]:
    """
    sections 내 item들에서 UploadedFile을 GitHub 업로드 후 metadata dict로 치환
    """
    out = []
    for sec in sections:
        sec2 = deepcopy(sec)
        items2 = []
        for it in sec2.get("items", []):
            it2 = deepcopy(it)
            t = it2.get("type")

            if t == "이미지":
                metas = []
                for f in (it2.get("files") or []):
                    if hasattr(f, "getvalue"):
                        metas.append(upload_streamlit_file_to_github(f, files_dir, msg_prefix))
                    elif isinstance(f, dict) and "path" in f:
                        metas.append(f)
                it2["files"] = metas

            elif t == "기타파일":
                f = it2.get("file")
                if hasattr(f, "getvalue"):
                    it2["file"] = upload_streamlit_file_to_github(f, files_dir, msg_prefix)
                elif isinstance(f, dict) and "path" in f:
                    pass
                else:
                    it2["file"] = None

            # 성경은 텍스트만 유지
            items2.append(it2)

        sec2["items"] = items2
        out.append(sec2)
    return out

# ---------------------------
# build_docx
# ---------------------------
def build_docx(worship_date: date, services: List[str], sections: List[Dict[str, Any]],
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
    doc.add_heading("구역(스토리보드)", level=1)

    if not sections:
        doc.add_paragraph("(추가된 구역이 없습니다)")
    else:
        for s_idx, sec in enumerate(sections, start=1):
            sec_title = (sec.get("title") or "").strip() or f"구역 {s_idx}"
            storyboard = sec.get("storyboard", "") or ""
            items = sec.get("items", []) or []

            doc.add_heading(f"{s_idx}. {sec_title}", level=2)

            # 구역 스토리보드
            p = doc.add_paragraph()
            p.add_run("스토리보드/설명: ").bold = True
            if storyboard.strip():
                add_rich_text(p, storyboard)
            else:
                p.add_run("(미입력)")

            doc.add_paragraph("")

            if not items:
                doc.add_paragraph("(이 구역에 추가된 자료가 없습니다)")
                doc.add_paragraph("")
                continue

            # 구역 내부 자료들
            for i_idx, it in enumerate(items, start=1):
                t = it.get("type")
                doc.add_heading(f"{s_idx}-{i_idx}. {t}", level=3)

                if t == "성경":
                    verse_text = it.get("verse_text", "") or ""
                    if verse_text.strip():
                        for line in verse_text.splitlines():
                            p2 = doc.add_paragraph()
                            add_rich_text(p2, line)
                        doc.add_paragraph("")
                    else:
                        doc.add_paragraph("(성경 구절 미입력)")
                        doc.add_paragraph("")

                elif t == "이미지":
                    files = it.get("files", []) or []
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
                                doc.add_paragraph(
                                    f"(이미지 삽입 실패) 파일: "
                                    f"{(f.get('name') if isinstance(f, dict) else getattr(f, 'name', 'unknown'))}"
                                )
                        doc.add_paragraph("")
                    else:
                        doc.add_paragraph("(이미지 파일 없음)")
                        doc.add_paragraph("")

                elif t == "기타파일":
                    f = it.get("file")
                    if isinstance(f, dict) and "name" in f:
                        doc.add_paragraph(f"첨부 파일: {f['name']} (문서에 직접 삽입되지 않습니다)")
                    elif f is not None and hasattr(f, "getvalue"):
                        doc.add_paragraph(f"첨부 파일: {getattr(f, 'name', '파일')} (문서에 직접 삽입되지 않습니다)")
                    else:
                        doc.add_paragraph("(첨부 파일 없음)")
                    doc.add_paragraph("")

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
        "sections": st.session_state.get("sections", []),
        "user_name": st.session_state.get("user_name"),
        "position": st.session_state.get("position"),
        "role": st.session_state.get("role"),
        "saved_at": datetime.now(timezone.utc).isoformat()
    }

def load_into_session(payload: dict):
    st.session_state.worship_date = date.fromisoformat(payload.get("worship_date"))
    st.session_state.services_selected = payload.get("services", [])
    st.session_state.sections = payload.get("sections", [])

def gh_paths(user_name: str, worship_date: date, submission_id: str = None):
    base = st.secrets.get("GITHUB_BASE_DIR", "worship_submissions")
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
# ② 구역(Section) 추가/편집
# ---------------------------
st.markdown("<div class='section-title'>② 구역(스토리보드) 만들기</div>", unsafe_allow_html=True)
st.caption("• 구역 단위로 순서 조정 가능 / 구역 안에 성경·이미지·파일 자료를 추가하세요.\n• 강조: **굵게**, ==형광펜== (Word 변환 시 반영)")

top_btns = st.columns([1, 5])
with top_btns[0]:
    if st.button("➕ 구역 추가", disabled=not can_edit):
        add_section()
        st.rerun()

if not st.session_state.sections:
    st.info("아직 구역이 없습니다. '구역 추가'를 눌러 시작하세요.")

for s_idx, sec in enumerate(st.session_state.sections):
    with st.container(border=True):
        head = st.columns([3.0, 0.35, 0.35, 0.5])
        with head[0]:
            sec_title_key = f"sec_title_{sec['id']}"
            if sec_title_key not in st.session_state:
                st.session_state[sec_title_key] = sec.get("title", "")
            sec["title"] = st.text_input("구역 제목", key=sec_title_key, disabled=not can_edit, placeholder="예: 찬양 / 광고 / 본문 / 설교 도입 등")

        with head[1]:
            st.write("")
            st.button("▲", key=f"sec_up_{sec['id']}", disabled=(not can_edit or s_idx == 0),
                      on_click=move_section, args=(sec["id"], "up"))
        with head[2]:
            st.write("")
            st.button("▼", key=f"sec_dn_{sec['id']}", disabled=(not can_edit or s_idx == len(st.session_state.sections)-1),
                      on_click=move_section, args=(sec["id"], "down"))
        with head[3]:
            st.write("")
            if st.button("삭제", key=f"sec_del_{sec['id']}", disabled=not can_edit):
                remove_section(sec["id"])
                st.rerun()

        # 구역 스토리보드
        sb_key = f"sec_sb_{sec['id']}"
        if sb_key not in st.session_state:
            st.session_state[sb_key] = sec.get("storyboard", "")
        sec["storyboard"] = st.text_area(
            "스토리보드/설명(구역 단위)",
            key=sb_key,
            height=120,
            disabled=not can_edit,
            placeholder="예: 이 구역은 1절~3절 읽고, 4절에서 강조. 영상 전환 타이밍은 00:35 등"
        )

        st.markdown("---")

        # 구역 내부 자료 추가 버튼
        btn_row = st.columns([1, 1, 1, 6])
        with btn_row[0]:
            if st.button("📖 성경 추가", key=f"add_b_{sec['id']}", disabled=not can_edit):
                add_item(sec, "성경")
                st.rerun()
        with btn_row[1]:
            if st.button("🖼️ 이미지 추가", key=f"add_i_{sec['id']}", disabled=not can_edit):
                add_item(sec, "이미지")
                st.rerun()
        with btn_row[2]:
            if st.button("📎 파일 추가", key=f"add_f_{sec['id']}", disabled=not can_edit):
                add_item(sec, "기타파일")
                st.rerun()

        items = sec.get("items", [])
        if not items:
            st.caption("이 구역에는 아직 자료가 없습니다.")
        else:
            for i_idx, it in enumerate(items):
                with st.container(border=True):
                    it_head = st.columns([2.2, 0.25, 0.25, 0.35])
                    with it_head[0]:
                        st.markdown(f"<span class='chip'>자료</span> **{it.get('type','')}**", unsafe_allow_html=True)
                    with it_head[1]:
                        st.button("▲", key=f"it_up_{sec['id']}_{it['id']}", disabled=(not can_edit or i_idx == 0),
                                  on_click=move_item, args=(sec, it["id"], "up"))
                    with it_head[2]:
                        st.button("▼", key=f"it_dn_{sec['id']}_{it['id']}", disabled=(not can_edit or i_idx == len(items)-1),
                                  on_click=move_item, args=(sec, it["id"], "down"))
                    with it_head[3]:
                        if st.button("삭제", key=f"it_del_{sec['id']}_{it['id']}", disabled=not can_edit):
                            remove_item(sec, it["id"])
                            st.rerun()

                    # 타입별 렌더링
                    if it.get("type") == "성경":
                        render_bible_item(sec, it, disabled=not can_edit)
                    elif it.get("type") == "이미지":
                        render_image_item(sec, it, disabled=not can_edit)
                    elif it.get("type") == "기타파일":
                        render_file_item(sec, it, disabled=not can_edit)

st.divider()

# ---------------------------
# ③ Word 파일 생성(로컬 다운로드)
# ---------------------------
st.markdown("<div class='section-title'>③ Word 저장 (로컬 미리 받기)</div>", unsafe_allow_html=True)
col1, _ = st.columns([1, 3])
with col1:
    do_save = st.button("📄 Word 만들기", type="primary", disabled=not can_edit)

if do_save and can_edit:
    try:
        docx_bytes = build_docx(
            worship_date=worship_date,
            services=services,
            sections=st.session_state.sections,
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
b1, b2, b3, _ = st.columns([1, 1, 1, 3])
with b1:
    save_draft = st.button("💾 임시 저장", disabled=not can_edit)
with b2:
    load_draft = st.button("↩️ 불러오기", disabled=not can_edit)
with b3:
    submit_now = st.button("✅ 제출", disabled=not can_edit)

if save_draft and can_edit:
    try:
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        sections_detached = detach_section_files_for_github(
            st.session_state.sections, p["files_dir"], msg_prefix="[draft-files]"
        )
        data = serialize_submission()
        data["sections"] = sections_detached
        gh_put_bytes(
            p["json"],
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            message=f"[draft] {st.session_state.user_name} {worship_date} 저장"
        )
        st.success("임시 저장되었습니다. (GitHub)")
    except Exception as e:
        st.error(f"임시 저장 실패: {e}")

if load_draft and can_edit:
    try:
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        draft_bytes = gh_get_bytes(p["json"])
        payload = json.loads(draft_bytes.decode("utf-8"))
        load_into_session(payload)
        st.success("임시 저장본을 불러왔습니다.")
        st.rerun()
    except Exception as e:
        st.error(f"불러오기 실패 또는 저장본 없음: {e}")

if submit_now and can_edit:
    try:
        sub_id = st.session_state.submission_id or datetime.now().strftime("%H%M%S") + "-" + uuid.uuid4().hex[:6]
        st.session_state.submission_id = sub_id
        p = gh_paths(st.session_state.user_name, worship_date, submission_id=sub_id)

        sections_detached = detach_section_files_for_github(
            st.session_state.sections, p["files_dir"], msg_prefix="[submit-files]"
        )

        docx_bytes = build_docx(
            worship_date=worship_date,
            services=st.session_state.services_selected,
            sections=st.session_state.sections,
            user_name=st.session_state.user_name,
            position=st.session_state.position,
            role=st.session_state.role
        )

        data = serialize_submission()
        data["status"] = "submitted"
        data["submission_id"] = sub_id
        data["sections"] = sections_detached

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
    base = st.secrets.get("GITHUB_BASE_DIR", "worship_submissions")
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
                                        f"- 구역수: {len(payload.get('sections', []))}\n"
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
                                if json_item and Document is not None:
                                    if st.button("📄 즉석 Word 생성", key=f"mk_{sel_day}_{u['name']}_{s['name']}"):
                                        try:
                                            payload = json.loads(gh_get_bytes(json_item["path"]).decode("utf-8"))
                                            docx_bytes2 = build_docx(
                                                worship_date=date.fromisoformat(sel_day),
                                                services=payload.get("services", []),
                                                sections=payload.get("sections", []),
                                                user_name=payload.get("user_name"),
                                                position=payload.get("position"),
                                                role=payload.get("role")
                                            )
                                            st.download_button(
                                                "⬇️ Word 다운로드(즉석)",
                                                data=docx_bytes2,
                                                file_name=f"설교자료_{sel_day}_{u['name']}_{s['name']}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"dl2_{sel_day}_{u['name']}_{s['name']}"
                                            )
                                        except Exception as e:
                                            st.error(f"생성 오류: {e}")

# ---------------------------
# 풋터
# ---------------------------
st.markdown(
    """
    <hr/>
    <div class='small-note'>
    ⚙️ 이미지/첨부 파일은 임시저장/제출 시 GitHub에 업로드되고, Word에는 이미지(가능한 경우)만 삽입됩니다.<br>
    ✍️ 강조법: **굵게**, ==형광펜== (Word 변환 시 자동 적용)<br>
    🔗 성경 본문: <code>bible_books_json/책이름.json</code> (없으면 <code>bible_books_json/books/책이름.json</code> 폴백)
    </div>
    """,
    unsafe_allow_html=True
)
