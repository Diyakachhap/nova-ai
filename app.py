import os
import re
import json
import uuid
import base64
from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components
from pypdf import PdfReader
import docx
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, SystemMessage


BOT_NAME = "NOVA AI"
BOT_SUBTITLE = "Multilingual AI workspace for chats, PDF / DOCX / TXT files, and images or notes."
HISTORY_FILE = Path(__file__).resolve().parent / "nova_ai_chats.json"

if "sidebar_hidden" not in st.session_state:
    st.session_state.sidebar_hidden = False

st.set_page_config(
    page_title=BOT_NAME,
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed" if st.session_state.sidebar_hidden else "expanded",
)

EMERALD_BG = "#0B1512"
EMERALD_SIDEBAR = "#0E1A16"
EMERALD_CARD = "#14241F"
EMERALD_ELEVATED = "#1B3029"
TEXT_PRIMARY = "#F1F5F9"
TEXT_MUTED = "#A7BEB6"
BORDER_EMERALD = "#254038"
BRASS_ACCENT = "#C9A96E"
BRASS_BRIGHT = "#E5C68B"

hide_sidebar_css = (
    """
    [data-testid="stSidebar"] {
        display: none !important;
    }
    """
    if st.session_state.sidebar_hidden
    else ""
)

st.markdown(
    f"""
    <style>
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        {hide_sidebar_css}

        [data-testid="stHeader"] {{
            background-color: {EMERALD_BG} !important;
            border-bottom: 1px solid rgba(201, 169, 110, 0.14);
        }}
        [data-testid="stBottom"],
        [data-testid="stBottom"] > div {{
            background-color: {EMERALD_BG} !important;
        }}

        .stApp {{
            background: radial-gradient(circle at top right, rgba(201, 169, 110, 0.08), transparent 45%), {EMERALD_BG};
            color: {TEXT_PRIMARY};
        }}

        [data-testid="stSidebar"] {{
            background-color: {EMERALD_SIDEBAR};
            border-right: 1px solid {BORDER_EMERALD};
        }}
        [data-testid="stSidebar"] .stMarkdown,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label {{
            color: {TEXT_PRIMARY};
        }}

        [data-testid="stSidebar"] [data-testid="stPopover"] > button {{
            background-color: {EMERALD_CARD} !important;
            border: 1px solid {BORDER_EMERALD} !important;
            color: {BRASS_BRIGHT} !important;
            border-radius: 9px !important;
            padding: 0.35rem 0.45rem !important;
            min-height: 0 !important;
            line-height: 1 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stPopover"] > button:hover {{
            background-color: {EMERALD_ELEVATED} !important;
            border-color: rgba(201, 169, 110, 0.55) !important;
        }}
        [data-testid="stPopoverBody"] {{
            background-color: {EMERALD_CARD} !important;
            border: 1px solid {BORDER_EMERALD} !important;
            border-radius: 10px !important;
            padding: 0.35rem !important;
        }}

        .main .block-container h1,
        .main .block-container h2,
        .main .block-container h3 {{
            color: {TEXT_PRIMARY};
            letter-spacing: -0.02em;
        }}
        .stCaption, .stCaption p {{
            color: {TEXT_MUTED} !important;
        }}

        [data-testid="stSidebar"] input,
        .main input[type="text"] {{
            background-color: {EMERALD_CARD} !important;
            color: {TEXT_PRIMARY} !important;
            border: 1px solid {BORDER_EMERALD} !important;
            border-radius: 10px !important;
        }}
        [data-testid="stSidebar"] input::placeholder,
        .main input[type="text"]::placeholder {{
            color: {TEXT_MUTED} !important;
            opacity: 1 !important;
        }}

        [data-testid="stSelectbox"] > div > div {{
            background-color: {EMERALD_CARD} !important;
            color: {TEXT_PRIMARY} !important;
            border: 1px solid {BORDER_EMERALD} !important;
            border-radius: 10px !important;
        }}
        [data-testid="stSelectbox"] svg {{
            fill: {BRASS_BRIGHT} !important;
        }}

        [data-testid="stChatMessage"] {{
            background-color: {EMERALD_CARD};
            border: 1px solid {BORDER_EMERALD};
            border-radius: 16px;
            padding: 0.85rem 1.15rem;
            margin-bottom: 0.75rem;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.35);
        }}
        [data-testid="stChatMessage"] p,
        [data-testid="stChatMessage"] li,
        [data-testid="stChatMessage"] span,
        [data-testid="stChatMessage"] code,
        [data-testid="stChatMessage"] pre,
        [data-testid="stChatMessage"] strong,
        [data-testid="stChatMessage"] em,
        [data-testid="stChatMessage"] td,
        [data-testid="stChatMessage"] th {{
            color: {TEXT_PRIMARY} !important;
            font-style: normal !important;
        }}
        [data-testid="stChatMessage"] table {{
            width: 100% !important;
            border-collapse: collapse !important;
            margin-top: 0.5rem !important;
            margin-bottom: 0.5rem !important;
            background-color: #0F1C18 !important;
        }}
        [data-testid="stChatMessage"] thead th {{
            background-color: {EMERALD_ELEVATED} !important;
            color: {BRASS_BRIGHT} !important;
            border: 1px solid rgba(201, 169, 110, 0.35) !important;
            padding: 0.65rem 0.85rem !important;
            font-weight: 700 !important;
        }}
        [data-testid="stChatMessage"] tbody td {{
            border: 1px solid {BORDER_EMERALD} !important;
            padding: 0.65rem 0.85rem !important;
            color: {TEXT_PRIMARY} !important;
        }}
        [data-testid="stChatMessage"] tbody tr:nth-child(even) td {{
            background-color: #13231E !important;
        }}

        [data-testid="stAlert"] {{
            background-color: #2B1518 !important;
            border: 1px solid #F87171 !important;
            border-radius: 12px !important;
        }}
        [data-testid="stAlert"] p,
        [data-testid="stAlert"] span,
        [data-testid="stAlert"] div {{
            color: #FFE4E6 !important;
        }}

        [data-testid="stChatInput"] textarea {{
            color: {TEXT_PRIMARY} !important;
            background-color: {EMERALD_CARD} !important;
            padding-right: 48px !important;
        }}
        [data-testid="stChatInput"] textarea::placeholder {{
            color: {TEXT_MUTED} !important;
            opacity: 1 !important;
        }}
        [data-testid="stChatInput"] > div {{
            position: relative !important;
            background-color: {EMERALD_CARD} !important;
            border: 1px solid {BORDER_EMERALD} !important;
            border-radius: 16px !important;
        }}
        [data-testid="stChatInput"] > div:focus-within {{
            border-color: {BRASS_ACCENT} !important;
            box-shadow: 0 0 0 2px rgba(201, 169, 110, 0.22) !important;
        }}

        [data-testid="stFileUploader"] section {{
            background-color: {EMERALD_CARD} !important;
            border: 1px dashed #32574C !important;
            border-radius: 14px !important;
        }}
        [data-testid="stFileUploader"] section,
        [data-testid="stFileUploader"] section span,
        [data-testid="stFileUploader"] section small,
        [data-testid="stFileUploader"] section div {{
            color: {TEXT_PRIMARY} !important;
        }}
        [data-testid="stFileUploader"] button {{
            background-color: {EMERALD_ELEVATED} !important;
            color: {BRASS_BRIGHT} !important;
            border: 1px solid rgba(201, 169, 110, 0.45) !important;
            border-radius: 9px !important;
        }}

        .stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, {BRASS_BRIGHT}, {BRASS_ACCENT});
            color: #0B1512 !important;
            border: none;
            border-radius: 11px;
            font-weight: 700;
        }}
        .stButton > button[kind="primary"]:hover {{
            background: linear-gradient(135deg, #EED6A6, {BRASS_BRIGHT});
            color: #0B1512 !important;
        }}

        .stButton > button[kind="secondary"] {{
            background-color: {EMERALD_CARD};
            color: {TEXT_PRIMARY} !important;
            border: 1px solid {BORDER_EMERALD};
            border-radius: 10px;
            text-align: left;
            justify-content: flex-start;
            font-weight: 400;
            padding: 0.5rem 0.75rem;
        }}
        .stButton > button[kind="secondary"]:hover {{
            background-color: {EMERALD_ELEVATED};
            border-color: rgba(201, 169, 110, 0.55);
            color: {BRASS_BRIGHT} !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


def get_openrouter_key() -> str | None:
    try:
        if "OPENROUTER_API_KEY" in st.secrets and st.secrets["OPENROUTER_API_KEY"]:
            return str(st.secrets["OPENROUTER_API_KEY"]).strip()
    except Exception:
        pass
    env_key = os.getenv("OPENROUTER_API_KEY")
    if env_key and env_key.strip() and env_key.strip() != "your_openrouter_api_key_here":
        return env_key.strip()
    return None


api_key = get_openrouter_key()
if not api_key:
    st.warning(
        "OPENROUTER_API_KEY is not set. Set a valid OpenRouter key in your terminal first:\n"
        '$env:OPENROUTER_API_KEY="sk-or-v1-..."'
    )
    st.stop()


@st.cache_resource
def create_llm(key: str) -> ChatOpenAI:
    return ChatOpenAI(
        model="openrouter/free",
        api_key=key,
        base_url="https://openrouter.ai/api/v1",
        temperature=0.7,
    )


llm = create_llm(api_key)

LANGUAGES = [
    "Auto-Detect (Match User)",
    "Hinglish (Hindi in Roman)",
    "Hindi (हिन्दी)",
    "English",
    "Bengali (বাংলা)",
    "Spanish (Español)",
    "French (Français)",
    "German (Deutsch)",
    "Japanese (日本語)",
    "Korean (한국어)",
    "Arabic (العربية)",
]

STYLES = ["Balanced", "Concise", "Detailed"]


def build_system_prompt(selected_lang: str, selected_style: str) -> str:
    if selected_lang == "Auto-Detect (Match User)":
        lang_rule = (
            "You are fluent in ALL world languages and dialects. Understand the user's message "
            "regardless of language and reply naturally in the exact same language (including "
            "Hinglish, Hindi, English, Bengali, or any other language the user speaks)."
        )
    else:
        lang_rule = (
            f"You are fluent in ALL languages. Always write your responses in {selected_lang}, "
            "while still accurately understanding input in any language."
        )

    style_rules = {
        "Balanced": "Keep responses clear, structured, and helpful.",
        "Concise": "Keep responses short, direct, and easy to scan.",
        "Detailed": "Give thorough, well-explained answers with helpful examples.",
    }

    return (
        f"You are {BOT_NAME}, a refined, intelligent, and polite AI assistant. "
        f"{lang_rule} {style_rules.get(selected_style, '')} "
        "When RAG document excerpts are provided, use them to answer accurately."
    )


def extract_file_text(uploaded_file) -> str:
    ext = uploaded_file.name.split(".")[-1].lower()
    if ext == "pdf":
        reader = PdfReader(uploaded_file)
        return "\n".join([page.extract_text() or "" for page in reader.pages]).strip()
    if ext == "docx":
        doc = docx.Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()]).strip()
    if ext == "txt":
        return uploaded_file.read().decode("utf-8", errors="ignore").strip()
    return ""


def encode_image(uploaded_file) -> str:
    return base64.b64encode(uploaded_file.getvalue()).decode("utf-8")


def chunk_text(text: str, chunk_size: int = 800, overlap: int = 120) -> list[str]:
    clean = re.sub(r"\s+", " ", text).strip()
    if not clean:
        return []
    if len(clean) <= chunk_size:
        return [clean]
    chunks = []
    start = 0
    step = max(1, chunk_size - overlap)
    while start < len(clean):
        part = clean[start : start + chunk_size].strip()
        if part:
            chunks.append(part)
        start += step
    return chunks


def tokenize(text: str) -> set[str]:
    return {w for w in re.findall(r"[\w\u0900-\u097F\u0980-\u09FF]+", text.lower()) if len(w) > 1}


def retrieve_relevant_chunks(doc_text: str, query: str, top_k: int = 4) -> list[tuple[int, str]]:
    chunks = chunk_text(doc_text)
    if not chunks:
        return []
    if len(chunks) <= top_k:
        return [(i + 1, c) for i, c in enumerate(chunks)]

    q_tokens = tokenize(query)
    if not q_tokens:
        return [(i + 1, c) for i, c in enumerate(chunks[:top_k])]

    scored = []
    for idx, ch in enumerate(chunks):
        c_tokens = tokenize(ch)
        overlap = len(q_tokens & c_tokens)
        phrase_bonus = 2.0 if query.strip().lower() in ch.lower() and len(query.strip()) > 2 else 0.0
        score = overlap + phrase_bonus
        scored.append((score, idx + 1, ch))

    scored.sort(key=lambda item: (-item[0], item[1]))
    top = [(chunk_num, ch) for score, chunk_num, ch in scored[:top_k] if score > 0]
    if not top:
        top = [(i + 1, c) for i, c in enumerate(chunks[:top_k])]
    top.sort(key=lambda item: item[0])
    return top


def search_in_document(doc_text: str, keyword: str) -> list[str]:
    kw = keyword.strip()
    if not kw or not doc_text:
        return []
    pattern = re.compile(re.escape(kw), re.IGNORECASE)
    matches = []
    for m in pattern.finditer(doc_text):
        start = max(0, m.start() - 90)
        end = min(len(doc_text), m.end() + 90)
        snippet = re.sub(r"\s+", " ", doc_text[start:end]).strip()
        if start > 0:
            snippet = "..." + snippet
        if end < len(doc_text):
            snippet = snippet + "..."
        matches.append(snippet)
        if len(matches) >= 5:
            break
    return matches


def make_title(text: str) -> str:
    cleaned = " ".join(text.strip().split())
    if not cleaned:
        return "New Chat"
    return cleaned[:26] + ("..." if len(cleaned) > 26 else "")


def load_persisted_state() -> tuple[dict, str | None]:
    if not HISTORY_FILE.exists():
        return {}, None
    try:
        raw = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        chats_data = raw.get("chats", {})
        active_id = raw.get("active_chat_id")
        restored = {}
        for cid, c in chats_data.items():
            msgs = []
            for m in c.get("messages", []):
                img_b64 = m.get("image_b64")
                img_bytes = base64.b64decode(img_b64) if img_b64 else None
                msgs.append(
                    {
                        "role": m.get("role", "user"),
                        "ui_text": m.get("ui_text", ""),
                        "lc_content": m.get("lc_content", ""),
                        "image_bytes": img_bytes,
                    }
                )
            restored[cid] = {
                "title": c.get("title", "New Chat"),
                "messages": msgs,
            }
        return restored, active_id
    except Exception:
        return {}, None


def save_persisted_state():
    try:
        serializable_chats = {}
        for cid, c in st.session_state.chats.items():
            s_msgs = []
            for m in c.get("messages", []):
                img_bytes = m.get("image_bytes")
                img_b64 = base64.b64encode(img_bytes).decode("utf-8") if img_bytes else None
                s_msgs.append(
                    {
                        "role": m.get("role", "user"),
                        "ui_text": m.get("ui_text", ""),
                        "lc_content": m.get("lc_content", ""),
                        "image_b64": img_b64,
                    }
                )
            serializable_chats[cid] = {
                "title": c.get("title", "New Chat"),
                "messages": s_msgs,
            }
        payload = {
            "active_chat_id": st.session_state.active_chat_id,
            "chats": serializable_chats,
        }
        HISTORY_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def start_new_chat(initial_prompt: str | None = None):
    chat_id = str(uuid.uuid4())
    st.session_state.chats[chat_id] = {
        "title": "New Chat",
        "messages": [],
    }
    st.session_state.active_chat_id = chat_id
    if initial_prompt:
        st.session_state.pending_prompt = initial_prompt
    save_persisted_state()


def delete_single_chat(cid: str):
    if cid in st.session_state.chats:
        del st.session_state.chats[cid]
    if st.session_state.active_chat_id == cid:
        remaining = list(st.session_state.chats.keys())
        if remaining:
            st.session_state.active_chat_id = remaining[-1]
            save_persisted_state()
        else:
            st.session_state.active_chat_id = None
            start_new_chat()
    else:
        save_persisted_state()


if "chats" not in st.session_state or "active_chat_id" not in st.session_state:
    loaded_chats, loaded_active = load_persisted_state()
    st.session_state.chats = loaded_chats
    st.session_state.active_chat_id = loaded_active

if "pending_prompt" not in st.session_state:
    st.session_state.pending_prompt = None

if not st.session_state.active_chat_id or st.session_state.active_chat_id not in st.session_state.chats:
    start_new_chat()

current_chat = st.session_state.chats[st.session_state.active_chat_id]


with st.sidebar:
    col_brand, col_hide = st.columns([5, 2])
    with col_brand:
        st.markdown(
            f"<div style='display:flex; align-items:center; gap:0.45rem; margin-bottom:0.4rem;'>"
            f"<span style='font-size:1.22rem; font-weight:700; color:{BRASS_BRIGHT};'>✦ {BOT_NAME}</span>"
            f"<span style='font-size:0.66rem; padding:0.16rem 0.48rem; border-radius:6px; "
            f"background:{EMERALD_CARD}; color:{TEXT_MUTED}; border:1px solid {BORDER_EMERALD}; "
            f"letter-spacing:0.05em;'>MULTI-LANG</span>"
            f"</div>",
            unsafe_allow_html=True,
        )
    with col_hide:
        if st.button("Hide", key="hide_sidebar_btn", help="Hide sidebar", use_container_width=True, type="secondary"):
            st.session_state.sidebar_hidden = True
            st.rerun()

    if st.button("+ New Chat", use_container_width=True, type="primary"):
        start_new_chat()
        st.rerun()

    st.markdown(
        f"<div style='margin-top:1.05rem; margin-bottom:0.35rem; color:{TEXT_MUTED}; "
        f"font-size:0.72rem; letter-spacing:0.08em; text-transform:uppercase;'>Language & Style</div>",
        unsafe_allow_html=True,
    )
    selected_language = st.selectbox(
        "Reply Language",
        options=LANGUAGES,
        index=0,
        key="nova_lang",
    )
    selected_style = st.selectbox(
        "Response Style",
        options=STYLES,
        index=0,
        key="nova_style",
    )

    st.markdown(
        f"<div style='margin-top:1.05rem; margin-bottom:0.4rem; color:{TEXT_MUTED}; "
        f"font-size:0.72rem; letter-spacing:0.08em; text-transform:uppercase;'>Quick Starters</div>",
        unsafe_allow_html=True,
    )
    starters = [
        ("📄 Summarize attached file", "Please summarize the attached file or image with key takeaways."),
        ("🌍 Translate or explain", "Explain how AI assistants understand multiple languages simply."),
        ("💡 Weekend project ideas", "Suggest 4 creative ideas I can build this weekend."),
        ("📋 Action checklist", "Give me a clear step-by-step checklist to plan a task."),
    ]
    for label, starter_text in starters:
        if st.button(label, key=f"starter_{label}", use_container_width=True, type="secondary"):
            if len(current_chat["messages"]) == 0:
                st.session_state.pending_prompt = starter_text
            else:
                start_new_chat(initial_prompt=starter_text)
            st.rerun()

    st.markdown(
        f"<div style='margin-top:1.1rem; margin-bottom:0.35rem; color:{TEXT_MUTED}; "
        f"font-size:0.72rem; letter-spacing:0.08em; text-transform:uppercase;'>Chat History</div>",
        unsafe_allow_html=True,
    )
    search_query = st.text_input(
        "Search chats",
        placeholder="Search chats...",
        label_visibility="collapsed",
        key="chat_search",
    ).strip().lower()

    for cid in reversed(list(st.session_state.chats.keys())):
        chat_item = st.session_state.chats[cid]
        if search_query and search_query not in chat_item["title"].lower():
            continue
        is_active = cid == st.session_state.active_chat_id

        col_chat, col_menu = st.columns([6, 1], vertical_alignment="center")
        with col_chat:
            if is_active:
                st.markdown(
                    f"<div style='background:{EMERALD_ELEVATED}; border-left:3px solid {BRASS_ACCENT}; "
                    f"color:{TEXT_PRIMARY}; padding:0.46rem 0.75rem; border-radius:10px; "
                    f"font-weight:500; margin-bottom:0.25rem; white-space:nowrap; "
                    f"overflow:hidden; text-overflow:ellipsis;'>{chat_item['title']}</div>",
                    unsafe_allow_html=True,
                )
            else:
                if st.button(
                    chat_item["title"],
                    key=f"open_{cid}",
                    use_container_width=True,
                    type="secondary",
                ):
                    st.session_state.active_chat_id = cid
                    save_persisted_state()
                    st.rerun()
        with col_menu:
            with st.popover("⋮", help="Chat options"):
                if st.button("Delete", key=f"delete_chat_{cid}", use_container_width=True, type="secondary"):
                    delete_single_chat(cid)
                    st.rerun()


col_head, col_toggle = st.columns([8, 2])
with col_head:
    st.markdown(
        f"<h1 style='color:{TEXT_PRIMARY}; margin-bottom:0.2rem;'>"
        f"<span style='color:{BRASS_BRIGHT};'>✦</span> {BOT_NAME}</h1>",
        unsafe_allow_html=True,
    )
with col_toggle:
    if st.session_state.sidebar_hidden:
        st.markdown("<div style='height:0.65rem;'></div>", unsafe_allow_html=True)
        if st.button("Show Sidebar", key="show_sidebar_btn", use_container_width=True, type="secondary"):
            st.session_state.sidebar_hidden = False
            st.rerun()

st.caption(BOT_SUBTITLE)

for msg in current_chat["messages"]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["ui_text"])
        if msg.get("image_bytes"):
            st.image(msg["image_bytes"], width=260)

uploaded_file = st.file_uploader(
    "Upload",
    type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
    key=f"uploader_{st.session_state.active_chat_id}",
    label_visibility="collapsed",
)

extracted_doc_text = ""
if uploaded_file is not None:
    ext = uploaded_file.name.split(".")[-1].lower()
    if ext in ["pdf", "docx", "txt"]:
        extracted_doc_text = extract_file_text(uploaded_file)
        if extracted_doc_text:
            total_chunks = len(chunk_text(extracted_doc_text))
            with st.expander(f"RAG Document Search ({uploaded_file.name} · {total_chunks} chunks)", expanded=False):
                doc_search_q = st.text_input(
                    "Search inside document",
                    placeholder="Search keyword or topic in file...",
                    key=f"rag_search_{st.session_state.active_chat_id}",
                )
                if doc_search_q.strip():
                    matches = search_in_document(extracted_doc_text, doc_search_q)
                    if matches:
                        for i_m, snip in enumerate(matches, 1):
                            st.markdown(f"**Match {i_m}:** {snip}")
                    else:
                        st.caption("No exact match found; relevant chunks will still be retrieved for your question.")

components.html(
    """
    <script>
    (function() {
        const parentDoc = window.parent.document;
        const parentWin = window.parent;
        const SpeechRecognition = parentWin.SpeechRecognition || parentWin.webkitSpeechRecognition;

        function mountMicInsideInput() {
            const chatInputWrap = parentDoc.querySelector('[data-testid="stChatInput"] > div');
            const textarea = parentDoc.querySelector('[data-testid="stChatInput"] textarea');
            if (!chatInputWrap || !textarea) return;

            let micBtn = parentDoc.getElementById("geminiStyleMicBtn");
            if (!micBtn) {
                micBtn = parentDoc.createElement("button");
                micBtn.id = "geminiStyleMicBtn";
                micBtn.type = "button";
                micBtn.title = "Voice search";
                micBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z"/><path d="M19 10v2a7 7 0 0 1-14 0v-2"/><line x1="12" x2="12" y1="19" y2="22"/></svg>`;
                micBtn.style.cssText = `
                    position: absolute;
                    right: 46px;
                    bottom: 8px;
                    width: 32px;
                    height: 32px;
                    border-radius: 50%;
                    border: 1px solid #223A33;
                    background: #1B3029;
                    color: #E5C68B;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    cursor: pointer;
                    z-index: 20;
                    transition: all 0.15s ease;
                `;
                chatInputWrap.appendChild(micBtn);

                if (!SpeechRecognition) {
                    micBtn.disabled = true;
                    micBtn.style.opacity = "0.45";
                    micBtn.title = "Voice search works in Chrome / Edge";
                    return;
                }

                let recognition = null;
                let listening = false;
                let baseText = "";

                function setInputValue(val) {
                    const setter = Object.getOwnPropertyDescriptor(parentWin.HTMLTextAreaElement.prototype, "value").set;
                    setter.call(textarea, val);
                    textarea.dispatchEvent(new Event("input", { bubbles: true }));
                    textarea.focus();
                }

                micBtn.addEventListener("click", (e) => {
                    e.preventDefault();
                    e.stopPropagation();

                    if (listening && recognition) {
                        recognition.stop();
                        return;
                    }

                    recognition = new SpeechRecognition();
                    recognition.interimResults = true;
                    recognition.maxAlternatives = 1;
                    baseText = textarea.value ? textarea.value.trim() + " " : "";

                    recognition.onstart = () => {
                        listening = true;
                        micBtn.style.background = "#E5C68B";
                        micBtn.style.color = "#0B1512";
                        micBtn.style.borderColor = "#E5C68B";
                        textarea.placeholder = "Listening... speak now";
                    };

                    recognition.onresult = (event) => {
                        let interim = "";
                        let finalText = "";
                        for (let i = 0; i < event.results.length; i++) {
                            const piece = event.results[i][0].transcript;
                            if (event.results[i].isFinal) {
                                finalText += piece + " ";
                            } else {
                                interim += piece;
                            }
                        }
                        setInputValue((baseText + finalText + interim).trimStart());
                    };

                    recognition.onerror = () => {
                        textarea.placeholder = "Ask NOVA AI in any language...";
                    };

                    recognition.onend = () => {
                        listening = false;
                        micBtn.style.background = "#1B3029";
                        micBtn.style.color = "#E5C68B";
                        micBtn.style.borderColor = "#223A33";
                        textarea.placeholder = "Ask NOVA AI in any language...";
                    };

                    try {
                        recognition.start();
                    } catch (err) {
                        listening = false;
                    }
                });
            }
        }

        mountMicInsideInput();
        const observer = new MutationObserver(() => mountMicInsideInput());
        observer.observe(parentDoc.body, { childList: true, subtree: true });
    })();
    </script>
    """,
    height=0,
)

typed_prompt = st.chat_input("Ask NOVA AI in any language...")
prompt = typed_prompt

if not prompt and st.session_state.pending_prompt:
    prompt = st.session_state.pending_prompt
    st.session_state.pending_prompt = None

if prompt:
    doc_context = ""
    img_b64 = None
    img_mime = None
    img_bytes = None
    attachment_label = ""

    if uploaded_file is not None:
        ext = uploaded_file.name.split(".")[-1].lower()
        attachment_label = f"\n\n📎 *{uploaded_file.name}*"
        if ext in ["pdf", "docx", "txt"] and extracted_doc_text:
            top_chunks = retrieve_relevant_chunks(extracted_doc_text, prompt, top_k=4)
            if top_chunks:
                formatted_chunks = "\n\n".join(
                    [f"[Chunk {c_num}]\n{c_body}" for c_num, c_body in top_chunks]
                )
                doc_context = (
                    f"\n\n--- Relevant Excerpts from {uploaded_file.name} (RAG) ---\n"
                    f"{formatted_chunks[:15000]}"
                )
        elif ext in ["png", "jpg", "jpeg"]:
            img_b64 = encode_image(uploaded_file)
            img_mime = uploaded_file.type or "image/jpeg"
            img_bytes = uploaded_file.getvalue()

    if len(current_chat["messages"]) == 0:
        title_seed = prompt if prompt.strip() else (
            uploaded_file.name if uploaded_file else "New Chat"
        )
        current_chat["title"] = make_title(title_seed)

    ui_text = prompt + attachment_label
    with st.chat_message("user"):
        st.markdown(ui_text)
        if img_bytes:
            st.image(img_bytes, width=260)

    if img_b64:
        lc_content = [
            {"type": "text", "text": prompt + doc_context},
            {
                "type": "image_url",
                "image_url": {"url": f"data:{img_mime};base64,{img_b64}"},
            },
        ]
    else:
        lc_content = prompt + doc_context

    current_chat["messages"].append(
        {
            "role": "user",
            "ui_text": ui_text,
            "lc_content": lc_content,
            "image_bytes": img_bytes,
        }
    )
    save_persisted_state()

    system_prompt = build_system_prompt(selected_language, selected_style)
    lc_messages = [SystemMessage(content=system_prompt)]
    for m in current_chat["messages"]:
        if m["role"] == "user":
            lc_messages.append(HumanMessage(content=m["lc_content"]))
        elif m["role"] == "assistant":
            lc_messages.append(AIMessage(content=m["lc_content"]))

    with st.chat_message("assistant"):
        try:
            reply = llm.invoke(lc_messages).content
            st.markdown(reply)
            current_chat["messages"].append(
                {
                    "role": "assistant",
                    "ui_text": reply,
                    "lc_content": reply,
                }
            )
            save_persisted_state()
        except Exception as err:
            st.error(f"Error: {err}")