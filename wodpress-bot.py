import base64
import html
import json
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    import docx  # type: ignore
except ImportError:
    docx = None

try:
    import pandas as pd  # type: ignore
except ImportError:
    pd = None

try:
    from PIL import Image  # type: ignore
    import pytesseract  # type: ignore
except ImportError:
    Image = None
    pytesseract = None

import os
import requests
from langchain_google_genai import ChatGoogleGenerativeAI
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()

CONFIG_PATH = Path(__file__).with_name("wodpress_settings.json")
MAX_REFERENCE_CHARS = 4000
DEFAULT_MIN_WORD_COUNT = 1000
MAX_GENERATION_ATTEMPTS = 2
ATTACHMENT_FILETYPES = (
    ("모든 지원 형식", "*.txt *.md *.docx *.xlsx *.xls *.csv *.png *.jpg *.jpeg"),
    ("텍스트", "*.txt *.md"),
    ("워드", "*.docx"),
    ("엑셀", "*.xlsx *.xls"),
    ("CSV", "*.csv"),
    ("이미지", "*.png *.jpg *.jpeg"),
)


def truncate_reference(text: str, limit: int = MAX_REFERENCE_CHARS) -> tuple[str, bool]:
    if len(text) <= limit:
        return text, False
    return text[:limit].rstrip() + "\n... (일부 내용 생략)", True


def extract_text_from_file(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in {".txt", ".md"}:
        return path.read_text(encoding="utf-8", errors="ignore")
    if ext == ".docx":
        if not docx:
            raise RuntimeError("python-docx 라이브러리가 필요합니다.")
        document = docx.Document(path)  # type: ignore[attr-defined]
        return "\n".join(p.text for p in document.paragraphs)
    if ext in {".xlsx", ".xls"}:
        if not pd:
            raise RuntimeError("pandas 및 openpyxl 라이브러리가 필요합니다.")
        sheets = pd.read_excel(path, sheet_name=None)  # type: ignore[attr-defined]
        return "\n\n".join(
            f"[Sheet: {sheet_name}]\n{df.head(50).to_csv(index=False)}"
            for sheet_name, df in sheets.items()
        )
    if ext == ".csv":
        if not pd:
            raise RuntimeError("pandas 라이브러리가 필요합니다.")
        return pd.read_csv(path).head(200).to_csv(index=False)  # type: ignore[attr-defined]
    if ext in {".png", ".jpg", ".jpeg"}:
        if not Image or not pytesseract:
            raise RuntimeError("Pillow 및 pytesseract 라이브러리가 필요합니다.")
        image = Image.open(path)  # type: ignore[attr-defined]
        return pytesseract.image_to_string(image, lang="kor+eng")  # type: ignore[attr-defined]
    raise ValueError(f"지원하지 않는 파일 형식입니다: {path.suffix}")


def create_seo_blog_prompt(topic: str, target_keyword: str, word_count: int = 2500) -> str:
    """Build a structured SEO blog prompt for Gemini."""
    return f"""당신은 구글 SEO 전문가입니다. Rank Math 점수 100점을 목표로 워드프레스용 HTML 포스팅을 작성합니다.

# 작성 대상 정보
- 주제: {topic}
- 타겟 키워드: {target_keyword} (토씨 하나 틀리지 않고 그대로 사용할 것)
- 목표 분량: 공백 포함 최소 {word_count}자 이상 (내용이 길고 풍부해야 함)

# 🚨 [형식 절대 규칙] (위반 시 워드프레스 오류 발생)
1. **[HTML만 사용]**: 마크다운(##, **, - 등)을 절대 사용하지 마십시오. 오직 HTML 태그(<h1>, <h2>, <h3>, <p>, <ul>, <li>, <strong>)만 사용하세요.
2. **[제목 태그]**: 
   - 대제목은 <h1> 태그를 사용하세요.
   - 소제목은 <h2>, <h3> 태그를 사용하세요.
3. **[강조]**: 중요한 단어는 <strong> 태그로 감싸세요.

# 🚨 [SEO 100점 달성 규칙]
1. **[키워드 반복]**: 본문 전체에서 **"{target_keyword}"**를 **최소 8~10회** 반복해서 자연스럽게 사용하세요.
2. **[첫 문장]**: H1 제목 바로 아래, 본문의 **맨 첫 문장은 무조건 "{target_keyword}"(으)로 시작**해야 합니다.
3. **[소제목 키워드]**: 
   - **<h2> 태그** 중 최소 1개 이상에 "{target_keyword}"를 포함하세요.
   - **<h3> 태그** 중 최소 1개 이상에 "{target_keyword}"를 포함하세요.

# 출력 형식 (이 HTML 구조를 그대로 유지할 것)
<div class="seo-overview">
  <p><strong>SEO 제목:</strong> [{target_keyword}와 숫자가 포함된 클릭을 부르는 제목]</p>
  <p><strong>메타 설명:</strong> [{target_keyword}가 포함된 160자 이내 요약]</p>
  <p><strong>타겟 키워드:</strong> {target_keyword}</p>
  <p><strong>추천 포커스 키워드 및 LSI 키워드(15개 이상):</strong></p>
  <ul>
    <li>[{target_keyword}]</li>
    <li>[연관 키워드 1]</li>
    <li>[연관 키워드 2]</li>
    <li>...</li>
  </ul>
</div>

<h1>[여기에도 {target_keyword}가 포함된 H1 제목]</h1>

<p><strong>{target_keyword}</strong> [여기서부터 서론 시작... 첫 문장은 반드시 타겟 키워드로 시작할 것]</p>
<p>[서론 내용 계속...]</p>

<h2>{target_keyword}의 중요성 및 특징</h2>
<p>[본문 내용...]</p>

<h3>{target_keyword} 활용 팁 1</h3>
<p>[상세 내용...]</p>

<h2>[또 다른 소제목]</h2>
<p>[본문 내용 - 중간중간 <strong>{target_keyword}</strong>를 자연스럽게 섞어주세요.]</p>

<h3>[상세 설명 소제목]</h3>
<p>[상세 내용...]</p>

<section class="faq">
  <h2>자주 묻는 질문 (FAQ)</h2>
  <dl>
    <dt>Q1: {target_keyword} 관련 질문?</dt>
    <dd>A: [답변]</dd>
    <dt>Q2: [질문]</dt>
    <dd>A: [답변]</dd>
  </dl>
</section>

<h2>결론</h2>
<p>[핵심 요약 및 마무리]</p>

<div class="links-and-media">
  <p><strong>내부 링크 제안:</strong></p>
  <ol><li>[관련 글 제목]</li></ol>
  <p><strong>외부 링크 제안:</strong></p>
  <ol><li>[나무위키/공공기관 등]</li></ol>
  <p><strong>이미지 alt 텍스트 제안:</strong></p>
  <ol><li>{target_keyword} 관련 사진</li></ol>
  <p><strong>SEO 최적화 해시태그:</strong></p>
  <p>#{target_keyword} #해시태그2 #해시태그3 ...</p>
</div>
"""

# Function to generate blog content
def generate_blog_content(sentence, llm):
    """
    Generates blog content using a given sentence and LLM.

    Args:
        sentence (str): The input sentence for the blog post.
        llm: The LLM model used for generating content.

    Returns:
        str: The generated blog content.
    """
    response = llm.invoke(sentence)
    if hasattr(response, "content"):
        content = response.content
        if isinstance(content, list):
            return "".join(
                part.get("text", "") if isinstance(part, dict) else str(part)
                for part in content
            )
        return content
    return str(response)


def _sanitize_meta_text(value: str) -> str:
    """
    Prepare extracted text for Rank Math metadata fields.
    Removes wrapping quotes, backticks, redundant whitespace,
    and ignores meaningless keywords like 'html'.
    """
    if not value:
        return ""

    cleaned = value.strip()
    cleaned = re.sub(r"<[^>]+>", "", cleaned)
    cleaned = cleaned.replace(""", '"').replace(""", '"')
    cleaned = cleaned.replace("'", "'").replace("'", "'")

    wrapping_quotes = {'"', "'", "`"}
    while len(cleaned) > 1 and cleaned[0] in wrapping_quotes and cleaned[-1] == cleaned[0]:
        cleaned = cleaned[1:-1].strip()

    cleaned = cleaned.strip("\"'`")

    cleaned = re.sub(r"[\s\u00a0]+", " ", cleaned).strip()

    if cleaned.lower() in {"html", "markdown", "text"}:
        return ""
    return cleaned


def extract_rank_math_metadata(content: str) -> dict:
    """
    Extract SEO metadata (title, description, focus keywords) from generated HTML
    or plain text. This is flexible enough to handle cases where the LLM ignores
    HTML formatting instructions.
    """

    def _strip_tags(text: str) -> str:
        return re.sub(r"<[^>]+>", "", text or "").strip()

    def _extract_flexible_line(label: str) -> str:
        """Extracts text following 'label:' on the same line."""
        pattern = rf"(?:<strong>)?\s*{re.escape(label)}\s*:\s*(?:</strong>)?\s*(.*?)\s*($|\n)"
        match = re.search(pattern, content, re.IGNORECASE)
        if not match:
            return ""
        return html.unescape(_strip_tags(match.group(1)))

    def _extract_flexible_keywords() -> list[str]:
        """Extracts keywords from either <li> tags or newline-separated text."""
        pattern = (
            r"(?:<strong>)?\s*추천\s+포커스\s+키워드.*?:.*?(?:</p>)?\s*(.*?)(?=\s*(?:<h[1-6]>|<div\s+class|<h2>결론</h2>|<section\s+class\s*=\s*(\"|')faq\2|<div\s+class\s*=\s*(\"|')final-checklist\3|$))"
        )
        match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
        if not match:
            return []

        block = match.group(1).strip()

        items = re.findall(r"<li>(.*?)</li>", block, re.DOTALL | re.IGNORECASE)

        if not items and "<li>" not in block:
            items = re.split(r"[\n,]", block)

        keywords = []
        for item in items:
            cleaned = _sanitize_meta_text(item)
            if cleaned and len(cleaned) > 1:
                keywords.append(cleaned)
        return keywords

    return {
        "seo_title": _extract_flexible_line("SEO 제목"),
        "meta_description": _extract_flexible_line("메타 설명"),
        "focus_keywords": _extract_flexible_keywords(),
    }


def strip_seo_overview(content: str) -> str:
    """Removes the SEO overview and checklist block from the content."""
    pattern_overview = r'^\s*<div\s+class\s*=\s*("|\')seo-overview\1.*?>.*?</div>'
    cleaned_content = re.sub(
        pattern_overview,
        "",
        content,
        count=1,
        flags=re.DOTALL | re.IGNORECASE,
    )

    pattern_checklist = r'<div\s+class\s*=\s*("|\')final-checklist\1.*?>.*?</div>\s*$'
    cleaned_content = re.sub(
        pattern_checklist,
        "",
        cleaned_content,
        count=1,
        flags=re.DOTALL | re.IGNORECASE,
    )

    marker = "최종 검토 체크리스트"
    marker_index = cleaned_content.find(marker)
    if marker_index != -1:
        cleaned_content = cleaned_content[:marker_index]

    hashtag_marker = "SEO 최적화 해시태그"
    hashtag_index = cleaned_content.find(hashtag_marker)
    if hashtag_index != -1:
        cleaned_content = cleaned_content[:hashtag_index]

    return cleaned_content.strip()


def extract_hashtags(content: str) -> list[str]:
    """Extracts #hashtags from the 'SEO 최적화 해시태그' section."""
    pattern_section = r"(?:<strong>)?\s*SEO\s+최적화\s+해시태그.*?:.*?(?:</p>)?\s*(.*)"
    match = re.search(pattern_section, content, re.DOTALL | re.IGNORECASE)
    if not match:
        return []

    block = match.group(1)

    block_limited = re.split(r"<h[1-6]>|<div", block, 1)[0]

    tags_raw = re.findall(r"#(\S+)", block_limited)

    sanitized_tags = []
    for tag in tags_raw:
        cleaned = _sanitize_meta_text(tag)
        if cleaned:
            sanitized_tags.append(cleaned)

    return sanitized_tags


def count_words_from_html(html_text: str) -> int:
    """Counts words in HTML by stripping tags and splitting on whitespace."""
    text = re.sub(r"<[^>]+>", " ", html_text or "")
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return 0
    return len(text.split(" "))

# Function to create a WordPress post
def create_post(wp_url, headers, title, content, meta_input=None, tag_names=None):
    """
    Creates a WordPress post using a CUSTOM API Endpoint to bypass Hostinger WAF.
    Endpoint: /wp-json/private/v1/post
    """
    # PHP 코드(전용 통로)가 받을 데이터 구조
    data = {
        'title': title,
        'content': content,
        'tags': tag_names if tag_names else [],

        # meta_input 안에 숨기지 않고, 밖으로 꺼내서 보냅니다.
        'rank_math_focus_keyword': meta_input.get('rank_math_focus_keyword', '') if meta_input else '',
        'rank_math_description': meta_input.get('rank_math_description', '') if meta_input else '',
        'rank_math_title': meta_input.get('rank_math_title', '') if meta_input else '',
    }

    # [핵심] 우리가 만든 전용 통로 주소 사용
    api_url = wp_url.rstrip('/') + '/wp-json/private/v1/post'

    # 디버그 출력
    print(f"[System] Sending to Custom Endpoint: {api_url}")

    try:
        response = requests.post(
            api_url,
            headers=headers,
            json=data,
            timeout=30,
        )
        return response
    except requests.exceptions.ConnectionError as e:
        # DNS 해석 실패 또는 연결 오류 처리
        error_msg = str(e)
        if 'getaddrinfo failed' in error_msg or 'NameResolutionError' in error_msg:
            from urllib.parse import urlparse
            parsed = urlparse(wp_url)
            host = parsed.netloc or parsed.path.split('/')[0]
            raise ConnectionError(
                f"DNS 해석 실패: '{host}' 도메인을 찾을 수 없습니다.\n\n"
                f"가능한 원인:\n"
                f"1. 인터넷 연결을 확인하세요\n"
                f"2. 도메인 이름이 올바른지 확인하세요\n"
                f"3. DNS 설정을 확인하세요\n"
                f"4. 방화벽이나 프록시 설정을 확인하세요"
            ) from e
        raise
    except requests.exceptions.Timeout:
        raise ConnectionError(
            "연결 시간 초과: 서버에 연결하는데 시간이 너무 오래 걸립니다.\n\n"
            "가능한 원인:\n"
            "1. 서버가 응답하지 않습니다\n"
            "2. 네트워크 연결이 느립니다\n"
            "3. 방화벽이 연결을 차단하고 있습니다"
        )
    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"네트워크 오류 발생: {str(e)}") from e


# Function to fetch WordPress posts
def fetch_posts(wp_url, headers, per_page=20, page=1, status='draft'):
    """
    Fetches WordPress posts using REST API.
    
    Args:
        wp_url: WordPress site URL
        headers: Authentication headers
        per_page: Number of posts per page (default: 20)
        page: Page number (default: 1)
        status: Post status ('draft', 'publish', 'any', etc.)
    
    Returns:
        Response object with posts data
    """
    api_url = wp_url.rstrip('/') + '/wp-json/wp/v2/posts'
    params = {
        'per_page': per_page,
        'page': page,
        'status': status,
        '_embed': 'true',
    }
    
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        return response
    except requests.exceptions.ConnectionError as e:
        error_msg = str(e)
        if 'getaddrinfo failed' in error_msg or 'NameResolutionError' in error_msg:
            from urllib.parse import urlparse
            parsed = urlparse(wp_url)
            host = parsed.netloc or parsed.path.split('/')[0]
            raise ConnectionError(
                f"DNS 해석 실패: '{host}' 도메인을 찾을 수 없습니다.\n\n"
                f"가능한 원인:\n"
                f"1. 인터넷 연결을 확인하세요\n"
                f"2. 도메인 이름이 올바른지 확인하세요\n"
                f"3. DNS 설정을 확인하세요\n"
                f"4. 방화벽이나 프록시 설정을 확인하세요"
            ) from e
        raise
    except requests.exceptions.Timeout:
        raise ConnectionError(
            "연결 시간 초과: 서버에 연결하는데 시간이 너무 오래 걸립니다."
        )
    except Exception as e:
        print(f"[Error] Failed to fetch posts: {e}")
        raise


# Function to fetch a single WordPress post
def fetch_post(wp_url, headers, post_id):
    """
    Fetches a single WordPress post by ID.
    
    Args:
        wp_url: WordPress site URL
        headers: Authentication headers
        post_id: Post ID
    
    Returns:
        Response object with post data
    """
    api_url = wp_url.rstrip('/') + f'/wp-json/wp/v2/posts/{post_id}'
    
    try:
        response = requests.get(api_url, headers=headers, timeout=30)
        return response
    except Exception as e:
        print(f"[Error] Failed to fetch post: {e}")
        return None

class WordPressBotGUI:
    """Tkinter-based GUI for generating and posting WordPress drafts."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("WordPress Bot with Gemini")
        self.root.state("zoomed")
        self.attached_files: list[Path] = []
        self.cancel_flag = False
        self.current_thread = None

        self._build_layout()
        self._load_preferences()

    def _build_layout(self) -> None:
        padding = {"padx": 10, "pady": 6}

        tk.Label(self.root, text="Gemini API Key").grid(row=0, column=0, sticky="w", **padding)
        self.gemini_key_entry = tk.Entry(self.root, width=54, show="*")
        self.gemini_key_entry.grid(row=0, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="WordPress URL").grid(row=1, column=0, sticky="w", **padding)
        self.wordpress_url_entry = tk.Entry(self.root, width=54)
        self.wordpress_url_entry.grid(row=1, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="Authentication").grid(row=2, column=0, sticky="nw", **padding)
        self.auth_method_var = tk.StringVar(value="application_password")
        auth_frame = tk.Frame(self.root)
        auth_frame.grid(row=2, column=1, sticky="w", **padding)
        tk.Radiobutton(
            auth_frame,
            text="Application Password (Basic Auth)",
            variable=self.auth_method_var,
            value="application_password",
            anchor="w",
        ).pack(fill="x", pady=1)
        tk.Radiobutton(
            auth_frame,
            text="Bearer Token (JWT 등)",
            variable=self.auth_method_var,
            value="bearer",
            anchor="w",
        ).pack(fill="x", pady=1)

        tk.Label(self.root, text="WordPress Username").grid(row=3, column=0, sticky="w", **padding)
        self.wordpress_user_entry = tk.Entry(self.root, width=54)
        self.wordpress_user_entry.grid(row=3, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="WordPress Token / App Password").grid(row=4, column=0, sticky="w", **padding)
        self.wordpress_secret_entry = tk.Entry(self.root, width=54, show="*")
        self.wordpress_secret_entry.grid(row=4, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="Post Title").grid(row=5, column=0, sticky="w", **padding)
        self.title_entry = tk.Entry(self.root, width=54)
        self.title_entry.grid(row=5, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="Blog Topic").grid(row=6, column=0, sticky="w", **padding)
        self.topic_entry = tk.Entry(self.root, width=54)
        self.topic_entry.grid(row=6, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="Target Keyword").grid(row=7, column=0, sticky="w", **padding)
        self.target_keyword_entry = tk.Entry(self.root, width=54)
        self.target_keyword_entry.grid(row=7, column=1, sticky="ew", **padding)

        tk.Label(self.root, text="Word Count (min)").grid(row=8, column=0, sticky="w", **padding)
        self.word_count_entry = tk.Entry(self.root, width=54)
        self.word_count_entry.insert(0, "1500")
        self.word_count_entry.grid(row=8, column=1, sticky="ew", **padding)

        attachment_frame = tk.LabelFrame(self.root, text="첨부 파일 (프롬프트 참고용)")
        attachment_frame.grid(row=9, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 6))
        attachment_buttons = tk.Frame(attachment_frame)
        attachment_buttons.pack(fill="x", padx=6, pady=(6, 0))
        self.attach_button = tk.Button(
            attachment_buttons,
            text="파일 추가",
            command=self._on_attach_files,
        )
        self.attach_button.pack(side="left")
        self.remove_attachment_button = tk.Button(
            attachment_buttons,
            text="선택 삭제",
            command=self._on_remove_attachment,
        )
        self.remove_attachment_button.pack(side="left", padx=(6, 0))
        self.attachment_list = tk.Listbox(attachment_frame, height=4)
        self.attachment_list.pack(fill="x", expand=True, padx=6, pady=(6, 6))

        tk.Label(self.root, text="Prompt / Sentence").grid(row=10, column=0, sticky="nw", **padding)
        self.build_prompt_button = tk.Button(
            self.root,
            text="Build Prompt",
            command=self._on_build_prompt_clicked,
        )
        self.build_prompt_button.grid(row=10, column=1, sticky="ne", padx=10, pady=6)

        self.prompt_text = ScrolledText(self.root, height=10, wrap="word")
        self.prompt_text.grid(row=11, column=1, sticky="nsew", **padding)

        tk.Label(self.root, text="Generated Content").grid(row=12, column=0, sticky="nw", **padding)
        self.output_text = ScrolledText(self.root, height=12, wrap="word")
        self.output_text.grid(row=12, column=1, sticky="nsew", **padding)
        self.output_text.configure(state="disabled")

        self.status_label = tk.Label(self.root, text="", anchor="w", fg="blue")
        self.status_label.grid(row=13, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 6))

        button_frame = tk.Frame(self.root)
        button_frame.grid(row=14, column=0, columnspan=2, pady=12)
        
        self.fetch_button = tk.Button(
            button_frame,
            text="WordPress 포스트 가져오기",
            command=self._on_fetch_clicked,
        )
        self.fetch_button.pack(side="left", padx=5)

        self.generate_button = tk.Button(
            button_frame,
            text="Generate & Create Draft",
            command=self._on_generate_clicked,
        )
        self.generate_button.pack(side="left", padx=5)

        self.cancel_button = tk.Button(
            button_frame,
            text="취소",
            command=self._on_cancel_clicked,
            state="disabled",
        )
        self.cancel_button.pack(side="left", padx=5)

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(11, weight=2)
        self.root.grid_rowconfigure(12, weight=2)

    def _on_attach_files(self) -> None:
        selected = filedialog.askopenfilenames(
            parent=self.root,
            title="참고 파일 선택",
            filetypes=ATTACHMENT_FILETYPES,
        )
        if not selected:
            return

        added = 0
        for path_str in selected:
            path = Path(path_str)
            if path not in self.attached_files:
                self.attached_files.append(path)
                self.attachment_list.insert("end", path.name)
                added += 1
        if added:
            self._set_status(f"{added}개의 파일을 첨부했습니다.", "blue")

    def _on_remove_attachment(self) -> None:
        selections = list(self.attachment_list.curselection())
        if not selections:
            return
        for index in reversed(selections):
            self.attachment_list.delete(index)
            self.attached_files.pop(index)
        self._set_status("선택한 첨부 파일을 제거했습니다.", "blue")

    def _prepare_reference_context(self) -> tuple[str, list[str], bool]:
        if not self.attached_files:
            return "", [], False

        collected: list[str] = []
        messages: list[str] = []
        had_error = False

        for path in self.attached_files:
            try:
                text = extract_text_from_file(path)
            except Exception as exc:  # noqa: BLE001
                had_error = True
                messages.append(f"{path.name}: {exc}")
                continue

            cleaned = text.strip()
            if not cleaned:
                had_error = True
                messages.append(f"{path.name}: 추출된 텍스트가 없습니다.")
                continue

            truncated_text, was_truncated = truncate_reference(cleaned)
            if was_truncated:
                messages.append(f"{path.name}: 내용이 길어 일부만 사용했습니다.")
            collected.append(f"[{path.name}]\n{truncated_text}")

        reference_text = "\n\n".join(collected).strip()
        if not reference_text:
            had_error = True
            if not messages:
                messages.append("첨부 파일에서 활용 가능한 내용을 찾지 못했습니다.")

        return reference_text, messages, had_error

    def _on_generate_clicked(self) -> None:
        # 환경 변수 우선, 없으면 GUI 입력값 사용
        gemini_key = os.getenv("GEMINI_API_KEY", "") or self.gemini_key_entry.get().strip()
        wordpress_user = os.getenv("WORDPRESS_USERNAME", "") or self.wordpress_user_entry.get().strip()
        wordpress_secret = os.getenv("WORDPRESS_SECRET", "") or self.wordpress_secret_entry.get().strip()
        wordpress_url = os.getenv("WORDPRESS_URL", "") or self.wordpress_url_entry.get().strip()
        title = self.title_entry.get().strip()
        prompt = self.prompt_text.get("1.0", "end").strip()
        auth_method = os.getenv("WORDPRESS_AUTH_METHOD", "") or self.auth_method_var.get()

        required_fields = [gemini_key, wordpress_secret, wordpress_url, title, prompt]
        if auth_method == "application_password":
            required_fields.append(wordpress_user)

        if not all(required_fields):
            messagebox.showwarning("Missing information", "Please fill in all fields before proceeding.")
            return

        word_count_raw = self.word_count_entry.get().strip()
        try:
            min_word_count = int(word_count_raw) if word_count_raw else DEFAULT_MIN_WORD_COUNT
        except ValueError:
            min_word_count = DEFAULT_MIN_WORD_COUNT
        if min_word_count < DEFAULT_MIN_WORD_COUNT:
            min_word_count = DEFAULT_MIN_WORD_COUNT

        self._set_status("Generating content...", "blue")
        self._set_button_state("disabled")
        self._set_cancel_button_state("normal")
        self.cancel_flag = False

        target_keywords_raw = self.target_keyword_entry.get().strip()
        reference_text, reference_messages, had_error = self._prepare_reference_context()

        if reference_messages:
            message = "\n".join(reference_messages)
            if had_error:
                messagebox.showwarning("첨부 파일 경고", message, parent=self.root)
                self._set_status("일부 첨부 파일을 활용하지 못했습니다.", "red")
            else:
                messagebox.showinfo("첨부 파일 알림", message, parent=self.root)
                self._set_status("일부 첨부 파일은 길이 제한으로 일부만 사용했습니다.", "orange")
        elif self.attached_files and not reference_text:
            messagebox.showwarning(
                "첨부 파일 경고",
                "첨부한 파일에서 사용할 수 있는 텍스트를 찾지 못했습니다.",
                parent=self.root,
            )
            self._set_status("첨부 파일에서 사용할 내용을 찾지 못했습니다.", "red")

        thread = threading.Thread(
            target=self._generate_and_post,
            args=(
                gemini_key,
                wordpress_user,
                wordpress_secret,
                wordpress_url,
                title,
                prompt,
                auth_method,
                target_keywords_raw,
                reference_text,
                min_word_count,
            ),
            daemon=True,
        )
        self.current_thread = thread
        thread.start()
        self._save_preferences()

    def _generate_and_post(
        self,
        gemini_key: str,
        wordpress_user: str,
        wordpress_secret: str,
        wordpress_url: str,
        title: str,
        prompt: str,
        auth_method: str,
        target_keywords_raw: str,
        reference_text: str,
        min_word_count: int,
    ) -> None:
        try:
            llm = ChatGoogleGenerativeAI(
                model="gemini-2.0-flash",
                api_key=gemini_key,
                convert_system_message_to_human=True,
            )

            prompt_with_refs = prompt
            if reference_text:
                prompt_with_refs = f"{prompt}\n\n# 참고 자료\n{reference_text}"

            base_prompt_with_refs = prompt_with_refs
            min_required_words = max(min_word_count, DEFAULT_MIN_WORD_COUNT)

            content = ""
            metadata: dict = {}
            hashtags: list[str] = []
            post_body = ""
            post_word_count = 0

            for attempt in range(MAX_GENERATION_ATTEMPTS):
                # 취소 플래그 확인
                if self.cancel_flag:
                    self._set_status_async("작업이 취소되었습니다.", "orange")
                    return
                
                content = generate_blog_content(prompt_with_refs, llm)
                
                # 취소 플래그 확인
                if self.cancel_flag:
                    self._set_status_async("작업이 취소되었습니다.", "orange")
                    return
                
                first_overview = content.lower().find('<div class="seo-overview">')
                if first_overview != -1:
                    content = content[first_overview:]
                metadata = extract_rank_math_metadata(content)
                hashtags = extract_hashtags(content)
                post_body = strip_seo_overview(content)
                post_word_count = count_words_from_html(post_body)
                if post_word_count >= min_required_words or attempt == MAX_GENERATION_ATTEMPTS - 1:
                    break
                prompt_with_refs = (
                    f"{base_prompt_with_refs}\n\n"
                    "# 분량 확장 지시\n"
                    f"현재 생성된 본문 분량이 약 {post_word_count}단어입니다. "
                    f"모든 섹션을 구체적인 사례, 데이터, 팁으로 확장하여 최소 {min_required_words}단어 이상이 되도록 다시 작성하세요."
                )

            meta_input = {}

            # 1. LLM이 생성한 포커스 키워드 목록을 가져옵니다. (target_keywords_raw 무시)
            focus_keywords_raw = metadata.get("focus_keywords") or []
            focus_keywords = []  # 최종적으로 Rank Math에 들어갈 리스트

            # 2. LLM이 생성한 목록만 필터링하여 최종 리스트에 추가합니다.
            for keyword in (_sanitize_meta_text(k) for k in focus_keywords_raw):
                if keyword and keyword not in focus_keywords:
                    focus_keywords.append(keyword)

            # 3. (LLM 목록의) 첫 번째 키워드를 Primary Keyword로 지정합니다.
            primary_keyword = focus_keywords[0] if focus_keywords else ""

            # 4. SEO 제목 자동 조정 및 Rank Math 메타 필드 생성
            seo_title = _sanitize_meta_text(metadata.get("seo_title")) or title
            if primary_keyword and primary_keyword.lower() not in seo_title.lower():
                seo_title = f"{seo_title} | {primary_keyword}"
            meta_input["rank_math_title"] = seo_title

            meta_description = _sanitize_meta_text(metadata.get("meta_description"))
            if not meta_description or meta_description.lower() in {"html", "'html", '"html"'}:
                if primary_keyword:
                    meta_description = f"{primary_keyword}에 대한 종합 가이드입니다."
                else:
                    meta_description = f"{title}에 대한 종합 가이드입니다."
            elif primary_keyword and primary_keyword.lower() not in meta_description.lower():
                meta_description = f"{meta_description} {primary_keyword}"
            if len(meta_description) > 160:
                meta_description = meta_description[:157].rstrip() + "..."
            meta_input["rank_math_description"] = meta_description

            if focus_keywords:
                # Rank Math 필드에 LLM이 생성한 키워드 목록을 쉼표로 구분하여 전송
                meta_input["rank_math_focus_keyword"] = ", ".join(focus_keywords)

            if auth_method == "application_password":
                token_raw = f"{wordpress_user}:{wordpress_secret}"
                token_encoded = base64.b64encode(token_raw.encode("utf-8")).decode("ascii")
                headers = {
                    'Authorization': f'Basic {token_encoded}',
                    'Content-Type': 'application/json',
                }
            else:
                headers = {
                    'Authorization': f'Bearer {wordpress_secret}',
                    'Content-Type': 'application/json',
                }

            print("---" * 10)
            print(f"[Debug] 전송할 제목: {seo_title}")
            print(f"[Debug] 전송할 메타데이터(meta_input): {meta_input}")
            print(f"[Debug] 전송할 태그(tag_names): {hashtags}")
            print("---" * 10)

            try:
                response = create_post(
                    wordpress_url,
                    headers,
                    seo_title,
                    post_body,
                    meta_input=meta_input or None,
                    tag_names=hashtags,
                )
            except ConnectionError as e:
                # 네트워크 연결 오류 처리
                self._show_error_async("연결 오류", str(e))
                self._set_status_async("WordPress 서버에 연결할 수 없습니다.", "red")
                return
            except requests.exceptions.ConnectionError as e:
                # requests ConnectionError 직접 처리
                error_msg = str(e)
                if 'getaddrinfo failed' in error_msg or 'NameResolutionError' in error_msg or 'Failed to resolve' in error_msg:
                    from urllib.parse import urlparse
                    try:
                        parsed = urlparse(wordpress_url)
                        host = parsed.netloc or parsed.path.split('/')[0]
                    except:
                        host = "알 수 없음"
                    error_msg = (
                        f"DNS 해석 실패: '{host}' 도메인을 찾을 수 없습니다.\n\n"
                        "가능한 원인:\n"
                        "1. 인터넷 연결을 확인하세요\n"
                        "2. WordPress URL이 올바른지 확인하세요 (예: https://example.com)\n"
                        "3. DNS 설정을 확인하세요\n"
                        "4. 방화벽이나 프록시 설정을 확인하세요\n"
                        "5. 도메인이 실제로 존재하는지 확인하세요"
                    )
                elif 'Max retries exceeded' in error_msg:
                    error_msg = (
                        "연결 실패: 서버에 연결할 수 없습니다.\n\n"
                        "가능한 원인:\n"
                        "1. 서버가 응답하지 않습니다\n"
                        "2. 네트워크 연결이 끊어졌습니다\n"
                        "3. 방화벽이 연결을 차단하고 있습니다"
                    )
                else:
                    error_msg = f"네트워크 연결 오류:\n{error_msg}\n\n가능한 해결 방법:\n1. 인터넷 연결을 확인하세요\n2. WordPress URL을 확인하세요"
                self._show_error_async("연결 오류", error_msg)
                self._set_status_async("WordPress 서버에 연결할 수 없습니다.", "red")
                return
            except requests.exceptions.Timeout as e:
                error_msg = (
                    "연결 시간 초과: 서버에 연결하는데 시간이 너무 오래 걸립니다.\n\n"
                    "가능한 원인:\n"
                    "1. 서버가 응답하지 않습니다\n"
                    "2. 네트워크 연결이 느립니다\n"
                    "3. 방화벽이 연결을 차단하고 있습니다"
                )
                self._show_error_async("연결 시간 초과", error_msg)
                self._set_status_async("연결 시간 초과", "red")
                return
            except Exception as e:
                # 기타 네트워크 오류
                error_msg = f"네트워크 오류 발생:\n{str(e)}\n\n"
                error_msg += "가능한 해결 방법:\n"
                error_msg += "1. 인터넷 연결을 확인하세요\n"
                error_msg += "2. WordPress URL이 올바른지 확인하세요\n"
                error_msg += "3. 방화벽 설정을 확인하세요"
                self._show_error_async("연결 실패", error_msg)
                self._set_status_async("연결 실패", "red")
                return

            if response.status_code == 200:
                self._update_output_async(content)
                self._set_status_async("Blog post draft created successfully!", "green")

            elif response.status_code == 401:
                message = "Unauthorized: Invalid credentials or blocked request."
                if auth_method == "application_password":
                    message += "\n• 사용자명과 Application Password를 다시 확인하세요."
                else:
                    message += "\n• 토큰이 유효한지, JWT 플러그인이 활성화되었는지 확인하세요."

                error_details = f"{message}\n\n[서버 응답 원문]\n{response.text}"
                self._show_error_async("WordPress Error (401)", error_details)
                self._set_status_async("Unauthorized request.", "red")

            elif response.status_code == 400:
                error_details = f"Bad Request (400)\n\n[서버 응답 원문]\n{response.text}"
                self._show_error_async("WordPress Error (400)", error_details)
                self._set_status_async("Bad request. Check the input data.", "red")

            else:
                error_details = (
                    f"Failed to create post: {response.status_code}\n\n"
                    f"[서버 응답 원문]\n{response.text}"
                )
                self._show_error_async("WordPress Error", error_details)
                self._set_status_async("Failed to create post.", "red")
        except ConnectionError as e:
            # 네트워크 연결 오류 처리
            self._show_error_async("연결 오류", str(e))
            self._set_status_async("WordPress 서버에 연결할 수 없습니다.", "red")
        except requests.exceptions.RequestException as e:
            # requests 라이브러리 관련 오류
            error_msg = str(e)
            if "getaddrinfo failed" in error_msg or "NameResolutionError" in error_msg or "Failed to resolve" in error_msg:
                from urllib.parse import urlparse
                parsed = urlparse(wordpress_url)
                host = parsed.netloc or parsed.path.split('/')[0]
                error_msg = (
                    f"DNS 해석 실패: '{host}' 도메인을 찾을 수 없습니다.\n\n"
                    "가능한 원인:\n"
                    "1. 인터넷 연결을 확인하세요\n"
                    "2. WordPress URL이 올바른지 확인하세요 (예: https://example.com)\n"
                    "3. DNS 설정을 확인하세요\n"
                    "4. 방화벽이나 프록시 설정을 확인하세요\n"
                    "5. 도메인이 실제로 존재하는지 확인하세요"
                )
            elif "Max retries exceeded" in error_msg or "timeout" in error_msg.lower():
                error_msg = (
                    "연결 시간 초과: 서버에 연결하는데 시간이 너무 오래 걸립니다.\n\n"
                    "가능한 원인:\n"
                    "1. 서버가 응답하지 않습니다\n"
                    "2. 네트워크 연결이 느립니다\n"
                    "3. 방화벽이 연결을 차단하고 있습니다\n"
                    "4. WordPress 서버가 다운되었을 수 있습니다"
                )
            else:
                error_msg = f"네트워크 오류 발생:\n{error_msg}\n\n가능한 해결 방법:\n1. 인터넷 연결을 확인하세요\n2. WordPress URL을 확인하세요"
            self._show_error_async("연결 실패", error_msg)
            self._set_status_async("연결 실패", "red")
        except Exception as exc:
            error_msg = f"오류 발생: {str(exc)}"
            # DNS 관련 오류 체크
            if "getaddrinfo failed" in str(exc) or "NameResolutionError" in str(exc) or "Failed to resolve" in str(exc):
                from urllib.parse import urlparse
                try:
                    parsed = urlparse(wordpress_url)
                    host = parsed.netloc or parsed.path.split('/')[0]
                except:
                    host = "알 수 없음"
                error_msg = (
                    f"DNS 해석 실패: '{host}' 도메인을 찾을 수 없습니다.\n\n"
                    "가능한 원인:\n"
                    "1. 인터넷 연결을 확인하세요\n"
                    "2. WordPress URL이 올바른지 확인하세요 (예: https://example.com)\n"
                    "3. DNS 설정을 확인하세요\n"
                    "4. 방화벽이나 프록시 설정을 확인하세요\n"
                    "5. 도메인이 실제로 존재하는지 확인하세요"
                )
            self._show_error_async("Error", error_msg)
            self._set_status_async("An error occurred while processing the request.", "red")
        finally:
            self._set_button_state_async("normal")
            self._set_cancel_button_state_async("disabled")
            self.cancel_flag = False
            self.current_thread = None

    def _on_fetch_clicked(self) -> None:
        """WordPress에서 포스트를 가져오는 기능"""
        # 환경 변수 우선, 없으면 GUI 입력값 사용
        wordpress_user = os.getenv("WORDPRESS_USERNAME", "") or self.wordpress_user_entry.get().strip()
        wordpress_secret = os.getenv("WORDPRESS_SECRET", "") or self.wordpress_secret_entry.get().strip()
        wordpress_url = os.getenv("WORDPRESS_URL", "") or self.wordpress_url_entry.get().strip()
        auth_method = os.getenv("WORDPRESS_AUTH_METHOD", "") or self.auth_method_var.get()

        required_fields = [wordpress_secret, wordpress_url]
        if auth_method == "application_password":
            required_fields.append(wordpress_user)

        if not all(required_fields):
            messagebox.showwarning("Missing information", "WordPress URL, 인증 정보를 입력해주세요.")
            return

        # 인증 헤더 생성
        if auth_method == "application_password":
            token_raw = f"{wordpress_user}:{wordpress_secret}"
            token_encoded = base64.b64encode(token_raw.encode("utf-8")).decode("ascii")
            headers = {
                'Authorization': f'Basic {token_encoded}',
                'Content-Type': 'application/json',
            }
        else:
            headers = {
                'Authorization': f'Bearer {wordpress_secret}',
                'Content-Type': 'application/json',
            }

        self._set_status("WordPress 포스트 가져오는 중...", "blue")
        threading.Thread(
            target=self._fetch_and_display_posts,
            args=(wordpress_url, headers),
            daemon=True,
        ).start()

    def _fetch_and_display_posts(self, wp_url: str, headers: dict) -> None:
        """WordPress 포스트를 가져와서 다이얼로그에 표시"""
        try:
            response = fetch_posts(wp_url, headers, per_page=50, status='any')
            
            if not response or response.status_code != 200:
                error_msg = "포스트를 가져오는데 실패했습니다."
                if response:
                    error_msg += f"\n상태 코드: {response.status_code}\n{response.text[:200]}"
                self._show_error_async("가져오기 실패", error_msg)
                self._set_status_async("포스트 가져오기 실패", "red")
                return

            posts = response.json()
            if not posts:
                self._show_info_async("알림", "가져올 포스트가 없습니다.")
                self._set_status_async("포스트가 없습니다", "orange")
                return

            # 포스트 선택 다이얼로그 표시
            self.root.after(0, self._show_post_selection_dialog, posts, wp_url, headers)
            self._set_status_async(f"{len(posts)}개의 포스트를 찾았습니다.", "green")

        except ConnectionError as e:
            # 네트워크 연결 오류 처리
            self._show_error_async("연결 오류", str(e))
            self._set_status_async("WordPress 서버에 연결할 수 없습니다.", "red")
        except Exception as exc:
            error_msg = f"포스트 가져오기 중 오류 발생: {str(exc)}"
            if "getaddrinfo failed" in str(exc) or "NameResolutionError" in str(exc):
                error_msg = (
                    "DNS 해석 실패: 도메인을 찾을 수 없습니다.\n\n"
                    "가능한 원인:\n"
                    "1. 인터넷 연결을 확인하세요\n"
                    "2. WordPress URL이 올바른지 확인하세요 (예: https://example.com)\n"
                    "3. DNS 설정을 확인하세요\n"
                    "4. 방화벽이나 프록시 설정을 확인하세요"
                )
            self._show_error_async("에러", error_msg)
            self._set_status_async("오류 발생", "red")

    def _show_post_selection_dialog(self, posts: list, wp_url: str, headers: dict) -> None:
        """포스트 선택 다이얼로그"""
        dialog = tk.Toplevel(self.root)
        dialog.title("WordPress 포스트 선택")
        dialog.geometry("800x600")
        dialog.transient(self.root)
        dialog.grab_set()

        # 포스트 목록
        list_frame = tk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("맑은 고딕", 10))
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        # 포스트 정보 저장
        post_data = {}
        for post in posts:
            post_id = post.get('id', 'N/A')
            title = post.get('title', {}).get('rendered', '제목 없음')
            status = post.get('status', 'unknown')
            # 연도 없이 월-일 형식으로 표시
            if post.get('date'):
                date_str = post.get('date', '')
                if len(date_str) >= 10:
                    # YYYY-MM-DD 형식에서 MM-DD만 추출
                    date = date_str[5:10]
                else:
                    date = '날짜 없음'
            else:
                date = '날짜 없음'
            display_text = f"[{post_id}] {title} ({status}) - {date}"
            listbox.insert("end", display_text)
            post_data[display_text] = post

        # 버튼 프레임
        button_frame = tk.Frame(dialog)
        button_frame.pack(fill="x", padx=10, pady=10)

        def load_selected_post():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("선택 필요", "포스트를 선택해주세요.", parent=dialog)
                return

            selected_text = listbox.get(selection[0])
            post = post_data[selected_text]
            
            # 선택한 포스트의 내용을 GUI에 로드
            self._load_post_to_gui(post)
            dialog.destroy()

        def refresh_posts():
            dialog.destroy()
            self._fetch_and_display_posts(wp_url, headers)

        tk.Button(button_frame, text="선택한 포스트 가져오기", command=load_selected_post).pack(side="left", padx=5)
        tk.Button(button_frame, text="새로고침", command=refresh_posts).pack(side="left", padx=5)
        tk.Button(button_frame, text="취소", command=dialog.destroy).pack(side="right", padx=5)

        # 더블클릭으로도 로드 가능
        listbox.bind("<Double-Button-1>", lambda e: load_selected_post())

    def _load_post_to_gui(self, post: dict) -> None:
        """가져온 포스트를 GUI에 로드"""
        try:
            # 제목
            title = post.get('title', {}).get('rendered', '')
            if title:
                self.title_entry.delete(0, "end")
                self.title_entry.insert(0, html.unescape(title))

            # 내용
            content = post.get('content', {}).get('rendered', '')
            if content:
                self.output_text.configure(state="normal")
                self.output_text.delete("1.0", "end")
                self.output_text.insert("1.0", html.unescape(content))
                self.output_text.configure(state="disabled")

            # 메타데이터에서 Rank Math 정보 추출
            meta = post.get('meta', {})
            if meta:
                rank_math_title = meta.get('rank_math_title', '')
                rank_math_description = meta.get('rank_math_description', '')
                rank_math_focus_keyword = meta.get('rank_math_focus_keyword', '')
                
                if rank_math_focus_keyword:
                    self.target_keyword_entry.delete(0, "end")
                    # 쉼표로 구분된 키워드 중 첫 번째 사용
                    first_keyword = rank_math_focus_keyword.split(',')[0].strip()
                    self.target_keyword_entry.insert(0, first_keyword)

            # 태그
            tags = post.get('tags', [])
            if tags:
                # 태그 정보는 나중에 사용할 수 있도록 저장
                pass

            self._set_status("포스트를 성공적으로 가져왔습니다.", "green")
            messagebox.showinfo("성공", "포스트를 가져왔습니다. 내용을 확인하고 필요시 수정하세요.", parent=self.root)

        except Exception as exc:
            self._show_error_async("에러", f"포스트 로드 중 오류 발생: {str(exc)}")

    def _show_info_async(self, title: str, message: str) -> None:
        """비동기적으로 정보 메시지 표시"""
        self.root.after(0, messagebox.showinfo, title, message)

    def _on_build_prompt_clicked(self) -> None:
        topic = self.topic_entry.get().strip()
        keyword = self.target_keyword_entry.get().strip()
        word_count_raw = self.word_count_entry.get().strip()

        if not topic or not keyword:
            messagebox.showwarning("Missing information", "Please provide both the blog topic and target keyword.")
            return

        try:
            word_count = int(word_count_raw) if word_count_raw else 1500
            if word_count <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid word count", "Word count must be a positive integer. Using default 1500.")
            word_count = 1500
            self.word_count_entry.delete(0, "end")
            self.word_count_entry.insert(0, str(word_count))

        prompt = create_seo_blog_prompt(topic, keyword, word_count)
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", prompt)
        self._save_preferences()

    def _update_output_async(self, content: str) -> None:
        self.root.after(0, self._update_output, content)

    def _update_output(self, content: str) -> None:
        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.insert("1.0", content)
        self.output_text.configure(state="disabled")

    def _set_status(self, message: str, color: str = "blue") -> None:
        self.status_label.config(text=message, fg=color)

    def _set_status_async(self, message: str, color: str = "blue") -> None:
        self.root.after(0, self._set_status, message, color)

    def _show_error_async(self, title: str, message: str) -> None:
        self.root.after(0, messagebox.showerror, title, message)

    def _set_button_state(self, state: str) -> None:
        self.generate_button.config(state=state)

    def _set_button_state_async(self, state: str) -> None:
        self.root.after(0, self._set_button_state, state)

    def _set_cancel_button_state(self, state: str) -> None:
        self.cancel_button.config(state=state)

    def _set_cancel_button_state_async(self, state: str) -> None:
        self.root.after(0, self._set_cancel_button_state, state)

    def _on_cancel_clicked(self) -> None:
        """취소 버튼 클릭 핸들러"""
        if self.cancel_flag:
            return  # 이미 취소 중
        
        result = messagebox.askyesno(
            "작업 취소",
            "진행 중인 작업을 취소하시겠습니까?",
            parent=self.root
        )
        
        if result:
            self.cancel_flag = True
            self._set_status("작업 취소 중...", "orange")
            self._set_cancel_button_state("disabled")

    def _load_preferences(self) -> None:
        # 환경 변수에서 민감 정보 로드 (GUI에는 표시하지 않음)
        env_gemini_key = os.getenv("GEMINI_API_KEY", "")
        env_wp_url = os.getenv("WORDPRESS_URL", "")
        env_wp_user = os.getenv("WORDPRESS_USERNAME", "")
        env_wp_secret = os.getenv("WORDPRESS_SECRET", "")
        env_auth_method = os.getenv("WORDPRESS_AUTH_METHOD", "")
        
        # 설정 파일에서 비민감 정보 로드
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                data = {}
        else:
            data = {}

        def _insert(entry: tk.Entry, value: str, *, default: str = "") -> None:
            entry.delete(0, "end")
            entry.insert(0, value if value else default)

        # WordPress URL은 환경 변수나 설정 파일에서 로드
        _insert(self.wordpress_url_entry, env_wp_url or data.get("wordpress_url", ""))
        
        # API 키와 토큰: 환경 변수가 있으면 GUI에는 표시하지 않음 (보안)
        # 환경 변수가 없을 때만 설정 파일에서 로드
        if env_gemini_key:
            # 환경 변수에서 읽었지만 GUI에는 표시하지 않음 (보안)
            self.gemini_key_entry.delete(0, "end")
        elif data.get("gemini_key"):
            _insert(self.gemini_key_entry, data.get("gemini_key", ""))
        else:
            self.gemini_key_entry.delete(0, "end")
        
        if env_wp_user:
            # 환경 변수에서 읽었지만 GUI에는 표시하지 않음 (보안)
            self.wordpress_user_entry.delete(0, "end")
        elif data.get("wordpress_user"):
            _insert(self.wordpress_user_entry, data.get("wordpress_user", ""))
        else:
            self.wordpress_user_entry.delete(0, "end")
        
        if env_wp_secret:
            # 환경 변수에서 읽었지만 GUI에는 표시하지 않음 (보안)
            self.wordpress_secret_entry.delete(0, "end")
        elif data.get("wordpress_secret"):
            _insert(self.wordpress_secret_entry, data.get("wordpress_secret", ""))
        else:
            self.wordpress_secret_entry.delete(0, "end")
        
        _insert(self.title_entry, data.get("title", ""))
        _insert(self.topic_entry, data.get("topic", ""))
        _insert(self.target_keyword_entry, data.get("target_keyword", ""))
        _insert(self.word_count_entry, data.get("word_count", "1500"), default="1500")
        prompt = data.get("prompt", "")
        if prompt:
            self.prompt_text.delete("1.0", "end")
            self.prompt_text.insert("1.0", prompt)
        
        # 인증 방식: 환경 변수 우선
        if env_auth_method in {"application_password", "bearer"}:
            self.auth_method_var.set(env_auth_method)
        elif data.get("auth_method") in {"application_password", "bearer"}:
            self.auth_method_var.set(data.get("auth_method"))

    def _save_preferences(self) -> None:
        # 민감 정보는 저장하지 않음 (환경 변수에서만 관리)
        data = {
            "wordpress_url": self.wordpress_url_entry.get().strip(),
            # gemini_key, wordpress_user, wordpress_secret은 저장하지 않음 (보안)
            "title": self.title_entry.get().strip(),
            "topic": self.topic_entry.get().strip(),
            "target_keyword": self.target_keyword_entry.get().strip(),
            "word_count": self.word_count_entry.get().strip() or "1500",
            "prompt": self.prompt_text.get("1.0", "end").strip(),
            "auth_method": self.auth_method_var.get(),
        }
        try:
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

def main() -> None:
    root = tk.Tk()
    WordPressBotGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
