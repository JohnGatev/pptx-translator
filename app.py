from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import json
import re
import time
import zipfile
from typing import Dict, Iterable, List, Sequence, Tuple
import xml.etree.ElementTree as ET

import requests
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_AUTO_SIZE


DEFAULT_BASE_URL = "https://ai-research-proxy.azurewebsites.net"
DEFAULT_MODEL = "gpt-oss-120b"
DEFAULT_TEMPERATURE = 0.2
DEFAULT_TIMEOUT = 120
LEXICON_PDF_NAME = "Finance 1 Lexicon ENG NL.pdf"



@dataclass(frozen=True)
class TranslationConfig:
    base_url: str
    model: str
    target_language: str
    temperature: float
    timeout: int
    api_key: str


@dataclass(frozen=True)
class LexiconEntry:
    term: str
    translation: str
    raw: str


@dataclass(frozen=True)
class LexiconStore:
    entries: List[LexiconEntry]
    token_index: Dict[str, List[int]]


def normalize_base_url(base_url: str) -> str:
    return base_url.rstrip("/")


def load_finance_lexicon() -> LexiconStore | None:
    lexicon_path = Path(__file__).parent / LEXICON_PDF_NAME
    if not lexicon_path.exists():
        return None
    try:
        from pypdf import PdfReader
    except ImportError:
        return None

    reader = PdfReader(str(lexicon_path))
    lines: List[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        lines.extend(text.splitlines())

    entries: List[LexiconEntry] = []
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if re.fullmatch(r"(english|engels)\s*(dutch|nederlands)", line, re.IGNORECASE):
            continue
        # prefer split on multi-space columns
        if re.search(r"\s{2,}", raw_line):
            parts = re.split(r"\s{2,}", raw_line.strip())
        else:
            parts = re.split(r"\s*(?:=>|->|–|—|-|:|\t)\s*", line, maxsplit=1)
        if len(parts) >= 2:
            english = parts[0].strip()
            dutch = parts[1].strip()
            if english and dutch:
                raw = f"{english} => {dutch}"
                entries.append(LexiconEntry(term=english, translation=dutch, raw=raw))
        else:
            entries.append(LexiconEntry(term=line, translation="", raw=line))

    if not entries:
        return None
    return build_lexicon_store(entries)


def build_lexicon_store(entries: List[LexiconEntry]) -> LexiconStore:
    token_index: Dict[str, List[int]] = {}
    for idx, entry in enumerate(entries):
        for token in tokenize(entry.term):
            token_index.setdefault(token, []).append(idx)
    return LexiconStore(entries=entries, token_index=token_index)


def tokenize(text: str) -> List[str]:
    return re.findall(r"[A-Za-zÀ-ÿ0-9]+", text.lower())


def build_rag_context(texts: Iterable[str], lexicon: LexiconStore, max_entries: int = 30) -> str:
    combined_text = re.sub(r"\[\[\[RUN_(?:\d+|END)\]\]\]", " ", " ".join(texts)).lower()
    doc_tokens = set(tokenize(combined_text))
    scores: Dict[int, int] = {}
    for token in doc_tokens:
        for idx in lexicon.token_index.get(token, []):
            scores[idx] = scores.get(idx, 0) + 1
    for idx, entry in enumerate(lexicon.entries):
        if entry.term.lower() in combined_text:
            scores[idx] = scores.get(idx, 0) + 5

    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    selected = [lexicon.entries[idx] for idx, _ in ranked[:max_entries]]
    if not selected:
        return ""
    lines = []
    for entry in selected:
        if entry.translation:
            lines.append(f"- {entry.term} => {entry.translation}")
        else:
            lines.append(f"- {entry.raw}")
    return "\n".join(lines)


def rag_instructions(rag_context: str | None) -> str:
    if not rag_context:
        return ""
    return (
        "\n\nFinance lexicon context. Use these preferred translations when relevant "
        "(including close variations). Do not output this list:\n"
        f"{rag_context}\n"
    )


def split_whitespace(text: str) -> tuple[str, str, str]:
    match = re.match(r"^(\s*)(.*?)(\s*)$", text, re.DOTALL)
    if not match:
        return "", text, ""
    return match.group(1), match.group(2), match.group(3)


def build_translation_messages(
    text: str,
    config: TranslationConfig,
    rag_context: str | None = None,
) -> List[Dict[str, str]]:
    rag_note = rag_instructions(rag_context)
    return [
        {
            "role": "system",
            "content": (
                "You are a professional translation engine. Translate the user's text "
                f"to {config.target_language}. Preserve meaning, punctuation, and line "
                "breaks. The text may include placeholders like [[[RUN_0]]]. "
                "Keep all placeholders unchanged and in the same order. Only translate the "
                "text between placeholders. Do not add commentary or quotes."
                f"{rag_note}"
            ),
        },
        {"role": "user", "content": text},
    ]


def request_chat_completion(messages: List[Dict[str, str]], config: TranslationConfig) -> str:
    base_url = normalize_base_url(config.base_url)
    payload = {
        "model": config.model,
        "messages": messages,
        "temperature": config.temperature,
    }
    headers = {
        "Authorization": f"Bearer {config.api_key}",
        "Content-Type": "application/json",
    }
    response = requests.post(
        f"{base_url}/v1/chat/completions",
        json=payload,
        headers=headers,
        timeout=config.timeout,
    )
    response.raise_for_status()
    data = response.json()
    content = data.get("choices", [{}])[0].get("message", {}).get("content")
    if content is None:
        raise RuntimeError("Translation response did not contain translated text.")
    return content


def translate_text(text: str, config: TranslationConfig, rag_context: str | None = None) -> str:
    return request_chat_completion(build_translation_messages(text, config, rag_context), config)


def cached_translate(text: str, config: TranslationConfig, rag_context: str | None = None) -> str:
    return translate_text(text, config, rag_context)


@dataclass
class TranslationSegment:
    segment_id: str
    text: str


@dataclass
class ParagraphInfo:
    runs: List
    prefixes: List[str]
    cores: List[str]
    suffixes: List[str]
    translate_mask: List[bool]
    shape: object
    original_length: int
    is_table: bool
    slide_index: int


def is_numeric_text(text: str) -> bool:
    return bool(re.fullmatch(r"[\d\s.,:%€$£¥+-]+", text.strip()))


def is_math_text(text: str) -> bool:
    if not text:
        return False
    math_symbols = set("=<>±√∑∫∏≈≠≤≥→←⇒⇔∞πσΔΩβμτ")
    if any(symbol in text for symbol in math_symbols):
        return True
    if re.search(r"[A-Za-z]\s*=\s*[\d]", text):
        return True
    if re.search(r"[A-Za-z]\s*[_^]", text):
        return True
    return False


def is_off_canvas(shape, slide_width: int, slide_height: int) -> bool:
    left = getattr(shape, "left", 0)
    top = getattr(shape, "top", 0)
    width = getattr(shape, "width", 0)
    height = getattr(shape, "height", 0)
    if width <= 0 or height <= 0:
        return True
    right = left + width
    bottom = top + height
    return right < 0 or bottom < 0 or left > slide_width or top > slide_height


def should_skip_shape(shape, slide_width: int, slide_height: int, seen_keys: set) -> bool:
    if getattr(shape, "visible", True) is False:
        return True
    if is_off_canvas(shape, slide_width, slide_height):
        return True
    if shape.is_placeholder:
        placeholder_type = shape.placeholder_format.type
        if placeholder_type in {
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        }:
            return True
    text = extract_shape_text(shape)
    if not text.strip():
        return False
    left = int(getattr(shape, "left", 0))
    top = int(getattr(shape, "top", 0))
    width = int(getattr(shape, "width", 0))
    height = int(getattr(shape, "height", 0))
    key = (left, top, width, height, text.strip())
    if key in seen_keys:
        return True
    seen_keys.add(key)
    return False


def extract_shape_text(shape) -> str:
    if shape.has_text_frame:
        return shape.text_frame.text or ""
    if shape.has_table:
        parts = []
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text or "")
        return "\n".join(parts)
    return ""


def iter_slide_shapes(shapes) -> Iterable:
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_slide_shapes(shape.shapes)
        else:
            yield shape


def iter_paragraph_runs(shape) -> Iterable[Tuple[List, bool]]:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            yield from iter_paragraph_runs(child)
        return

    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if paragraph.runs:
                yield list(paragraph.runs), False

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    if paragraph.runs:
                        yield list(paragraph.runs), True


def collect_paragraphs(presentation: Presentation) -> List[ParagraphInfo]:
    paragraphs: List[ParagraphInfo] = []
    slide_width = presentation.slide_width
    slide_height = presentation.slide_height
    for slide_index, slide in enumerate(presentation.slides):
        seen_keys: set = set()
        for shape in iter_slide_shapes(slide.shapes):
            if should_skip_shape(shape, slide_width, slide_height, seen_keys):
                continue
            for runs, is_table in iter_paragraph_runs(shape):
                prefixes = []
                cores = []
                suffixes = []
                translate_mask = []
                original_length = 0
                for run in runs:
                    prefix, core, suffix = split_whitespace(run.text or "")
                    can_translate = bool(core)
                    if core:
                        font_name = getattr(getattr(run, "font", None), "name", None)
                        if font_name and "math" in font_name.lower():
                            can_translate = False
                        if is_math_text(core) or is_numeric_text(core):
                            can_translate = False
                    prefixes.append(prefix)
                    cores.append(core)
                    suffixes.append(suffix)
                    translate_mask.append(can_translate)
                    if can_translate:
                        original_length += len(core)
                if any(translate_mask):
                    paragraphs.append(
                        ParagraphInfo(
                            runs=runs,
                            prefixes=prefixes,
                            cores=cores,
                            suffixes=suffixes,
                            translate_mask=translate_mask,
                            shape=shape,
                            original_length=original_length,
                            is_table=is_table,
                            slide_index=slide_index,
                        )
                    )
    return paragraphs


def build_segments_json(segments: Sequence[TranslationSegment]) -> str:
    payload = {"segments": [{"id": seg.segment_id, "text": seg.text} for seg in segments]}
    return json.dumps(payload, ensure_ascii=False)


def parse_segments_json(raw: str) -> List[Dict[str, str]]:
    raw = raw.strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw, re.DOTALL)
        if not match:
            match = re.search(r"\[.*\]", raw, re.DOTALL)
            if not match:
                raise
            data = json.loads(match.group(0))
        else:
            data = json.loads(match.group(0))
    if isinstance(data, list):
        return data
    segments = data.get("segments")
    if not isinstance(segments, list):
        raise ValueError("Invalid translation response format.")
    return segments


def translate_segments(
    segments: Sequence[TranslationSegment],
    config: TranslationConfig,
    rag_context: str | None = None,
) -> Dict[str, str]:
    rag_note = rag_instructions(rag_context)
    system_prompt = (
        "You are a professional translation engine. Translate the provided JSON segments "
        f"to {config.target_language}. Return ONLY valid JSON with the same 'segments' array. "
        "Each item must have the same 'id' and the translated 'text'. Preserve all placeholders "
        "like [[[RUN_0]]], keep them unchanged and in the same order. "
        "Do not add commentary."
        f"{rag_note}"
    )
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": build_segments_json(segments)},
    ]
    content = request_chat_completion(messages, config)
    parsed = parse_segments_json(content)
    results: Dict[str, str] = {}
    for item in parsed:
        segment_id = str(item.get("id", ""))
        text = item.get("text")
        if not segment_id or text is None:
            raise ValueError("Translation output missing id/text.")
        results[segment_id] = text
    if len(results) != len(segments):
        raise ValueError("Translation output segment count mismatch.")
    return results


def translate_segments_cached(
    segments: Tuple[Tuple[str, str], ...],
    config: TranslationConfig,
    rag_context: str | None = None,
) -> Dict[str, str]:
    segment_objs = [TranslationSegment(segment_id=seg_id, text=text) for seg_id, text in segments]
    return translate_segments(segment_objs, config, rag_context)


def batch_segments(segments: Sequence[TranslationSegment], max_chars: int = 4000, max_items: int = 20):
    batch: List[TranslationSegment] = []
    total_chars = 0
    for segment in segments:
        segment_len = len(segment.text)
        if batch and (len(batch) >= max_items or total_chars + segment_len > max_chars):
            yield batch
            batch = []
            total_chars = 0
        batch.append(segment)
        total_chars += segment_len
    if batch:
        yield batch


def batch_limits() -> Tuple[int, int]:
    return 30000, 120


def count_translation_batches(paragraphs: List[ParagraphInfo]) -> int:
    grouped: Dict[int, List[ParagraphInfo]] = {}
    for paragraph in paragraphs:
        grouped.setdefault(paragraph.slide_index, []).append(paragraph)
    max_chars, max_items = batch_limits()
    total_batches = 0
    for slide_paragraphs in grouped.values():
        segments = []
        for idx, paragraph in enumerate(slide_paragraphs):
            translatable_cores = [
                core for core, can_translate in zip(paragraph.cores, paragraph.translate_mask) if can_translate
            ]
            if not translatable_cores:
                continue
            segments.append(
                TranslationSegment(segment_id=str(idx), text=build_segmented_text(translatable_cores))
            )
        total_batches += sum(1 for _ in batch_segments(segments, max_chars=max_chars, max_items=max_items))
    return total_batches


def build_segmented_text(parts: Sequence[str]) -> str:
    chunks = []
    for idx, part in enumerate(parts):
        chunks.append(f"[[[RUN_{idx}]]]")
        chunks.append(part)
    chunks.append("[[[RUN_END]]]")
    return "".join(chunks)


def split_segmented_text(text: str, count: int) -> List[str]:
    pattern = re.compile(r"\[\[\[RUN_(\d+|END)\]\]\]")
    matches = list(pattern.finditer(text))
    if not matches:
        raise ValueError("No run markers found in translation.")
    segments = {}
    for index, match in enumerate(matches):
        token = match.group(1)
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        if token == "END":
            continue
        segments[int(token)] = text[start:end]
    if len(segments) != count:
        raise ValueError("Run marker count mismatch after translation.")
    return [segments[i] for i in range(count)]


@dataclass
class ProgressTracker:
    total: int
    container: object
    completed: int = 0

    def advance(self) -> None:
        self.completed += 1
        fraction = self.completed / self.total if self.total else 1.0
        render_progress(self.container, fraction)


def translate_paragraphs(
    paragraphs: List[ParagraphInfo],
    config: TranslationConfig,
    tracker: ProgressTracker | None = None,
    lexicon: LexiconStore | None = None,
) -> None:
    grouped: Dict[int, List[ParagraphInfo]] = {}
    for paragraph in paragraphs:
        grouped.setdefault(paragraph.slide_index, []).append(paragraph)

    max_chars, max_items = batch_limits()

    for slide_index in sorted(grouped.keys()):
        slide_paragraphs = grouped[slide_index]
        paragraph_maps: List[List[int | None]] = []
        segments: List[TranslationSegment] = []

        for idx, paragraph in enumerate(slide_paragraphs):
            run_map: List[int | None] = []
            translatable_cores: List[str] = []
            for core, can_translate in zip(paragraph.cores, paragraph.translate_mask):
                if can_translate:
                    run_map.append(len(translatable_cores))
                    translatable_cores.append(core)
                else:
                    run_map.append(None)
            segmented = build_segmented_text(translatable_cores)
            segments.append(TranslationSegment(segment_id=str(idx), text=segmented))
            paragraph_maps.append(run_map)

        translated_map: Dict[str, str] = {}
        for batch in batch_segments(segments, max_chars=max_chars, max_items=max_items):
            batch_key = tuple((seg.segment_id, seg.text) for seg in batch)
            rag_context = build_rag_context([seg.text for seg in batch], lexicon) if lexicon else None
            try:
                translated_map.update(translate_segments_cached(batch_key, config, rag_context))
            except Exception:
                for seg in batch:
                    translated_map[seg.segment_id] = cached_translate(seg.text, config, rag_context)

        for idx, paragraph in enumerate(slide_paragraphs):
            translated = translated_map.get(str(idx), "")
            run_map = paragraph_maps[idx]
            translatable_count = sum(1 for item in run_map if item is not None)
            try:
                translated_parts = split_segmented_text(translated, translatable_count)
            except Exception:
                translated_parts = []
                for core, can_translate in zip(paragraph.cores, paragraph.translate_mask):
                    if can_translate:
                        translated_parts.append(cached_translate(core, config))

            translated_length = 0
            for run, prefix, suffix, core, map_index, can_translate in zip(
                paragraph.runs,
                paragraph.prefixes,
                paragraph.suffixes,
                paragraph.cores,
                run_map,
                paragraph.translate_mask,
            ):
                if not can_translate:
                    run.text = f"{prefix}{core}{suffix}"
                    continue
                if map_index is not None and map_index < len(translated_parts):
                    translated_core = translated_parts[map_index]
                else:
                    translated_core = core
                translated_length += len(translated_core)
                run.text = f"{prefix}{translated_core}{suffix}"

            if not paragraph.is_table:
                if paragraph.original_length and translated_length > paragraph.original_length * 1.15:
                    text_frame = paragraph.shape.text_frame
                    text_frame.word_wrap = True
                    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            if tracker:
                tracker.advance()


def translate_xml_text_nodes(
    xml_bytes: bytes,
    config: TranslationConfig,
    lexicon: LexiconStore | None = None,
) -> bytes:
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    root = tree.getroot()
    text_nodes = []
    for elem in root.iter():
        if elem.tag.endswith("}t"):
            text = elem.text or ""
            if not text.strip():
                continue
            if is_numeric_text(text) or is_math_text(text):
                continue
            text_nodes.append(elem)

    if not text_nodes:
        return xml_bytes

    segments = [TranslationSegment(segment_id=str(i), text=node.text or "") for i, node in enumerate(text_nodes)]
    translated_map: Dict[str, str] = {}
    max_chars, max_items = batch_limits()
    for batch in batch_segments(segments, max_chars=max_chars, max_items=max_items):
        batch_key = tuple((seg.segment_id, seg.text) for seg in batch)
        rag_context = build_rag_context([seg.text for seg in batch], lexicon) if lexicon else None
        try:
            translated_map.update(translate_segments_cached(batch_key, config, rag_context))
        except Exception:
            for seg in batch:
                translated_map[seg.segment_id] = cached_translate(seg.text, config, rag_context)

    for idx, node in enumerate(text_nodes):
        translated = translated_map.get(str(idx))
        if translated:
            node.text = translated
    output = BytesIO()
    tree.write(output, encoding="utf-8", xml_declaration=True)
    return output.getvalue()


def count_xml_parts(pptx_bytes: bytes) -> int:
    with zipfile.ZipFile(BytesIO(pptx_bytes), "r") as zin:
        return sum(
            1
            for item in zin.infolist()
            if item.filename.endswith(".xml")
            and (
                item.filename.startswith("ppt/charts/")
                or item.filename.startswith("ppt/diagrams/")
                or item.filename.startswith("ppt/diagramData/")
                or item.filename.startswith("ppt/diagramLayout/")
                or item.filename.startswith("ppt/diagramStyles/")
            )
        )


def translate_xml_parts(
    pptx_bytes: bytes,
    config: TranslationConfig,
    tracker: ProgressTracker | None = None,
    lexicon: LexiconStore | None = None,
) -> bytes:
    input_buffer = BytesIO(pptx_bytes)
    output_buffer = BytesIO()
    with zipfile.ZipFile(input_buffer, "r") as zin, zipfile.ZipFile(
        output_buffer, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and (
                item.filename.startswith("ppt/charts/")
                or item.filename.startswith("ppt/diagrams/")
                or item.filename.startswith("ppt/diagramData/")
                or item.filename.startswith("ppt/diagramLayout/")
                or item.filename.startswith("ppt/diagramStyles/")
            ):
                try:
                    data = translate_xml_text_nodes(data, config, lexicon)
                except Exception:
                    pass
                if tracker:
                    tracker.advance()
            zout.writestr(item, data)
    output_buffer.seek(0)
    return output_buffer.read()


def set_status(container, message: str) -> None:
    container.markdown(
        "<div style='border:1px solid #a8a29f; border-left:4px solid #bc0031; "
        "padding:12px; background:#ffffff; color:#1B1918;'>"
        f"{message}"
        "</div>",
        unsafe_allow_html=True,
    )


def render_progress(container, fraction: float) -> None:
    percent = max(0.0, min(1.0, fraction)) * 100
    container.markdown(
        f"<div class='progress-track'><div class='progress-fill' style='width:{percent:.2f}%;'></div></div>",
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="PPTX Translator", page_icon="🈯", layout="wide")
    st.title("PPTX Translator")
    st.markdown(
        """
        <style>
        .stApp { background-color: #ffffff; color: #1B1918; }
        .block-container { padding-top: 2rem; max-width: 100% !important; padding-left: 2rem; padding-right: 2rem; }
        [data-testid="stSidebar"] { background-color: #f2f2f2; }
        h1, h2, h3, h4, h5, h6 { color: #1B1918; }
        * { border-radius: 0 !important; }
        .stButton > button {
            background-color: #bc0031;
            color: #ffffff;
            border-radius: 0;
            border: none;
        }
        .stButton > button:hover { background-color: #a1002a; color: #ffffff; }
        .stTextInput > div > div > input,
        .stFileUploader,
        .stProgress > div > div,
        [data-testid="stAlert"],
        .stTextInput,
        .stSelectbox,
        .stRadio {
            border-radius: 0 !important;
        }
        [data-testid="stAlert"] {
            background: #ffffff !important;
            border: 1px solid #a8a29f !important;
            border-left: 4px solid #bc0031 !important;
            color: #1B1918 !important;
        }
        [data-testid="stAlert"] svg { color: #bc0031 !important; }
        [data-testid="stAlert"] [data-testid="stMarkdownContainer"] { color: #1B1918 !important; }
        .progress-track {
            width: 100%;
            height: 14px;
            background: #e5e5e5;
            border: 1px solid #a8a29f;
        }
        .progress-fill {
            height: 100%;
            background: #bc0031;
            width: 0%;
        }
        div[data-testid="stNotification"] { display: none; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        "Upload a PowerPoint deck, translate the text slide-by-slide, "
        "and download a PPTX with the original layout preserved."
    )

    with st.sidebar:
        st.subheader("Translation Settings")
        target_language = st.text_input("Target language", value="Dutch")
        api_key = st.text_input("API key", type="password")
        st.caption("Uses LiteLLM endpoint with model gpt-oss-120b.")

    uploaded_file = st.file_uploader("Drop your PPTX file here", type=["pptx"])
    if not uploaded_file:
        st.markdown(
            "<div style='border:1px solid #a8a29f; border-left:4px solid #bc0031; "
            "padding:12px; background:#ffffff; color:#1B1918;'>"
            "Upload a .pptx file to get started."
            "</div>",
            unsafe_allow_html=True,
        )
        return

    lexicon_store = None
    if target_language.strip().lower() in {"dutch", "nederlands", "nl"}:
        lexicon_store = load_finance_lexicon()
        if not lexicon_store:
            set_status(
                st.empty(),
                f"Finance lexicon PDF not found or unreadable: {LEXICON_PDF_NAME}.",
            )
            return
    config = TranslationConfig(
        base_url=DEFAULT_BASE_URL,
        model=DEFAULT_MODEL,
        target_language=target_language,
        temperature=DEFAULT_TEMPERATURE,
        timeout=DEFAULT_TIMEOUT,
        api_key=api_key,
    )

    file_id = f"{uploaded_file.name}:{uploaded_file.size}"
    if st.session_state.get("file_id") != file_id:
        st.session_state["file_id"] = file_id
        st.session_state.pop("translated_pptx", None)

    action_col, download_col, _spacer = st.columns([1, 1, 6], gap="small")
    translate_clicked = action_col.button("Translate Deck", type="primary")
    download_placeholder = download_col.empty()
    if "translated_pptx" in st.session_state:
        download_placeholder.download_button(
            "Download PPTX",
            data=st.session_state["translated_pptx"],
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_{target_language}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    if translate_clicked:
        if not config.api_key:
            set_status(st.empty(), "Please provide an API key.")
            return
        raw_bytes = uploaded_file.read()
        if not raw_bytes:
            set_status(st.empty(), "The uploaded file was empty.")
            return

        presentation = Presentation(BytesIO(raw_bytes))
        paragraphs = collect_paragraphs(presentation)

        if not paragraphs:
            set_status(st.empty(), "No translatable paragraphs were found in this deck.")
            return

        xml_count = count_xml_parts(raw_bytes)
        total_steps = len(paragraphs) + xml_count
        if total_steps == 0:
            total_steps = 1

        progress_container = st.empty()
        status_text = st.empty()
        render_progress(progress_container, 0.0)
        tracker = ProgressTracker(total=total_steps, container=progress_container)
        start_time = time.time()

        try:
            set_status(status_text, "Translating text…")
            translate_paragraphs(paragraphs, config, tracker=tracker, lexicon=lexicon_store)
        except requests.RequestException as exc:
            set_status(st.empty(), f"Translation request failed: {exc}")
            return
        except RuntimeError as exc:
            set_status(st.empty(), str(exc))
            return
        except Exception as exc:
            set_status(st.empty(), f"Translation failed: {exc}")
            return
        output = BytesIO()
        presentation.save(output)
        pptx_bytes = output.getvalue()
        if xml_count:
            set_status(status_text, "Translating chart/SmartArt text…")
        pptx_bytes = translate_xml_parts(pptx_bytes, config, tracker=tracker, lexicon=lexicon_store)
        output = BytesIO(pptx_bytes)
        output.seek(0)

        render_progress(progress_container, 1.0)

        elapsed = time.time() - start_time
        set_status(status_text, f"Translation complete in {elapsed:.1f}s.")

        st.session_state["translated_pptx"] = output.getvalue()
        download_placeholder.download_button(
            "Download PPTX",
            data=st.session_state["translated_pptx"],
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_{target_language}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

if __name__ == "__main__":
    main()
