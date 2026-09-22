"""
Evrensel cevap anahtarı çıkarımı.

Birden fazla PDF/OCR stratejisini dener; en yüksek doluluk skorunu seçer.
Desteklenen başlıca yerleşimler:
- A|B yan yana tablo (SORU NO + ders sütunları)
- Sayfa başı kitapçık, satırda tekrarlı soru no (Barış vb.)
- Bölüm başlıklı harf/numara akışı (Acil / Yayın Denizi vb.)
- Özde TYT İlk Prova: '1 - C 11 - D …' 4×10 tire ızgarası (A+B)
- Izgara: 1.A 2.B ... (görüntü/OCR dahil, A|B yan yana satırlar)
"""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path

from pypdf import PdfReader

TESTS = [
    ("Türkçe", 40),
    ("Sosyal", 20),
    ("Matematik", 40),
    ("Fen", 20),
]
DEFAULT_TEST_LENGTHS = dict(TESTS)
WIDE_TEST_COUNTS = {
    "Türkçe": 40,
    "Sosyal": 46,
    "Matematik": 40,
    "Fen": 40,
}

OCR_ANSWER_MAP = str.maketrans(
    {
        "0": "D",
        "O": "D",
        "Q": "D",
        "8": "B",
        "€": "E",
        "£": "E",
        "§": "E",
        "Þ": "B",
        "ß": "B",
        "С": "C",  # Cyrillic
        "А": "A",
        "В": "B",
        "Е": "E",
    }
)

PAIR_RE = re.compile(
    # "1.A" "1. A" "1-A" veya "1 A" / "1A" — cevap yalnızca A-E (rakam değil)
    r"(?<!\d)(\d{1,2})\s*[.\-]\s*([A-Ea-e])(?![0-9A-Za-z])"
    r"|"
    r"(?<!\d)(\d{1,2})(?:\s+|)([A-Ea-e])(?![0-9A-Za-z])",
)

PAIR_RE_OCR = re.compile(
    # OCR: 8→B, 0→D karışımları; nokta veya boşluk ayırıcı zorunlu
    r"(?<!\d)(\d{1,2})\s*[.\-]\s*([A-E0O8€£§ÞßСАВЕa-e])(?![0-9A-Za-z])"
    r"|"
    r"(?<!\d)(\d{1,2})\s+([A-E0O8€£§ÞßСАВЕa-e])(?![0-9A-Za-z])",
)


def to_ascii_upper(s: str) -> str:
    n = unicodedata.normalize("NFKD", s)
    n = "".join(ch for ch in n if not unicodedata.combining(ch))
    return n.upper()


def norm_test_name(line: str):
    t = to_ascii_upper(line)
    if "TURK DILI VE EDEBIYATI" in t or "TURK DILI EDEBIYATI" in t:
        return "Türkçe"
    if "TURKCE" in t or "TORKCE" in t:
        return "Türkçe"
    if "SOSYAL BILIMLER" in t or "SOSYAL B." in t or "SOSYAL B " in t:
        return "Sosyal"
    if "SOSYAL" in t:
        return "Sosyal"
    if "TEMEL MATEMATIK" in t or "MATEMATIK" in t:
        return "Matematik"
    if "FEN BILIMLERI" in t or "FEN B." in t:
        return "Fen"
    if re.search(r"\bFEN\b", t):
        return "Fen"
    return None


def normalize_ocr_answer(ch: str) -> str:
    ch = ch.upper().translate(OCR_ANSWER_MAP)
    return ch if ch in "ABCDE" else ""


def key_quality(keys) -> int:
    total = 0
    for booklet in ("A", "B"):
        for test_name, _ in TESTS:
            v = keys.get(booklet, {}).get(test_name, "")
            if isinstance(v, dict):
                total += sum(1 for ch in v.values() if isinstance(ch, str) and ch in "ABCDE")
            else:
                total += sum(1 for ch in v if ch in "ABCDE")
    return total


def candidate_score(keys) -> int:
    """Doluluk + A/B dengesi (tek taraflı yanlış parse'ları cezalandır)."""
    base = key_quality(keys)
    per = {}
    for booklet in ("A", "B"):
        per[booklet] = sum(
            sum(1 for ch in keys.get(booklet, {}).get(test_name, "") if ch in "ABCDE")
            if not isinstance(keys.get(booklet, {}).get(test_name, ""), dict)
            else sum(
                1
                for ch in keys[booklet][test_name].values()
                if isinstance(ch, str) and ch in "ABCDE"
            )
            for test_name, _ in TESTS
        )
    a, b = per["A"], per["B"]
    if a >= 80 and b >= 80:
        return base + 40
    if min(a, b) < 20 and max(a, b) >= 80:
        return base - 30
    return base


def empty_dict_keys():
    return {"A": {name: {} for name, _ in TESTS}, "B": {name: {} for name, _ in TESTS}}


def finalize_keys(keys, prefer_defaults=True):
    final = {"A": {}, "B": {}}
    for booklet in ("A", "B"):
        for test_name, default_len in TESTS:
            section = keys[booklet].get(test_name, {})
            if isinstance(section, str):
                final[booklet][test_name] = section
                continue
            max_q = max(section.keys()) if section else 0
            # GİS 30/30/30/30 gibi gerçek uzunluğu koru; boşsa TYT varsayılanı.
            if max_q <= 0:
                length = default_len
            else:
                length = max_q
            final[booklet][test_name] = "".join(section.get(i, " ") for i in range(1, length + 1))
    return final


def booklet_quality(keys, booklet: str) -> int:
    total = 0
    for test_name, _ in TESTS:
        v = keys.get(booklet, {}).get(test_name, "")
        if isinstance(v, dict):
            total += sum(1 for ch in v.values() if isinstance(ch, str) and ch in "ABCDE")
        else:
            total += sum(1 for ch in v if ch in "ABCDE")
    return total


def fill_missing_booklet(keys):
    """Tek kitapçıklı anahtarlarda boş olan tarafı doldur (ÇAP Maarif vb.)."""
    if not keys:
        return keys
    qa, qb = booklet_quality(keys, "A"), booklet_quality(keys, "B")
    if qa >= 80 and qb < 40:
        keys["B"] = {k: v for k, v in keys["A"].items()}
    elif qb >= 80 and qa < 40:
        keys["A"] = {k: v for k, v in keys["B"].items()}
    return keys


def _subjects_for_tyt_table_row(q_num):
    if q_num <= 20:
        return ["Türkçe", "Sosyal", "Matematik", "Fen"]
    return ["Türkçe", "Matematik"]


def assign_tyt_table_answers(parsed, booklet, q_num, answers):
    if q_num <= 20:
        for subject, ans in zip(_subjects_for_tyt_table_row(q_num), answers):
            ans = normalize_ocr_answer(ans) or to_ascii_upper(ans)
            if ans in "ABCDE":
                parsed[booklet][(subject, q_num)] = ans
        return
    if not answers:
        return
    tur = normalize_ocr_answer(answers[0]) or to_ascii_upper(answers[0])
    mat = normalize_ocr_answer(answers[-1]) or to_ascii_upper(answers[-1])
    if tur in "ABCDE":
        parsed[booklet][("Türkçe", q_num)] = tur
    if mat in "ABCDE":
        parsed[booklet][("Matematik", q_num)] = mat


def find_q_ans_pairs(text: str, ocr: bool = False):
    rx = PAIR_RE_OCR if ocr else PAIR_RE
    out = []
    for m in rx.finditer(text):
        groups = m.groups()
        # iki alternatifli regex: (q,ans,None,None) veya (None,None,q,ans)
        if groups[0] is not None:
            q_s, ans_s = groups[0], groups[1]
        else:
            q_s, ans_s = groups[2], groups[3]
        q = int(q_s)
        ans = normalize_ocr_answer(ans_s) if ocr else ans_s.upper()
        if not ocr:
            ans = ans if ans in "ABCDE" else ""
        if 1 <= q <= 46 and ans in "ABCDE":
            out.append((q, ans))
    return out


def detect_booklet_marker(text: str):
    up = to_ascii_upper(text)
    if re.search(r"\bA\s*KITAPC", up) or re.search(r"\(\s*A\s*\)", up):
        return "A"
    if re.search(r"\bB\s*KITAPC", up) or re.search(r"\(\s*B\s*\)", up):
        return "B"
    if re.search(r"SINAVI\s*-\s*\d+\s*A\b", up):
        return "A"
    if re.search(r"SINAVI\s*-\s*\d+\s*B\b", up):
        return "B"
    # MSÜ: DENEME-1A / DENEME-1B
    if re.search(r"DENEME\s*-?\s*\d+A\b", up) or re.search(r"\b1A\b", up):
        return "A"
    if re.search(r"DENEME\s*-?\s*\d+B\b", up) or re.search(r"\b1B\b", up):
        return "B"

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in reversed(lines[-10:]):
        up_ln = to_ascii_upper(ln)
        if up_ln in ("A", "B"):
            return up_ln
        if re.fullmatch(r"[AB]\s*KITAPCIGI", up_ln):
            return up_ln[0]
        m = re.search(r"\b([AB])$", up_ln)
        if m and ("SINAV" in up_ln or "DENEME" in up_ln or "KITAP" in up_ln or "CEVAP" in up_ln):
            return m.group(1)
    return None


# ---------------------------------------------------------------------------
# Text extraction (+ OCR)
# ---------------------------------------------------------------------------


def extract_native_page_texts(pdf_path: str):
    # Özde vb. ızgaralarda pdftotext -layout daha düzenli satır üretir.
    if shutil.which("pdftotext"):
        try:
            proc = subprocess.run(
                ["pdftotext", "-layout", "-enc", "UTF-8", pdf_path, "-"],
                check=True,
                capture_output=True,
                text=True,
            )
            layout_text = proc.stdout or ""
            if layout_text.count("-") >= 40 and len(layout_text.strip()) > 200:
                # Sayfa sonu form feed ile böl
                pages = [p for p in re.split(r"\f+", layout_text) if p.strip()]
                return pages or [layout_text]
        except (subprocess.CalledProcessError, FileNotFoundError, OSError):
            pass
    reader = PdfReader(pdf_path)
    return [(page.extract_text() or "") for page in reader.pages]


def ocr_pdf_pages(pdf_path: str, dpi: int = 280):
    if not shutil.which("pdftoppm") or not shutil.which("tesseract"):
        return None
    texts = []
    convert = shutil.which("magick") or shutil.which("convert")
    with tempfile.TemporaryDirectory() as td:
        prefix = str(Path(td) / "page")
        try:
            subprocess.run(
                ["pdftoppm", "-png", "-r", str(dpi), pdf_path, prefix],
                check=True,
                capture_output=True,
            )
        except (subprocess.CalledProcessError, FileNotFoundError):
            return None
        images = sorted(Path(td).glob("page*.png"))
        if not images:
            return None
        for img in images:
            work = img
            if convert:
                enhanced = Path(td) / f"enh_{img.name}"
                try:
                    subprocess.run(
                        [
                            convert,
                            str(img),
                            "-colorspace",
                            "Gray",
                            "-normalize",
                            "-contrast-stretch",
                            "2%x2%",
                            str(enhanced),
                        ],
                        check=True,
                        capture_output=True,
                    )
                    work = enhanced
                except (subprocess.CalledProcessError, FileNotFoundError):
                    work = img
            try:
                # İki PSM dene; daha uzun olanı al
                chunks = []
                for psm in ("6", "4"):
                    proc = subprocess.run(
                        [
                            "tesseract",
                            str(work),
                            "stdout",
                            "-l",
                            "eng",
                            "--psm",
                            psm,
                        ],
                        capture_output=True,
                        text=True,
                        check=False,
                    )
                    chunks.append(proc.stdout or "")
                texts.append(max(chunks, key=lambda s: len(find_q_ans_pairs(s, ocr=True))))
            except FileNotFoundError:
                return None
    return texts


def load_page_texts(pdf_path: str, force_ocr: bool = False):
    native = extract_native_page_texts(pdf_path)
    native_chars = sum(len(t.strip()) for t in native)
    meta = {"source": "native", "native_chars": native_chars, "ocr": False}

    if force_ocr or native_chars < 80:
        ocr = ocr_pdf_pages(pdf_path)
        if ocr and sum(len(t.strip()) for t in ocr) > native_chars:
            meta["source"] = "ocr"
            meta["ocr"] = True
            return ocr, meta
    return native, meta


# ---------------------------------------------------------------------------
# Strategies
# ---------------------------------------------------------------------------


def parse_side_by_side_table(page_texts):
    keys = empty_dict_keys()
    found_rows = 0

    for page_text in page_texts:
        up = to_ascii_upper(page_text)
        if "TURKCE" not in up or "MATEMATIK" not in up:
            continue
        if not ("A KITAP" in up and "B KITAP" in up) and "SORU NO" not in up:
            # yine de satır kalıbını dene
            pass

        for line in page_text.splitlines():
            tokens = line.split()
            if not tokens:
                continue
            # "1 C D C C 1 A C C A"
            if not tokens[0].isdigit():
                continue
            q_num = int(tokens[0])
            if q_num < 1 or q_num > 40:
                continue
            split_idx = None
            for idx in range(1, len(tokens)):
                if tokens[idx] == str(q_num):
                    split_idx = idx
                    break
            if split_idx is None:
                continue
            parsed = {"A": {}, "B": {}}
            for booklet, answers in (
                ("A", tokens[1:split_idx]),
                ("B", tokens[split_idx + 1 :]),
            ):
                assign_tyt_table_answers(parsed, booklet, q_num, answers)
            if not parsed["A"] and not parsed["B"]:
                continue
            found_rows += 1
            for booklet in ("A", "B"):
                for (subject, q), ans in parsed[booklet].items():
                    keys[booklet][subject][q] = ans

    if found_rows < 10:
        return None
    return finalize_keys(keys)


def parse_columnar_booklet(page_texts):
    keys = empty_dict_keys()
    pages_used = 0
    used = set()

    for page_text in page_texts:
        rows = []
        for line in page_text.splitlines():
            pairs = re.findall(r"(\d+)\.\s*([A-E])", line, flags=re.IGNORECASE)
            if len(pairs) < 2:
                continue
            q_nums = [int(q) for q, _ in pairs]
            if len(set(q_nums)) != 1:
                continue
            q_num = q_nums[0]
            if q_num < 1 or q_num > 40:
                continue
            answers = [a.upper() for _, a in pairs]
            bucket = {"X": {}}
            assign_tyt_table_answers(bucket, "X", q_num, answers)
            rows.append(bucket["X"])

        if len(rows) < 10:
            continue

        booklet = detect_booklet_marker(page_text)
        if booklet is None:
            booklet = "A" if "A" not in used else ("B" if "B" not in used else None)
        if booklet is None:
            continue

        for subject_map in rows:
            for (subject, q), ans in subject_map.items():
                limit = WIDE_TEST_COUNTS.get(subject, 40)
                if q <= limit:
                    keys[booklet][subject][q] = ans
        pages_used += 1
        used.add(booklet)

    if pages_used < 1:
        return None
    return finalize_keys(keys)


def parse_heading_sequence(text: str):
    keys = {"A": {}, "B": {}}
    booklet = None
    current_test = None
    test_counts = dict(WIDE_TEST_COUNTS)

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    first_test_idx = next((i for i, ln in enumerate(lines) if norm_test_name(ln)), -1)
    a_header_idx = next(
        (
            i
            for i, ln in enumerate(lines)
            if "A KITAP" in to_ascii_upper(ln) or re.search(r"\(A\)", to_ascii_upper(ln))
        ),
        -1,
    )
    b_header_idx = next(
        (
            i
            for i, ln in enumerate(lines)
            if "B KITAP" in to_ascii_upper(ln) or re.search(r"\(B\)", to_ascii_upper(ln))
        ),
        -1,
    )
    alternating_mode = (
        first_test_idx != -1
        and a_header_idx != -1
        and b_header_idx != -1
        and a_header_idx < first_test_idx
        and b_header_idx < first_test_idx
    )
    section_occurrence = {name: 0 for name, _ in TESTS}

    def add_by_sequence(booklet_name, test_name, letters):
        d = keys[booklet_name].setdefault(test_name, {})
        limit = test_counts[test_name]
        next_q = len(d) + 1
        for ch in letters:
            if next_q > limit:
                break
            d[next_q] = ch
            next_q += 1

    for raw_line in lines:
        line = raw_line.strip()
        up = to_ascii_upper(line)
        if not alternating_mode and ("A KITAP" in up or re.search(r"\(A\)", up)):
            booklet = "A"
            current_test = None
            continue
        if not alternating_mode and ("B KITAP" in up or re.search(r"\(B\)", up)):
            booklet = "B"
            current_test = None
            continue
        if alternating_mode and ("A KITAP" in up or "B KITAP" in up or re.search(r"\([AB]\)", up)):
            continue
        if not alternating_mode and booklet is None:
            # Kitapçık başlığı "TYT ... A KİTAPÇIĞI" tek satırda da olabilir
            if "KITAPC" in up and re.search(r"\bA\b", up):
                booklet = "A"
                current_test = None
                continue
            if "KITAPC" in up and re.search(r"\bB\b", up):
                booklet = "B"
                current_test = None
                continue
            continue

        test_name = norm_test_name(line)
        if test_name:
            current_test = test_name
            if alternating_mode:
                idx = section_occurrence[test_name]
                booklet = "A" if idx % 2 == 0 else "B"
                section_occurrence[test_name] += 1
            continue

        if current_test is None or booklet is None:
            continue

        numbered = find_q_ans_pairs(line, ocr=False)
        if numbered:
            for q, ans in numbered:
                if ans and q <= test_counts[current_test]:
                    keys[booklet].setdefault(current_test, {})[q] = ans
            continue

        # Saf sayılar satırı (1 2 3 ...): atla
        if re.fullmatch(r"[\d\s]+", line):
            continue

        # Numara satırı + sonraki harf satırı (Acil): sadece harfler
        if not re.search(r"[A-E]", up):
            continue
        letters = re.findall(r"[A-E]", up)
        if letters and not re.search(r"\d", line):
            add_by_sequence(booklet, current_test, letters)

    # Stream fallback
    if key_quality(finalize_keys(keys)) < 120:
        keys2 = {"A": {}, "B": {}}
        booklet2 = None
        test2 = None
        for line in lines:
            up = to_ascii_upper(line)
            if "A KITAP" in up or re.search(r"\(A\)", up) or (
                "KITAPC" in up and re.search(r"\bA\b", up) and "B KITAP" not in up
            ):
                booklet2 = "A"
                test2 = None
                continue
            if "B KITAP" in up or re.search(r"\(B\)", up) or (
                "KITAPC" in up and re.search(r"\bB\b", up)
            ):
                booklet2 = "B"
                test2 = None
                continue
            if booklet2 is None:
                continue
            ntest = norm_test_name(line)
            if ntest:
                test2 = ntest
                continue
            if test2 is None:
                continue
            letters = re.findall(r"[A-E]", up)
            if not letters:
                continue
            d = keys2[booklet2].setdefault(test2, {})
            next_q = len(d) + 1
            for ch in letters:
                if next_q > test_counts[test2]:
                    break
                d[next_q] = ch
                next_q += 1
        if key_quality(finalize_keys(keys2)) > key_quality(finalize_keys(keys)):
            keys = keys2

    final = finalize_keys(keys)
    return final if key_quality(final) >= 40 else None


def parse_section_numbered_pairs(text: str, ocr: bool = False):
    """Başlık bağlamında 1.A 2.B ... topla (tek kitapçık veya sırayla A/B)."""
    keys = empty_dict_keys()
    booklet = None
    current_test = None
    filled = 0

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        up = to_ascii_upper(line)

        b = detect_booklet_marker(line)
        if b and ("KITAP" in up or "CEVAP" in up or "DENEME" in up or "SINAV" in up or len(line) < 40):
            if "A KITAP" in up and "B KITAP" in up:
                continue
            booklet = b
            current_test = None
            continue

        tname = norm_test_name(line)
        if tname and len(find_q_ans_pairs(line, ocr=ocr)) < 3:
            current_test = tname
            continue

        pairs = find_q_ans_pairs(line, ocr=ocr)
        if not pairs:
            continue
        if booklet is None:
            booklet = "A"
        if current_test is None:
            continue

        limit = WIDE_TEST_COUNTS[current_test]
        for q, ans in pairs:
            if q <= limit:
                keys[booklet][current_test][q] = ans
                filled += 1

    if filled < 40:
        return None
    return finalize_keys(keys)


def _split_dual_booklet_pairs(pairs):
    """Aynı satırda 1..N sonra tekrar 1..M → (A pairs, B pairs)."""
    if len(pairs) < 4:
        return None
    restart = None
    for i in range(1, len(pairs)):
        if pairs[i][0] == 1 and pairs[i - 1][0] >= 2:
            restart = i
            break
    if restart is None:
        # Orta noktadan böl (eşit uzunluk)
        if len(pairs) >= 8 and len(pairs) % 2 == 0:
            mid = len(pairs) // 2
            left_qs = [q for q, _ in pairs[:mid]]
            right_qs = [q for q, _ in pairs[mid:]]
            if left_qs == right_qs:
                return pairs[:mid], pairs[mid:]
        return None
    return pairs[:restart], pairs[restart:]


def parse_dual_column_grid(text: str, ocr: bool = False):
    """
    ÇAP vb. izgara: satırda A blok + B blok (soru numaraları yeniden başlar).
    Ders başlıkları satır aralarında gelir; yoksa TYT sırasıyla ilerler.
    """
    keys = empty_dict_keys()
    subject_plan = [
        ("Türkçe", 40),
        ("Sosyal", 25),
        ("Matematik", 40),
        ("Fen", 20),
    ]
    plan_idx = 0
    current_test = subject_plan[0][0]
    filled = 0
    max_seen = 0

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue

        tname = norm_test_name(line)
        pairs = find_q_ans_pairs(line, ocr=ocr)
        if tname and len(pairs) < 3:
            current_test = tname
            for i, (name, _) in enumerate(subject_plan):
                if name == tname:
                    plan_idx = i
                    break
            max_seen = 0
            continue
        if not pairs:
            continue

        qs = [q for q, _ in pairs]
        if min(qs) <= 3 and max_seen >= 18:
            if plan_idx < len(subject_plan) - 1:
                plan_idx += 1
                current_test = subject_plan[plan_idx][0]
                max_seen = 0

        split = _split_dual_booklet_pairs(pairs)
        limit = 25 if current_test == "Sosyal" else WIDE_TEST_COUNTS.get(current_test, 40)
        if split:
            for booklet, chunk in (("A", split[0]), ("B", split[1])):
                for q, ans in chunk:
                    if q <= limit:
                        keys[booklet][current_test][q] = ans
                        filled += 1
                        max_seen = max(max_seen, q)
        else:
            for q, ans in pairs:
                if q <= limit:
                    keys["A"][current_test][q] = ans
                    filled += 1
                    max_seen = max(max_seen, q)

    final = finalize_keys(keys)
    if key_quality(final) < 40:
        return None
    return final


def parse_global_q_stream(text: str, ocr: bool = False):
    """
    Son çare: tüm 1.A kalıplarını sırayla ders uzunluklarına böl.
    A ve B kitapçıklarını metindeki sıraya göre ayırır.
    """
    up = to_ascii_upper(text)
    parts = re.split(r"(?=[AB]\s*KITAPC|\([AB]\)|KITAPCIGI)", up)
    if len(parts) < 2:
        parts = [text]

    keys = empty_dict_keys()
    booklet_order = []
    for part in parts:
        b = detect_booklet_marker(part)
        if b:
            booklet_order.append((b, part))
    if not booklet_order:
        booklet_order = [("A", text)]

    for booklet, part in booklet_order:
        segments = []
        current = None
        buf = []
        for line in part.splitlines():
            tname = norm_test_name(line)
            if tname and len(find_q_ans_pairs(line, ocr=ocr)) < 3:
                if current and buf:
                    segments.append((current, "\n".join(buf)))
                current = tname
                buf = []
            else:
                buf.append(line)
        if current and buf:
            segments.append((current, "\n".join(buf)))

        if not segments:
            pairs = find_q_ans_pairs(part, ocr=ocr)
            cursor = 0
            for test_name, length in [
                ("Türkçe", 40),
                ("Sosyal", 25),
                ("Matematik", 40),
                ("Fen", 20),
            ]:
                chunk = pairs[cursor : cursor + length]
                cursor += length
                for i, (_, ans) in enumerate(chunk, start=1):
                    keys[booklet][test_name][i] = ans
            continue

        for test_name, body in segments:
            pairs = find_q_ans_pairs(body, ocr=ocr)
            by_q = {}
            for q, ans in pairs:
                by_q[q] = ans
            limit = 25 if test_name == "Sosyal" else WIDE_TEST_COUNTS[test_name]
            for q, ans in by_q.items():
                if q <= limit:
                    keys[booklet][test_name][q] = ans

    final = finalize_keys(keys)
    return final if key_quality(final) >= 40 else None


# ---------------------------------------------------------------------------
# Excel / image
# ---------------------------------------------------------------------------

GIS_BRANCH_TO_TEST = {
    "TUR": "Türkçe",
    "TURK": "Türkçe",
    "TÜR": "Türkçe",
    "MAT": "Matematik",
    "GEO": "Matematik",
    "TAR": "Sosyal",
    "COG": "Sosyal",
    "FEL": "Sosyal",
    "DIN": "Sosyal",
    "FIZ": "Fen",
    "KIM": "Fen",
    "BIO": "Fen",
}


def _cell_text(val) -> str:
    if val is None:
        return ""
    try:
        if pd_isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()


def pd_isna(val):
    try:
        import pandas as pd

        return pd.isna(val)
    except Exception:
        return val != val


def _norm_header(val) -> str:
    return re.sub(r"\s+", " ", to_ascii_upper(_cell_text(val)))


def map_branch_to_test(name: str):
    raw = _cell_text(name)
    t = to_ascii_upper(raw)
    compact = re.sub(r"[^A-Z]", "", t)
    if compact in GIS_BRANCH_TO_TEST:
        return GIS_BRANCH_TO_TEST[compact]
    n = norm_test_name(raw)
    if n:
        return n
    if compact[:3] in GIS_BRANCH_TO_TEST:
        return GIS_BRANCH_TO_TEST[compact[:3]]
    return None


def _df_to_matrix(df):
    return [[df.iat[r, c] if c < df.shape[1] else None for c in range(df.shape[1])] for r in range(df.shape[0])]


def parse_tozok_excel_df(df):
    """Tözok: Kitapcık / Ders grubu / Ders soru no / Cevap."""

    if df is None or df.empty:
        return None
    # header satırını bul
    header_row = None
    for r in range(min(8, len(df))):
        vals = [_norm_header(df.iat[r, c]) for c in range(df.shape[1])]
        joined = " ".join(vals)
        if "CEVAP" in joined and ("DERS" in joined or "KITAP" in joined):
            header_row = r
            break
    if header_row is None:
        return None

    headers = [_norm_header(df.iat[header_row, c]) for c in range(df.shape[1])]

    def col_idx(*needles):
        for i, h in enumerate(headers):
            if any(n in h for n in needles):
                return i
        return None

    i_book = col_idx("KITAPC")
    i_group = col_idx("DERS GRUBU", "[DERS GRUBU]")
    i_q = col_idx("DERS SORU")
    if i_q is None:
        i_q = col_idx("SORU NUMARASI")
    if i_q is None:
        i_q = col_idx("SORU NO")
    i_ans = col_idx("CEVAP")
    i_ders = None
    for i, h in enumerate(headers):
        if h == "DERS" or (h.startswith("DERS") and "SORU" not in h and "GRUP" not in h):
            i_ders = i
            break
    if i_ans is None or i_q is None:
        return None

    keys = empty_dict_keys()
    filled = 0
    for r in range(header_row + 1, len(df)):
        ans = to_ascii_upper(_cell_text(df.iat[r, i_ans]))[:1]
        if ans not in "ABCDE":
            continue
        try:
            q = int(float(_cell_text(df.iat[r, i_q])))
        except (TypeError, ValueError):
            continue
        booklet = "A"
        if i_book is not None:
            btxt = to_ascii_upper(_cell_text(df.iat[r, i_book]))
            if btxt.startswith("B"):
                booklet = "B"
            elif btxt.startswith("A"):
                booklet = "A"
        test = None
        if i_group is not None:
            test = map_branch_to_test(_cell_text(df.iat[r, i_group]))
        if test is None and i_ders is not None:
            test = map_branch_to_test(_cell_text(df.iat[r, i_ders]))
        if test is None:
            continue
        limit = 46 if test == "Sosyal" else 40
        if 1 <= q <= limit:
            keys[booklet][test][q] = ans
            filled += 1
    if filled < 40:
        return None
    return finalize_keys(keys)


def parse_gis_capraz_df(df):
    """GİS çapraz: A Soru No, B Soru No, Branş, Cevap."""
    header_row = None
    for r in range(min(10, len(df))):
        vals = [_norm_header(df.iat[r, c]) for c in range(min(df.shape[1], 8))]
        joined = " ".join(vals)
        if "SORU NO" in joined and "CEVAP" in joined:
            header_row = r
            break
        if "BRANS" in joined and "CEVAP" in joined:
            header_row = r
            break
    if header_row is None:
        return None

    headers = [_norm_header(df.iat[header_row, c]) for c in range(df.shape[1])]

    def col_idx(pred):
        for i, h in enumerate(headers):
            if pred(h):
                return i
        return None

    i_a = col_idx(lambda h: "A SORU" in h)
    i_b = col_idx(lambda h: "B SORU" in h)
    i_br = col_idx(lambda h: "BRANS" in h)
    i_ans = col_idx(lambda h: h == "CEVAP" or h.startswith("CEVAP "))
    if i_a is None:
        i_a = 0
    if i_br is None or i_ans is None:
        return None

    keys = empty_dict_keys()
    filled = 0
    for r in range(header_row + 1, len(df)):
        ans = to_ascii_upper(_cell_text(df.iat[r, i_ans]))[:1]
        if ans not in "ABCDE":
            continue
        test = map_branch_to_test(_cell_text(df.iat[r, i_br]))
        if test is None:
            continue
        try:
            qa = int(float(_cell_text(df.iat[r, i_a])))
        except (TypeError, ValueError):
            continue
        qb = None
        if i_b is not None:
            try:
                qb = int(float(_cell_text(df.iat[r, i_b])))
            except (TypeError, ValueError):
                qb = None
        if 1 <= qa <= 40:
            keys["A"][test][qa] = ans
            filled += 1
        if qb is not None and 1 <= qb <= 40:
            keys["B"][test][qb] = ans
            filled += 1
    if filled < 40:
        return None
    return finalize_keys(keys)


def parse_gis_grid_df(df):
    """GİS 'Cevap Anahtarı' ızgarası: soru, '-', harf blokları."""
    keys = empty_dict_keys()
    booklet = "A"
    current_tests = []  # (col_start, test_name)
    filled = 0

    for r in range(len(df)):
        row_txt = " ".join(_cell_text(df.iat[r, c]) for c in range(df.shape[1]))
        up = to_ascii_upper(row_txt)
        if re.search(r"\(B\)", up) or re.search(r"\bB\s*KITAPC", up):
            booklet = "B"
            current_tests = []
            continue
        if re.search(r"\(A\)", up) or re.search(r"\bA\s*KITAPC", up):
            booklet = "A"
            current_tests = []
            continue

        # ders başlıkları
        found = []
        for c in range(df.shape[1]):
            tname = norm_test_name(_cell_text(df.iat[r, c]))
            if tname:
                found.append((c, tname))
        if found and not re.match(r"^\d", _cell_text(df.iat[r, 1] if df.shape[1] > 1 else "")):
            # başlık satırı (sayı yok / az)
            nums = sum(1 for c in range(df.shape[1]) if re.fullmatch(r"\d+(\.0)?", _cell_text(df.iat[r, c])))
            if nums < 3:
                current_tests = found
                continue

        if not current_tests:
            continue

        # her 3 kolonda (q, -, ans) tekrarları
        for c in range(df.shape[1] - 2):
            qtxt = _cell_text(df.iat[r, c])
            ans = to_ascii_upper(_cell_text(df.iat[r, c + 2]))[:1]
            try:
                q = int(float(qtxt))
            except (TypeError, ValueError):
                continue
            if ans not in "ABCDE" or q < 1 or q > 40:
                continue
            # hangi ders bloğu
            test = None
            for start, name in reversed(current_tests):
                if c >= start:
                    test = name
                    break
            if test is None:
                test = current_tests[0][1]
            keys[booklet][test][q] = ans
            filled += 1

    if filled < 40:
        return None
    return finalize_keys(keys)


def merge_final_keys(parts):
    acc = empty_dict_keys()
    for keys in parts:
        if not keys:
            continue
        for booklet in ("A", "B"):
            for test_name, _ in TESTS:
                v = keys.get(booklet, {}).get(test_name, "")
                if isinstance(v, dict):
                    for q, ans in v.items():
                        if isinstance(ans, str) and ans in "ABCDE":
                            acc[booklet][test_name][int(q)] = ans
                else:
                    for i, ch in enumerate(str(v), start=1):
                        if ch in "ABCDE":
                            acc[booklet][test_name][i] = ch
    return finalize_keys(acc)


def parse_excel_key(path: str):
    import pandas as pd

    try:
        xl = pd.ExcelFile(path)
    except Exception as exc:
        return None, {"strategy": None, "quality": 0, "source": "excel", "error": str(exc)}

    candidates = []
    for sheet in xl.sheet_names:
        df0 = pd.read_excel(path, sheet_name=sheet, header=None)
        for name, fn in (
            ("excel_gis_capraz", parse_gis_capraz_df),
            ("excel_gis_grid", parse_gis_grid_df),
            ("excel_tozok", parse_tozok_excel_df),
        ):
            try:
                keys = fn(df0)
            except Exception:
                keys = None
            if keys and key_quality(keys) > 0:
                candidates.append((candidate_score(keys), key_quality(keys), f"{name}:{sheet.strip()}", keys))

    if not candidates:
        return finalize_keys(empty_dict_keys()), {
            "strategy": None,
            "quality": 0,
            "source": "excel",
            "candidates": [],
        }

    merged = fill_missing_booklet(merge_final_keys([k for _s, _q, _n, k in candidates]))
    return merged, {
        "strategy": "excel_merged",
        "quality": key_quality(merged),
        "source": "excel",
        "candidates": [{"strategy": n, "quality": q} for _s, q, n, _ in sorted(candidates, reverse=True)[:8]],
    }


def ocr_image_variants(image_path: str):
    """Birkaç PSM + isteğe bağlı netleştirme; parse aşamasında en iyisi seçilir."""
    if not shutil.which("tesseract"):
        return []
    sources = [image_path]
    convert = shutil.which("magick") or shutil.which("convert")
    tmp_enh = None
    if convert:
        tmp_enh = Path(tempfile.mkdtemp()) / "enh.png"
        try:
            subprocess.run(
                [
                    convert,
                    image_path,
                    "-colorspace",
                    "Gray",
                    "-normalize",
                    "-resize",
                    "200%",
                    "-sharpen",
                    "0x1",
                    str(tmp_enh),
                ],
                check=True,
                capture_output=True,
            )
            sources.append(str(tmp_enh))
        except (subprocess.CalledProcessError, FileNotFoundError):
            tmp_enh = None

    texts = []
    for src in sources:
        for psm in ("6", "4"):
            try:
                proc = subprocess.run(
                    ["tesseract", src, "stdout", "-l", "eng", "--psm", psm],
                    capture_output=True,
                    text=True,
                    check=False,
                )
            except FileNotFoundError:
                return texts
            if proc.stdout and proc.stdout.strip():
                texts.append(proc.stdout)
    return texts


def ocr_image_file(image_path: str):
    variants = ocr_image_variants(image_path)
    if not variants:
        return ""
    return max(variants, key=lambda s: len(find_q_ans_pairs(s, ocr=True)) + len(s) // 50)


def parse_glued_ocr_grid(text: str, ocr: bool = False):
    """ÇAP görüntü OCR: '10'='1.D', '28'='2.B', '12.8'='12.B' yapışık hücreler."""
    keys = empty_dict_keys()
    subject_plan = [("Türkçe", 40), ("Sosyal", 25), ("Matematik", 40), ("Fen", 20)]
    plan_idx = 0
    current_test = subject_plan[0][0]
    max_seen = 0
    filled = 0

    def parse_token(tok: str):
        tok = to_ascii_upper(tok)
        tok = tok.replace("O", "D")
        tok = re.sub(r"[^0-9A-E]", "", tok)
        m = re.fullmatch(r"(\d{1,2})([A-E0-9])", tok)
        if not m:
            return None
        q = int(m.group(1))
        ans = normalize_ocr_answer(m.group(2))
        if not ans:
            return None
        if 1 <= q <= 46:
            return q, ans
        return None

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        tname = norm_test_name(line)
        tokens = re.split(r"[|\s,/]+", line)
        parsed = [parse_token(t) for t in tokens]
        parsed = [p for p in parsed if p]
        if tname and len(parsed) < 3:
            current_test = tname
            for i, (name, _) in enumerate(subject_plan):
                if name == tname:
                    plan_idx = i
                    break
            max_seen = 0
            continue
        if len(parsed) < 4:
            continue
        qs = [q for q, _ in parsed]
        if min(qs) <= 3 and max_seen >= 18 and plan_idx < len(subject_plan) - 1:
            plan_idx += 1
            current_test = subject_plan[plan_idx][0]
            max_seen = 0
        split = _split_dual_booklet_pairs(parsed)
        limit = 25 if current_test == "Sosyal" else WIDE_TEST_COUNTS.get(current_test, 40)
        if split:
            for booklet, chunk in (("A", split[0]), ("B", split[1])):
                for q, ans in chunk:
                    if q <= limit:
                        keys[booklet][current_test][q] = ans
                        filled += 1
                        max_seen = max(max_seen, q)
        else:
            for q, ans in parsed:
                if q <= limit:
                    keys["A"][current_test][q] = ans
                    filled += 1
                    max_seen = max(max_seen, q)

    final = finalize_keys(keys)
    return final if key_quality(final) >= 60 else None


def parse_ozde_tyt_dash_blocks(text: str, ocr: bool = False):
    """Özde / TYT İlk Prova: '1 - C 11 - D … 1 - A 11 - C' 4×10 satır blokları.

    Blok sırası (A sonra B, pdftotext -layout ile doğrulanmış):
      0: A Türkçe(40) + Sosyal(≤25)
      1: A Matematik(40) + Fen(20)
      2: B Türkçe + Sosyal
      3: B Matematik + Fen
    """
    dash_re = re.compile(
        r"(?<!\d)(\d{1,2})\s*[-–—=]\s*([A-Ea-eİıI0O8€£§Þß])",
    )
    rows = []
    for raw in text.splitlines():
        pairs = []
        for m in dash_re.finditer(raw):
            q = int(m.group(1))
            ans = normalize_ocr_answer(m.group(2))
            if not ans and m.group(2).upper() in "ABCDE":
                ans = m.group(2).upper()
            if ans and 1 <= q <= 40:
                pairs.append((q, ans))
        if len(pairs) >= 4:
            rows.append(pairs)

    if len(rows) < 40:
        return None
    rows = rows[:40]

    def accumulate(row_slice):
        left, right = {}, {}
        for pairs in row_slice:
            split_at = None
            for i in range(1, len(pairs)):
                if pairs[i][0] < pairs[i - 1][0]:
                    split_at = i
                    break
            left_pairs = pairs if split_at is None else pairs[:split_at]
            right_pairs = [] if split_at is None else pairs[split_at:]
            for q, ans in left_pairs:
                left[q] = ans
            for q, ans in right_pairs:
                right[q] = ans
        return left, right

    # pdftotext -layout: önce A kitapçık (Türkçe/Sosyal + Mat/Fen), sonra B.
    # pypdf etiketleri sonda karıştırsa da 40 satırlık gövde aynı sırada gelir.
    keys = empty_dict_keys()
    for bi, booklet in enumerate(("A", "B")):
        base = bi * 20
        tur, sos = accumulate(rows[base : base + 10])
        mat, fen = accumulate(rows[base + 10 : base + 20])
        keys[booklet]["Türkçe"] = {q: a for q, a in tur.items() if q <= 40}
        keys[booklet]["Sosyal"] = {q: a for q, a in sos.items() if q <= 25}
        keys[booklet]["Matematik"] = {q: a for q, a in mat.items() if q <= 40}
        keys[booklet]["Fen"] = {q: a for q, a in fen.items() if q <= 20}

    final = finalize_keys(keys)
    return final if key_quality(final) >= 160 else None


def parse_dash_number_grid(text: str, ocr: bool = False):
    """MSÜ: '2- E  9- D  16- C' sütun-major ızgara."""
    dash_re = re.compile(
        r"(?<!\d)(\d{1,2})\s*[-–—=]\s*([A-Ea-e0O8€£§8])",
    )
    keys = empty_dict_keys()
    booklet = None
    subjects = []
    filled = 0

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        up = to_ascii_upper(line)
        b = detect_booklet_marker(line)
        if b and ("DENEME" in up or "CEVAP" in up or "OTURUM" in up or "1A" in up or "1B" in up):
            booklet = b
            subjects = []
            continue
        names = []
        for part in re.split(r"\s{1,}|\t|\|", line):
            n = norm_test_name(part)
            if n and n not in names:
                names.append(n)
        if len(names) >= 2 and dash_re.search(line) is None:
            subjects = names[:2]
            continue
        if len(names) == 1 and dash_re.search(line) is None:
            subjects = names
            continue
        pairs = []
        for m in dash_re.finditer(line):
            q = int(m.group(1))
            ans = normalize_ocr_answer(m.group(2))
            if ans and 1 <= q <= 46:
                pairs.append((q, ans))
        if not pairs:
            pairs = find_q_ans_pairs(line, ocr=True)
        if not pairs or booklet is None or not subjects:
            continue
        split = _split_dual_booklet_pairs(pairs)
        if split and len(subjects) >= 2:
            for test, chunk in zip(subjects, split):
                limit = 25 if test == "Sosyal" else WIDE_TEST_COUNTS.get(test, 40)
                for q, ans in chunk:
                    if q <= limit:
                        keys[booklet][test][q] = ans
                        filled += 1
        else:
            test = subjects[0]
            limit = 25 if test == "Sosyal" else WIDE_TEST_COUNTS.get(test, 40)
            for q, ans in pairs:
                if q <= limit:
                    keys[booklet][test][q] = ans
                    filled += 1

    final = finalize_keys(keys)
    return final if key_quality(final) >= 40 else None


def parse_side_by_side_subjects(text: str, ocr: bool = False):
    """MSÜ: satırda Türkçe|Sosyal (soru no yeniden başlar), kitapçık 1A/1B."""
    keys = empty_dict_keys()
    booklet = None
    subjects = []
    filled = 0

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        b = detect_booklet_marker(line)
        up = to_ascii_upper(line)
        if b and (
            "DENEME" in up
            or "CEVAP" in up
            or "KITAP" in up
            or "1A" in up.replace(" ", "")
            or "1B" in up.replace(" ", "")
        ):
            booklet = b
            subjects = []
            continue

        names = []
        for part in re.split(r"\s{2,}|\t|\|", line):
            n = norm_test_name(part)
            if n and n not in names:
                names.append(n)
        # tek satırda birden fazla ders adı
        whole = norm_test_name(line)
        if len(names) >= 2:
            subjects = names[:2]
            continue
        if whole and len(find_q_ans_pairs(line, ocr=ocr)) < 3:
            subjects = [whole]
            continue

        pairs = find_q_ans_pairs(line, ocr=ocr)
        if not pairs or booklet is None:
            continue
        split = _split_dual_booklet_pairs(pairs)
        if split and len(subjects) >= 2:
            chunks = [split[0], split[1]]
            for test, chunk in zip(subjects, chunks):
                limit = 25 if test == "Sosyal" else WIDE_TEST_COUNTS.get(test, 40)
                for q, ans in chunk:
                    if q <= limit:
                        keys[booklet][test][q] = ans
                        filled += 1
        elif subjects:
            test = subjects[0]
            limit = 25 if test == "Sosyal" else WIDE_TEST_COUNTS.get(test, 40)
            for q, ans in pairs:
                if q <= limit:
                    keys[booklet][test][q] = ans
                    filled += 1

    final = finalize_keys(keys)
    return final if key_quality(final) >= 40 else None


def collect_text_candidates(page_texts, ocr: bool = False, prefix: str = ""):
    full_text = "\n".join(page_texts)
    pfx = f"{prefix}_" if prefix else ""
    strategies = [
        (f"{pfx}ozde_tyt_dash_blocks", lambda: parse_ozde_tyt_dash_blocks(full_text, ocr=ocr)),
        (f"{pfx}side_by_side_table", lambda: parse_side_by_side_table(page_texts)),
        (f"{pfx}columnar_booklet", lambda: parse_columnar_booklet(page_texts)),
        (f"{pfx}heading_sequence", lambda: parse_heading_sequence(full_text)),
        (f"{pfx}section_numbered_pairs", lambda: parse_section_numbered_pairs(full_text, ocr=ocr)),
        (f"{pfx}dual_column_grid", lambda: parse_dual_column_grid(full_text, ocr=ocr)),
        (f"{pfx}glued_ocr_grid", lambda: parse_glued_ocr_grid(full_text, ocr=ocr)),
        (f"{pfx}dash_number_grid", lambda: parse_dash_number_grid(full_text, ocr=ocr)),
        (f"{pfx}side_by_side_subjects", lambda: parse_side_by_side_subjects(full_text, ocr=ocr)),
        (f"{pfx}global_q_stream", lambda: parse_global_q_stream(full_text, ocr=ocr)),
    ]
    candidates = []
    for name, fn in strategies:
        try:
            keys = fn()
        except Exception:
            continue
        if not keys:
            continue
        q = key_quality(keys)
        if q <= 0:
            continue
        candidates.append((candidate_score(keys), q, name, keys))
    return candidates


def pick_best_candidate(candidates, text_meta):
    if not candidates:
        empty = finalize_keys(empty_dict_keys())
        return empty, {**text_meta, "strategy": None, "quality": 0, "candidates": []}
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    _score, best_q, best_name, best_keys = candidates[0]
    best_keys = fill_missing_booklet(best_keys)
    return best_keys, {
        **text_meta,
        "strategy": best_name,
        "quality": key_quality(best_keys),
        "candidates": [{"strategy": n, "quality": q} for _s, q, n, _ in candidates[:8]],
    }


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------


def parse_answer_key_pdf(pdf_path: str):
    """
    Returns (keys, meta) where meta includes strategy, quality, text_source.
    """
    page_texts, text_meta = load_page_texts(pdf_path)
    candidates = collect_text_candidates(page_texts, ocr=text_meta.get("source") == "ocr")

    if (not candidates or max(c[0] for c in candidates) < 160) and not text_meta.get("ocr"):
        ocr_pages = ocr_pdf_pages(pdf_path)
        if ocr_pages and sum(len(t) for t in ocr_pages) > 50:
            candidates.extend(collect_text_candidates(ocr_pages, ocr=True, prefix="ocr"))
            text_meta = {**text_meta, "ocr_attempted": True}

    return pick_best_candidate(candidates, text_meta)


def parse_image_key(image_path: str):
    meta = {"source": "ocr_image", "ocr": True, "native_chars": 0}
    variants = ocr_image_variants(image_path)
    if not variants:
        return finalize_keys(empty_dict_keys()), {**meta, "strategy": None, "quality": 0, "candidates": []}
    candidates = []
    for i, text in enumerate(variants):
        candidates.extend(collect_text_candidates([text], ocr=True, prefix=f"img{i}"))
    return pick_best_candidate(candidates, meta)


def parse_answer_key(path: str):
    """PDF / Excel / görüntü / JSON cevap anahtarını oku."""
    suf = Path(path).suffix.lower()
    if suf == ".json":
        import json

        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        keys = fill_missing_booklet(data.get("booklets", data))
        return keys, {"strategy": "json", "quality": key_quality(keys), "source": "json"}
    if suf in {".xlsx", ".xls"}:
        return parse_excel_key(path)
    if suf in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
        return parse_image_key(path)
    return parse_answer_key_pdf(path)
