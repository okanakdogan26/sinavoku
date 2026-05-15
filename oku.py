import re
import json
import unicodedata
import pandas as pd
from pypdf import PdfReader
from openpyxl.styles import Alignment, Font


TXT_PATH = "orbi.txt"
PDF_PATH = "cevaporbital.pdf"
OUTPUT_XLSX = "sinav_sonuclari_orbital_tyt1.xlsx"

TESTS = [
    ("Türkçe", 40),
    ("Sosyal", 20),
    ("Matematik", 40),
    ("Fen", 20),
]
DEFAULT_TEST_LENGTHS = dict(TESTS)


def normalize_answers(s):
    return "".join(ch if ch in "ABCDE" else " " for ch in s)


def parse_pdf_keys(pdf_path):
    text = ""
    for page in PdfReader(pdf_path).pages:
        text += (page.extract_text() or "") + "\n"

    keys = {"A": {}, "B": {}}
    booklet = None
    current_test = None
    # Sequence-style PDF'lerde soru numarası olmadan harf akışı geliyor.
    # Bu durumda üst sınırlar TYT/AYT ortak kapsayacak şekilde geniş tutulur.
    test_counts = {
        "Türkçe": 40,
        "Sosyal": 46,      # AYT Sosyal-2: 46
        "Matematik": 40,
        "Fen": 40,         # AYT Fen: 40
    }

    def to_ascii_upper(s):
        n = unicodedata.normalize("NFKD", s)
        n = "".join(ch for ch in n if not unicodedata.combining(ch))
        return n.upper()

    def norm_test_name(line):
        t = to_ascii_upper(line)
        if "TURK DILI VE EDEBIYATI" in t or "TURK DILI EDEBIYATI" in t:
            return "Türkçe"
        if "TURKCE" in t:
            return "Türkçe"
        if "SOSYAL BILIMLER" in t:
            return "Sosyal"
        if "SOSYAL" in t:
            return "Sosyal"
        if "MATEMATIK" in t:
            return "Matematik"
        if "FEN BILIMLERI" in t:
            return "Fen"
        if "FEN" in t:
            return "Fen"
        return None

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    first_test_idx = next((i for i, ln in enumerate(lines) if norm_test_name(ln)), -1)
    a_header_idx = next((i for i, ln in enumerate(lines) if "A KITAP" in to_ascii_upper(ln) or re.search(r"\(A\)", to_ascii_upper(ln))), -1)
    b_header_idx = next((i for i, ln in enumerate(lines) if "B KITAP" in to_ascii_upper(ln) or re.search(r"\(B\)", to_ascii_upper(ln))), -1)
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
        if not line:
            continue

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
            continue

        test_name = norm_test_name(line)
        if test_name:
            current_test = test_name
            if alternating_mode:
                idx = section_occurrence[test_name]
                booklet = "A" if idx % 2 == 0 else "B"
                section_occurrence[test_name] += 1
            continue

        if current_test is None:
            continue

        # Format 1: "1. A 2. B ..."
        for q, ans in re.findall(r"(\d+)\.\s*([ABCDE])", line):
            keys[booklet].setdefault(current_test, {})[int(q)] = ans

        # Format 2 fallback: separate answer lines like "A D E C ..."
        if not re.search(r"\d", line):
            letters = re.findall(r"[ABCDE]", up)
            if letters:
                add_by_sequence(booklet, current_test, letters)

    # Format 3 fallback: chunk by booklet/test headings and stream letters.
    # This catches PDFs where numbering is broken or punctuation lost.
    if key_quality(keys) < 160:
        keys2 = {"A": {}, "B": {}}
        booklet2 = None
        test2 = None
        for line in lines:
            up = to_ascii_upper(line)
            if "A KITAP" in up or re.search(r"\(A\)", up):
                booklet2 = "A"
                test2 = None
                continue
            if "B KITAP" in up or re.search(r"\(B\)", up):
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

            letters = re.findall(r"[ABCDE]", up)
            if letters:
                d = keys2[booklet2].setdefault(test2, {})
                limit = test_counts[test2]
                next_q = len(d) + 1
                for ch in letters:
                    if next_q > limit:
                        break
                    d[next_q] = ch
                    next_q += 1

        # Pick better parse by filled-answer quality.
        def to_final(kdict):
            out = {"A": {}, "B": {}}
            for b in ["A", "B"]:
                for test_name, default_len in TESTS:
                    d = kdict[b].get(test_name, {})
                    max_q = max(d.keys()) if d else 0
                    length = max_q if max_q > 0 else default_len
                    out[b][test_name] = "".join(d.get(i, " ") for i in range(1, length + 1))
            return out

        final1 = {"A": {}, "B": {}}
        for b in ["A", "B"]:
            for test_name, default_len in TESTS:
                d = keys[b].get(test_name, {})
                max_q = max(d.keys()) if d else 0
                length = max_q if max_q > 0 else default_len
                final1[b][test_name] = "".join(d.get(i, " ") for i in range(1, length + 1))
        final2 = to_final(keys2)

        if key_quality(final2) > key_quality(final1):
            return final2

    final_keys = {"A": {}, "B": {}}
    for booklet in ["A", "B"]:
        for test_name, default_len in TESTS:
            d = keys[booklet].get(test_name, {})
            max_q = max(d.keys()) if d else 0
            length = max_q if max_q > 0 else default_len
            key = "".join(d.get(i, " ") for i in range(1, length + 1))
            final_keys[booklet][test_name] = key

    return final_keys


def normalize_keys(keys):
    out = {"A": {}, "B": {}}
    for booklet in ["A", "B"]:
        for test_name, default_len in TESTS:
            raw = keys.get(booklet, {}).get(test_name, "")
            if isinstance(raw, dict):
                max_q = max(raw.keys()) if raw else 0
                length = max_q if max_q > 0 else default_len
                s = "".join(raw.get(i, " ") for i in range(1, length + 1))
            else:
                s = str(raw)
                length = len(s) if s.strip() else default_len
            s = normalize_answers(s)
            s = (s + (" " * length))[:length]
            out[booklet][test_name] = s
    return out


def save_standard_key_json(keys, path):
    tests = []
    for test_name, default_len in TESTS:
        length = (
            max(
                len(keys.get("A", {}).get(test_name, "")),
                len(keys.get("B", {}).get(test_name, "")),
            )
            or default_len
        )
        tests.append({"name": test_name, "length": length})

    payload = {
        "version": 1,
        "exam_type": "TYT",
        "tests": tests,
        "booklets": keys,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def load_standard_key_json(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return normalize_keys(data.get("booklets", {}))


def key_quality(keys):
    # Count non-space answers for TYT tests.
    total = 0
    for booklet in ["A", "B"]:
        for test_name, _ in TESTS:
            v = keys[booklet].get(test_name, "")
            if isinstance(v, dict):
                total += sum(1 for ch in v.values() if isinstance(ch, str) and ch in "ABCDE")
            else:
                total += sum(1 for ch in v if ch in "ABCDE")
    return total


def calculate_score(student_ans, key_str):
    dogru = 0
    yanlis = 0
    bos = 0

    for s, k in zip(student_ans, key_str):
        if k == " ":
            continue
        if s == " ":
            bos += 1
        elif s == k:
            dogru += 1
        else:
            yanlis += 1

    net = dogru - (yanlis / 4.0)
    return dogru, yanlis, bos, net


def get_test_lengths(keys):
    def observed_len(s):
        return len(s) if any(ch in "ABCDE" for ch in s) else 0

    return {
        test_name: (
            max(
                observed_len(keys.get("A", {}).get(test_name, "")),
                observed_len(keys.get("B", {}).get(test_name, "")),
            )
            or DEFAULT_TEST_LENGTHS[test_name]
        )
        for test_name, _ in TESTS
    }


def answers_from_layout(line, s1, s2, o1, o2, lengths):
    tur_l = lengths["Türkçe"]
    sos_l = lengths["Sosyal"]
    mat_l = lengths["Matematik"]
    fen_l = lengths["Fen"]
    g1 = tur_l + sos_l
    g2 = mat_l + fen_l

    b1 = normalize_answers(line[s1:s1 + g1])
    b2 = normalize_answers(line[s2:s2 + g2])

    if o1 == 0:
        tur = b1[:tur_l]
        sos = b1[tur_l:tur_l + sos_l]
    else:
        sos = b1[:sos_l]
        tur = b1[sos_l:sos_l + tur_l]

    if o2 == 0:
        mat = b2[:mat_l]
        fen = b2[mat_l:mat_l + fen_l]
    else:
        fen = b2[:fen_l]
        mat = b2[fen_l:fen_l + mat_l]

    return {
        "Türkçe": tur,
        "Sosyal": sos,
        "Matematik": mat,
        "Fen": fen,
    }


def answers_from_four_blocks(line, starts, lengths):
    t_start, s_start, m_start, f_start = starts
    tur_l = lengths["Türkçe"]
    sos_l = lengths["Sosyal"]
    mat_l = lengths["Matematik"]
    fen_l = lengths["Fen"]
    return {
        "Türkçe": normalize_answers(line[t_start:t_start + tur_l]),
        "Sosyal": normalize_answers(line[s_start:s_start + sos_l]),
        "Matematik": normalize_answers(line[m_start:m_start + mat_l]),
        "Fen": normalize_answers(line[f_start:f_start + fen_l]),
    }


def extract_answers(line, layout, lengths):
    if layout["mode"] == "four_blocks":
        return answers_from_four_blocks(line, layout["starts"], lengths)
    return answers_from_layout(
        line,
        layout["s1"],
        layout["s2"],
        layout["o1"],
        layout["o2"],
        lengths,
    )


def _evaluate_layout(lines, keys, lengths, layout):
    totals = []
    valid = True
    for line in lines:
        if layout["mode"] == "two_blocks":
            g1 = lengths["Türkçe"] + lengths["Sosyal"]
            g2 = lengths["Matematik"] + lengths["Fen"]
            if len(line) < max(layout["s1"] + g1, layout["s2"] + g2):
                valid = False
                break
        else:
            t_start, s_start, m_start, f_start = layout["starts"]
            if len(line) < max(
                t_start + lengths["Türkçe"],
                s_start + lengths["Sosyal"],
                m_start + lengths["Matematik"],
                f_start + lengths["Fen"],
            ):
                valid = False
                break

        ans = extract_answers(line, layout, lengths)
        a = score_with_booklet(ans, keys["A"])
        b = score_with_booklet(ans, keys["B"])
        best = a if a["Toplam Net"] >= b["Toplam Net"] else b
        totals.append(best["Toplam Net"])

    if not valid or not totals:
        return None
    return (sum(totals) / len(totals), min(totals))


def _detect_layout_two_blocks(lines, keys, lengths):
    lengths = get_test_lengths(keys)
    g1 = lengths["Türkçe"] + lengths["Sosyal"]
    g2 = lengths["Matematik"] + lengths["Fen"]
    min_line_len = max(len(x) for x in lines) if lines else 0
    max_s1 = max(45, min(120, min_line_len - g1))
    min_s2 = 80
    max_s2 = max(min_s2, min(220, min_line_len - g2))

    # Auto detect answer blocks.
    candidates = []
    for s1 in range(45, max_s1 + 1):
        for s2 in range(min_s2, max_s2 + 1):
            for o1 in [0, 1]:
                for o2 in [0, 1]:
                    layout = {"mode": "two_blocks", "s1": s1, "s2": s2, "o1": o1, "o2": o2}
                    metrics = _evaluate_layout(lines, keys, lengths, layout)
                    if metrics is None:
                        continue
                    mean_total, min_total = metrics
                    candidates.append((mean_total, min_total, layout))

    if not candidates:
        return None

    # Prefer higher mean total; tie-breaker higher minimum.
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0]


def _detect_layout_four_blocks(lines, keys, lengths):
    if not lines:
        return None

    min_start = 45
    max_line_len = max(len(x) for x in lines)
    if max_line_len <= min_start:
        return None

    letters_per_col = [0] * max_line_len
    for line in lines:
        for i, ch in enumerate(line):
            if ch in "ABCDE":
                letters_per_col[i] += 1

    test_order = ["Türkçe", "Sosyal", "Matematik", "Fen"]
    test_lens = [lengths[t] for t in test_order]
    max_starts = [max_line_len - l for l in test_lens]
    if any(ms < min_start for ms in max_starts):
        return None

    window_scores = []
    for l in test_lens:
        prefix = [0]
        for v in letters_per_col:
            prefix.append(prefix[-1] + v)
        scores = [0] * max_line_len
        for s in range(min_start, max_line_len - l + 1):
            scores[s] = prefix[s + l] - prefix[s]
        window_scores.append(scores)

    gap = 2
    l1, l2, l3, l4 = test_lens
    s1_range = range(min_start, max_line_len - (l1 + l2 + l3 + l4 + 3 * gap) + 1)

    best4_from = [(-1, -1)] * (max_line_len + 2)
    best_score = -1
    best_start = -1
    for s in range(max_line_len - l4, min_start - 1, -1):
        score = window_scores[3][s]
        if score >= best_score:
            best_score = score
            best_start = s
        best4_from[s] = (best_score, best_start)

    best3_from = [(-1, -1, -1)] * (max_line_len + 2)
    best_score = -1
    best_s3 = -1
    best_s4 = -1
    for s3 in range(max_line_len - l3, min_start - 1, -1):
        s4_min = s3 + l3 + gap
        if s4_min > max_line_len - l4:
            best3_from[s3] = (best_score, best_s3, best_s4)
            continue
        s4_score, s4 = best4_from[s4_min]
        cand = window_scores[2][s3] + s4_score
        if cand >= best_score:
            best_score = cand
            best_s3 = s3
            best_s4 = s4
        best3_from[s3] = (best_score, best_s3, best_s4)

    best2_from = [(-1, -1, -1, -1)] * (max_line_len + 2)
    best_score = -1
    best_s2 = -1
    best_s3 = -1
    best_s4 = -1
    for s2 in range(max_line_len - l2, min_start - 1, -1):
        s3_min = s2 + l2 + gap
        if s3_min > max_line_len - l3:
            best2_from[s2] = (best_score, best_s2, best_s3, best_s4)
            continue
        s3_score, s3, s4 = best3_from[s3_min]
        cand = window_scores[1][s2] + s3_score
        if cand >= best_score:
            best_score = cand
            best_s2 = s2
            best_s3 = s3
            best_s4 = s4
        best2_from[s2] = (best_score, best_s2, best_s3, best_s4)

    best_candidate = None
    for s1 in s1_range:
        s2_min = s1 + l1 + gap
        if s2_min > max_line_len - l2:
            continue
        s2_score, s2, s3, s4 = best2_from[s2_min]
        if s2 == -1:
            continue
        starts = (s1, s2, s3, s4)
        layout = {"mode": "four_blocks", "starts": starts}
        metrics = _evaluate_layout(lines, keys, lengths, layout)
        if metrics is None:
            continue
        mean_total, min_total = metrics
        cand = (mean_total, min_total, layout)
        if best_candidate is None or (cand[0], cand[1]) > (best_candidate[0], best_candidate[1]):
            best_candidate = cand

    return best_candidate


def detect_layout(lines, keys):
    lengths = get_test_lengths(keys)
    best_two = _detect_layout_two_blocks(lines, keys, lengths)
    best_four = _detect_layout_four_blocks(lines, keys, lengths)

    if best_two is None and best_four is None:
        raise ValueError("TXT içinde cevap blokları bulunamadı.")
    if best_two is None:
        return best_four[2]
    if best_four is None:
        return best_two[2]
    return best_four[2] if (best_four[0], best_four[1]) > (best_two[0], best_two[1]) else best_two[2]


def parse_student_line(line, layout, lengths):
    prefix = line[:55]
    ans = extract_answers(line, layout, lengths)

    tc_match = re.search(r"\b\d{11}\b", prefix)
    tc_no = tc_match.group(0) if tc_match else ""

    if tc_no:
        name_part = prefix[:tc_match.start()]
        class_part = prefix[tc_match.end():]
    else:
        name_part = prefix
        class_part = ""

    # Remove symbols used in first columns.
    ad_soyad = re.sub(r"[*0-9]", " ", name_part)
    ad_soyad = re.sub(r"\s+", " ", ad_soyad).strip()
    sinif = class_part.strip()

    return tc_no, ad_soyad, sinif, ans


def score_with_booklet(ans_map, keyset):
    out = {}
    total = 0.0
    for test_name, _ in TESTS:
        d, y, b, n = calculate_score(ans_map[test_name], keyset[test_name])
        out[f"{test_name} Doğru"] = d
        out[f"{test_name} Yanlış"] = y
        out[f"{test_name} Boş"] = b
        out[f"{test_name} Net"] = n
        total += n
    out["Toplam Net"] = total
    return out


def d_y_n(ans, key):
    d, y, b, n = calculate_score(ans, key)
    return d, y, n


def run_pipeline(txt_path, pdf_path, output_xlsx):
    print("Cevap anahtarı okunuyor...")
    if str(pdf_path).lower().endswith(".json"):
        keys = load_standard_key_json(pdf_path)
    else:
        keys = normalize_keys(parse_pdf_keys(pdf_path))

    per_booklet_quality = {
        booklet: sum(
            sum(1 for ch in keys.get(booklet, {}).get(test_name, "") if ch in "ABCDE")
            for test_name, _ in TESTS
        )
        for booklet in ["A", "B"]
    }
    if max(per_booklet_quality.values()) < 80:
        raise ValueError(
            "PDF cevap anahtarı yeterli doğrulukta okunamadı. "
            "Muhtemelen farklı başlık/yerleşim var; lütfen kontrol edin."
        )
    for booklet in ["A", "B"]:
        lens = {k: len(v) for k, v in keys[booklet].items()}
        print(f"{booklet} anahtarı uzunlukları: {lens}")

    print("Öğrenci yanıtları okunuyor...")
    results = []
    with open(txt_path, "r", encoding="windows-1254", errors="replace") as f:
        lines = [line.rstrip("\n") for line in f if line.strip()]

    lengths = get_test_lengths(keys)
    layout = detect_layout(lines, keys)
    if layout["mode"] == "four_blocks":
        print(f"Tespit edilen layout (4 blok): starts={layout['starts']}")
    else:
        print(
            "Tespit edilen layout (2 blok): "
            f"s1={layout['s1']}, s2={layout['s2']}, o1={layout['o1']}, o2={layout['o2']}"
        )

    for line in lines:
        if layout["mode"] == "four_blocks":
            t_start, s_start, m_start, f_start = layout["starts"]
            min_len = max(
                t_start + lengths["Türkçe"],
                s_start + lengths["Sosyal"],
                m_start + lengths["Matematik"],
                f_start + lengths["Fen"],
            )
        else:
            g1 = lengths["Türkçe"] + lengths["Sosyal"]
            g2 = lengths["Matematik"] + lengths["Fen"]
            min_len = max(layout["s1"] + g1, layout["s2"] + g2)

        if len(line) < min_len:
            continue

        tc_no, ad_soyad, sinif, answers = parse_student_line(line, layout, lengths)
        score_a = score_with_booklet(answers, keys["A"])
        score_b = score_with_booklet(answers, keys["B"])

        best_booklet = "A" if score_a["Toplam Net"] >= score_b["Toplam Net"] else "B"
        best_score = score_a if best_booklet == "A" else score_b

        results.append(
            {
                "TC No": tc_no,
                "Ad Soyad": ad_soyad,
                "Sınıf": sinif,
                "Tahmini Kitapçık": best_booklet,
                "A Toplam Net": score_a["Toplam Net"],
                "B Toplam Net": score_b["Toplam Net"],
                "Ans Türkçe": answers["Türkçe"],
                "Ans Sosyal": answers["Sosyal"],
                "Ans Matematik": answers["Matematik"],
                "Ans Fen": answers["Fen"],
                **best_score,
            }
        )

    df = pd.DataFrame(results)
    if df.empty:
        print("Veri okunamadı.")
        return

    df = df.sort_values("Toplam Net", ascending=False).reset_index(drop=True)
    df.insert(0, "Sıra", range(1, len(df) + 1))

    stats_rows = []
    for sinif, group in df.groupby("Sınıf"):
        stats_rows.append(
            {
                "Sınıf/Şube": sinif if sinif else "Belirtilmemiş",
                "Öğrenci Sayısı": len(group),
                "Türkçe Ort. Net": group["Türkçe Net"].mean(),
                "Sosyal Ort. Net": group["Sosyal Net"].mean(),
                "Matematik Ort. Net": group["Matematik Net"].mean(),
                "Fen Ort. Net": group["Fen Net"].mean(),
                "Toplam Ort. Net": group["Toplam Net"].mean(),
            }
        )
    stats_rows.append(
        {
            "Sınıf/Şube": "GENEL ORTALAMA",
            "Öğrenci Sayısı": len(df),
            "Türkçe Ort. Net": df["Türkçe Net"].mean(),
            "Sosyal Ort. Net": df["Sosyal Net"].mean(),
            "Matematik Ort. Net": df["Matematik Net"].mean(),
            "Fen Ort. Net": df["Fen Net"].mean(),
            "Toplam Ort. Net": df["Toplam Net"].mean(),
        }
    )
    df_stats = pd.DataFrame(stats_rows)

    # Reference-like format sheets (TYT / NET SIRALI)
    is_tyt_profile = (
        lengths["Türkçe"] == 40
        and lengths["Sosyal"] == 20
        and lengths["Matematik"] == 40
        and lengths["Fen"] == 20
    )
    has_tyt_subbreakdown = is_tyt_profile
    tyt_rows = []
    for _, r in df.iterrows():
        bk = r["Tahmini Kitapçık"]
        key = keys[bk]

        sos = r["Ans Sosyal"]
        fen = r["Ans Fen"]
        if has_tyt_subbreakdown:
            tar_d, tar_y, tar_n = d_y_n(sos[0:5], key["Sosyal"][0:5])
            cog_d, cog_y, cog_n = d_y_n(sos[5:10], key["Sosyal"][5:10])
            fel_d, fel_y, fel_n = d_y_n(sos[10:15], key["Sosyal"][10:15])
            din_d, din_y, din_n = d_y_n(sos[15:20], key["Sosyal"][15:20])

            fiz_d, fiz_y, fiz_n = d_y_n(fen[0:7], key["Fen"][0:7])
            kim_d, kim_y, kim_n = d_y_n(fen[7:14], key["Fen"][7:14])
            bio_d, bio_y, bio_n = d_y_n(fen[14:20], key["Fen"][14:20])
        else:
            tar_d = tar_y = tar_n = ""
            cog_d = cog_y = cog_n = ""
            fel_d = fel_y = fel_n = ""
            din_d = din_y = din_n = ""
            fiz_d = fiz_y = fiz_n = ""
            kim_d = kim_y = kim_n = ""
            bio_d = bio_y = bio_n = ""

        toplam_d = r["Türkçe Doğru"] + r["Sosyal Doğru"] + r["Matematik Doğru"] + r["Fen Doğru"]
        toplam_y = r["Türkçe Yanlış"] + r["Sosyal Yanlış"] + r["Matematik Yanlış"] + r["Fen Yanlış"]

        if is_tyt_profile:
            tyt_rows.append(
                {
                    "NO": int(r["Sıra"]),
                    "ŞUBE": r["Sınıf"] if pd.notna(r["Sınıf"]) else "",
                    "Numara": 0,
                    "AD VE SOYAD": r["Ad Soyad"],
                    "ALAN": "",
                    "Türkçe D": r["Türkçe Doğru"],
                    "Türkçe Y": r["Türkçe Yanlış"],
                    "Türkçe N": r["Türkçe Net"],
                    "Tarih D": tar_d, "Tarih Y": tar_y, "Tarih N": tar_n,
                    "Coğrafya D": cog_d, "Coğrafya Y": cog_y, "Coğrafya N": cog_n,
                    "Felsefe D": fel_d, "Felsefe Y": fel_y, "Felsefe N": fel_n,
                    "Din D": din_d, "Din Y": din_y, "Din N": din_n,
                    "Matematik D": r["Matematik Doğru"], "Matematik Y": r["Matematik Yanlış"], "Matematik N": r["Matematik Net"],
                    "Fizik D": fiz_d, "Fizik Y": fiz_y, "Fizik N": fiz_n,
                    "Kimya D": kim_d, "Kimya Y": kim_y, "Kimya N": kim_n,
                    "Biyoloji D": bio_d, "Biyoloji Y": bio_y, "Biyoloji N": bio_n,
                    "Toplam D": toplam_d,
                    "Toplam Y": toplam_y,
                    "NET": r["Toplam Net"],
                    "TYT PUANI": "",
                    "GENEL": "",
                    "KURUM": int(r["Sıra"]),
                    "ŞUBE SIRA": "",
                    "SINIF SIRA": "",
                    "OSYM23": "",
                    "OSYM24": "",
                    "OSYM25": "",
                }
            )
        else:
            tyt_rows.append(
                {
                    "NO": int(r["Sıra"]),
                    "ŞUBE": r["Sınıf"] if pd.notna(r["Sınıf"]) else "",
                    "Numara": 0,
                    "AD VE SOYAD": r["Ad Soyad"],
                    "ALAN": "",
                    "Türk Dili ve Edebiyatı D": r["Türkçe Doğru"],
                    "Türk Dili ve Edebiyatı Y": r["Türkçe Yanlış"],
                    "Türk Dili ve Edebiyatı N": r["Türkçe Net"],
                    "Sosyal Bilimler D": r["Sosyal Doğru"],
                    "Sosyal Bilimler Y": r["Sosyal Yanlış"],
                    "Sosyal Bilimler N": r["Sosyal Net"],
                    "Matematik D": r["Matematik Doğru"],
                    "Matematik Y": r["Matematik Yanlış"],
                    "Matematik N": r["Matematik Net"],
                    "Fen Bilimleri D": r["Fen Doğru"],
                    "Fen Bilimleri Y": r["Fen Yanlış"],
                    "Fen Bilimleri N": r["Fen Net"],
                    "Toplam D": toplam_d,
                    "Toplam Y": toplam_y,
                    "NET": r["Toplam Net"],
                    "PUAN": "",
                    "GENEL": "",
                    "KURUM": int(r["Sıra"]),
                    "ŞUBE SIRA": "",
                    "SINIF SIRA": "",
                }
            )

    df_tyt = pd.DataFrame(tyt_rows)

    df_net = df_tyt.copy()
    df_net = df_net.sort_values("NET", ascending=False).reset_index(drop=True)
    df_net["NO"] = range(1, len(df_net) + 1)

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        df_tyt.to_excel(writer, sheet_name="TYT", index=False)
        df_net.to_excel(writer, sheet_name="NET SIRALI", index=False)
        df.to_excel(writer, sheet_name="Sonuçlar", index=False)
        df_stats.to_excel(writer, sheet_name="İstatistikler", index=False)

        def style_sheet(ws):
            ws.insert_rows(1, amount=2)
            if is_tyt_profile:
                ws["F1"] = "TYT-TÜRKÇE"
                ws["I1"] = "TYT-SOSYAL BİLİMLER"
                ws["U1"] = "TYT-MATEMATİK"
                ws["X1"] = "TYT-FEN BİLİMLERİ"

                ws.merge_cells("F1:H1")
                ws.merge_cells("I1:T1")
                ws.merge_cells("U1:W1")
                ws.merge_cells("X1:AF1")

                ws["F2"] = "TÜRKÇE"
                ws.merge_cells("F2:H2")

                ws["I2"] = "TARİH"
                ws.merge_cells("I2:K2")
                ws["L2"] = "COĞRAFYA"
                ws.merge_cells("L2:N2")
                ws["O2"] = "FELSEFE"
                ws.merge_cells("O2:Q2")
                ws["R2"] = "DİN KÜLTÜRÜ"
                ws.merge_cells("R2:T2")

                ws["U2"] = "MATEMATİK"
                ws.merge_cells("U2:W2")

                ws["X2"] = "FİZİK"
                ws.merge_cells("X2:Z2")
                ws["AA2"] = "KİMYA"
                ws.merge_cells("AA2:AC2")
                ws["AD2"] = "BİYOLOJİ"
                ws.merge_cells("AD2:AF2")

                ws["AG2"] = "TOPLAM"
                ws.merge_cells("AG2:AI2")
                ws["AJ2"] = "TYT"
            else:
                ws["F1"] = "TÜRK DİLİ VE EDEBİYATI"
                ws["I1"] = "SOSYAL BİLİMLER"
                ws["L1"] = "MATEMATİK"
                ws["O1"] = "FEN BİLİMLERİ"
                ws["R1"] = "TOPLAM"

                ws.merge_cells("F1:H1")
                ws.merge_cells("I1:K1")
                ws.merge_cells("L1:N1")
                ws.merge_cells("O1:Q1")
                ws.merge_cells("R1:T1")

                ws["F2"] = "TÜRK DİLİ VE EDEBİYATI"
                ws["I2"] = "SOSYAL BİLİMLER"
                ws["L2"] = "MATEMATİK"
                ws["O2"] = "FEN BİLİMLERİ"
                ws["R2"] = "TOPLAM"

                ws.merge_cells("F2:H2")
                ws.merge_cells("I2:K2")
                ws.merge_cells("L2:N2")
                ws.merge_cells("O2:Q2")
                ws.merge_cells("R2:T2")
                ws["U2"] = "PUAN"

            max_col = ws.max_column
            for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=max_col):
                for c in row:
                    c.font = Font(bold=True)
                    c.alignment = Alignment(horizontal="center", vertical="center")

            # Basic alignment
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=max_col):
                for c in row:
                    c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions["D"].width = 28
            ws.column_dimensions["B"].width = 10
            ws.column_dimensions["C"].width = 8

        style_sheet(writer.book["TYT"])
        style_sheet(writer.book["NET SIRALI"])

    print(f"Tamamlandı: {output_xlsx}")
    print("Tahmini kitapçık dağılımı:")
    print(df["Tahmini Kitapçık"].value_counts().to_string())
    print(f"En yüksek net: {df['Toplam Net'].max():.2f}")

    # Save canonical answer-key next to output for reproducible re-runs.
    key_json_path = str(output_xlsx).rsplit(".", 1)[0] + "_answer_key.standard.json"
    save_standard_key_json(keys, key_json_path)
    print(f"Standart cevap anahtarı kaydedildi: {key_json_path}")
    return df


def main():
    run_pipeline(TXT_PATH, PDF_PATH, OUTPUT_XLSX)


if __name__ == "__main__":
    main()
