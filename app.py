import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from oku import normalize_keys, parse_pdf_keys_detailed, run_pipeline, TESTS, key_quality


KEY_FILE_TYPES = ["pdf", "json", "xlsx", "xls", "png", "jpg", "jpeg"]

st.set_page_config(page_title="Sınav Oku", page_icon="📊", layout="centered")
st.title("Sınav Sonuç Oluşturucu")
st.caption(
    "TXT öğrenci cevapları + PDF / Excel / görüntü / JSON cevap anahtarı. "
    "Sistem birden fazla yerleşimi ve gerekirse OCR dener."
)


def _serialize_standard_keys(keys):
    default_lengths = {
        "Türkçe": 40,
        "Sosyal": 20,
        "Matematik": 40,
        "Fen": 20,
    }
    tests = []
    for name in ["Türkçe", "Sosyal", "Matematik", "Fen"]:
        length = (
            max(
                len(keys.get("A", {}).get(name, "")),
                len(keys.get("B", {}).get(name, "")),
            )
            or default_lengths[name]
        )
        tests.append({"name": name, "length": length})

    payload = {
        "version": 1,
        "exam_type": "TYT",
        "tests": tests,
        "booklets": keys,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def _keys_from_editor_state():
    keys = {"A": {}, "B": {}}
    for booklet in ("A", "B"):
        for test_name, default_len in TESTS:
            raw = st.session_state.get(f"key_{booklet}_{test_name}", "")
            cleaned = "".join(ch if ch in "ABCDE " else " " for ch in raw.upper())
            if len(cleaned) < default_len:
                cleaned = cleaned + (" " * (default_len - len(cleaned)))
            keys[booklet][test_name] = cleaned
    return normalize_keys(keys)


def _load_keys_into_editor(keys):
    for booklet in ("A", "B"):
        for test_name, _ in TESTS:
            st.session_state[f"key_{booklet}_{test_name}"] = keys.get(booklet, {}).get(
                test_name, ""
            )


tab_sonuc, tab_anahtar = st.tabs(["Sonuç Üret", "Anahtar Düzenleyici"])


with tab_anahtar:
    st.subheader("Cevap anahtarını çıkar / düzelt")
    st.write(
        "PDF, Excel veya fotoğraf yükleyin; sistem yerleşimleri ve gerekirse OCR’ı dener. "
        "Eksik/yanlış harfleri aşağıdan düzeltip standart JSON indirin."
    )
    edit_pdf = st.file_uploader(
        "Cevap anahtarı (PDF / Excel / görüntü / JSON)",
        type=KEY_FILE_TYPES,
        key="editor_key_file",
    )
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Dosyadan otomatik çıkar", type="primary", use_container_width=True):
            if edit_pdf is None:
                st.error("Önce cevap anahtarı dosyasını yükleyin.")
            else:
                with tempfile.TemporaryDirectory() as tmpdir:
                    path = Path(tmpdir) / edit_pdf.name
                    path.write_bytes(edit_pdf.getvalue())
                    if path.suffix.lower() == ".json":
                        data = json.loads(path.read_text(encoding="utf-8"))
                        keys = normalize_keys(data.get("booklets", data))
                        meta = {"strategy": "json", "quality": key_quality(keys)}
                    else:
                        keys, meta = parse_pdf_keys_detailed(str(path))
                        keys = normalize_keys(keys)
                    _load_keys_into_editor(keys)
                    st.session_state["key_meta"] = meta
                    st.success(
                        f"Okundu · strateji: {meta.get('strategy')} · "
                        f"doluluk: {meta.get('quality')} · kaynak: {meta.get('source')}"
                    )
                    cands = meta.get("candidates") or []
                    if cands:
                        st.caption(
                            "Adaylar: "
                            + ", ".join(f"{c['strategy']}={c['quality']}" for c in cands[:5])
                        )
    with col_b:
        if st.button("Editörü temizle", use_container_width=True):
            for booklet in ("A", "B"):
                for test_name, default_len in TESTS:
                    st.session_state[f"key_{booklet}_{test_name}"] = " " * default_len

    if "key_meta" in st.session_state:
        meta = st.session_state["key_meta"]
        q = meta.get("quality") or 0
        if q < 160:
            st.warning(
                "Otomatik okuma zayıf görünüyor. Harfleri kontrol edin veya "
                "görüntü PDF ise sistemde `tesseract` + `pdftoppm` kurulu olsun."
            )

    for booklet in ("A", "B"):
        st.markdown(f"#### Kitapçık {booklet}")
        for test_name, default_len in TESTS:
            key_id = f"key_{booklet}_{test_name}"
            if key_id not in st.session_state:
                st.session_state[key_id] = " " * default_len
            st.text_input(
                f"{booklet} · {test_name} ({default_len})",
                key=key_id,
                help="Sadece A–E ve boşluk. Uzunluk ders soru sayısına göre ayarlanır.",
            )

    edited = _keys_from_editor_state()
    st.caption(f"Editör doluluk skoru: {key_quality(edited)}")
    st.download_button(
        "Standart Anahtar JSON İndir",
        data=_serialize_standard_keys(edited),
        file_name="answer_key.standard.json",
        mime="application/json",
        use_container_width=True,
    )


with tab_sonuc:
    txt_file = st.file_uploader("Öğrenci dosyası (.txt)", type=["txt"])
    key_file = st.file_uploader(
        "Cevap anahtarı (.pdf / .xlsx / görüntü / .json)",
        type=KEY_FILE_TYPES,
    )
    kazanim_file = st.file_uploader(
        "Kazanım tablosu (.csv / .xlsx) - Opsiyonel",
        type=["csv", "xlsx"],
    )
    use_editor_key = st.checkbox(
        "Anahtar olarak düzenleyicideki metni kullan",
        value=False,
        help="Anahtar Düzenleyici sekmesindeki A/B cevaplarını kullanır.",
    )

    if st.button("Sonuç Oluştur", type="primary", use_container_width=True):
        if txt_file is None:
            st.error("Lütfen TXT öğrenci dosyasını yükleyin.")
        elif key_file is None and not use_editor_key:
            st.error("Lütfen cevap anahtarı yükleyin veya düzenleyiciden kullanın.")
        else:
            with st.spinner("Dosyalar işleniyor..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmp = Path(tmpdir)
                    txt_path = tmp / txt_file.name
                    key_json_path = tmp / "answer_key.standard.json"
                    out_path = tmp / "sinav_sonuclari.xlsx"
                    kazanim_path = None

                    txt_path.write_bytes(txt_file.getvalue())
                    if kazanim_file is not None:
                        kazanim_path = tmp / kazanim_file.name
                        kazanim_path.write_bytes(kazanim_file.getvalue())

                    try:
                        if use_editor_key:
                            standard_keys = _keys_from_editor_state()
                            key_json_path.write_bytes(_serialize_standard_keys(standard_keys))
                            st.info(f"Düzenleyici anahtarı kullanıldı (doluluk {key_quality(standard_keys)}).")
                        else:
                            key_path = tmp / key_file.name
                            key_path.write_bytes(key_file.getvalue())
                            if key_path.suffix.lower() == ".json":
                                key_json_path.write_bytes(key_path.read_bytes())
                            else:
                                raw_keys, meta = parse_pdf_keys_detailed(str(key_path))
                                standard_keys = normalize_keys(raw_keys)
                                key_json_path.write_bytes(_serialize_standard_keys(standard_keys))
                                q = key_quality(standard_keys)
                                st.caption(
                                    f"Anahtar: {meta.get('strategy')} · "
                                    f"doluluk {q} · "
                                    f"kaynak {meta.get('source')}"
                                )
                                if q < 160:
                                    st.warning(
                                        "PDF anahtarı zayıf okundu. "
                                        "Anahtar Düzenleyici ile düzeltip JSON kullanın."
                                    )

                        df = run_pipeline(
                            str(txt_path),
                            str(key_json_path),
                            str(out_path),
                            str(kazanim_path) if kazanim_path else None,
                        )
                    except ValueError as exc:
                        st.error(str(exc))
                    except Exception as exc:
                        st.exception(exc)
                    else:
                        st.success(f"Tamamlandı. {len(df)} öğrenci işlendi.")
                        st.dataframe(
                            df[["Sıra", "Ad Soyad", "Sınıf", "Toplam Net"]].head(20),
                            use_container_width=True,
                        )
                        st.download_button(
                            "Excel İndir",
                            data=out_path.read_bytes(),
                            file_name="sinav_sonuclari.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                        st.download_button(
                            "Otomatik Üretilen Standart Anahtarı İndir (JSON)",
                            data=key_json_path.read_bytes(),
                            file_name="answer_key.standard.json",
                            mime="application/json",
                            use_container_width=True,
                        )
                        if "kazanim_weak_preview" in df.attrs:
                            st.markdown("### Kazanım Analizi (Zayıf Alanlar)")
                            st.caption(
                                "Detaylar indirilen Excel içinde: "
                                "Kazanım Zayıf / Kazanım Özet / Kazanım Detay."
                            )
                            st.dataframe(
                                pd.DataFrame(df.attrs["kazanim_weak_preview"]),
                                use_container_width=True,
                            )

                        st.markdown("### Detaylı Analiz")

                        df_view = df.copy()
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            class_opts = ["Tümü"] + sorted(
                                [str(x) for x in df_view["Sınıf"].dropna().unique().tolist()]
                            )
                            selected_class = st.selectbox("Sınıf", class_opts, index=0)
                        with col2:
                            book_opts = ["Tümü"] + sorted(
                                df_view["Tahmini Kitapçık"].dropna().astype(str).unique().tolist()
                            )
                            selected_book = st.selectbox("Kitapçık", book_opts, index=0)
                        with col3:
                            top_n = st.slider(
                                "İlk N Öğrenci (sıralama grafiği)",
                                min_value=5,
                                max_value=max(5, len(df_view)),
                                value=min(20, len(df_view)),
                            )

                        if selected_class != "Tümü":
                            df_view = df_view[df_view["Sınıf"].astype(str) == selected_class]
                        if selected_book != "Tümü":
                            df_view = df_view[
                                df_view["Tahmini Kitapçık"].astype(str) == selected_book
                            ]

                        if df_view.empty:
                            st.warning("Filtreye uygun veri bulunamadı.")
                        else:
                            st.caption("Toplam Net Dağılımı")
                            hist = (
                                df_view["Toplam Net"]
                                .round(0)
                                .value_counts()
                                .sort_index()
                                .rename_axis("Net")
                                .reset_index(name="Öğrenci Sayısı")
                            )
                            st.bar_chart(hist.set_index("Net"))

                            st.caption("Ders Bazlı Ortalama Net")
                            lesson_means = pd.DataFrame(
                                {
                                    "Ders": ["Türkçe", "Sosyal", "Matematik", "Fen"],
                                    "Ortalama Net": [
                                        df_view["Türkçe Net"].mean(),
                                        df_view["Sosyal Net"].mean(),
                                        df_view["Matematik Net"].mean(),
                                        df_view["Fen Net"].mean(),
                                    ],
                                }
                            )
                            st.bar_chart(lesson_means.set_index("Ders"))

                            st.caption("Sınıf Bazlı Ortalama Toplam Net")
                            class_means = (
                                df[df["Sınıf"].astype(str).str.strip() != ""]
                                .groupby("Sınıf", dropna=True)["Toplam Net"]
                                .mean()
                                .sort_values(ascending=False)
                                .reset_index()
                            )
                            if not class_means.empty:
                                st.bar_chart(class_means.set_index("Sınıf"))
                            else:
                                st.info("Sınıf bilgisi yeterli değil.")

                            st.caption("Doğru / Yanlış / Boş Dağılımı (%)")
                            agg_counts = pd.DataFrame(
                                {
                                    "Metrik": ["Doğru", "Yanlış", "Boş"],
                                    "Türkçe": [
                                        df_view["Türkçe Doğru"].sum(),
                                        df_view["Türkçe Yanlış"].sum(),
                                        df_view["Türkçe Boş"].sum(),
                                    ],
                                    "Sosyal": [
                                        df_view["Sosyal Doğru"].sum(),
                                        df_view["Sosyal Yanlış"].sum(),
                                        df_view["Sosyal Boş"].sum(),
                                    ],
                                    "Matematik": [
                                        df_view["Matematik Doğru"].sum(),
                                        df_view["Matematik Yanlış"].sum(),
                                        df_view["Matematik Boş"].sum(),
                                    ],
                                    "Fen": [
                                        df_view["Fen Doğru"].sum(),
                                        df_view["Fen Yanlış"].sum(),
                                        df_view["Fen Boş"].sum(),
                                    ],
                                }
                            ).set_index("Metrik")
                            agg_pct = agg_counts.copy().astype(float)
                            for col in agg_pct.columns:
                                total = agg_pct[col].sum()
                                if total > 0:
                                    agg_pct[col] = (agg_pct[col] / total) * 100.0
                            st.bar_chart(agg_pct)

                            st.caption("Sıralama Eğrisi (Toplam Net)")
                            rank_curve = (
                                df_view.sort_values("Toplam Net", ascending=False)
                                .head(top_n)[["Ad Soyad", "Toplam Net"]]
                                .reset_index(drop=True)
                            )
                            rank_curve.index = rank_curve.index + 1
                            st.line_chart(rank_curve["Toplam Net"])
