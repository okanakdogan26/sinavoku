import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from oku import normalize_keys, parse_pdf_keys, run_pipeline


st.set_page_config(page_title="Sınav Oku", page_icon="📊", layout="centered")
st.title("Sınav Sonuç Oluşturucu")
st.caption("TXT öğrenci cevapları + PDF cevap anahtarı yükleyin, sistem anahtarı otomatik standartlaştırıp sonucu üretsin.")


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


txt_file = st.file_uploader("Öğrenci dosyası (.txt)", type=["txt"])
key_file = st.file_uploader("Cevap anahtarı (.pdf / .json)", type=["pdf", "json"])
kazanim_file = st.file_uploader(
    "Kazanım tablosu (.csv / .xlsx) - Opsiyonel",
    type=["csv", "xlsx"],
)

if st.button("Sonuç Oluştur", type="primary", use_container_width=True):
    if txt_file is None or key_file is None:
        st.error("Lütfen hem TXT hem cevap anahtarı dosyasını yükleyin.")
    else:
        with st.spinner("Dosyalar işleniyor..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp = Path(tmpdir)
                txt_path = tmp / txt_file.name
                key_path = tmp / key_file.name
                key_json_path = tmp / "answer_key.standard.json"
                out_path = tmp / "sinav_sonuclari.xlsx"
                kazanim_path = None

                txt_path.write_bytes(txt_file.getvalue())
                key_path.write_bytes(key_file.getvalue())
                if kazanim_file is not None:
                    kazanim_path = tmp / kazanim_file.name
                    kazanim_path.write_bytes(kazanim_file.getvalue())

                try:
                    if key_path.suffix.lower() == ".json":
                        key_json_path.write_bytes(key_path.read_bytes())
                    else:
                        standard_keys = normalize_keys(parse_pdf_keys(str(key_path)))
                        key_json_path.write_bytes(_serialize_standard_keys(standard_keys))

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
                        st.caption("Detaylar indirilen Excel içinde: Kazanım Zayıf / Kazanım Özet / Kazanım Detay.")
                        st.dataframe(pd.DataFrame(df.attrs["kazanim_weak_preview"]), use_container_width=True)

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
                        # 1) Toplam net histogramı
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

                        # 2) Ders ortalama netleri
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

                        # 3) Sınıf bazlı ortalama toplam net
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

                        # 4) Doğru-Yanlış-Boş dağılımı (yüzde)
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
                        # Convert each lesson to percentage over (D+Y+Boş)
                        agg_pct = agg_counts.copy().astype(float)
                        for col in agg_pct.columns:
                            total = agg_pct[col].sum()
                            if total > 0:
                                agg_pct[col] = (agg_pct[col] / total) * 100.0
                        st.bar_chart(agg_pct)

                        # 5) Sıralama eğrisi
                        st.caption("Sıralama Eğrisi (Toplam Net)")
                        rank_curve = (
                            df_view.sort_values("Toplam Net", ascending=False)
                            .head(top_n)[["Ad Soyad", "Toplam Net"]]
                            .reset_index(drop=True)
                        )
                        rank_curve.index = rank_curve.index + 1
                        st.line_chart(rank_curve["Toplam Net"])
