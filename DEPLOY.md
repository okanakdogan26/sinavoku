# Sınavoku Deploy ve Çalıştırma

## Yerelde Çalıştırma

```bash
cd /Users/okanakdogan/Documents/WebProjects/koc_ogrenci/sinavoku
python3 -m pip install --user -r requirements.txt
python3 -m streamlit run app.py
```

Tarayıcı: `http://localhost:8501`

## Streamlit Community Cloud Deploy

1. Projeyi GitHub'a push et.
2. `https://share.streamlit.io` üzerinden `New app` aç.
3. Ayarlar:
   - Repository: `<kendi-repon>`
   - Branch: `main` (veya aktif branch)
   - Main file path: `sinavoku/app.py`
4. Deploy et.

## Uygulama Akışı

1. `Anahtar Düzenleyici` sekmesinde PDF'den taslak anahtar çıkar.
2. Gerekirse A/B kitapçık ve ders cevaplarını düzelt.
3. `Standart JSON İndir` ile `answer_key.standard.json` kaydet.
4. `Sonuç Üret` sekmesinde:
   - TXT yükle
   - Anahtar olarak PDF veya tercihen standart JSON yükle
   - Excel sonucu indir.

## Notlar

- Her yayınevinin PDF formatı farklı olabilir; bu yüzden standart JSON kullanımı önerilir.
- Üretilen `xlsx` ve `*_answer_key.standard.json` dosyaları `.gitignore` ile versiyon kontrolü dışında tutulur.
