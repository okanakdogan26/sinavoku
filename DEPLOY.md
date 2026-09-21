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
   - Main file path: `app.py`
4. Deploy et.

## Desteklenen cevap anahtarı dosyaları

- **PDF** (metinli): Barış sütunlu, Yayın Denizi / Maarif başlıklı, yan yana A|B tablo
- **Excel** (`.xlsx`): Tözok satır tablosu, GİS ızgara / çapraz sayfa
- **JSON**: standart `answer_key.standard.json`
- **Görüntü** (`.jpg` / `.png`): OCR dener; WhatsApp ekran görüntüsü ve taranmış tablo çoğu zaman zayıf kalır — **Anahtar Düzenleyici** veya hazır JSON kullanın

Öğrenci dosyası: optik form **TXT** (windows-1254).

`samples/` altındaki örnek çiftler bu formatları temsil eder.

## Notlar

- Sistem PDF anahtarını birden fazla yerleşim stratejisiyle okur; gerekirse OCR (`pdftoppm` + `tesseract`) dener.
- Okuma zayıfsa **Anahtar Düzenleyici** sekmesinden düzeltip standart JSON indirin; sonuç üretiminde JSON tercih edin.
- Üretilen `xlsx` ve `*_answer_key.standard.json` dosyaları `.gitignore` ile versiyon kontrolü dışında tutulur.

### OCR (görüntü / taranmış PDF)

```bash
# macOS
brew install poppler tesseract tesseract-lang
```

Kurulu değilse görüntü PDF’lerde otomatik okuma zayıf kalır; düzenleyici + JSON kullanın.
