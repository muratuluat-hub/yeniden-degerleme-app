# Yeniden Değerleme ve Amortisman Hesaplama Programı

Sabit kıymetler için yeniden değerleme ve amortisman hesaplama yapan web uygulaması.

## Özellikler

- 📊 Excel dosyası yükleme (sürükle-bırak destekli)
- 🧮 Otomatik yeniden değerleme hesaplaması
- 💰 Amortisman hesaplama (Normal/Hızlı yöntem)
- 📝 Muhasebe fişleri oluşturma
- 📥 Excel çıktı dosyası indirme
- 🎨 Modern ve kullanıcı dostu arayüz

## Yerel Kullanım

```bash
pip install -r requirements.txt
python web_app.py
```

Tarayıcıda `http://localhost:8080` adresini açın.

## Kullanım

1. Sabit kıymet listesi Excel dosyanızı yükleyin
2. İşlem yılını girin
3. Dönemi seçin (1. Dönem / 2. Dönem / 3. Dönem / Yıllık)
4. Yeniden değerleme oranını girin
5. Hesaplamayı başlatın
6. Sonuç Excel dosyasını indirin

## Excel Dosyası Formatı

Giriş dosyanız şu kolonları içermelidir:

| Kolon | Açıklama |
|-------|----------|
| sabit kıymet | Hesap kodu (254, 255, vb.) |
| sabit kıymet açıklama | Kıymet açıklaması |
| aktife giriş tarihi | Tarihi (GG.AA.YYYY) |
| amortisman oranı | Oran (0.2 = %20) |
| amortisman yöntemi | "Normal" veya "Hızlı" |
| defter son değeri | Tutar |
| defter birikmiş amort | Tutar |
| defter net değeri | Formül veya tutar |

## Lisans

© 2025 - Tüm hakları saklıdır.
