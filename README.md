# Yeniden Değerleme ve Amortisman Uygulaması

Excel sabit kıymet listesini alır, VUK uygulamasına uygun yeniden değerleme ve amortisman hesaplamalarını yapar, sonucu muhasebe fişleriyle birlikte Excel olarak indirir.

## Özellikler

- Eski uygulamadaki 8 kolonlu sabit kıymet şablonunu destekler.
- Yeniden değerleme tablosu üretir.
- Muhasebe fişlerini hesap kodu bazında oluşturur.
- `254` hesap kodundaki taşıtları binek/taşıt kabul ederek ilk yıl kıst amortisman uygular.
- Son amortisman yılına gelen kıymetlere `Son yıl`, 254 taşıtlara `Son yıl dikkat` uyarısı yazar.
- Render gibi Python web servislerinde çalışmaya hazırdır.

## Yerelde Çalıştırma

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Tarayıcıdan `http://localhost:5000` adresini açın. Bu port doluysa `PORT=5001 python app.py` ile 5001 portunda çalıştırabilirsiniz.

## Beklenen Excel Kolonları

Şablonu uygulama içindeki `Şablon Excel Dosyasını İndir` bağlantısından alabilirsiniz. Eski uygulamadaki şablon formatı desteklenir:

- sabit kıymet
- sabit kıymet açıklama
- aktife giriş tarihi
- amortisman oranı
- amortisman yöntemi
- defter son değeri
- defter birikmiş amort
- defter net değeri

## Örnek Dosya

`ornek_son_yil_binek_test.xlsx` dosyası son yıl ve binek kıst amortisman kontrolleri için örnek veri içerir.

## Not

Bu uygulama hesaplama yardımcısıdır. Vergi uygulamalarında nihai kontrol için ilgili VUK hükümleri, tebliğler ve meslek mensubu değerlendirmesi esas alınmalıdır.
