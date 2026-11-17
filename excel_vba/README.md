# Sarf Malzeme Takip · Excel VBA

[Uygulamayı başlatın](<../outputs/excel-vba/Sarf Malzeme Takip.lnk>). Başlatıcı Excel’i arka planda çalıştırır ve ana ekranı açar. Yerel servis bu bilgisayardaki Windows oturumuna bağlı bir anahtarla uygulamada kayıtlı hesabın yetkilerini uygular. Formlar açılırken Excel'in bulunduğu monitörün kullanılabilir alanına göre ölçeklenir.

## Yerel pilot

Pilot adresi `http://localhost:8768`. Veriler `data/ExcelVbaPilot/records.json` dosyasındadır. Bu çalışma kitabına örnek ürünler, personel, girişler, teslimler ve bir açık sayım eklenmiştir. Pilot yönetici hesabı, dosya servisi ilk başlatılırken Windows'tan alınmıştır.

Servisi durdurduktan sonra proje kökünde yeniden başlatabilirsiniz. Python 3.12, `openpyxl`, `reportlab` ve `pywin32` gerekir.

```powershell
python -m pip install openpyxl==3.1.5 reportlab==4.4.9 pywin32==312
python -m excel_vba.service --data data/ExcelVbaPilot --admin (whoami)
```

Boş bir kayıt klasörüyle ilk kurulumda gerçek yönetici hesabını belirtin. Yalnızca deneme kayıtları için `--sample` kullanın. Ana web uygulamasının kayıtları Excel pilotuna aktarılmamıştır. Excel şablonuyla toplu tanım veya hareket aktarımı yapılabilir.

## Kullanım

- Sol menüden ilgili ekranı açın. Bir sütun başlığına tıklayarak değer arayın ve onay kutularıyla birden fazla değer seçin. Diğer sütunların seçenekleri etkin filtrelere göre daralır. Tablonun bütün sütunları aynı ekrandadır. Uzun bir değerin tamamını görmek için hücrenin üzerinde bekleyin.
- Filtre penceresindeki **Tümünü seç** kutusu sütunun seçeneklerini topluca seçer. Kutunun işaretini kaldırıp **Uygula** düğmesine basarak o sütunun filtresini temizleyebilirsiniz.
- Gridde bir satıra tıklayarak seçebilirsiniz. **Sonraki satırlar** düğmesi bir sonraki kayıt grubunu gösterir.
- Stok ve hareket ekranlarında miktar ile değer için toplam, ortalama, minimum ve maksimum seçilebilir. **İşlemi aç** düğmesi seçilen işlemin formunu açar. Düzenlemede kaydı seçip **Kaydı yükle** düğmesini kullanın.
- Günlük giriş ve teslim formları tek varyant kaydeder. Çok satırlı fişler Excel içe aktarma şablonuyla yüklenebilir.
- **Sayım > Yeni sayım** ile ad ve tarih girin. Üstteki **Kayıtlı sayımlar** listesinden bir ad seçtiğinizde aşağıdaki ürün tablosu güncellenir. Yeni sayım formunda **Elle giriş** ürün ve miktar formunu, **Excel ile giriş** ise sonuç dosyası seçicisini açar. **Sayım şablonunu indir** ile ürün listesini Excel’e alın, **Sayılan miktar** sütununu doldurun ve **Excel sonuçlarını yükle** ile dosyayı seçin. Yükleme sayımı kaydeder. Yetkili kullanıcı **Sayımı stoğa işle** düğmesini onayladığında farklar stok hareketlerine **Sayım düzeltmesi** olarak yazılır. Açık ve düzenlenebilir sayımlarda bir ürün seçip **Miktarı düzelt** düğmesini kullanın. Dönem ve sayım durumu kilitleri uygulanır.
- Filtrelenmiş liste Excel, PDF ve CSV olarak kaydedilebilir. Excel ve PDF rapor çıktıları hesap özetlerini de içerir. Tüketim miktarı, giriş miktarı ve son kalan bakiye filtreli görünümden yeniden hesaplanır.

## Dosya güvenliği ve şirket ağı

- Servis kullanıcının Windows SID'sini doğrular. Ekran, işlem, maliyet ve departman kuralları servis tarafında uygulanır. Her kullanıcı kendi çalışma kitabı kopyasıyla çalışır.
- Veri dizini yalnızca servis hesabına ve yetkili yedekleme yöneticilerine açık olmalıdır. Excel kullanıcıları servis adresine erişir. Kayıtlar tek yazıcı kilidi altında geçici dosyaya yazılır ve atomik olarak değiştirilir. Yedek almak için servisi durdurup `records.json` dosyasını kopyalayın.
- Kurumsal kurulum için `windows_service.py` Windows Servis Denetim Yöneticisi uyarlamasını içerir. Yönetici, `%ProgramData%/SarfMalzemeExcel/service.json` dosyasında `data`, `host`, `port`, `certificate` ve `private_key` ayarlarını tanımlar. Ağ erişimi HTTPS gerektirir. Servis hesabını kurumun Windows ve sertifika yapılandırmasıyla belirleyin.
- Yerel pilot arka planda çalışır. Şirket ağında yayın ve Windows servisi kaydı bu pilotta etkinleştirilmemiştir. Ağ kurulumunda iki farklı Windows hesabıyla oturum ve yetki ayrımı doğrulanmalıdır.

## VBA kaynakları

`vba` klasöründe form, sınıf ve modül kaynakları vardır. `Build-Workbook.ps1` bunları yerel Microsoft Excel ile çalışma kitabına ekler. Normal kullanım için VBA proje nesne modeline erişim izni gerekmez.
