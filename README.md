# Sarf Malzeme Takip

Satın Alma deposu için Türkçe web uygulaması. Ana uygulama SQL Server ve Windows oturum doğrulaması kullanır. Yerel demo ayrı bir SQLite veritabanıyla çalışır.

## Çalıştırma

PowerShell'de proje klasöründen çalıştırın. Python 3.12, Microsoft ODBC Driver 18 ve Windows üzerinde `requirements.txt` bağımlılıkları gerekir.

```powershell
python -m pip install -r requirements.txt
python server.py
```

Tarayıcıda [http://localhost:8765](http://localhost:8765) adresini açıp **Windows hesabıyla devam et** seçeneğini kullanın. Demo için `python server.py --demo --port 8766` çalıştırıp [http://127.0.0.1:8766](http://127.0.0.1:8766) adresini açın.

## SQL Server ve erişim

- Bu bilgisayarda ana veritabanı `SarfMalzemeTakip`. Dosyalar proje klasöründeki `data/sqlserver` dizininde. Yerel bağlantı varsayılan SQL Server örneğine Windows kimliğiyle kurulur. Sunucu, veritabanı ve ODBC sürücüsü gerektiğinde `SARF_SQL_SERVER`, `SARF_SQL_DATABASE` ve `SARF_SQL_DRIVER` ortam değişkenleriyle ayarlanır.
- Yeni bir kurulumda veritabanını oluşturduktan sonra `python configure_sqlserver.py --admin (whoami) --name "Ad Soyad"` komutunu çalıştırın. Eski ana SQLite kayıtları taşınacaksa `--migrate-legacy` ekleyin. Hedef tablolar doluysa taşıma durur.
- **Kullanıcılar ve yetkiler** ekranı, Windows hesabına ekran, işlem, tedarikçi, maliyet, fiş açıklaması ve çıktı yetkisi atar. Departman kapsamı olan hesabın hareket ve personel kayıtları o departmana göre sınırlandırılır. Ortak stok görüntüleme ayrı ekran yetkisidir. Alış ve sayım için şirket geneli işlem ve ilgili alan yetkileri gerekir. Toplu Excel aktarımı yönetici yetkisi gerektirir.
- Uygulama, Windows hesabını SID ile eşler. SQL Server bağlantısı uygulama sürecinin Windows kimliğini kullanır. İşlemi yapan uygulama hesabı ayrıca işlem günlüğüne kaydedilir.
- SQL Server'da eşzamanlı sorgular kayıt işlemlerinden bağımsız bir veri görünümü kullanır. Yazma işlemleri sıraya girer. Stok kontrolü, fiş ve maliyet değişiklikleri tek işlemde tamamlanır.

## Excel içe aktarma

**Excel içe aktar** ekranından şablonu indirin. İhtiyacınız olan sayfaları doldurup dosyayı seçin. **Dosyayı kontrol et** işlemi kayıt eklemeden sonucu gösterir. **İçe aktar** başarılı olursa tüm dosya tek işlemde kaydedilir. Aktarım yeni tanımlar ve fişler ekler. Mevcut satırların düzenlenmesi uygulamadaki ilgili ekrandan yapılır.

## İlk kayıtlar

1. Departmanları ve teslim alacak personeli ekleyin.
2. **Ürün kartları → Ürün grubu** ekranında grubu ve kod önekini tanımlayın. Örneğin **Tişört** grubu için `TSH` girin. Bu grupta yeni ürünler `TSH-0001`, `TSH-0002` şeklinde kodlanır.
3. Ürün kartını oluştururken grubu seçin. Renk ve beden gerekiyorsa, örneğin **Lacivert** ve `M, L` girin. Diğer kombinasyonları aynı karta **Varyant ekle** ile ekleyin. Stok her kombinasyon için ayrı tutulur. Eldiven hem kişiye hem departmana veriliyorsa **Her ikisi** kapsamını seçin.
4. Alış girişiyle miktarı ve stok birimi başına TL fiyatı kaydedin. Başlangıç sayımı için **Sayım → Yeni sayım** ekranını kullanın.
5. Departman tesliminde departmanı ve teslim alan yetkiliyi seçin. İşlem Satın Alma deposundan çıkar ve ilgili departmanın tüketim raporuna girer.

## Stok ve maliyet

- **Stok Durumu** tablosu renk, beden ve ayakkabı numarasını ayrı sütunlarda gösterir. Ürün grubunun **Ölçü türü**, varyantın beden ya da ayakkabı numarası sütununda yer almasını belirler.
- Eldeki miktar ve stok değeri için tablonun üstünden toplam, ortalama, en düşük ve en yüksek değer seçilebilir. Hesaplar görünür satırları kullanır. Miktarlar her stok birimi için ayrı hesaplanır.
- **Stok Hareketleri** ekranında aynı hesaplar miktar ve stok değeri farkı için kullanılabilir. Giriş ve çıkışlar işaretleriyle hesaba katılır. Seçilen hesaplar Excel ve PDF çıktılarında yer alır.

- Miktarlar stok biriminde ve üç ondalık basamağa kadar kaydedilir. Örneğin alış kutu, stok birimi çift ise miktarı ve birim fiyatı girişten önce çift üzerinden hesaplayın.
- Çıkış maliyeti hareketli ağırlıklı ortalamayla hesaplanır. İade, kaynak teslim fişine bağlıdır ve o çıkışın maliyetini geri alır. Tüketim raporu iadeyi kaydedildiği ayda mahsup eder.
- Aylık ortalama seçili tarih aralığının dokunduğu takvim aylarına, sıfır hareketli aylar dahil, bölünür. Depo giriş değeri depo giriş tarihine göre gösterilir.
- Bir varyantın son hareketinden önceki tarihe yeni hareket eklenemez. Aynı gün birden fazla fiş kaydedilebilir ve maliyet kayıt sırasıyla hesaplanır.
- Ana SQL Server veritabanını SQL Server yedekleme işlemiyle yedekleyin. Demo SQLite veritabanını dosya olarak yedeklemek için demo sunucusunu kapatıp `data/demo.sqlite3` dosyasını kopyalayın.

## Tanım geçmişi

- Ürün, departman ve personel kayıtlarında **Geçmiş** düğmesi tam tanım sürümlerini gösterir. Ürün sürümünde renk, beden ve kritik stok bilgileri de yer alır. Stok hareketleri her zaman aynı ürün koduna bağlı kalır.
- Düzenlemede **Geçerlilik tarihi** değişikliğin başladığı günü belirtir. Aynı tarihli değişikliklerin tamamı saklanır. Gün bazlı görünümde o gün için en son kaydedilen sürüm geçerli olur.
- Hareket ve tüketim raporları fiş tarihindeki adları kullanır. Personel listesinde **Çalışma durumu tarihi** alanıyla seçilen günün bilgisi görülebilir.
- Geçmiş özelliği devreye alınırken eldeki tanımlar başlangıç sürümü olarak alınır. Daha önce kaydedilmemiş sürümler otomatik olarak geri oluşturulamaz.

## Fiziksel sayım

- Tarih seçerek sayım açın. Varyantlar için sistem karşılaştırma bakiyesi o tarihten alınır. Satırları düzenleyebilir, kapsam dışı ürünleri silebilir ve sayım Excel şablonundaki miktarları topluca yükleyebilirsiniz.
- Sayımı kilitlemek yöneticiye devreder. Reddedilen sayım raporu saklanır. Onay, sayılan satırlardaki farkları stok düzeltme fişlerine işler. Sonraki hareketler varsa fark güncel bakiyeye yansıtılır ve fiş tarihi ilgili son hareketin tarihine alınır.
- Boş miktar sayılmadı anlamına gelir. Sıfır, ürünün sayıldığı ve elde bulunmadığı anlamına gelir. Pozitif farkta sistem bilinen maliyeti kullanır. Maliyet bulunamazsa yetkili kullanıcı satıra TL birim maliyeti girer.
- Geçen takvim ayının ilk gününden önceki işlemler dönem kilidine tabidir. Yönetici **Dönem kilidi** ekranından bir Windows kullanıcısı için tarih aralığı açabilir ve izni sonradan kapatabilir. İzin kaydında açan ve kapatan yönetici saklanır. Kilitli sayımın yönetici onayı ayrı bir kontroldür. Sayım liste ve detayında CSV, Excel ve PDF alınabilir.

## IIS ile şirket içi yayın

Yerel pilot, Windows oturumunu loopback adresinde doğrudan doğrular. Şirket içi yayın için IIS'te ayrı bir HTTPS site adresi ve Windows Authentication kullanılabilir. Üretimde IIS'in doğruladığı kullanıcı kimliği, yalnızca IIS'in erişebildiği bir bağlantıyla uygulamaya aktarılmalı. Mevcut loopback giriş yöntemi bu IIS bağlantısını içermez. Uygulama hizmet hesabına veritabanında günlük çalışma yetkileri verilmeli ve IIS üzerinden iki farklı Windows hesabıyla yetki ve departman ayrımı doğrulanmalı.

Bu bilgisayardaki SQL Server Developer Edition yerel geliştirme ve kabul testi içindir. Üretimde kurumun lisansına ve kapasite ihtiyacına göre SQL Server Express veya lisanslı bir SQL Server örneği seçilir.
