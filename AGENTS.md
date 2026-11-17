# Sarf Malzeme Takip uygulama kararları

Bu kurallar ana ve demo uygulamalarında geçerlidir. Ekranlar yeniden tasarlanırken, taşınırken veya başka bir bileşenle değiştirilirken bu davranışları koru. Yeni veri tablolarında da aynı kuralları uygula.

## Yerleşim ve sütun filtreleri

- **Ürün ara** kutusu ile **+ Ürün kartı** düğmesini aynı satırda ve aynı alt hizada tut. Dar ekranlarda kutu küçülebilir ve düğme aynı satırda kalır.
- Web ana uygulamasında ve demoda **Ürün ara** etiketini kutunun solunda ve dikey ortasında göster. **+ Ürün kartı** düğmesinde ana işlem rengini kullan ve arama kutusuyla hizasını koru.
- Ekranlar ve diyaloglar dahil bütün veri tablolarını ortak tablo oluşturucusu ve `TableFilters` ile sun. Her veri sütununda sütun içi arama, onay kutularıyla çoklu seçim ve temizleme işlemi olmalı. Yalnızca satır düğmelerini içeren sütunları `data-action-column` olarak işaretle.
- Web ana uygulaması, demo ve Excel VBA dahil tüm ekran ve diyaloglarda, filtrelenebilir her veri sütununun seçenek listesinin üstünde **Tümünü seç** onay kutusu bulunur. Kutuyu işaretlemek, diğer sütunların filtrelerinden kalan tüm seçenekleri seçer. Sütun içi arama, toplu seçimin kapsamını daraltmaz. Kutunun işaretini kaldırıp **Uygula** düğmesine basmak o sütunun filtresini temizler. Diğer sütunların seçimleri korunur. Bu davranışı ortak filtre bileşeninde uygula ve yeni ekranlarda koru.
- Aynı sütundaki seçimleri VEYA, farklı sütunlardaki seçimleri VE mantığıyla birleştir. Bir sütunun seçeneklerini, diğer sütunların uygulanmış filtrelerinden kalan satırlara göre hesapla. Etkin seçimleri tablo yeniden çizildiğinde ve veri yenilendiğinde koru.
- Web ana uygulamasında ve demoda sol menüyü, üst araç çubuğunu ve ekran başlığını tablo satırları kaydırılırken görünür tut. Her tabloda sütun başlıkları ve varsa sütun hesapları satırlarla birlikte kaymasın. Özet kartları ve araçları kompakt yerleştir, dar ekranda menüyü üstte erişilebilir tut ve çok kısa ekranlarda içerik kaydırmasına izin ver. Otomatik yenilemede aynı ekranın kaydırma konumlarını koru.

## Stok durumu ve stok hareketleri

- Stok Durumu tablosunda **Renk**, **Beden** ve **Ayakkabı numarası** ayrı filtrelenebilir sütunlar olarak kalır. Varyantın beden veya ayakkabı numarası sütununa yazılmasını ürün grubunun **Ölçü türü** belirler.
- Stok Durumu ekranında **Eldeki miktar** ve **Stok değeri**, Stok Hareketleri ekranında **Miktar** ve **Stok değeri farkı** sütunlarının üstünde **Toplam**, **Ortalama**, **En düşük (Min)** ve **En yüksek (Max)** işlemleri bulunur.
- Sütun hesapları filtrede görünür satırları kullanır. Miktarları adet, çift ve kilogram gibi her stok birimi için ayrı hesapla. Stok hareketlerindeki artı ve eksi işaretlerini koru.
- Web ana uygulamasında ve demoda hesap satırını kompakt tut. Varsayılan işlem **Toplam** olsun. Filtre yokken **Tüm satırlar**, arama, malzeme grubu, ekran veya sütun filtresi uygulanmışken **Filtreli satırlar** yaz. Birim sonuçlarını yan yana göster ve dar alanda gerektiğinde satıra böl. Kullanıcı diğer hesapları menüden seçebilsin.

## Tüketim raporu kartları ve çıktılar

- Tablo filtresi uygulandığında veya temizlendiğinde rapor kartlarını aynı anda güncelle. **Net tüketim maliyeti** görünür satırların toplamıdır. **Aylık ortalama** filtreli görünümde kalan ay sayısını, filtreler temizlendiğinde seçili takvim dönemini kullanır.
- Rapor maliyet kartlarına ek olarak **Giriş miktarı**, **Tüketim miktarı** ve **Son kalan bakiye miktarı** kartları bulunur. İadeleri net tüketimden düş. Son bakiyeye önceki dönemden devreden stok ve bakiye tarihine kadar olan depo hareketleri dahildir. Bakiye tarihini kartta göster. Ay sütunu filtrelenirse görünür son ayın son gününü, rapor bitiş tarihi daha erkense bitiş tarihini kullan. Adet, çift ve kilogram gibi stok birimlerini ayrı göster. Depo giriş ve bakiye kartlarında mevcut depo kapsamı yetkilerini uygula.
- Filtreli **Depo giriş değeri** kartı görünür ürünlerin ilgili aylardaki depo alışlarını toplar. Aynı ay ve ürün varyantı birden fazla tüketim satırında görünse de ilgili alış hareketlerini bir kez hesaba kat. Kartın kapsamını görünür etikette belirt. Mevcut maliyet ve departman erişim yetkilerini uygula.
- Excel ve PDF çıktılarında görünür veri satırlarını, etkin filtreleri, güncel kart değerlerini ve seçili sütun hesaplarını kullan.

## Fiziksel sayım ve dönem kilidi

- Fiziksel sayımın liste ve ürün detayında standart sütun filtrelerini ve CSV, Excel, PDF çıktısını koru. Sayım şablonu taslaktaki ürünleri içerir. Excel yüklemesi ve ekrandaki ürün bazlı ekleme, düzeltme ve silme işlemleri sayım kaydını günceller.
- Düzenlenebilir sayım satırlarında miktar alanı ve satıra ait **Kaydet** düğmesi bulunur. Filtre ve çıktılarda kayıtlı miktarı kullan. Kilitleme ve onaydan önce ekrandaki bekleyen miktar değişikliklerinin kaydedilmesini iste.
- Sayım tarihi sistem bakiyesinin karşılaştırma tarihidir. Sayılmamış satırla sıfır olarak sayılmış satırı ayrı tut. Taslak, kilitli veya reddedilmiş sayımlarda stok hareketi oluşturma. Onayda farkı tek veritabanı işlemiyle stok düzeltme fişlerine işle. Sonraki hareketler varsa geçmiş maliyeti koruyarak sayım farkını son hareket tarihinde kaydet.
- Kilitlenen sayımdaki değişiklikler ve kararlar yöneticilere açıktır. Onaylanmış sayımı geçmişiyle sakla ve sonraki düzeltmeyi yeni sayımla işle.
- Geçen takvim ayının ilk gününden önceki hareket, fiş referansı, tanım geçerlilik tarihi ve sayım yazma işlemlerinde dönem kilidini İstanbul tarihine göre sunucuda uygula. Yönetici belirli bir Windows kullanıcısına başlangıç ve bitiş tarihleri dahil yazma istisnası verebilir ve geri alabilir. İstisna mevcut ekran, işlem ve departman yetkileriyle birlikte uygulanır. Kilitli veya onaylı sayım durumunu değiştirmez. Yetkili kullanıcılar önceki dönem kayıtlarını okuyabilir ve çıktı alabilir.

## Yönetim ekranları

- Kullanıcılar ve Yetkiler, Dönem Kilidi ve yeni yönetim ekranlarını yalnızca doğrulanmış yönetici oturumunda göster. Menüyü, ekranı, yönetim verisini ve yazma işlemlerini yönetici yetkisine bağla. Yönetici olmayan kullanıcıya, tarih işlemleri için yalnızca kendi açık tarih aralıklarını gönder.

## Değişiklik sonrası doğrulama

- Dar ve geniş ekranda ürün aramasıyla düğmenin hizasını, sütunlarda çoklu seçim ve diğer sütun seçeneklerinin daralmasını, boş tabloyu ve tanım geçmişi penceresini doğrula.
- L beden tişört ile 42 numara ayakkabıyı ayrı filtrele. Stok hesaplarının görünür satırlardan geldiğini, hareket toplamının işaretleri koruduğunu ve farklı birimlerin ayrı göründüğünü doğrula.
- Tüketim raporunda ay ve departman filtresi uygula, sonra temizle. Kartların her iki adımda güncellendiğini ve Excel/PDF çıktısının seçili görünümü yansıttığını doğrula.

## Excel VBA arayüzü

- Uygulamayı başlatıcıdan aç. Excel arka planda çalışsın ve kullanıcıya yalnızca UserForm'lar görünsün. Excel çıktısı istendiğinde ayrı bir çalışma kitabı dosyaya kaydedilsin.
- Makrolar etkinken çalışma kitabı açılışında Windows hesap adı otomatik doldurulmuş parola UserForm'u göster. Parolayı Windows doğrulaması için kullan ve hiçbir kalıcı kayda yazma.
- Doğrulama sonrası sol menülü ana UserForm'u göster. Listeleri, sütun filtrelerini, rapor özetlerini, aktarımı ve kayıt işlemlerini UserForm üzerinden sun. Liste çalışma sayfasını arka planda gizli tut. Yönetim ve dönem kilidi menülerine yalnızca yöneticiler erişir.
- Excel VBA ekranlarına ürün arama ve ürün kartı düğmesinin hizası, çoklu sütun filtreleri, filtreye bağlı özetler ve görünür satırların dışa aktarımı kurallarını uygula.

### UserForm ekran uyumu

- Bu bilgisayarda Windows ekran ölçeği **%125**. Çalışma anında Excel’in bulunduğu monitörün kullanılabilir alanını ve DPI değerini aynı koordinat sisteminde ölç. Formu bu alanın içinde ortala ve kenarlarda 12 punto boşluk bırak.
- Ana UserForm kullanılabilir alanı kaplasın ve genişliği tabloya ayırsın. Sol menü, ekran başlığı ve birincil işlem düğmeleri görünür kalsın. Tablo başlığı ve alt gezinme düğmeleri sabit kalsın, satırlar tablo alanında kaydırılsın veya sayfalansın.
- Veri sütunlarını ana UserForm'un genişliğine sığdır. Uzun metinleri hücre içinde satıra böl ve tam içeriği hücre ipucunda göster. Yatay sütun sayfalaması kullanma.
- Giriş, filtre ve işlem formları içeriklerine göre boyutlansın. Ana işlemler formun görünür alanında, uzun içerik ise kaydırılabilir bir bölümde kalsın. Standart yerleşimde metin en az 10 punto, birincil düğmeler en az 24 punto yüksekliğinde olsun.
- Kullanılabilir form genişliği **760 puntonun** veya yüksekliği **470 puntonun** altına düşerse sıkışık yerleşime geç. Kenar boşluklarını azalt, araçları gerekirse ikinci satıra al ve tabloya ayrılan görünür satır sayısını azalt. Metni okunamayacak kadar küçültmek yerine içerik alanını kaydır veya sayfala.
- Doğrulamada bu bilgisayarın %125 ölçeğini ve daha küçük bir çalışma alanını kullan. Alt düğmelerin, tablo başlıklarının, filtre penceresinin ve birincil işlemlerin görev çubuğu arkasında veya form sınırları dışında kalmadığını kontrol et.
