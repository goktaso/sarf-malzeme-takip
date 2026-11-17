import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.argv[2];
const columns = JSON.parse(await fs.readFile(path.join(root, "import_columns.json"), "utf8"));
const workbook = Workbook.create();
const guide = workbook.worksheets.add("Kullanım");
const instructions = [
  ["TOPLU EXCEL AKTARIMI", "Sarf Malzeme Takip"],
  ["Başlangıç", "İhtiyacınız olan sayfaları doldurun. Verileri 2. satırdan başlayarak girin ve sütun başlıklarını koruyun."],
  ["İşlem", "Excel içe aktar ekranında önce Dosyayı kontrol et, ardından İçe aktar düğmesini kullanın."],
  ["Ürün Grupları", "Grup adı: Tişört  ·  Kod öneki: TSH  ·  Modül: KKD veya Sarf."],
  ["Departmanlar", "Örnek: Üretim. Personel ve departman tesliminde departman adını aynı şekilde yazın."],
  ["Personel", "Örnek: 0017  ·  Deniz Kaya  ·  Üretim  ·  Çalışıyor. Durum boşsa Çalışıyor kabul edilir."],
  ["Ürün Kartları", "Örnek: POLO  ·  TSH  ·  Polo iş tişörtü  ·  adet  ·  Personel  ·  M,L  ·  Lacivert  ·  2."],
  ["Teslim kapsamı", "Personel, Departman veya Her ikisi. Ürün kodu, seçilen grubun önekiyle otomatik oluşturulur."],
  ["Varyantlar", "Örnek: @POLO  ·  Siyah  ·  M  ·  2. Mevcut kart için sistemdeki ürün kodunu yazın."],
  ["Kart referansı", "Aynı dosyada oluşturulan kartı varyant veya hareket satırında @POLO gibi bir referansla kullanın."],
  ["Fiş referansı", "Her fişe benzersiz bir referans verin. Birden çok ürün satırında fiş referansını ve üst bilgileri tekrarlayın."],
  ["İşlem türleri", "Alış, Personel teslimi, Departman teslimi, İade veya Sayım. Tarih için Excel tarihi ya da YYYY-AA-GG kullanın."],
  ["Alış", "İşlem tarihi, alış tarihi, tedarikçi, belge numarası, ürün, miktar ve stok birimi başına TL maliyet girin."],
  ["Teslim", "Personel sicili girin. Departman tesliminde departman adını ve teslim alan yetkilinin sicilini birlikte girin."],
  ["İade", "Kaynak teslim fişi için uygulamadaki #12 gibi fiş numarasını veya Excel teslim fişinin @referansını yazın."],
  ["Sayım / açılış stoğu", "Sayılan toplam miktarı ve düzeltme gerekçesini girin. Stok artıyorsa TL birim maliyeti de yazın."],
  ["Hareket sırası", "Fişler tarih sırasıyla, aynı gün ise dosyadaki ilk satır sırasıyla işlenir. Önce girişi, sonra çıkışı yazın."],
  ["Renk ve beden", "Hareket satırında ürün kartındaki renk ve bedeni yazın. Ürün bu özellikleri izlemiyorsa alanı boş bırakın."],
  ["Kontrol", "Hatalı veya mükerrer satır varsa dosya kaydedilmez. İçe aktarma yeni kayıt ekler."],
];

function header(range) {
  range.format.fill = "#164C40";
  range.format.font = { name: "Arial", size: 10, bold: true, color: "#FFFFFF" };
  range.format.rowHeight = 25;
}

guide.showGridLines = false;
guide.getRange(`A1:B${instructions.length}`).values = instructions;
guide.getRange(`A1:B${instructions.length}`).format.font = { name: "Arial", size: 10, color: "#19302A" };
guide.getRange("A1:A19").format.columnWidth = 28;
guide.getRange("B1:B19").format.columnWidth = 105;
guide.getRange("A2:A19").format.fill = "#E6F2EC";
guide.getRange("A2:A19").format.font = { name: "Arial", size: 10, bold: true, color: "#19302A" };
guide.getRange("B2:B19").format.wrapText = true;
guide.getRange("A2:B19").format.rowHeight = 29;
header(guide.getRange("A1:B1"));

const out = path.join(root, "outputs", "excel-import");
await fs.mkdir(out, { recursive: true });
for (const [name, mapping] of Object.entries(columns)) {
  const sheet = workbook.worksheets.add(name);
  const entries = Object.entries(mapping);
  const last = String.fromCharCode(64 + entries.length);
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange(`A1:${last}1`).values = [Object.values(mapping)];
  header(sheet.getRange(`A1:${last}1`));
  const body = sheet.getRange(`A2:${last}21`);
  body.format.fill = "#F0F7F3";
  body.format.font = { name: "Arial", size: 10, color: "#19302A" };
  for (const [index, [key, title]] of entries.entries()) {
    const col = String.fromCharCode(65 + index);
    sheet.getRange(`${col}1:${col}21`).format.columnWidth = Math.max(17, title.length + 5);
    const numeric = ["qty", "minimum", "unit_price"].includes(key);
    if (["day", "purchase_day"].includes(key)) {
      sheet.getRange(`${col}2:${col}501`).format.numberFormat = "yyyy-mm-dd";
    } else if (!numeric) {
      sheet.getRange(`${col}2:${col}501`).format.numberFormat = "@";
    }
  }
}

for (const [index, name] of ["Kullanım", ...Object.keys(columns)].entries()) {
  const end = name === "Kullanım" ? "B19" : String.fromCharCode(64 + Object.keys(columns[name]).length) + "5";
  const render = await workbook.render({ sheetName: name, range: "A1:" + end, scale: 1 });
  await fs.writeFile(path.join(path.dirname(fileURLToPath(import.meta.url)), `preview-${index}.png`),
    new Uint8Array(await render.arrayBuffer()));
}
const file = await SpreadsheetFile.exportXlsx(workbook);
await file.save(path.join(out, "toplu-aktarim-sablonlari.xlsx"));
