**[TR]**

Merhabalar,

Reconciliation süreci için robot ilk olarak ACME System sitesini açıyor. Ardından "Internal Invoices" butonuna ve "Download Full Invoice Report" butonuna tıklıyor.
Sonra indirilen "export_total.xlsx" excelini arıyor. İndirilenler klasöründe ".xlsx" uzantılı dosyaları bir diziye atıyor ve bu exceller arasında dosya isminde "total"
bulunduran bir dosya bulduğu takdirde taşınmak istenen klasöre taşıyor. Eğer taşınmak istenen klasörde bu dosya mevcutsa ilk önce mevcut dosyayı siliyor ve taşınmak istenen
dosyayı taşıyor. "export_total.xlsx" exceline, süreç içerisinde yapacağımız işlemleri yazdırmak için "Reconcile" sütunu ekleniyor. ACME System sitesine geri dönüş yapılıyor.
"Internal Invoices" butonuna ve ardından "Download Monthly Invoices" butonuna tıklanıyor. Bir For Each Excel Row oluşturuluyor. ACME System kısmında karşımıza çıkan "Vendor TaxID"
kısmı için gerekli olan bilgi, "export_total.xlsx" excelinde bulunan "Vendor" sütunundan alınıyor ve yazdırılıyor. İkinci olarak karşımıza çıkan "Month" kısmı için gerekli olan bilgi
yine "export_total.xlsx" excelinde bulunan "Month" kısmından sayı olarak alınıyor. Sayı olarak alınan ay, DateTime ay ismine çevriliyor ve site için kullanılabilecek hale getiriliyor.
Ay kısmını girdikten sonra "Download Report" butonuna tıklanarak aylık excel indiriliyor. İndirilen excel, aynı dinamik yöntem ile taşınıyor. Taşınma sonrası excel okunuyor.
Bir For Each Excel Row oluşturuluyor. Bu sayede İndirilen aylık excel içerisinde bulunan "Total" sütunundaki verilerin toplamı alınıyor. Sonrasında "export_total.xlsx" excelinde 
bulunan "Total" sütunundaki veriden, aylık excelinden alınan toplam sayı çıkartılıyor. Aradaki fark, "export_total.xlsx" excelinde oluşturmuş olduğumuz "Reconcile" sütununa yazdırılıyor.
Bu işlemler, "export_total.xlsx" exceli içerisindeki tüm sütunlar için tekrarlanıyor. Tüm sütunlar için bu işlemler gerçekleştirildikten sonra süreç sonlanıyor.

Teşekkürler.

**[ENG]**

Hello,

Hello,

For the reconciliation process, the robot first opens the ACME System website. Then it clicks on the "Internal Invoices" button and the "Download Full Invoice Report" button.
Next, it searches for the downloaded "export_total.xlsx" Excel file. It lists all files with the ".xlsx" extension in the Downloads folder into an array and checks for a file containing "total" in its name. If such a file is found, it moves the file to the desired folder. If the file already exists in the target folder, it first deletes the existing file and then moves the new file.
A "Reconcile" column is added to the "export_total.xlsx" file to log the operations performed during the process. The robot then navigates back to the ACME System website, clicks on the "Internal Invoices" button again, and then clicks the "Download Monthly Invoices" button.
A For Each Excel Row loop is created. The necessary information for the "Vendor TaxID" field on the ACME System is retrieved from the "Vendor" column in the "export_total.xlsx" file and entered into the website. The required information for the "Month" field is also retrieved from the "Month" column in the same file as a numeric value. This numeric month value is converted to the month name using DateTime, making it suitable for use on the site.
After entering the month field, the "Download Report" button is clicked to download the monthly Excel file. The downloaded file is moved using the same dynamic method. After moving the file, the Excel file is read.
Another For Each Excel Row loop is created. This allows summing up the data in the "Total" column of the downloaded monthly Excel file. Subsequently, the total value from the "Total" column in the "export_total.xlsx" file is subtracted by the total from the monthly Excel file. The difference is written into the "Reconcile" column that was added to the "export_total.xlsx" file.
These operations are repeated for all rows in the "export_total.xlsx" file. Once the operations are completed for all rows, the process ends.

Thank you.

Thanks.
