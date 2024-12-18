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

1. ACME System sitesi açılır.
2. "Internal Invoices" butonuna tıklanır.
3. "Download Full Invoice Report" butonuna tıklanır.
4. İndirilen "export_total" exceli bulunur ve taşınır.
5. İndirilen excele "Reconcile" sütunu eklenir.
6. ACME System sitesine geri dönüş yapılır.
7. "Internal Invoices" butonuna tıklanır.
8. "Download Monthly Invoices" butonuna tıklanır.
9. İndirilen excel içni For Each Excel Row oluşturulur.
10. İndirilen excelden "Vendor TaxID" ve ay bilgileri alınır.
11. Sayı olarak gelen ay bilgisi, ay ismine çevrilir.
12. "Vendor TaxID" kısmına veri girilir.
13. İşlem yapılacak ay girilir.
14. "Download Report" butonuna tıklanır.
15. İndirilen aylık excel taşınır.
16. Aylık excel için For Each Excel Row oluşturulur.
17. "Total" sütunundaki verilerin toplamı alınır.
18. Total excelindeki "Total" sütunundaki veriden, aylık excelinden alınan "Total" sütununun toplamı olan sayı çıkartılır.
19. Aradaki fark, Total excelinde oluşturulmuş olan "Reconcile" sütununa yazdırılır.
20. Bu işlemler Total excelindeki tüm sütunlar için tekrarlanır.

--Robot Yolu--

![image](https://github.com/user-attachments/assets/810972b0-9c47-4f84-8c95-c6974d1e3248)

--Total Excel Örneği--

![image](https://github.com/user-attachments/assets/02bd3a64-3ca5-4dc1-81c9-d4a720822423)

--Aylık Excel Örneği--

![image](https://github.com/user-attachments/assets/e0937f3b-363b-4062-9539-d6316d168c00)

--Düzenlenmiş Total Excel Örneği--

![image](https://github.com/user-attachments/assets/87e7458e-87a6-4ae0-ad06-ee1d8088ab06)

**[ENG]**

Hello,

For the reconciliation process, the robot first opens the ACME System website. Then it clicks on the "Internal Invoices" button and the "Download Full Invoice Report" button.
Next, it searches for the downloaded "export_total.xlsx" Excel file. It lists all files with the ".xlsx" extension in the Downloads folder into an array and checks for a file containing "total" in its name. If such a file is found, it moves the file to the desired folder. If the file already exists in the target folder, it first deletes the existing file and then moves the new file.
A "Reconcile" column is added to the "export_total.xlsx" file to log the operations performed during the process. The robot then navigates back to the ACME System website, clicks on the "Internal Invoices" button again, and then clicks the "Download Monthly Invoices" button.
A For Each Excel Row loop is created. The necessary information for the "Vendor TaxID" field on the ACME System is retrieved from the "Vendor" column in the "export_total.xlsx" file and entered into the website. The required information for the "Month" field is also retrieved from the "Month" column in the same file as a numeric value. This numeric month value is converted to the month name using DateTime, making it suitable for use on the site.
After entering the month field, the "Download Report" button is clicked to download the monthly Excel file. The downloaded file is moved using the same dynamic method. After moving the file, the Excel file is read.
Another For Each Excel Row loop is created. This allows summing up the data in the "Total" column of the downloaded monthly Excel file. Subsequently, the total value from the "Total" column in the "export_total.xlsx" file is subtracted by the total from the monthly Excel file. The difference is written into the "Reconcile" column that was added to the "export_total.xlsx" file.
These operations are repeated for all rows in the "export_total.xlsx" file. Once the operations are completed for all rows, the process ends.

Thanks.

1. The ACME System website is opened.
2. The "Internal Invoices" button is clicked.
3. The "Download Full Invoice Report" button is clicked.
4. The downloaded "export_total" Excel file is located and moved.
5. A "Reconcile" column is added to the downloaded Excel file.
6. The ACME System website is revisited.
7. The "Internal Invoices" button is clicked again.
8. The "Download Monthly Invoices" button is clicked.
9. A For Each Excel Row loop is created for the downloaded Excel file.
10. "Vendor TaxID" and month information are retrieved from the downloaded Excel file.
11. The numeric month information is converted to the month name.
12. Data is entered into the "Vendor TaxID" field.
13. The target month is entered.
14. The "Download Report" button is clicked.
15. The downloaded monthly Excel file is moved.
16. A For Each Excel Row loop is created for the monthly Excel file.
17. The sum of the data in the "Total" column is calculated.
18. The value from the "Total" column in the Total Excel file is subtracted by the sum of the "Total" column from the monthly Excel file.
19. The difference is written into the "Reconcile" column in the Total Excel file.
20. These operations are repeated for all columns in the Total Excel file.

--Robot Path--

![image](https://github.com/user-attachments/assets/83cdb181-620f-46cf-9fd1-160c709ddefb)

--Total Excel Example--

![image](https://github.com/user-attachments/assets/7718ff4a-e997-4f0a-ba1c-7c64d19e5ed3)

--Monthly Excel Example--

![image](https://github.com/user-attachments/assets/e0937f3b-363b-4062-9539-d6316d168c00)

--Modified Total Excel Example--

![image](https://github.com/user-attachments/assets/87e7458e-87a6-4ae0-ad06-ee1d8088ab06)
