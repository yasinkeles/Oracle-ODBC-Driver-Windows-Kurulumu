# Oracle-ODBC-Driver-Windows-Kurulumu
Oracle ODBC Driver Windows Kurulum adımları

Öncelikle **"oracle odbc error 126 sqoras32 dll"** hatası almamak için https://www.techpowerup.com/download/visual-c-redistributable-runtime-package-all-in-one adresinden Visual-C-Runtimes paketini indirip zipten çıkardıktan sonra **"install_all"** dosyasını yönetici olarak çalıştıralım ve kurulumun bitmesini bekleyelim.

https://www.oracle.com/tr/database/technologies/instant-client/winx64-64-downloads.html adresinden (21 sürümü tavsiye edilir) **"Basic Package"** ve **"ODBC Package"** zip dosyalarını indirip
**"C:\oracle\instantclient_xx_xx"** konumuna çıkaralım *(xx_xx sürüm numarasıdır)*

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows.png)

Kurulum bittikten sonra **Komut satırını (cmd)** yönetici olarak çalıştırıp sırasıyla aşağıdaki komutları giriyoruz
```
cd C:\oracle\instantclient_xx_xx
```
```
odbc_install.exe
```

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_2.png)

**"oracle odbc driver is installed successfully"** ibaresini görmeliyiz, eğer bu ibareyi göremezseniz **"basic"** ve **"odbc"** zip dosyalarını kontrol edebilirsiniz

Sıradki işlemimiz ortam değişkenlerini ayarlamak olacaktır, ortam değişkenlerini aşağıdaki görseldeki gibi ayarlamalıyız (başlata ortam değişkenleri yazarak ilgili ekranı açabilirsiniz)

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_3.png)

Bu işlemide tamamladıktan sonra başlata **"ODBC Veri Kaynakları (64-bit)"** yazarak ekrana gelecek programı açalım

Sürücüler sekmesi altında **"instantclient_xx_xx"** in yüklenmiş olduğunu kontrol edelim

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_4.png)

Şimdi **"SISTEM DSN"** sekmesi altından **"EKLE"** butonuna basarak **"instantclient_xx_xx"** i bulup **"SON"** a tıklayalım

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_5.png)

Sıradaki işlemimiz sunucu bilgilerimiz ile bir dsn oluşturmak, bilgileri girelim ve **"TEST CONNECTION"** e basalım, aşağıdaki gibi bir bilgi mesajı ekrana gelmeli

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_6.png)

Artık ODBC kurulumumuz ve sunucu bağlantımız hazır, aşağıdaki gibi bir connecting string (VBA) ile driverimizi kullanabiliriz (Sistem DSN hatası alırsanız ofis sürümünüzü kontrol ediniz, yalnızca 64 bit ofis programları ile çalışır)

```
strCon = "DSN=ODBS; " & _
"CONNECTSTRING=(DESCRIPTION = (ADDRESS = (COMMUNITY = tcp.world)(PROTOCOL = UDP)(Host = sunucu_ip_adresi)(Port = 1521))) " & _
"(CONNECT_DATA=(SID=ORCL))); uid=sql_kullanıcı_adı; pwd=sql_kullanıcı_şifresi;"
Con.Open (strCon)
```

"DSN" alanına yazdığımız isimle connecting stringde yer alan DSN ismi aynı olmalıdır

**Sorgularda Türkçe veya farklı bir dilde karakter sorunu yaşayanlar aşağıdaki düzeltmeyi de yapabilirler**
Bu ayar yapıldıktan sonra bilgisayarı yeniden başlatmalısınız.

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_7.png)
```
NLS_LANG
```
```
AMERICAN_TURKEY.WE8ISO8859P9
veya
TURKISH_TURKEY.WE8ISO8859P9 
veya
https://www.sdcc.bnl.gov/phobos/Detectors/Computing/Orant/doc/database.804/a55928/ape.htm adresinden uygun olanı seçebilirsiniz
```
**Kurulumda yer alan öğeler internetten indirilebildiği için indirilen kaynak tarafından eklenmiş olabilecek zararlı yazılımlara karşı dikkatli olun.**

