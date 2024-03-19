# Oracle-ODBC-Driver-Windows-Kurulumu
Oracle ODBC Driver Windows Kurulum adımları

https://www.oracle.com/tr/database/technologies/instant-client/winx64-64-downloads.html adresinden **"Basic Package"** ve **"ODBC Package"** zip dosyalarını indirip
**"C:\oracle\instantclient_xx_xx"** konumuna çıkaralım *(xx_xx sürüm numarasıdır)*

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows.png)

**Komut satırını (cmd)** yönetici olarak çalıştırıp sırasıyla aşağıdaki komutları giriyoruz
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
(Eğer **"oracle odbc error 126 sqoras32 dll"** hatası alırsanız https://dosya.co/53ryept2kobp/Visual-C-Runtimes-All-in-One-YK.rar.html adresinden Visual-C-Runtimes paketini indirip zipten çıkardıktan sonra **"install_all"** a tıklayarak tümünü kurulalım)

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_5.png)

Sıradaki işlemimiz sunucu bilgilerimiz ile bir dsn oluşturmak, bilgileri girelim ve **"TEST CONNECTION"** e basalım, aşağıdaki gibi bir bilgi mesajı ekrana gelmeli

![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_6.png)

Artık ODBC kurulumumuz ve sunucu bağlantımız hazır, aşağıdaki gibi bir connecting string (VBA) ile driverimizi kullanabiliriz

```
strCon = "DSN=ODBS; " & _
"CONNECTSTRING=(DESCRIPTION = (ADDRESS = (COMMUNITY = tcp.world)(PROTOCOL = UDP)(Host = sunucu_ip_adresi)(Port = 1521))) " & _
"(CONNECT_DATA=(SID=ORCL))); uid=sql_kullanıcı_adı; pwd=sql_kullanıcı_şifresi;"
Con.Open (strCon)
```

**"DSN" alanına yazdığımız isimle connecting stringde yer alan DSN ismi aynı olmalıdır**

**Sorgularda Türkçe karakter sorunu yaşayanlar aşağıdaki düzeltmeyi de yapabilirler**
```
NLS_LANG
```
```
AMERICAN_TURKEY.WE8ISO8859P9
```
![instantclient_xx_xx](https://github.com/yasinkeles/Oracle-ODBC-Driver-Windows-Kurulumu/blob/main/1_odbc_windows_7.png)

**kurulumda yer alan öğeler internetten indirilebildiği için indirilen kaynak tarafından eklenmiş olabilecek zararlı yazılımlara karşı dikkatli olun**

