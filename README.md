
## EN | English

### Overview
This script reads devices from an Excel file, creates them in **WALLIX Bastion PAM**, adds the proper service (**SSH/RDP**), and optionally assigns each device+service pair to a **Target Group**.

### Features
- Device names like `ServerName_<lastTwoOctets>`, e.g. `MySrv_34.213`
- Adds service based on the `Service` column (currently **SSH**, **RDP**)
- Group assignment using **device & service names**
- Handles `204/409` and resolves IDs via pagination
- Fully configurable via CLI args or env vars
- Optional CSV run log

### Prerequisites
- Python 3.9+  
- Install deps:
```bash
pip install -r requirements.txt
```

### Excel columns (defaults)
- `Destination ip`
- `Server Name`
- `Service` (`SSH` or `RDP`)

### Quick start
```bash
python pam_updated.py   --host bastion.example.local   --api-version v3.12   --username admin   --excel "PAM Access 4.xlsx"   --group "IT Services"   --csv-log output.csv
```
If `--password` is omitted, you'll be prompted.  
You can also use env vars:
```
BASTION_HOST, BASTION_API_VERSION, BASTION_USERNAME, BASTION_PASSWORD
```

### Do I have to keep the Excel filename?
No. Pass any path you want:
```bash
python pam_updated.py --excel "Path_to\MyServers.xlsx"
```
If you have multiple sheets:
```bash
python pam_updated.py --excel "MyServers.xlsx" --sheet "Sheet1"
```

### Common options (CLI)
- `--excel`, `--sheet`
- `--ip-column`, `--name-column`, `--service-column`
- `--group "IT Services"` (use `--no-group` to disable)
- `--description "IT Services"`
- `--insecure` (testing only)
- `--csv-log out.csv`

### Notes
- To support more protocols, extend `build_service_payload()`.
- Group assignment is done with **device/service names**.
- Use `--no-group` if you don't want any group assignment.


# WALLIX Bastion PAM – Excel Import (Devices + Services + Target Group)

**FA | فارسی** — پایین همین فایل نسخهٔ انگلیسی هم هست.

## معرفی
این اسکریپت دستگاه‌ها را از یک فایل Excel می‌خواند، در **WALLIX Bastion PAM** می‌سازد، سرویس‌های مربوطه (SSH/RDP) را اضافه می‌کند و در صورت تمایل، هر دستگاه+سرویس را به یک **Target Group** اضافه می‌کند.

### قابلیت‌ها
- ساخت دستگاه با نامی به‌صورت `ServerName_<دو بخش آخر IP>` مثل `MySrv_34.213`
- افزودن سرویس بر اساس ستون `Service` (پشتیبانی فعلی: **SSH** و **RDP**)
- افزودن به Target Group با **نام دستگاه** و **نام سرویس** (همان روشی که تست شد)
- مدیریت پاسخ‌های `204/409` و واکشی شناسه‌ها با pagination
- پیکربندی کامل از طریق آرگومان‌های CLI یا متغیرهای محیطی
- خروجی CSV اختیاری از نتایج هر ردیف

### پیش‌نیاز
- Python 3.9+  
- نصب وابستگی‌ها:
```bash
pip install -r requirements.txt
```

### ورودی Excel (پیش‌فرض ستون‌ها)
- `Destination ip`
- `Server Name`
- `Service` (مثل `SSH` یا `RDP`)

### اجرای سریع
```bash
python pam_updated.py   --host bastion.example.local   --api-version v3.12   --username admin   --excel "PAM Access 4.xlsx"   --group "IT Services"   --csv-log output.csv
```
اگر `--password` ندهید، هنگام اجرا از شما پرسیده می‌شود.  
می‌توانید از متغیرهای محیطی هم استفاده کنید:
```
BASTION_HOST, BASTION_API_VERSION, BASTION_USERNAME, BASTION_PASSWORD
```

### آیا لازم است اسم فایل Excel حتماً همین باشد؟
خیر. هر نام/مسیر دلخواه را می‌توانید بدهید:
```bash
python pam_updated.py --excel "Path_to\MyServers.xlsx"
```
اگر چند شیت دارید:
```bash
python pam_updated.py --excel "MyServers.xlsx" --sheet "Sheet1"
```

### گزینه‌های پرکاربرد (CLI)
- `--excel` مسیر فایل اکسل
- `--sheet` نام شیت (اختیاری)
- `--ip-column`, `--name-column`, `--service-column` برای تطبیق نام ستون‌ها
- `--group "IT Services"` برای افزودن به گروه (با `--no-group` غیرفعال می‌شود)
- `--description "IT Services"` توضیح دستگاه
- `--insecure` غیرفعال‌کردن بررسی TLS (فقط برای تست)
- `--csv-log out.csv` خروجی لاگ

### یادداشت‌ها
- برای افزودن سرویس‌های جدید، تابع `build_service_payload()` را گسترش دهید.
- افزودن به گروه با **اسم Device** و **اسم Service** انجام می‌شود.
- اگر گروه نمی‌خواهید، `--no-group` را بدهید.

---

