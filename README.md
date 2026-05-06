# الحزب الاشتراكى — Finance App v2.0
## نظام مالي متكامل مع قاعدة بيانات دائمة

---

## 📁 الملفات
```
hizb_finance.html   ← الواجهة الأمامية (Frontend)
server.py           ← الخادم + API + قاعدة البيانات
gen_excel.py        ← توليد ملفات Excel
requirements.txt    ← المكتبات المطلوبة
Procfile            ← إعداد الـ deployment
```

---

## 🚀 تشغيل محلي (Localhost)

```bash
# 1. تثبيت المكتبات
pip install -r requirements.txt

# 2. تشغيل الخادم
python server.py

# 3. افتح المتصفح
http://localhost:5000
```

**المستخدمون الافتراضيون:**
| اسم المستخدم | كلمة المرور |
|---|---|
| admin | hizb2024 |
| محاسب | hizb2024 |

---

## ☁️ رفع على دومين حقيقي

### الخيار 1: Railway.app (مجاني - الأسهل)

1. سجّل على https://railway.app
2. اعمل GitHub repo وحط فيه الملفات
3. Railway → New Project → Deploy from GitHub
4. في الـ Variables أضف:
   - `SECRET_KEY` = أي نص عشوائي طويل
   - `DEBUG` = `false`
5. سيعطيك دومين تلقائي مجاني

### الخيار 2: Render.com (مجاني)

1. سجّل على https://render.com
2. New → Web Service → اربطه بـ GitHub repo
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `gunicorn server:app --bind 0.0.0.0:$PORT`
5. في Environment Variables:
   - `DEBUG` = `false`
   - `DB_PATH` = `/var/data/hizb_finance.db`
6. في Disks: أضف disk على `/var/data` لحفظ قاعدة البيانات

### الخيار 3: VPS (DigitalOcean/Hostinger/etc.)

```bash
# على الـ VPS
git clone <your-repo>
cd hizb-finance
pip install -r requirements.txt

# تشغيل مع gunicorn
gunicorn server:app --bind 0.0.0.0:5000 --daemon

# أو مع nginx + SSL
# استخدم Certbot للـ HTTPS
```

---

## 🔒 إضافة مستخدمين جدد

بعد تشغيل الخادم، من API:
```bash
curl -X POST http://localhost:5000/api/admin/users \
  -H "Content-Type: application/json" \
  -H "X-Auth-Token: <admin-token>" \
  -d '{"username":"user2","password":"pass123","role":"user"}'
```

أو عدّل `server.py` في `init_db()` وأضف المستخدمين في القائمة.

---

## 🗄️ قاعدة البيانات

- تستخدم **SQLite** — ملف واحد `hizb_finance.db`
- كل يوزر بياناته منفصلة تماماً
- نسخ احتياطي:
  ```bash
  cp hizb_finance.db hizb_finance_backup_$(date +%Y%m%d).db
  ```

---

## ⚙️ متغيرات البيئة

| المتغير | القيمة الافتراضية | الوصف |
|---|---|---|
| `PORT` | `5000` | رقم البورت |
| `DB_PATH` | `hizb_finance.db` | مسار قاعدة البيانات |
| `DEBUG` | `true` | وضع التطوير (عطّله في Production) |
