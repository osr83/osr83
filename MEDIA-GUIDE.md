# 📸🎬 دليل استبدال الصور والفيديو

## 📦 مجلد الميديا

```
assets/
├── images/   ← ضع الصور هنا
└── videos/   ← ضع الفيديوهات هنا
```

---

## 📸 استبدال صور المعرض

### 1) جهّز 6 صور للمنتج (1200×900 px أو أعلى)

| الملف | الوصف |
|------|-------|
| `lamp-warm.jpg`   | المصباح بالضوء الذهبي/البرتقالي |
| `lamp-rose.jpg`   | المصباح بالضوء الوردي |
| `lamp-purple.jpg` | المصباح بالضوء الأرجواني |
| `lamp-gold.jpg`   | المصباح بالضوء الذهبي |
| `lamp-room.jpg`   | المصباح في غرفة (Lifestyle) |
| `lamp-box.jpg`    | علبة المنتج / التغليف |

### 2) ضعهم في `assets/images/`

### 3) في ملف `assets/css/style.css`، استبدل قسم `Mood backgrounds`:

```css
.mood-warm   { background: url('../images/lamp-warm.jpg')   center/cover; }
.mood-rose   { background: url('../images/lamp-rose.jpg')   center/cover; }
.mood-purple { background: url('../images/lamp-purple.jpg') center/cover; }
.mood-gold   { background: url('../images/lamp-gold.jpg')   center/cover; }
.mood-room   { background: url('../images/lamp-room.jpg')   center/cover; }
.mood-box    { background: url('../images/lamp-box.jpg')    center/cover; }
```

ثم احذف `.gallery-lamp` من HTML داخل `gallery-main` لأن صورتك الحقيقية كافية.

---

## 🎬 استبدال الفيديو الرئيسي

### 1) جهّز فيديو MP4 بحجم مضغوط (تحت 10MB) — اسمه `demo.mp4`

### 2) ضعه في `assets/videos/`

### 3) في `index.html` ابحث عن `<div class="video-player" id="mainVideo">` واستبدله بـ:

```html
<video class="video-player real-video"
       src="assets/videos/demo.mp4"
       poster="assets/images/lamp-warm.jpg"
       controls
       playsinline
       preload="metadata">
</video>
```

ثم احذف الكود القديم اللي فيه `play-btn` و `vp-bg`.

### 4) في `assets/js/script.js` احذف بلوك `mainVideo.addEventListener` لأن الفيديو راح يشتغل تلقائي.

---

## 🎵 استبدال فيديوهات تيك توك

### الطريقة الأسهل: Embed مباشر من تيك توك

1. روح للفيديو على تيك توك
2. اضغط **Share → Embed**
3. انسخ كود الـ embed
4. الصقه مكان كل `<div class="tt-card">` في HTML

### الطريقة الثانية: استخدام صور تمبنيلز + لينك خارجي

1. حمّل صور تمبنيل من فيديوهاتك (أو استخدم سكرين شوت)
2. ضعها في `assets/images/tiktok/`
3. في CSS:

```css
.tt-card.mood-rose   { background: url('../images/tiktok/video1.jpg') center/cover; }
.tt-card.mood-purple { background: url('../images/tiktok/video2.jpg') center/cover; }
.tt-card.mood-warm   { background: url('../images/tiktok/video3.jpg') center/cover; }
.tt-card.mood-gold   { background: url('../images/tiktok/video4.jpg') center/cover; }
```

4. في HTML اربط الكارد بفيديو تيك توك الحقيقي:

```html
<a href="https://www.tiktok.com/@username/video/123" target="_blank" class="tt-card mood-rose">
  ...
</a>
```

---

## 🛠️ نصائح لتجهيز الصور والفيديو

### للصور:
- **الأبعاد**: 1200×900 px (نسبة 4:3)
- **الصيغة**: JPG (أصغر) أو WebP (أفضل جودة/حجم)
- **الحجم**: تحت 200KB لكل صورة
- **أداة ضغط مجانية**: [tinypng.com](https://tinypng.com)

### للفيديو:
- **الأبعاد**: 1920×1080 (Full HD)
- **المدة**: 30-60 ثانية فقط
- **الصيغة**: MP4 (H.264)
- **الحجم**: تحت 10MB
- **أداة ضغط مجانية**: [handbrake.fr](https://handbrake.fr)

---

## 💡 من وين تحصل صور احترافية للمنتج؟

1. **اطلب من المورّد**: غالباً عند سبلايرز AliExpress / CJ Dropshipping صور احترافية مجانية
2. **صوّر بنفسك**: استخدم آيفون + إضاءة طبيعية + خلفية بيضاء
3. **اطلب UGC**: ادفع 100-300 درهم لكريتور إماراتي على Instagram يصوّر لك
4. **AI Generators**: استخدم Midjourney أو DALL-E بأمر مثل:
   > "Sunset projection lamp on a luxury beige nightstand, warm orange light, aesthetic bedroom, professional product photography"

---

## ✅ Checklist قبل الإطلاق

- [ ] استبدلت 6 صور للمعرض بصور حقيقية
- [ ] أضفت فيديو ديمو 30-60 ثانية
- [ ] ربطت كروت تيك توك بفيديوهات حقيقية
- [ ] ضغطت كل الصور تحت 200KB
- [ ] جربت الموقع على الجوال (Safari + Chrome)
- [ ] ربطت نموذج الطلب بـ Google Sheets أو Shopify
- [ ] غيّرت رقم الواتساب في `index.html`
- [ ] أضفت Pixel تيك توك ومنصات الإعلانات

---

محتاج مساعدة؟ اطلب وأنا أكمل لك أي خطوة 🚀
