"""Generate Excel and PDF files for AKSH GPS C10 device guide (Arabic)."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

import arabic_reshaper
from bidi.algorithm import get_display


# ----- Arabic shaping helper -----
def ar(text: str) -> str:
    """Reshape and bidi-reorder Arabic text for correct rendering in PDF."""
    if not text:
        return text
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


# =====================================================================
# DATA
# =====================================================================
DEVICE_INFO = [
    ("الموديل", "AKSH GPS C10"),
    ("النوع", "جهاز تتبع GPS للمواشي يعمل بالطاقة الشمسية"),
    ("الأبعاد", "137 × 90 × 43 ملم"),
    ("الوزن", "400 جرام"),
    ("البطارية", "16,500 mAh ليثيوم"),
    ("طريقة الشحن", "طاقة شمسية + USB احتياطي (منفذ مغناطيسي)"),
    ("أنظمة التموضع", "Beidou + GPS + LBS"),
    ("درجة الحماية", "IP67 (مقاوم للماء والغبار)"),
    ("درجة حرارة التشغيل", "من -20°C إلى 60°C"),
    ("الشبكة المدعومة", "4G LTE"),
    ("متطلبات الشريحة", "SIM card مع باقة بيانات 100-300 ميجا شهرياً"),
    ("التطبيق", "AKSH GPS (متوفر في App Store و Google Play)"),
    ("كلمة المرور الافتراضية", "123456"),
    ("الموقع الرسمي", "www.akshgps.com"),
]

FEATURES = [
    "تتبع مباشر للموقع في الوقت الفعلي",
    "سجل المسارات السابقة (Historical Route)",
    "تحديد منطقة جغرافية وتنبيه عند الخروج (Geofence)",
    "تنبيه البطارية الضعيفة",
    "تنبيه الحركة (Vibration / Shock Alert)",
    "إرسال أوامر للجهاز عن بُعد",
    "حساب المسافة المقطوعة (Mileage)",
    "تسجيل مواقع التوقف (Parking)",
    "إعدادات التنبيهات المتعددة",
    "شحن ذاتي بالطاقة الشمسية",
]

ACTIVATION_STEPS = [
    ("1", "تركيب شريحة SIM",
     "افتح غطاء فتحة الشريحة بالمفك الصغير المرفق، أدخل الشريحة في مكانها الصحيح، ثم اضغط على الدرج بإحكام وأعد إغلاق الغطاء."),
    ("2", "التأكد من الشريحة",
     "تأكد أن الشريحة مفعّلة، فيها رصيد، باقة بيانات 100-300 ميجا، وملغي قفل الـ PIN (اختبرها بالجوال أولاً)."),
    ("3", "تشغيل الجهاز",
     "اضغط على زر الطاقة لمدة 3 ثوانٍ حتى يضيء المؤشر. لإيقاف التشغيل: اضغط مطولاً 5 ثوانٍ حتى يومض المؤشر 3 مرات ببطء."),
    ("4", "شحن الجهاز",
     "اشحن الجهاز قبل الاستخدام الأول. الضوء الأحمر يعني الجهاز قيد الشحن. اتركه تحت الشمس للشحن المستمر."),
    ("5", "تثبيت الجهاز على الحيوان",
     "ثبّت الجهاز على ظهر أو رقبة الحيوان باستخدام الأحزمة. تأكد أن اللوح الشمسي متجه للأعلى. شدّ الأحزمة جيداً حتى لا ينزلق."),
    ("6", "تحميل التطبيق",
     "ابحث في App Store أو Google Play عن: AKSH GPS، أو امسح رمز QR الموجود في الكتيب."),
    ("7", "تسجيل الدخول",
     "افتح التطبيق، اختر Scan لمسح QR على الجهاز، أو أدخل Device ID يدوياً. كلمة المرور الافتراضية: 123456."),
    ("8", "ربط الجهاز",
     "اختر Add device من الشاشة الرئيسية، أدخل رقم الجهاز أو امسح الـ QR لربطه."),
    ("9", "تغيير كلمة المرور",
     "بعد أول دخول، يُنصح بتغيير كلمة المرور 123456 إلى كلمة مرور قوية والاحتفاظ بها في مكان آمن."),
    ("10", "اختبار التتبع",
     "ضع الجهاز تحت السماء المفتوحة 2-5 دقائق ليلتقط إشارة GPS، ثم تحقق من ظهور موقعه في التطبيق."),
]

# SMS Commands
SMS_COMMANDS = [
    {
        "no": "1",
        "name": "التحقق من حالة الجهاز",
        "command": "status123456",
        "alt": "check123456",
        "description": "يرد عليك الجهاز برسالة فيها: حالة GPS، البطارية، الشبكة، APN، رقم IMEI."
    },
    {
        "no": "2",
        "name": "ضبط APN (إعدادات الإنترنت)",
        "command": "APN123456 اسم_المشغل",
        "alt": "pw,123456,APN,اسم_المشغل#",
        "description": "مهم جداً! بدونه الجهاز لا يقدر يتصل بالإنترنت ولن يظهر في التطبيق."
    },
    {
        "no": "3",
        "name": "ضبط رقم المسؤول (Admin)",
        "command": "admin123456 966xxxxxxxxx",
        "alt": "SOS,A,رقمك#",
        "description": "يجعل رقمك هو المسؤول، وتصلك التنبيهات (بطارية ضعيفة، خروج من المنطقة، حركة)."
    },
    {
        "no": "4",
        "name": "طلب الموقع الحالي عبر SMS",
        "command": "where123456",
        "alt": "url123456",
        "description": "يرسل لك الجهاز رابط Google Maps فيه موقعه الحالي."
    },
    {
        "no": "5",
        "name": "ضبط الخادم (Server)",
        "command": "server123456 1 www.akshgps.com 8001 0",
        "alt": "adminip,123456,IP,Port#",
        "description": "يربط الجهاز بسيرفر التطبيق. استخدمه إذا الجهاز ما اتصل بالتطبيق."
    },
    {
        "no": "6",
        "name": "ضبط فاصل تحديث الموقع",
        "command": "timer123456 30",
        "alt": "upload,30#",
        "description": "الرقم بالثواني. مثال: 30 يعني تحديث كل 30 ثانية. للتوفير: 300 (5 دقائق)."
    },
    {
        "no": "7",
        "name": "تغيير كلمة المرور",
        "command": "password123456 الجديدة",
        "alt": "pw,123456,الجديدة#",
        "description": "غيّر كلمة المرور الافتراضية لحماية الجهاز من العبث."
    },
    {
        "no": "8",
        "name": "إعادة ضبط المصنع",
        "command": "factory123456",
        "alt": "FACTORY",
        "description": "يرجّع كل الإعدادات للوضع الافتراضي. استخدمه عند المشاكل المعقدة."
    },
    {
        "no": "9",
        "name": "إعادة تشغيل الجهاز عن بُعد",
        "command": "reset123456",
        "alt": "RESET#",
        "description": "يعيد تشغيل الجهاز بدون مسح الإعدادات."
    },
    {
        "no": "10",
        "name": "تفعيل تنبيه الحركة",
        "command": "shock123456",
        "alt": "sensoralarm,on#",
        "description": "يرسل تنبيه إذا تحرك الجهاز (مفيد إذا الحيوان سُرق)."
    },
    {
        "no": "11",
        "name": "تفعيل وضع الطاقة المنخفضة",
        "command": "sleep123456 on",
        "alt": "sleep,1#",
        "description": "يوفر استهلاك البطارية. للإلغاء: sleep123456 off"
    },
    {
        "no": "12",
        "name": "ضبط Geofence (المنطقة الجغرافية)",
        "command": "stockade123456 lat,lng,radius",
        "alt": "fence,on,lat,lng,radius#",
        "description": "ينبهك إذا خرج الجهاز من المنطقة المحددة. radius بالأمتار."
    },
    {
        "no": "13",
        "name": "الحصول على رقم IMEI",
        "command": "imei123456",
        "alt": "param#",
        "description": "يرسل لك رقم IMEI للجهاز (يستخدم للتسجيل في بعض المنصات)."
    },
    {
        "no": "14",
        "name": "ضبط المنطقة الزمنية",
        "command": "time123456 zone 3",
        "alt": "GMT,E,3,0#",
        "description": "ضبط التوقيت. للسعودية: 3+ (يعني GMT+3)."
    },
    {
        "no": "15",
        "name": "إيقاف التتبع مؤقتاً",
        "command": "stop123456",
        "alt": "tracker#",
        "description": "يوقف إرسال الموقع. لإعادة التشغيل: tracker123456"
    },
]

APN_LIST = [
    ("STC السعودية", "jawalnet.com.sa", "السعودية"),
    ("موبايلي", "web1", "السعودية"),
    ("زين السعودية", "zain", "السعودية"),
    ("اتصالات الإمارات", "etisalat.ae", "الإمارات"),
    ("دو الإمارات", "du", "الإمارات"),
    ("Ooredoo قطر", "web.ooredoo.com.qa", "قطر"),
    ("Vodafone مصر", "internet.vodafone.net", "مصر"),
    ("Orange مصر", "internet.orange", "مصر"),
    ("اتصالات مصر", "etisalat", "مصر"),
    ("Zain الكويت", "pps", "الكويت"),
    ("Ooredoo الكويت", "action.web", "الكويت"),
    ("Asiacell العراق", "internet", "العراق"),
    ("Zain الأردن", "internet.zain.jo", "الأردن"),
    ("Orange الأردن", "net.orange.jo", "الأردن"),
]

FAQ = [
    ("هل يحتاج الجهاز شريحة SIM؟",
     "نعم، يحتاج شريحة SIM مع باقة بيانات 100-300 ميجا شهرياً للتتبع المستمر."),
    ("كيف يتم شحن الجهاز؟",
     "بالطاقة الشمسية بشكل أساسي + كابل USB مغناطيسي احتياطي. البطارية 16,500 mAh."),
    ("ما هي تقنيات التموضع المدعومة؟",
     "Beidou + GPS + LBS (تموضع عبر أبراج الشبكة) لضمان دقة أعلى."),
    ("هل الجهاز مقاوم للماء؟",
     "نعم، بدرجة حماية IP67 (مقاوم للماء والغبار) ومناسب للبيئات الصعبة."),
    ("كيف أتابع موقع الماشية؟",
     "عبر تطبيق AKSH GPS على الجوال، يعرض الموقع المباشر، السجل، ويسمح بضبط Geofence."),
    ("ماذا أفعل إذا الجهاز ما اشتغل؟",
     "تأكد من شحن البطارية عبر USB أولاً، ثم اضغط زر التشغيل 3 ثوانٍ."),
    ("كيف أحل مشكلة تنبيهات Geofence المتكررة؟",
     "تحقق من إعدادات المنطقة الجغرافية، وتأكد من قوة إشارة الـ GPS."),
    ("هل في رسوم اشتراك شهرية؟",
     "لا يوجد اشتراك للجهاز نفسه، فقط تكلفة باقة البيانات للشريحة."),
    ("ماذا أفعل لو الجهاز ما اتصل بالتطبيق؟",
     "تأكد من ضبط APN صحيح، وأعد تشغيل الجهاز والتطبيق، وتحقق من قوة الإشارة."),
    ("ما هي كلمة المرور الافتراضية؟",
     "123456 — يُنصح بتغييرها بعد أول تسجيل دخول."),
]

INDICATOR_LIGHTS = [
    ("ضوء أحمر ثابت", "الجهاز قيد الشحن"),
    ("ضوء أحمر مطفأ بعد الشحن", "الشحن مكتمل"),
    ("ضوء يضيء عند الضغط القصير", "تأكيد التشغيل"),
    ("3 ومضات بطيئة عند الضغط 5 ثوانٍ", "تأكيد إيقاف التشغيل"),
    ("ضوء يومض بسرعة", "البحث عن إشارة GPS"),
    ("ضوء ثابت", "تم الاتصال بالشبكة"),
]

TROUBLESHOOTING = [
    ("الجهاز لا يستجيب للأوامر",
     "تأكد من شحن البطارية، الشريحة فيها رصيد، PIN ملغي، وقم بتجربة صيغ مختلفة للأوامر."),
    ("الجهاز يستهلك البطارية بسرعة",
     "زيادة فاصل التحديث (timer123456 300)، تفعيل وضع النوم، التأكد من تعرض اللوح الشمسي للشمس."),
    ("الموقع غير دقيق",
     "تأكد أن الجهاز في منطقة مفتوحة، وأن إشارة GPS قوية. تجنب الأماكن المغلقة."),
    ("لا يظهر الجهاز في التطبيق",
     "تأكد من ضبط APN، تأكد من باقة البيانات، أعد تشغيل الجهاز، تحقق من إعدادات Server."),
    ("اللوح الشمسي لا يشحن",
     "نظّف اللوح الشمسي، تأكد من تعرضه للشمس مباشرة، استخدم الشحن بـ USB كحل بديل."),
    ("الأحزمة تنزلق من الحيوان",
     "اختر مقاس مناسب، شدّها جيداً، ثبّت الجهاز على الرقبة بدلاً من الظهر للحيوانات النشطة."),
    ("نسيت كلمة المرور",
     "أرسل أمر إعادة ضبط المصنع: factory123456 (تجربة) أو راجع البائع."),
]


# =====================================================================
# EXCEL GENERATION
# =====================================================================
def create_excel(filepath: str):
    wb = Workbook()

    # Styles
    header_font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    sub_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    cell_font = Font(name="Arial", size=11)
    bold_font = Font(name="Arial", size=11, bold=True)
    thin = Side(border_style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    right_align = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True, readingOrder=2)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ========== Sheet 1: Device Information ==========
    ws1 = wb.active
    ws1.title = "معلومات الجهاز"
    ws1.sheet_view.rightToLeft = True

    ws1.merge_cells("A1:B1")
    ws1["A1"] = "AKSH GPS C10 - دليل جهاز تتبع المواشي"
    ws1["A1"].font = title_font
    ws1["A1"].fill = title_fill
    ws1["A1"].alignment = center_align
    ws1.row_dimensions[1].height = 35

    ws1["A3"] = "البند"
    ws1["B3"] = "المواصفات"
    for cell in ("A3", "B3"):
        ws1[cell].font = header_font
        ws1[cell].fill = header_fill
        ws1[cell].alignment = center_align
        ws1[cell].border = border

    row = 4
    for label, value in DEVICE_INFO:
        ws1.cell(row=row, column=1, value=label).font = bold_font
        ws1.cell(row=row, column=2, value=value).font = cell_font
        for col in (1, 2):
            ws1.cell(row=row, column=col).alignment = right_align
            ws1.cell(row=row, column=col).border = border
            ws1.cell(row=row, column=col).fill = sub_fill if row % 2 == 0 else PatternFill()
        row += 1

    ws1.column_dimensions["A"].width = 30
    ws1.column_dimensions["B"].width = 60

    # ========== Sheet 2: Features ==========
    ws2 = wb.create_sheet("المميزات")
    ws2.sheet_view.rightToLeft = True

    ws2.merge_cells("A1:B1")
    ws2["A1"] = "مميزات الجهاز"
    ws2["A1"].font = title_font
    ws2["A1"].fill = title_fill
    ws2["A1"].alignment = center_align
    ws2.row_dimensions[1].height = 35

    ws2["A3"] = "#"
    ws2["B3"] = "الميزة"
    for cell in ("A3", "B3"):
        ws2[cell].font = header_font
        ws2[cell].fill = header_fill
        ws2[cell].alignment = center_align
        ws2[cell].border = border

    for i, feature in enumerate(FEATURES, 1):
        ws2.cell(row=i + 3, column=1, value=i).font = bold_font
        ws2.cell(row=i + 3, column=2, value=feature).font = cell_font
        for col in (1, 2):
            ws2.cell(row=i + 3, column=col).alignment = right_align if col == 2 else center_align
            ws2.cell(row=i + 3, column=col).border = border

    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 70

    # ========== Sheet 3: Activation Steps ==========
    ws3 = wb.create_sheet("خطوات التفعيل")
    ws3.sheet_view.rightToLeft = True

    ws3.merge_cells("A1:C1")
    ws3["A1"] = "خطوات تفعيل الجهاز بالترتيب"
    ws3["A1"].font = title_font
    ws3["A1"].fill = title_fill
    ws3["A1"].alignment = center_align
    ws3.row_dimensions[1].height = 35

    headers = ["#", "الخطوة", "التفاصيل"]
    for i, h in enumerate(headers, 1):
        c = ws3.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, (no, name, desc) in enumerate(ACTIVATION_STEPS, 1):
        ws3.cell(row=i + 3, column=1, value=no).font = bold_font
        ws3.cell(row=i + 3, column=2, value=name).font = bold_font
        ws3.cell(row=i + 3, column=3, value=desc).font = cell_font
        for col in (1, 2, 3):
            ws3.cell(row=i + 3, column=col).alignment = right_align if col > 1 else center_align
            ws3.cell(row=i + 3, column=col).border = border
        ws3.row_dimensions[i + 3].height = 50

    ws3.column_dimensions["A"].width = 6
    ws3.column_dimensions["B"].width = 25
    ws3.column_dimensions["C"].width = 80

    # ========== Sheet 4: SMS Commands ==========
    ws4 = wb.create_sheet("أوامر SMS")
    ws4.sheet_view.rightToLeft = True

    ws4.merge_cells("A1:E1")
    ws4["A1"] = "أوامر SMS الكاملة لإعداد الجهاز"
    ws4["A1"].font = title_font
    ws4["A1"].fill = title_fill
    ws4["A1"].alignment = center_align
    ws4.row_dimensions[1].height = 35

    headers = ["#", "اسم الأمر", "الأمر الأساسي", "صيغة بديلة", "الشرح"]
    for i, h in enumerate(headers, 1):
        c = ws4.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, cmd in enumerate(SMS_COMMANDS, 1):
        ws4.cell(row=i + 3, column=1, value=cmd["no"]).font = bold_font
        ws4.cell(row=i + 3, column=2, value=cmd["name"]).font = bold_font
        ws4.cell(row=i + 3, column=3, value=cmd["command"]).font = Font(name="Consolas", size=11, color="C00000")
        ws4.cell(row=i + 3, column=4, value=cmd["alt"]).font = Font(name="Consolas", size=10, color="666666")
        ws4.cell(row=i + 3, column=5, value=cmd["description"]).font = cell_font
        for col in range(1, 6):
            cell = ws4.cell(row=i + 3, column=col)
            cell.border = border
            if col in (3, 4):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col == 1:
                cell.alignment = center_align
            else:
                cell.alignment = right_align
            if i % 2 == 0:
                cell.fill = sub_fill
        ws4.row_dimensions[i + 3].height = 55

    ws4.column_dimensions["A"].width = 6
    ws4.column_dimensions["B"].width = 28
    ws4.column_dimensions["C"].width = 35
    ws4.column_dimensions["D"].width = 30
    ws4.column_dimensions["E"].width = 55

    # ========== Sheet 5: APN Settings ==========
    ws5 = wb.create_sheet("إعدادات APN")
    ws5.sheet_view.rightToLeft = True

    ws5.merge_cells("A1:C1")
    ws5["A1"] = "إعدادات APN حسب المشغل"
    ws5["A1"].font = title_font
    ws5["A1"].fill = title_fill
    ws5["A1"].alignment = center_align
    ws5.row_dimensions[1].height = 35

    headers = ["المشغل", "APN", "الدولة"]
    for i, h in enumerate(headers, 1):
        c = ws5.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, (operator, apn, country) in enumerate(APN_LIST, 1):
        ws5.cell(row=i + 3, column=1, value=operator).font = bold_font
        ws5.cell(row=i + 3, column=2, value=apn).font = Font(name="Consolas", size=11, color="C00000")
        ws5.cell(row=i + 3, column=3, value=country).font = cell_font
        for col in (1, 2, 3):
            cell = ws5.cell(row=i + 3, column=col)
            cell.border = border
            cell.alignment = right_align if col != 2 else Alignment(horizontal="left", vertical="center")
            if i % 2 == 0:
                cell.fill = sub_fill

    ws5.column_dimensions["A"].width = 30
    ws5.column_dimensions["B"].width = 35
    ws5.column_dimensions["C"].width = 20

    # ========== Sheet 6: Indicator Lights ==========
    ws6 = wb.create_sheet("مؤشر الإضاءة")
    ws6.sheet_view.rightToLeft = True

    ws6.merge_cells("A1:B1")
    ws6["A1"] = "حالات مؤشر الإضاءة على الجهاز"
    ws6["A1"].font = title_font
    ws6["A1"].fill = title_fill
    ws6["A1"].alignment = center_align
    ws6.row_dimensions[1].height = 35

    headers = ["الحالة", "المعنى"]
    for i, h in enumerate(headers, 1):
        c = ws6.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, (state, meaning) in enumerate(INDICATOR_LIGHTS, 1):
        ws6.cell(row=i + 3, column=1, value=state).font = bold_font
        ws6.cell(row=i + 3, column=2, value=meaning).font = cell_font
        for col in (1, 2):
            cell = ws6.cell(row=i + 3, column=col)
            cell.border = border
            cell.alignment = right_align
            if i % 2 == 0:
                cell.fill = sub_fill

    ws6.column_dimensions["A"].width = 45
    ws6.column_dimensions["B"].width = 50

    # ========== Sheet 7: FAQ ==========
    ws7 = wb.create_sheet("الأسئلة الشائعة")
    ws7.sheet_view.rightToLeft = True

    ws7.merge_cells("A1:B1")
    ws7["A1"] = "الأسئلة الشائعة"
    ws7["A1"].font = title_font
    ws7["A1"].fill = title_fill
    ws7["A1"].alignment = center_align
    ws7.row_dimensions[1].height = 35

    headers = ["السؤال", "الإجابة"]
    for i, h in enumerate(headers, 1):
        c = ws7.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, (q, a) in enumerate(FAQ, 1):
        ws7.cell(row=i + 3, column=1, value=q).font = bold_font
        ws7.cell(row=i + 3, column=2, value=a).font = cell_font
        for col in (1, 2):
            cell = ws7.cell(row=i + 3, column=col)
            cell.border = border
            cell.alignment = right_align
            if i % 2 == 0:
                cell.fill = sub_fill
        ws7.row_dimensions[i + 3].height = 50

    ws7.column_dimensions["A"].width = 45
    ws7.column_dimensions["B"].width = 65

    # ========== Sheet 8: Troubleshooting ==========
    ws8 = wb.create_sheet("حل المشاكل")
    ws8.sheet_view.rightToLeft = True

    ws8.merge_cells("A1:B1")
    ws8["A1"] = "حل المشاكل الشائعة"
    ws8["A1"].font = title_font
    ws8["A1"].fill = title_fill
    ws8["A1"].alignment = center_align
    ws8.row_dimensions[1].height = 35

    headers = ["المشكلة", "الحل"]
    for i, h in enumerate(headers, 1):
        c = ws8.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
        c.border = border

    for i, (problem, solution) in enumerate(TROUBLESHOOTING, 1):
        ws8.cell(row=i + 3, column=1, value=problem).font = bold_font
        ws8.cell(row=i + 3, column=2, value=solution).font = cell_font
        for col in (1, 2):
            cell = ws8.cell(row=i + 3, column=col)
            cell.border = border
            cell.alignment = right_align
            if i % 2 == 0:
                cell.fill = sub_fill
        ws8.row_dimensions[i + 3].height = 60

    ws8.column_dimensions["A"].width = 40
    ws8.column_dimensions["B"].width = 70

    wb.save(filepath)
    print(f"[OK] Excel file created: {filepath}")


# =====================================================================
# PDF GENERATION
# =====================================================================
def create_pdf(filepath: str):
    # Register Arabic fonts
    pdfmetrics.registerFont(TTFont(
        "ArabicFont", "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf"
    ))
    pdfmetrics.registerFont(TTFont(
        "ArabicFontBold", "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf"
    ))

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        rightMargin=1.5 * cm, leftMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )
    story = []

    title_style = ParagraphStyle(
        "Title", fontName="ArabicFontBold", fontSize=22,
        alignment=TA_CENTER, textColor=colors.HexColor("#1F4E78"),
        spaceAfter=10, leading=28,
    )
    subtitle_style = ParagraphStyle(
        "Subtitle", fontName="ArabicFont", fontSize=14,
        alignment=TA_CENTER, textColor=colors.HexColor("#555555"),
        spaceAfter=20, leading=20,
    )
    h1_style = ParagraphStyle(
        "H1", fontName="ArabicFontBold", fontSize=18,
        alignment=TA_RIGHT, textColor=colors.white,
        backColor=colors.HexColor("#1F4E78"),
        borderPadding=(8, 8, 8, 8),
        spaceBefore=15, spaceAfter=12, leading=24,
    )
    h2_style = ParagraphStyle(
        "H2", fontName="ArabicFontBold", fontSize=14,
        alignment=TA_RIGHT, textColor=colors.HexColor("#1F4E78"),
        spaceBefore=10, spaceAfter=8, leading=20,
    )
    body_style = ParagraphStyle(
        "Body", fontName="ArabicFont", fontSize=11,
        alignment=TA_RIGHT, leading=18, spaceAfter=6,
    )
    note_style = ParagraphStyle(
        "Note", fontName="ArabicFont", fontSize=10,
        alignment=TA_RIGHT, leading=16, spaceAfter=4,
        textColor=colors.HexColor("#C00000"),
    )

    # ----- Cover -----
    story.append(Spacer(1, 4 * cm))
    story.append(Paragraph(ar("AKSH GPS C10"), title_style))
    story.append(Paragraph(ar("الدليل الكامل لجهاز تتبع المواشي"), title_style))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(
        ar("جهاز تتبع GPS يعمل بالطاقة الشمسية - الإصدار V1.0"), subtitle_style
    ))
    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph(ar("يحتوي هذا الدليل على:"), h2_style))
    story.append(Paragraph(ar("• معلومات الجهاز ومواصفاته"), body_style))
    story.append(Paragraph(ar("• المميزات والوظائف"), body_style))
    story.append(Paragraph(ar("• خطوات التفعيل خطوة بخطوة"), body_style))
    story.append(Paragraph(ar("• أوامر SMS الكاملة وشرحها"), body_style))
    story.append(Paragraph(ar("• إعدادات APN لكل المشغلين"), body_style))
    story.append(Paragraph(ar("• حل المشاكل الشائعة"), body_style))
    story.append(Paragraph(ar("• الأسئلة الشائعة"), body_style))
    story.append(PageBreak())

    # ----- Section 1: Device Info -----
    story.append(Paragraph(ar("1. معلومات الجهاز ومواصفاته"), h1_style))
    data = [[Paragraph(ar("المواصفات"), ParagraphStyle(
        "h", fontName="ArabicFontBold", fontSize=11, alignment=TA_RIGHT, textColor=colors.white)),
             Paragraph(ar("البند"), ParagraphStyle(
        "h", fontName="ArabicFontBold", fontSize=11, alignment=TA_RIGHT, textColor=colors.white))]]
    for label, value in DEVICE_INFO:
        data.append([
            Paragraph(ar(value), body_style),
            Paragraph(ar(label), ParagraphStyle("l", fontName="ArabicFontBold", fontSize=11, alignment=TA_RIGHT))
        ])
    t = Table(data, colWidths=[11 * cm, 6 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#DDEBF7")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ----- Section 2: Features -----
    story.append(Paragraph(ar("2. المميزات والوظائف"), h1_style))
    for i, feat in enumerate(FEATURES, 1):
        story.append(Paragraph(ar(f"{i}. {feat}"), body_style))
    story.append(Spacer(1, 0.5 * cm))

    # ----- Section 3: Activation Steps -----
    story.append(Paragraph(ar("3. خطوات التفعيل بالترتيب"), h1_style))
    for no, name, desc in ACTIVATION_STEPS:
        story.append(Paragraph(ar(f"الخطوة {no}: {name}"), h2_style))
        story.append(Paragraph(ar(desc), body_style))
    story.append(PageBreak())

    # ----- Section 4: SMS Commands -----
    story.append(Paragraph(ar("4. أوامر SMS الكاملة"), h1_style))
    story.append(Paragraph(
        ar("ملاحظة: كلمة المرور الافتراضية 123456. أرسل الأوامر إلى رقم الشريحة الموجودة في الجهاز."),
        note_style
    ))
    story.append(Spacer(1, 0.3 * cm))

    cmd_header_style = ParagraphStyle(
        "ch", fontName="ArabicFontBold", fontSize=11, alignment=TA_CENTER, textColor=colors.white
    )
    cmd_cell_style = ParagraphStyle(
        "cc", fontName="ArabicFont", fontSize=9, alignment=TA_RIGHT, leading=14
    )
    cmd_code_style = ParagraphStyle(
        "code", fontName="Helvetica-Bold", fontSize=9, alignment=TA_LEFT,
        textColor=colors.HexColor("#C00000"), leading=14
    )

    cmd_data = [[
        Paragraph(ar("الشرح"), cmd_header_style),
        Paragraph(ar("الأمر"), cmd_header_style),
        Paragraph(ar("الاسم"), cmd_header_style),
        Paragraph("#", cmd_header_style),
    ]]
    for cmd in SMS_COMMANDS:
        cmd_data.append([
            Paragraph(ar(cmd["description"]), cmd_cell_style),
            Paragraph(cmd["command"], cmd_code_style),
            Paragraph(ar(cmd["name"]), ParagraphStyle(
                "n", fontName="ArabicFontBold", fontSize=9, alignment=TA_RIGHT, leading=14)),
            Paragraph(cmd["no"], ParagraphStyle(
                "no", fontName="Helvetica-Bold", fontSize=10, alignment=TA_CENTER)),
        ])
    ct = Table(cmd_data, colWidths=[7 * cm, 4.5 * cm, 4.5 * cm, 1 * cm], repeatRows=1)
    ct.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#DDEBF7")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(ct)
    story.append(PageBreak())

    # ----- Section 5: APN Settings -----
    story.append(Paragraph(ar("5. إعدادات APN حسب المشغل"), h1_style))
    apn_data = [[
        Paragraph(ar("الدولة"), cmd_header_style),
        Paragraph(ar("APN"), cmd_header_style),
        Paragraph(ar("المشغل"), cmd_header_style),
    ]]
    for op, apn, country in APN_LIST:
        apn_data.append([
            Paragraph(ar(country), body_style),
            Paragraph(apn, cmd_code_style),
            Paragraph(ar(op), ParagraphStyle(
                "n", fontName="ArabicFontBold", fontSize=11, alignment=TA_RIGHT)),
        ])
    apt = Table(apn_data, colWidths=[4 * cm, 6 * cm, 7 * cm])
    apt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#DDEBF7")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(apt)
    story.append(Spacer(1, 0.5 * cm))

    # ----- Section 6: Indicator Lights -----
    story.append(Paragraph(ar("6. حالات مؤشر الإضاءة"), h1_style))
    light_data = [[
        Paragraph(ar("المعنى"), cmd_header_style),
        Paragraph(ar("الحالة"), cmd_header_style),
    ]]
    for state, meaning in INDICATOR_LIGHTS:
        light_data.append([
            Paragraph(ar(meaning), body_style),
            Paragraph(ar(state), ParagraphStyle(
                "n", fontName="ArabicFontBold", fontSize=11, alignment=TA_RIGHT)),
        ])
    lt = Table(light_data, colWidths=[9 * cm, 8 * cm])
    lt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#DDEBF7")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(lt)
    story.append(PageBreak())

    # ----- Section 7: Troubleshooting -----
    story.append(Paragraph(ar("7. حل المشاكل الشائعة"), h1_style))
    for problem, solution in TROUBLESHOOTING:
        story.append(Paragraph(ar(f"المشكلة: {problem}"), h2_style))
        story.append(Paragraph(ar(f"الحل: {solution}"), body_style))
    story.append(PageBreak())

    # ----- Section 8: FAQ -----
    story.append(Paragraph(ar("8. الأسئلة الشائعة"), h1_style))
    for q, a in FAQ:
        story.append(Paragraph(ar(f"س: {q}"), h2_style))
        story.append(Paragraph(ar(f"ج: {a}"), body_style))

    story.append(Spacer(1, 1 * cm))
    story.append(Paragraph(ar("--- نهاية الدليل ---"), subtitle_style))

    doc.build(story)
    print(f"[OK] PDF file created: {filepath}")


if __name__ == "__main__":
    out_dir = "/home/user/osr83"
    create_excel(os.path.join(out_dir, "AKSH_GPS_C10_Guide.xlsx"))
    create_pdf(os.path.join(out_dir, "AKSH_GPS_C10_Guide.pdf"))
    print("\n[DONE] All files generated successfully.")
