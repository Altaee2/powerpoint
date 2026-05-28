import os
import telebot
from telebot import types
from pypdf import PdfReader, PdfWriter
import aspose.slides as slides

# 👤 حقوق المطور
DEVELOPER_NAME = "Ali Altaee"
DEVELOPER_URL = "https://www.ali-altaee.free.nf"

# 🤖 توكن البوت الخاص بك
BOT_TOKEN = "8898684943:AAEG6Ow2BJNEGhrR4BIm8Sw4Ua9i8FTJfOM"
bot = telebot.TeleBot(BOT_TOKEN)

# قاموس مؤفت لتخزين مسارات ملفات المستخدمين وخياراتهم
user_data = {}

# دالة تحويل الباوربوينت إلى PDF السحابية (متوافقة مع لينكس 100%)
def ppt_to_pdf(input_path, output_path):
    with slides.Presentation(input_path) as presentation:
        presentation.save(output_path, slides.export.SaveFormat.PDF)

# دالة دمج كل شريحتين في صفحة واحدة موسطة ومع هوامش أمان
def merge_slides_two_per_page(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    num_pages = len(reader.pages)
    
    # أبعاد صفحة A4 القياسية بالنقاط
    page_width = 595
    page_height = 842
    
    for i in range(0, num_pages, 2):
        # إنشاء صفحة A4 فارغة جديدة
        blank_page = writer.add_blank_page(width=page_width, height=page_height)
        
        # 1. الشريحة الأولى (العلوية)
        page1 = reader.pages[i]
        # الحجم المخصص من قبلك لتناسق الهوامش (0.50)
        page1.scale_by(0.50)
        w1 = float(page1.mediabox.width)
        h1 = float(page1.mediabox.height)
        
        # حساب التوسيط الذكي للشريحة العلوية
        tx1 = (page_width - w1) / 2
        ty1 = ((page_height / 2) - h1) / 2 + (page_height / 2) # النصف العلوي مع توسيط
        
        blank_page.merge_translated_page(page1, tx=tx1, ty=ty1)
        
        # 2. الشريحة الثانية (السفلية) - إن وجدت
        if i + 1 < num_pages:
            page2 = reader.pages[i+1]
            page2.scale_by(0.50)
            w2 = float(page2.mediabox.width)
            h2 = float(page2.mediabox.height)
            
            # حساب التوسيط الذكي للشريحة السفلية
            tx2 = (page_width - w2) / 2
            ty2 = ((page_height / 2) - h2) / 2 # النصف السفلي مع توسيط
            
            blank_page.merge_translated_page(page2, tx=tx2, ty=ty2)
            
    with open(output_pdf, "wb") as f:
        writer.write(f)

# عند إرسال /start أو /help
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    welcome_text = (
        f"✨ *مرحباً بك في بوت تحويل الباوربوينت إلى PDF المحترف* ✨\n\n"
        f"أرسل لي أي ملف يخص عروضك التقديمية بنسخة (`.pptx` أو `.pptm`) وسأقوم بالواجب فوراً وبكل تنسيق وأناقة.\n\n"
        f"🛡️ تم التطوير بواسطة: [{DEVELOPER_NAME}]({DEVELOPER_URL})"
    )
    # زر شفاف لزيارة موقعك مباشرة من رسالة الترحيب
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("🌐 زيارة موقع المطور", url=DEVELOPER_URL))
    
    bot.reply_to(message, welcome_text, parse_mode="Markdown", reply_markup=markup, disable_web_page_preview=True)

# استقبال ملفات الباوربوينت
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name.lower()
    
    if file_name.endswith('.pptx') or file_name.endswith('.pptm'):
        sent_msg = bot.reply_to(message, "⏳ جاري تحميل عرض الباوربوينت، يرجى الانتظار لحين المعالجة السحابية...")
        
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        os.makedirs("downloads", exist_ok=True)
        local_input_path = f"downloads/{message.chat.id}_{message.document.file_name}"
        with open(local_input_path, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        user_data[message.chat.id] = {
            'input_path': local_input_path,
            'file_name': message.document.file_name,
            'msg_id': sent_msg.message_id
        }
        
        # أزرار اختيار نوعية الإخراج
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("📄 تحويل عادي (كل شريحة في صفحة)", callback_data="convert_normal")
        btn2 = types.InlineKeyboardButton("👥 دمج شريحتين in صفحة واحدة (موسطة)", callback_data="convert_merged")
        markup.add(btn1, btn2)
        
        bot.edit_message_text(
            chat_id=message.chat.id,
            message_id=sent_msg.message_id,
            text="🎯 *كيف تفضل تفاصيل تنسيق ملف الـ PDF الناتج؟*",
            reply_markup=markup,
            parse_mode="Markdown"
        )
    else:
        bot.reply_to(message, "❌ عذراً، يرجى إرسال ملف باوربوينت حصراً بصيغة `.pptx` أو `.pptm`.")

# معالجة الضغط على الأزرار الشفافة لخيارات التحويل
@bot.callback_query_handler(func=lambda call: call.data in ["convert_normal", "convert_merged"])
def callback_inline(call):
    chat_id = call.message.chat.id
    if chat_id not in user_data:
        bot.answer_callback_query(call.id, "❌ انتهت الجلسة، يرجى إعادة إرسال الملف.")
        return

    data = user_data[chat_id]
    input_path = data['input_path']
    base_name = os.path.splitext(data['file_name'])[0]
    
    pdf_normal_path = f"downloads/{chat_id}_{base_name}_normal.pdf"
    pdf_final_path = f"downloads/{chat_id}_{base_name}_final.pdf"

    bot.edit_message_text(chat_id=chat_id, message_id=data['msg_id'], text="⚙️ يرجى الانتظار، جاري ضبط المقاسات وتنسيق الهوامش والتحويل السحابي...")

    try:
        # استدعاء دالة التحويل السحابية الجديدة
        ppt_to_pdf(input_path, pdf_normal_path)
        
        if call.data == "convert_normal":
            output_file = pdf_normal_path
            mode_text = "التحويل العادي (كل شريحة بصفحة)"
        else:
            # تطبيق دالة الدمج الموسطة المحدثة
            merge_slides_two_per_page(pdf_normal_path, pdf_final_path)
            output_file = pdf_final_path
            mode_text = "دمج شريحتين في صفحة واحدة متناسقة"

        # قراءة تفاصيل الـ PDF الناتج لمعرفة عدد الصفحات ديماميكياً
        reader = PdfReader(output_file)
        pages_count = len(reader.pages)
        
        # تحضير وصف الملف النهائي الاحترافي
        caption_text = (
            f"✅ *تم تجهيز الملف بنجاح عبر السيرفر!*\n\n"
            f"📊 *نوع التنسيق:* {mode_text}\n"
            f"📄 *عدد الصفحات الناتجة:* {pages_count} صفحة\n\n"
            f"💡 مبرمج البوت: [{DEVELOPER_NAME}]({DEVELOPER_URL})"
        )
        
        # زر شفاف يحتوي على رابط موقعك أسفل المستند المرسل
        download_markup = types.InlineKeyboardMarkup()
        download_markup.add(types.InlineKeyboardButton("🌐 موقع المطور الرسمي", url=DEVELOPER_URL))

        with open(output_file, 'rb') as doc:
            bot.send_document(
                chat_id, 
                doc, 
                caption=caption_text, 
                parse_mode="Markdown",
                reply_markup=download_markup,
                reply_to_message_id=call.message.reply_to_message.message_id
            )
            
        bot.delete_message(chat_id, data['msg_id'])

    except Exception as e:
        bot.edit_message_text(chat_id=chat_id, message_id=data['msg_id'], text=f"❌ خطأ في معالجة السيرفر: {str(e)}")
    
    finally:
        # تنظيف مخلفات السيرفر بذكاء وحذف الملفات المؤقتة
        for path in [input_path, pdf_normal_path, pdf_final_path]:
            if os.path.exists(path):
                try: os.remove(path)
                except: pass
        if chat_id in user_data:
            del user_data[chat_id]

if __name__ == "__main__":
    print(f"🤖 Bot is running cloud-mode perfectly... Developed by {DEVELOPER_NAME}")
    bot.infinity_polling()
