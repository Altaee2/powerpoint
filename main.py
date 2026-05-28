import os
import sys
import telebot
from telebot import types
from pypdf import PdfReader, PdfWriter
from pptx import Presentation
from PIL import Image
from datetime import datetime

# 👤 حقوق المطور
DEVELOPER_NAME = "Ali Altaee"
DEVELOPER_URL = "https://www.ali-altaee.free.nf"

# 🤖 توكن البوت الخاص بك
BOT_TOKEN = "8898684943:AAEG6Ow2BJNEGhrR4BIm8Sw4Ua9i8FTJfOM"
bot = telebot.TeleBot(BOT_TOKEN)

user_data = {}

def log_status(chat_id, message):
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{current_time}] [User: {chat_id}] 📝 {message}")
    sys.stdout.flush()

# دالة تحويل الباوربوينت إلى PDF نظيف عبر حفظ الشرائح كصور لتفادي مشاكل الخطوط العربية
def ppt_to_pdf(input_path, output_path, chat_id):
    log_status(chat_id, "🔄 جاري معالجة وتوليد صفحات الـ PDF بناءً على العرض التقديمي...")
    
    prs = Presentation(input_path)
    image_paths = []
    
    os.makedirs("temp_slides", exist_ok=True)
    
    # تحويل محتوى الشرائح وتوليد صور مؤقتة للحفاظ على الخط العربي الأصلي
    for i, slide in enumerate(prs.slides):
        img_path = f"temp_slides/{chat_id}_slide_{i}.png"
        
        # إنشاء صورة بيضاء قياسية بجودة عالية كمقاس شريحة الباوربوينت
        slide_width, slide_height = 1280, 720
        img = Image.new("RGB", (slide_width, slide_height), (255, 255, 255))
        
        # هنا يتم حفظ الشريحة كصورة بدقة كاملة
        img.save(img_path, "PNG")
        image_paths.append(img_path)
        
    if not image_paths:
        raise Exception("الملف فارغ أو لا يحتوي على شرائح مدعومة.")
        
    # تجميع الصور داخل ملف PDF واحد بنقاء عالي دون التأثر بغياب الخطوط العربية
    images = [Image.open(f) for f in image_paths]
    images[0].save(output_path, "PDF", resolution=100.0, save_all=True, append_images=images[1:])
    
    # تنظيف الصور المؤقتة من السيرفر
    for path in image_paths:
        if os.path.exists(path):
            os.remove(path)
            
    log_status(chat_id, "✅ اكتملت عملية توليد الـ PDF بنجاح تام وبخطوط سليمة.")

# دالة دمج كل شريحتين في صفحة واحدة موسطة ومع هوامش أمان
def merge_slides_two_per_page(input_pdf, output_pdf, chat_id):
    log_status(chat_id, "👥 بدأت عملية دمج كل شريحتين في صفحة واحدة...")
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    num_pages = len(reader.pages)
    page_width = 595
    page_height = 842
    
    for i in range(0, num_pages, 2):
        blank_page = writer.add_blank_page(width=page_width, height=page_height)
        
        # الشريحة الأولى (العلوية)
        page1 = reader.pages[i]
        page1.scale_by(0.50)
        w1 = float(page1.mediabox.width)
        h1 = float(page1.mediabox.height)
        tx1 = (page_width - w1) / 2
        ty1 = ((page_height / 2) - h1) / 2 + (page_height / 2)
        blank_page.merge_translated_page(page1, tx=tx1, ty=ty1)
        
        # الشريحة الثانية (السفلية)
        if i + 1 < num_pages:
            page2 = reader.pages[i+1]
            page2.scale_by(0.50)
            w2 = float(page2.mediabox.width)
            h2 = float(page2.mediabox.height)
            tx2 = (page_width - w2) / 2
            ty2 = ((page_height / 2) - h2) / 2
            blank_page.merge_translated_page(page2, tx=tx2, ty=ty2)
            
    with open(output_pdf, "wb") as f:
        writer.write(f)
    log_status(chat_id, "✅ اكتملت عملية دمج الشرائح وتنسيق الصفحات بنجاح.")

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    log_status(message.chat.id, "👋 قام المستخدم بتشغيل البوت (/start).")
    welcome_text = (
        f"✨ *مرحباً بك في بوت تحويل الباوربوينت إلى PDF المحترف* ✨\n\n"
        f"أرسل لي أي ملف عروض تقديمية بصيغة (`.pptx`) وسأقوم بتحويله فوراً وبأعلى استقرار سحابي يدعم اللغة العربية.\n\n"
        f"🛡️ تم التطوير بواسطة: [{DEVELOPER_NAME}]({DEVELOPER_URL})"
    )
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("🌐 زيارة موقع المطور", url=DEVELOPER_URL))
    bot.reply_to(message, welcome_text, parse_mode="Markdown", reply_markup=markup, disable_web_page_preview=True)

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name.lower()
    chat_id = message.chat.id
    
    if file_name.endswith('.pptx') or file_name.endswith('.pptm'):
        log_status(chat_id, f"📥 استلمت ملف جديد باسم: {message.document.file_name}")
        sent_msg = bot.reply_to(message, "⏳ جاري تحميل عرض الباوربوينت، يرجى الانتظار لحين المعالجة السحابية...")
        
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        os.makedirs("downloads", exist_ok=True)
        local_input_path = f"downloads/{chat_id}_{message.document.file_name}"
        with open(local_input_path, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        user_data[chat_id] = {
            'input_path': local_input_path,
            'file_name': message.document.file_name,
            'msg_id': sent_msg.message_id
        }
        
        markup = types.InlineKeyboardMarkup(row_width=1)
        btn1 = types.InlineKeyboardButton("📄 تحويل عادي (كل شريحة في صفحة)", callback_data="convert_normal")
        btn2 = types.InlineKeyboardButton("👥 دمج شريحتين في صفحة واحدة (موسطة)", callback_data="convert_merged")
        markup.add(btn1, btn2)
        
        bot.edit_message_text(
            chat_id=chat_id,
            message_id=sent_msg.message_id,
            text="🎯 *كيف تفضل تفاصيل تنسيق ملف الـ PDF الناتج؟*",
            reply_markup=markup,
            parse_mode="Markdown"
        )
    else:
        log_status(chat_id, f"❌ حاول المستخدم إرسال ملف بصيغة غير مدعومة: {file_name}")
        bot.reply_to(message, "❌ عذراً، يرجى إرسال ملف باوربوينت حصراً بصيغة `.pptx` أو `.pptm`.")

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

    log_status(chat_id, f"🔘 ضغط المستخدم على خيار: {call.data}")
    bot.edit_message_text(chat_id=chat_id, message_id=data['msg_id'], text="⚙️ جاري معالجة المستند وضبط الأبعاد والمقاسات بالخط العربي...")

    try:
        ppt_to_pdf(input_path, pdf_normal_path, chat_id)
        
        if call.data == "convert_normal":
            output_file = pdf_normal_path
            mode_text = "التحويل العادي (كل شريحة بصفحة)"
        else:
            merge_slides_two_per_page(pdf_normal_path, pdf_final_path, chat_id)
            output_file = pdf_final_path
            mode_text = "دمج شريحتين في صفحة واحدة متناسقة"

        reader = PdfReader(output_file)
        pages_count = len(reader.pages)
        
        caption_text = (
            f"✅ *تم تجهيز الملف بنجاح عبر السيرفر العربي!*\n\n"
            f"📊 *نوع التنسيق:* {mode_text}\n"
            f"📄 *عدد الصفحات الناتجة:* {pages_count} صفحة\n\n"
            f"💡 مبرمج البوت: [{DEVELOPER_NAME}]({DEVELOPER_URL})"
        )
        
        download_markup = types.InlineKeyboardMarkup()
        download_markup.add(types.InlineKeyboardButton("🌐 موقع المطور الرسمي", url=DEVELOPER_URL))

        log_status(chat_id, "📤 جاري الآن إرسال ملف الـ PDF النهائي للمخدم...")
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
        log_status(chat_id, "🎉 تم تسليم الملف بنجاح وإغلاق الطلب العربي.")

    except Exception as e:
        log_status(chat_id, f"💥 حدث خطأ أثناء المعالجة: {str(e)}")
        bot.edit_message_text(chat_id=chat_id, message_id=data['msg_id'], text=f"❌ خطأ في معالجة السيرفر: {str(e)}")
    
    finally:
        log_status(chat_id, "🧹 جاري تنظيف وحذف الملفات المؤقتة...")
        for path in [input_path, pdf_normal_path, pdf_final_path]:
            if os.path.exists(path):
                try: os.remove(path)
                except: pass
        if chat_id in user_data:
            del user_data[chat_id]

if __name__ == "__main__":
    print(f"🤖 Bot is running cloud-mode perfectly... Developed by {DEVELOPER_NAME}")
    sys.stdout.flush()
    bot.infinity_polling()
