import os
import datetime
import io
import pandas as pd
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, CallbackQueryHandler, filters

TOKEN = "9098342871:AAH589i1e876M3_NYgvlue2YSG3ZSthjbFM"

# Biến toàn cục lưu trữ dữ liệu trực tiếp từ các file bạn gửi qua Telegram (dạng bộ nhớ tạm)
UPLOADED_FILES = {
    'thc': None,  # Lưu nội dung file Tổng hợp chung
    'kh': None,   # Lưu nội dung file Kế hoạch
    'sc': None    # Lưu nội dung file Sự cố
}

def get_main_keyboard():
    keyboard = [
        ["Báo cáo thực hiện ĐTC PCHT", "💡 Biện pháp & Kiến nghị"],
        ["Xuất Excel OMS", "Gửi file Excel"],
        ["📊 Theo dõi bất thường", "Trạng thái"]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Xin chào! Hãy gửi 3 file Excel báo cáo vào khung chat này, tôi sẽ tự động ghi nhớ và xử lý.",
        reply_markup=get_main_keyboard()
    )

def find_month_columns(df_raw, target_month):
    for r_idx in [6, 7, 5]:
        if r_idx < len(df_raw):
            row_vals = df_raw.iloc[r_idx].values
            for c_idx, val in enumerate(row_vals):
                try:
                    if int(val) == target_month:
                        return c_idx, c_idx + 1
                except:
                    pass
    month_mapping = {6: 6, 7: 9, 8: 12, 9: 15, 10: 18, 11: 21, 12: 24}
    base_col = month_mapping.get(target_month, 21)
    return base_col, base_col + 1

def generate_report(mode='thang', filter_type='all'):
    """
    mode: 'thang' hoặc 'luy_ke'
    filter_type: 'all' (Tất cả), 'red' (Vượt KH), 'yellow' (Cảnh báo)
    """
    if UPLOADED_FILES['thc'] is None:
        return "❌ Bạn chưa gửi file Tổng hợp chung (THC) lên Telegram! Vui lòng gửi đủ file trước khi xem báo cáo."
    
    try:
        df_thc = pd.read_excel(io.BytesIO(UPLOADED_FILES['thc']), sheet_name=0, header=None)
        df_kh = pd.read_excel(io.BytesIO(UPLOADED_FILES['kh']), sheet_name=0, header=None) if UPLOADED_FILES['kh'] is not None else None
        df_sc = pd.read_excel(io.BytesIO(UPLOADED_FILES['sc']), sheet_name=0, header=None) if UPLOADED_FILES['sc'] is not None else None
        
        target_month = datetime.datetime.now().month
        th_col, kh_col = (3, 4) if mode == 'luy_ke' else find_month_columns(df_thc, target_month)
        
        period_str = f"LŨY KẾ NĂM" if mode == 'luy_ke' else f"THÁNG {target_month}"
        filter_titles = {
            'all': 'TẤT CẢ CHỈ TIÊU',
            'red': 'DANH SÁCH VƯỢT KẾ HOẠCH (>= 100%)',
            'yellow': 'DANH SÁCH CẢNH BÁO (90% - < 100%)'
        }
        title = f"BÁO CÁO {period_str} - {filter_titles.get(filter_type, '')} - OMS PCHT"
        
        datasets = [
            ("THC lưới điện phân phối", df_thc, ['MAIFI', 'SAIDI', 'SAIFI']),     # 14 đơn vị x 3 = 42
            ("Sự cố trước miễn trừ", df_sc, ['MAIFI', 'SAIDI', 'SAIFI']),        # 14 đơn vị x 3 = 42
            ("Kế hoạch trước miễn trừ", df_kh, ['SAIDI', 'SAIFI'])               # 14 đơn vị x 2 = 28
        ]
        
        category_reports = []
        
        for cat_name, df_raw, allowed_indicators in datasets:
            if df_raw is None:
                category_reports.append(f"📁 **Bộ chỉ tiêu: {cat_name}**\n  ❌ (Chưa nhận được file tương ứng)\n")
                continue
                
            df = df_raw.iloc[8:].reset_index(drop=True)
            red_list = []
            yellow_list = []
            cat_total = 0
            cat_good = 0
            current_unit = "Công ty Điện lực Hà Tĩnh"
            
            for index, row in df.iterrows():
                col1 = row.get(1)
                chi_tieu = str(row.get(2, '')).strip()
                
                if pd.notna(col1) and str(col1).strip() != '':
                    current_unit = str(col1).strip()
                
                if chi_tieu not in allowed_indicators:
                    continue
                
                try:
                    th = float(row.get(th_col, 0))
                    kh = float(row.get(kh_col, 0))
                except:
                    continue
                    
                if kh == 0:
                    continue
                    
                cat_total += 1
                rate = (th / kh) * 100
                item_str = f"{current_unit} - {chi_tieu}: TH {th} / KH {kh} ({rate:.1f}%)"
                
                if rate >= 100.0:
                    excess = round(th - kh, 4)
                    red_list.append(f"• {item_str}, vượt KH {excess}")
                elif 90.0 <= rate < 100.0:
                    diff = round(kh - th, 4)
                    yellow_list.append(f"• {item_str}, còn {diff}")
                else:
                    cat_good += 1
                    
            # Lọc danh sách hiển thị theo tùy chọn của người dùng
            display_red = red_list if filter_type in ['all', 'red'] else []
            display_yellow = yellow_list if filter_type in ['all', 'yellow'] else []
            
            # Nếu chọn xem riêng Vượt KH hoặc Cảnh báo mà mục đó trống, có thể bỏ qua hoặc thông báo
            if filter_type == 'red' and not red_list:
                continue
            if filter_type == 'yellow' and not yellow_list:
                continue

            cat_rate = (cat_good / cat_total * 100) if cat_total > 0 else 0
            
            cat_msg = f"📁 **Bộ chỉ tiêu: {cat_name}**\n"
            if filter_type == 'all':
                cat_msg += f"  • Tổng số: {cat_total} chỉ tiêu | Đạt tốt: {cat_good} ({cat_rate:.1f}%) | Vượt KH: {len(red_list)} | Cảnh báo: {len(yellow_list)}\n"
            
            if display_red:
                cat_msg += "  🔴 Vượt KH (>= 100%):\n    " + "\n    ".join(display_red) + "\n"
            if display_yellow:
                cat_msg += "  🟡 Cảnh báo (90% - < 100%):\n    " + "\n    ".join(display_yellow) + "\n"
            
            if filter_type == 'all' and not red_list and not yellow_list:
                cat_msg += "  🟢 Đạt tốt tất cả các chỉ tiêu (< 90%)\n"
                
            category_reports.append(cat_msg + "\n")

        if not category_reports:
            return f"=== {title} ===\n\n✨ Tuyệt vời! Không có đơn vị nào nằm trong danh mục lọc này."

        msg = f"=== {title} ===\n\n"
        msg += "".join(category_reports)
        return msg
    except Exception as e:
        return f"Lỗi xử lý file dữ liệu: {str(e)}"

def get_recommendation_report():
    if UPLOADED_FILES['thc'] is None:
        return "❌ Chưa nhận được file Tổng hợp chung! Vui lòng gửi file lên Telegram."
        
    try:
        df_thc = pd.read_excel(io.BytesIO(UPLOADED_FILES['thc']), sheet_name=0, header=None)
        target_month = datetime.datetime.now().month
        th_col, kh_col = find_month_columns(df_thc, target_month)
        
        df = df_thc.iloc[8:].reset_index(drop=True)
        units_status = {}
        current_unit = "Công ty Điện lực Hà Tĩnh"
        
        for index, row in df.iterrows():
            col1 = row.get(1)
            chi_tieu = str(row.get(2, '')).strip()
            
            if pd.notna(col1) and str(col1).strip() != '':
                current_unit = str(col1).strip()
            
            if chi_tieu not in ['MAIFI', 'SAIDI', 'SAIFI']:
                continue
            
            try:
                th = float(row.get(th_col, 0))
                kh = float(row.get(kh_col, 0))
            except:
                continue
                
            if kh == 0:
                continue
                
            rate = (th / kh) * 100
            
            if rate >= 90.0:
                if current_unit not in units_status:
                    units_status[current_unit] = {'MAIFI': False, 'SAIDI': False, 'SAIFI': False, 'details': []}
                
                units_status[current_unit][chi_tieu.upper()] = True
                status_text = "Vượt kế hoạch" if rate >= 100.0 else "Cảnh báo sát ngưỡng"
                units_status[current_unit]['details'].append(f"{chi_tieu}: {rate:.1f}% ({status_text} - TH: {th} / KH: {kh})")

        recommendations = []
        for unit, data in units_status.items():
            has_maifi = data.get('MAIFI', False)
            has_said = data.get('SAIDI', False)
            has_saifi = data.get('SAIFI', False)
            
            if has_maifi and (has_said or has_saifi):
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ trên 0 phút để được tối ưu chỉ số MAIFI, SAIDI"
            elif has_maifi and not has_said and not has_saifi:
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ dưới 5 phút để tối ưu hóa chỉ số MAIFI."
            elif (has_said or has_saifi) and not has_maifi:
                solution = "Kiến nghị thay đổi mã nguyên nhân đối với các Mã LMĐ trên 5 phút để tối ưu hóa chỉ số SAIDI, SAIFI."
            else:
                solution = "Rà soát toàn diện và kiến nghị mã lịch miễn trừ theo quy định."

            details_lines = "\n     • ".join(data['details'])
            recommendations.append(f"**{unit}**:\n     • {details_lines}\n     💡 Đề xuất giải pháp: {solution}")

        msg = f"📋 RÀ SOÁT & KIẾN NGHỊ ĐỘ TIN CẬY - THÁNG {target_month}\n"
        msg += "-------------------------------------------------------\n"
        msg += f"⚠️ Tổng số đơn vị cần thực hiện kiến nghị (điều kiện duyệt tự động >= 90%): {len(recommendations)} đơn vị\n\n\n"
        
        if recommendations:
            for i, item in enumerate(recommendations, 1):
                msg += f"{i}. {item}\n\n"
        else:
            msg += f"(Tháng {target_month} hiện tại không có đơn vị nào đạt từ 90% trở lên).\n"
            
        return msg
    except Exception as e:
        return f"Lỗi xử lý kiến nghị: {str(e)}"

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name.lower()
    
    if not file_name.endswith(('.xls', '.xlsx')):
        await update.message.reply_text("⚠️ Vui lòng gửi định dạng file Excel (.xls hoặc .xlsx)!")
        return
        
    file = await context.bot.get_file(document.file_id)
    file_bytes = await file.download_as_bytearray()
    
    if "chung" in file_name or "tong" in file_name or "thc" in file_name:
        UPLOADED_FILES['thc'] = bytes(file_bytes)
        await update.message.reply_text(f"✅ Đã nhận thành công file **Tổng hợp chung (THC)**: `{document.file_name}`")
    elif "ke" in file_name or "hoach" in file_name or "kế hoạch" in file_name:
        UPLOADED_FILES['kh'] = bytes(file_bytes)
        await update.message.reply_text(f"✅ Đã nhận thành công file **Kế hoạch**: `{document.file_name}`")
    elif "su" in file_name or "co" in file_name or "sự cố" in file_name:
        UPLOADED_FILES['sc'] = bytes(file_bytes)
        await update.message.reply_text(f"✅ Đã nhận thành công file **Sự cố**: `{document.file_name}`")
    else:
        UPLOADED_FILES['thc'] = bytes(file_bytes)
        await update.message.reply_text(f"⚠️ Đã nhận file `{document.file_name}` (được gán mặc định làm file THC).")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "Báo cáo thực hiện ĐTC PCHT":
        # Tạo bảng chọn đầy đủ: Cả Tháng & Lũy kế kết hợp với 3 lựa chọn (Tất cả, Vượt KH, Cảnh báo)
        keyboard = [
            [InlineKeyboardButton("📊 Tất cả (Tháng)", callback_data="rep_thang_all"),
             InlineKeyboardButton("📊 Tất cả (Lũy kế)", callback_data="rep_luyke_all")],
            [InlineKeyboardButton("🔴 Vượt KH (Tháng)", callback_data="rep_thang_red"),
             InlineKeyboardButton("🔴 Vượt KH (Lũy kế)", callback_data="rep_luyke_red")],
            [InlineKeyboardButton("🟡 Cảnh báo (Tháng)", callback_data="rep_thang_yellow"),
             InlineKeyboardButton("🟡 Cảnh báo (Lũy kế)", callback_data="rep_luyke_yellow")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("👉 **Vui lòng chọn tiêu chí xem báo cáo ĐTC PCHT:**", reply_markup=reply_markup, parse_mode="Markdown")
        
    elif text == "💡 Biện pháp & Kiến nghị":
        rec_msg = get_recommendation_report()
        await update.message.reply_text(rec_msg, parse_mode="Markdown")
        
    elif text == "Xuất Excel OMS":
        await update.message.reply_text("📁 Tính năng xuất file Excel đang sẵn sàng...")
    elif text == "Gửi file Excel":
        thc_status = "✅ Đã có" if UPLOADED_FILES['thc'] else "❌ Chưa có"
        kh_status = "✅ Đã có" if UPLOADED_FILES['kh'] else "❌ Chưa có"
        sc_status = "✅ Đã có" if UPLOADED_FILES['sc'] else "❌ Chưa có"
        await update.message.reply_text(
            f"📎 **Trạng thái file trong bộ nhớ bot:**\n"
            f"- Tổng hợp chung (THC): {thc_status}\n"
            f"- Kế hoạch: {kh_status}\n"
            f"- Sự cố: {sc_status}\n\n"
            f"👉 Bạn hãy kéo thả trực tiếp 3 file Excel vào khung chat này để cập nhật dữ liệu mới nhất nhé!"
        )
    elif text == "📊 Theo dõi bất thường":
        await update.message.reply_text("⚠️ Đang kiểm tra các điểm vận hành bất thường...")
    elif text == "Trạng thái":
        loaded = sum(1 for v in UPLOADED_FILES.values() if v is not None)
        await update.message.reply_text(f"🟢 Bot đang hoạt động. Số file Excel đã được nạp vào bộ nhớ tạm: **{loaded}/3 file**.")

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    data = query.data
    if data == "rep_thang_all":
        report_content = generate_report(mode='thang', filter_type='all')
    elif data == "rep_thang_red":
        report_content = generate_report(mode='thang', filter_type='red')
    elif data == "rep_thang_yellow":
        report_content = generate_report(mode='thang', filter_type='yellow')
    elif data == "rep_luyke_all":
        report_content = generate_report(mode='luy_ke', filter_type='all')
    elif data == "rep_luyke_red":
        report_content = generate_report(mode='luy_ke', filter_type='red')
    elif data == "rep_luyke_yellow":
        report_content = generate_report(mode='luy_ke', filter_type='yellow')
    else:
        report_content = "⚠️ Lựa chọn không hợp lệ."
        
    await query.message.reply_text(report_content)

if __name__ == '__main__':
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    
    print("Bot đã được cập nhật tính năng phân loại báo cáo (Tất cả / Vượt KH / Cảnh báo)...")
    app.run_polling()