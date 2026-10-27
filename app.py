
import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
import openpyxl

app = Flask(__name__, template_folder='template')
app.secret_key = 'teacher-salary-secret-key-12345'

UPLOAD_FOLDER = 'uploads'
EXCEL_FILENAME = 'salary_data.xlsx'
PASS_FILE = 'password.txt'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

if not os.path.exists(PASS_FILE):
    with open(PASS_FILE, 'w') as f:
        f.write('Saber456519')

def get_password():
    with open(PASS_FILE, 'r') as f:
        return f.read().strip()

def get_excel_path():
    return os.path.join(app.config['UPLOAD_FOLDER'], EXCEL_FILENAME)

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    error = None
    search_code = None
    
    if request.method == 'POST':
        search_code = request.form.get('personnel_code', '').strip()
        excel_path = get_excel_path()
        
        if not os.path.exists(excel_path):
            error = "هنوز فایل اطلاعات توسط مدیریت بارگذاری نشده است."
        elif not search_code:
            error = "لطفاً کد پرسنلی را وارد کنید."
        else:
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=True)
                sheet = wb.active
                
                # خواندن سرستون‌ها
                headers = [str(cell.value).strip() if cell.value is not None else '' for cell in sheet[1]]
                
                if 'کد پرسنلی' not in headers:
                    error = "ستون 'کد پرسنلی' در فایل اکسل یافت نشد."
                else:
                    found = False
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        row_dict = {headers[i]: (row[i] if i < len(row) and row[i] is not None else '') for i in range(len(headers))}
                        current_code = str(row_dict.get('کد پرسنلی', '')).strip()
                        # حذف اعشار احتمالی از کدهای عددی
                        if current_code.endswith('.0'):
                            current_code = current_code[:-2]
                            
                        if current_code == search_code:
                            result = {
                                'نام': row_dict.get('نام', 'نامشخص'),
                                'نام_خانوادگی': row_dict.get('نام خانوادگی', 'نامشخص'),
                                'ساعت_محاسبه_شده': row_dict.get('ساعت محاسبه شده', 'نامشخص'),
                                'ساعت_ماهانه': row_dict.get('ساعت ماهانه حق التدریس', 'نامشخص'),
                                'نرخ': row_dict.get('نرخ حق التدریس', 'نامشخص'),
                                'مقدار': row_dict.get('مقدار حق التدریس', 'نامشخص'),
                                'خالص': row_dict.get('خالص پرداختی', 'نامشخص')
                            }
                            found = True
                            break
                    if not found:
                        error = "کد پرسنلی مورد نظر یافت نشد."
            except Exception as e:
                error = f"خطا در پردازش فایل: {str(e)}"
                
    return render_template('index.html', result=result, error=error, search_code=search_code)

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        entered_pass = request.form.get('password')
        if entered_pass == get_password():
            session['admin_logged'] = True
            return redirect(url_for('admin_panel'))
        else:
            flash('رمز عبور اشتباه است.', 'danger')
    return render_template('admin_login.html')

@app.route('/admin', methods=['GET', 'POST'])
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
        
    if request.method == 'POST':
        action = request.form.get('action')
        
        if action == 'change_pass':
            new_pass = request.form.get('new_password', '').strip()
            if len(new_pass) >= 4:
                with open(PASS_FILE, 'w') as f:
                    f.write(new_pass)
                flash('رمز عبور با موفقیت تغییر کرد.', 'success')
            else:
                flash('رمز عبور باید حداقل 4 کاراکتر باشد.', 'danger')
                
        elif action == 'upload_excel':
            if 'excel_file' not in request.files:
                flash('فایلی انتخاب نشده است.', 'danger')
            else:
                file = request.files['excel_file']
                if file.filename == '':
                    flash('فایلی انتخاب نشده است.', 'danger')
                elif file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                    file.save(get_excel_path())
                    flash('فایل اکسل با موفقیت آپلود و جایگزین شد.', 'success')
                else:
                    flash('فقط فرمت اکسل (.xlsx یا .xls) مجاز است.', 'danger')
                    
        return redirect(url_for('admin_panel'))
        
    return render_template('admin.html')

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
