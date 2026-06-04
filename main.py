"""
PDF发票名称提取+邮件批量发送 v2.0
基于原版exe反编译逻辑，使用 PyQt5 重构UI，修复正则表达式BUG
"""

import sys
import os
import re
import csv
import zipfile
import shutil
import time
import random
import smtplib
import traceback
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
import pdfplumber
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QGridLayout, QLabel, QLineEdit, QPushButton, QProgressBar,
    QTextEdit, QFileDialog, QMessageBox, QFrame, QScrollArea,
    QSizePolicy, QSpacerItem, QToolTip
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QTimer
from PyQt5.QtGui import QFont, QIcon, QColor

# ============ 样式表 ============
QSS = """
/* 全局 */
QMainWindow {
    background-color: #f0f2f5;
}
QWidget {
    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
    font-size: 13px;
}

/* 标题栏 */
#titleBar {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #667eea, stop:1 #764ba2);
    border-radius: 0px;
    padding: 12px 20px;
}
#titleIcon { font-size: 22px; }
#titleLabel {
    color: white;
    font-size: 16px;
    font-weight: bold;
}
#versionLabel {
    color: rgba(255,255,255,0.7);
    font-size: 11px;
}

/* 分组框 */
QGroupBox {
    background: white;
    border: 1px solid #e8eaed;
    border-radius: 10px;
    margin-top: 14px;
    padding: 20px 16px 16px 16px;
    font-weight: bold;
    color: #333;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 4px 12px;
    background: white;
    border: 1px solid #e8eaed;
    border-radius: 6px;
    left: 16px;
}

/* 输入框 */
QLineEdit {
    background: #f8f9fa;
    border: 1px solid #dcdfe6;
    border-radius: 6px;
    padding: 8px 12px;
    font-size: 13px;
    color: #333;
}
QLineEdit:focus {
    border-color: #667eea;
    background: white;
}
QLineEdit:disabled {
    background: #eee;
    color: #999;
}

/* 按钮 */
QPushButton {
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-size: 13px;
    font-weight: 500;
}
QPushButton:hover { opacity: 0.9; }

#btnPrimary {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #667eea, stop:1 #764ba2);
    color: white;
    padding: 10px 28px;
    font-size: 14px;
}
#btnPrimary:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #5a6ed6, stop:1 #6a4190);
}
#btnPrimary:disabled {
    background: #ccc;
    color: #999;
}

#btnOutline {
    background: white;
    border: 1px solid #667eea;
    color: #667eea;
}
#btnOutline:hover {
    background: #667eea;
    color: white;
}

#btnDanger {
    background: white;
    border: 1px solid #f56c6c;
    color: #f56c6c;
}
#btnDanger:hover {
    background: #f56c6c;
    color: white;
}

#btnBrowse {
    background: #667eea;
    color: white;
    padding: 8px 12px;
    min-width: 60px;
}

#btnAdd {
    background: #67c23a;
    color: white;
    padding: 6px 14px;
}
#btnAdd:hover { background: #5daf34; }

#btnRemove {
    background: #f56c6c;
    color: white;
    padding: 4px 10px;
    font-size: 12px;
    min-width: 36px;
}
#btnRemove:hover { background: #e04545; }

/* 进度条 */
QProgressBar {
    border: none;
    border-radius: 8px;
    background: #e8eaed;
    height: 22px;
    text-align: center;
    font-size: 11px;
    color: #333;
}
QProgressBar::chunk {
    border-radius: 8px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #667eea, stop:1 #764ba2);
}

/* 日志 */
#logArea {
    background: #1e1e2e;
    color: #cdd6f4;
    border: 1px solid #313244;
    border-radius: 8px;
    padding: 10px;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 12px;
}

/* 状态标签 */
#statusLabel {
    color: #666;
    font-size: 12px;
    padding: 4px 0;
}

/* 分割线 */
#separator {
    background: #e8eaed;
    max-height: 1px;
}

/* 滚动区域 */
QScrollArea {
    border: none;
    background: transparent;
}

/* 窗口控制按钮 */
#btnMinimize, #btnMaximize, #btnClose {
    background: transparent;
    color: white;
    border: none;
    font-size: 16px;
    padding: 6px 10px;
    border-radius: 0px;
    min-width: 34px;
}
#btnMinimize:hover, #btnMaximize:hover {
    background: rgba(255,255,255,0.15);
}
#btnClose:hover {
    background: #e81123;
    color: white;
}
"""

# ============ 核心处理线程 ============
class ProcessThread(QThread):
    log_signal = pyqtSignal(str, str)          # (message, level)
    progress_signal = pyqtSignal(int)          # 0-100
    status_signal = pyqtSignal(str)            # 状态文字
    finished_signal = pyqtSignal(bool, str)    # (success, message)

    def __init__(self, zip_path, output_dir, excel_path, mail_accounts):
        super().__init__()
        self.zip_path = zip_path
        self.output_dir = output_dir
        self.excel_path = excel_path
        self.mail_accounts = mail_accounts

    def log(self, msg, level='info'):
        self.log_signal.emit(msg, level)

    def run(self):
        try:
            self.process()
        except Exception as e:
            self.log(f'处理出错: {e}', 'error')
            traceback.print_exc()
            self.finished_signal.emit(False, str(e))

    def process(self):
        # 步骤1: 解压
        self.log('━━━ 步骤1: 解压ZIP文件 ━━━', 'title')
        temp_dir = os.path.join(self.output_dir, 'temp_extract')
        os.makedirs(temp_dir, exist_ok=True)
        with zipfile.ZipFile(self.zip_path, 'r') as z:
            z.extractall(temp_dir)
        pdf_files = [f for f in os.listdir(temp_dir) if f.lower().endswith('.pdf')]
        total_pdf = len(pdf_files)
        self.log(f'解压完成，共 {total_pdf} 个PDF文件')

        # 步骤2: 提取购买方名称
        self.log('━━━ 步骤2: 提取购买方名称 ━━━', 'title')
        self.progress_signal.emit(5)
        self.status_signal.emit('正在提取购买方名称...')

        # 修复后的正则：匹配 "购 名称：XXX" / "购 称：XXX" / "购买方名称: XXX" 等格式
        # 捕获组允许空格，解决"赵 亮"截断问题（后续用 re.sub 清理空格）
        purchaser_pattern = re.compile(
            r'购\s*(?:买\s*方\s*)?名\s*[称:：]\s*[：:]?\s*'
            r'([\u4e00-\u9fa5a-zA-Z0-9（）\s]+?)(?=\s*销|\s*$)'
        )

        purchaser_names = []
        for i, pdf_file in enumerate(pdf_files):
            pdf_path = os.path.join(temp_dir, pdf_file)
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    found = False
                    for page in pdf.pages:
                        text = page.extract_text() or ''
                        # 合并换行方便匹配
                        text_merged = text.replace('\n', ' ')
                        match = purchaser_pattern.search(text_merged)
                        if match:
                            name = match.group(1).strip()
                            # 清理多余空格
                            name = re.sub(r'\s+', '', name)
                            purchaser_names.append((name, pdf_file))
                            self.log(f'  [{i+1}/{total_pdf}] {name} ← {pdf_file}')
                            found = True
                            break
                    if not found:
                        self.log(f'  [{i+1}/{total_pdf}] ⚠ 未提取到购买方名称: {pdf_file}', 'warn')
            except Exception as e:
                self.log(f'  [{i+1}/{total_pdf}] ❌ 处理出错: {pdf_file} | {e}', 'error')

        self.log(f'成功提取 {len(purchaser_names)} 个购买方名称')
        self.progress_signal.emit(30)

        # 步骤3: 匹配Excel数据
        self.log('━━━ 步骤3: 匹配Excel数据 ━━━', 'title')
        self.status_signal.emit('正在匹配Excel数据...')
        excel_data = pd.read_excel(self.excel_path)
        target_col_index = 12  # 第13列（索引12）是邮箱

        self.log(f'Excel: {excel_data.shape[0]}行 × {excel_data.shape[1]}列 | 邮箱列: 第{target_col_index+1}列')

        csv_path = os.path.join(self.output_dir, '购买方名称.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['购买方名称', '匹配结果'])
            for name, pdf_file in purchaser_names:
                search_value = name.strip()
                last_match = None
                # 从后往前搜索
                for idx in reversed(excel_data.index):
                    excel_row = excel_data.loc[idx]
                    if any(search_value in str(cell) for cell in excel_row):
                        last_match = excel_row
                        break
                target_value = '未找到匹配内容'
                if last_match is not None and target_col_index < len(last_match):
                    raw_val = last_match.iloc[target_col_index]
                    if pd.notna(raw_val):
                        target_value = str(raw_val)
                writer.writerow([search_value, target_value])
                status_icon = '✅' if '@' in target_value else '⚠'
                self.log(f'  {status_icon} {search_value} → {target_value}')

        # 读取收件人列表
        with open(csv_path, 'r', encoding='utf-8-sig') as f2:
            reader = csv.reader(f2)
            next(reader)
            recipient_list = [(row[0].strip(), row[1].strip()) for row in reader
                              if '@' in row[1] and row[1] != '未找到匹配内容']

        self.log(f'有效收件人: {len(recipient_list)} 人')
        self.progress_signal.emit(50)

        if not recipient_list:
            self.log('没有找到有效收件人，处理结束', 'warn')
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.progress_signal.emit(100)
            self.status_signal.emit('完成 — 无有效收件人')
            self.finished_signal.emit(True, '完成 — 请检查Excel登记表和PDF的购买方名称是否匹配')
            return

        # 步骤4: 发送邮件
        self.log('━━━ 步骤4: 发送邮件 ━━━', 'title')
        batch_size = 9
        pause_time = 60
        total_emails = len(recipient_list)

        for batch_num, i in enumerate(range(0, total_emails, batch_size)):
            batch = recipient_list[i:i + batch_size]
            # 每批次随机选取一个账号，打散避免单账号密集发送
            mail_config = random.choice(self.mail_accounts)

            self.log(f'--- 批次 {batch_num+1}: 使用账号 {mail_config["email_from"]} '
                     f'发送 {len(batch)} 封 ---')

            try:
                server = smtplib.SMTP_SSL(mail_config['smtp_server'], mail_config['smtp_port'], timeout=30)
                server.login(mail_config['email_from'], mail_config['auth_code'])
            except Exception as e:
                self.log(f'❌ SMTP登录失败 (账号 {mail_config["email_from"]}): {e}', 'error')
                # 随机尝试其他账号
                remaining = [a for a in self.mail_accounts if a != mail_config]
                if remaining:
                    mail_config = random.choice(remaining)
                    self.log(f'切换至备用账号: {mail_config["email_from"]}')
                    try:
                        server = smtplib.SMTP_SSL(mail_config['smtp_server'], mail_config['smtp_port'], timeout=30)
                        server.login(mail_config['email_from'], mail_config['auth_code'])
                    except Exception as e2:
                        self.log(f'❌ 备用账号也失败: {e2}', 'error')
                        continue
                else:
                    continue

            for j, (name, email_addr) in enumerate(batch):
                global_idx = i + j + 1
                progress = 50 + int((global_idx / total_emails) * 45)
                self.progress_signal.emit(progress)
                self.status_signal.emit(f'正在发送邮件 ({global_idx}/{total_emails})...')

                # 找对应PDF
                matched_files = []
                for fname in os.listdir(temp_dir):
                    if fname.lower().endswith('.pdf') and name in fname:
                        matched_files.append(os.path.join(temp_dir, fname))
                if not matched_files:
                    # 模糊匹配
                    for fname in os.listdir(temp_dir):
                        if fname.lower().endswith('.pdf'):
                            matched_files.append(os.path.join(temp_dir, fname))
                            break

                if not matched_files:
                    self.log(f'  [{global_idx}] ⚠ {name} → 找不到PDF文件', 'warn')
                    continue

                # 构建邮件
                msg = MIMEMultipart()
                msg['From'] = mail_config['email_from']
                msg['To'] = email_addr
                msg['Subject'] = '发票文件 - 请查收'
                body = MIMEText('您好，附件是您的发票文件，请查收。', 'plain', 'utf-8')
                msg.attach(body)

                for file_path in matched_files:
                    with open(file_path, 'rb') as fp:
                        part = MIMEApplication(fp.read(), Name=os.path.basename(file_path))
                    part['Content-Disposition'] = (
                        f'attachment; filename="{os.path.basename(file_path)}"'
                    )
                    msg.attach(part)

                try:
                    server.sendmail(mail_config['email_from'], email_addr, msg.as_string())
                    self.log(f'  [{global_idx}] ✅ 发送成功: {name} → {email_addr}', 'success')
                except Exception as e:
                    self.log(f'  [{global_idx}] ❌ 发送失败: {name} → {email_addr} | {e}', 'error')

            server.quit()

            # 批次间暂停
            if i + batch_size < total_emails:
                self.log(f'已发送 {min(i + batch_size, total_emails)} 封，暂停 {pause_time} 秒并切换账号...')
                time.sleep(5)  # 测试时用5秒

        # 清理
        shutil.rmtree(temp_dir, ignore_errors=True)
        self.progress_signal.emit(100)
        self.status_signal.emit('完成 — 所有处理已完成！')
        self.log('━━━ ✅ 全部完成！━━━', 'success')
        self.finished_signal.emit(True, '所有处理已完成！')


# ============ 邮件账号行组件 ============
class EmailAccountRow(QWidget):
    removed = pyqtSignal(object)

    def __init__(self, index, email='', auth='', parent=None):
        super().__init__(parent)
        self.index = index
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 4)
        layout.setSpacing(8)

        self.num_label = QLabel(f'<b>#{index + 1}</b>')
        self.num_label.setFixedWidth(30)
        self.num_label.setAlignment(Qt.AlignCenter)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText('QQ邮箱地址')
        self.email_input.setText(email)
        self.email_input.setMinimumWidth(180)

        self.auth_input = QLineEdit()
        self.auth_input.setPlaceholderText('授权码')
        self.auth_input.setEchoMode(QLineEdit.Password)
        self.auth_input.setText(auth)

        self.status_label = QLabel('⏳')
        self.status_label.setFixedWidth(30)
        self.status_label.setAlignment(Qt.AlignCenter)
        if email and auth:
            self.status_label.setText('🟢')
            self.status_label.setToolTip('已填写')

        self.remove_btn = QPushButton('✕')
        self.remove_btn.setObjectName('btnRemove')
        self.remove_btn.setToolTip('删除此行')
        self.remove_btn.clicked.connect(lambda: self.removed.emit(self))

        self.clear_btn = QPushButton('清空')
        self.clear_btn.setObjectName('btnDanger')
        self.clear_btn.setToolTip('清空此行的邮箱和授权码')
        self.clear_btn.setFixedWidth(40)
        self.clear_btn.setStyleSheet('font-size:11px; padding:2px 6px;')
        self.clear_btn.clicked.connect(self.clear_row)

        layout.addWidget(self.num_label)
        layout.addWidget(self.email_input)
        layout.addWidget(self.auth_input)
        layout.addWidget(self.status_label)
        layout.addWidget(self.remove_btn)
        layout.addWidget(self.clear_btn)

    def clear_row(self):
        self.email_input.clear()
        self.auth_input.clear()
        self.status_label.setText('⏳')
        self.status_label.setToolTip('')

    def get_data(self):
        return {
            'email_from': self.email_input.text().strip(),
            'smtp_server': 'smtp.qq.com',
            'smtp_port': 465,
            'auth_code': self.auth_input.text().strip()
        }

    def is_valid(self):
        email = self.email_input.text().strip()
        auth = self.auth_input.text().strip()
        return bool(email and auth and '@' in email)


# ============ 主窗口 ============
class InvoiceMailerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PDF发票名称提取+邮件批量发送 v2.0')
        self.setMinimumSize(720, 800)
        self.resize(760, 860)
        self.setStyleSheet(QSS)

        self.account_rows = []
        self.is_processing = False
        self.output_dir = os.path.dirname(os.path.abspath(__file__))

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- 标题栏（自定义无边框窗口） ---
        self.setWindowFlags(Qt.FramelessWindowHint)
        # 添加阴影边框效果
        self.setAttribute(Qt.WA_TranslucentBackground, False)

        title_bar = QFrame()
        title_bar.setObjectName('titleBar')
        title_bar.setFixedHeight(60)
        # 标题栏支持拖拽移动窗口
        title_bar.mousePressEvent = self.title_bar_mouse_press
        title_bar.mouseMoveEvent = self.title_bar_mouse_move

        title_layout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(20, 0, 8, 0)

        icon_label = QLabel('📧')
        icon_label.setObjectName('titleIcon')
        title_label = QLabel('PDF发票批量发送工具')
        title_label.setObjectName('titleLabel')
        title_label.setCursor(Qt.PointingHandCursor)
        title_label.setMouseTracking(True)
        # 自定义悬浮/点击行为
        title_label.enterEvent = self.title_label_enter
        title_label.leaveEvent = self.title_label_leave
        title_label.mouseReleaseEvent = self.title_label_click
        version_label = QLabel('v2.0')
        version_label.setObjectName('versionLabel')

        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_label)
        title_layout.addWidget(version_label)
        title_layout.addStretch()

        # 悬浮信息卡片（默认隐藏）
        self.hover_card = QFrame(self)
        self.hover_card.setObjectName('hoverCard')
        self.hover_card.setStyleSheet("""
            #hoverCard {
                background: #ffffff;
                border: 1px solid #dcdfe6;
                border-radius: 6px;
                padding: 4px 20px;
            }
            #hoverCard QLabel {
                color: #333;
                font-size: 12px;
                font-family: "Microsoft YaHei";
            }
        """)
        hover_layout = QHBoxLayout(self.hover_card)
        hover_layout.setContentsMargins(10, 2, 10, 2)
        hover_layout.setSpacing(0)
        hover_label_card = QLabel('by-陈鑫  <b style="color:#667eea;">xuandong__happy</b>')
        hover_label_card.setStyleSheet('font-size:12px;')
        hover_layout.addWidget(hover_label_card)
        self.hover_card.hide()

        # Toast 提示浮层（默认隐藏）
        self.toast = QLabel(self)
        self.toast.setAlignment(Qt.AlignCenter)
        self.toast.setStyleSheet("""
            background: #323c4e;
            color: #d4e4ff;
            border-radius: 6px;
            padding: 8px 20px;
            font-size: 12px;
            font-family: "Microsoft YaHei";
        """)
        self.toast.hide()
        self.toast_timer = QTimer()
        self.toast_timer.setSingleShot(True)
        self.toast_timer.timeout.connect(self.toast.hide)

        # 窗口控制按钮
        btn_min = QPushButton('─')
        btn_min.setObjectName('btnMinimize')
        btn_min.setToolTip('最小化')
        btn_min.clicked.connect(self.showMinimized)

        btn_max = QPushButton('□')
        btn_max.setObjectName('btnMaximize')
        btn_max.setToolTip('最大化/还原')
        btn_max.clicked.connect(self.toggle_maximize)

        btn_close = QPushButton('✕')
        btn_close.setObjectName('btnClose')
        btn_close.setToolTip('关闭')
        btn_close.clicked.connect(self.close)

        title_layout.addWidget(btn_min)
        title_layout.addWidget(btn_max)
        title_layout.addWidget(btn_close)

        main_layout.addWidget(title_bar)

        # --- 内容区域（可滚动） ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(16, 16, 16, 16)
        content_layout.setSpacing(12)

        # --- 文件选择区域 ---
        file_group = QGroupBox('📋 文件路径设置')
        file_grid = QGridLayout()
        file_grid.setSpacing(10)

        labels = ['📦 压缩包路径:', '📁 输出目录:', '📊 发票登记表:']
        self.file_entries = []
        self.file_browse_btns = []

        for i, (label_text, browse_method) in enumerate([
            ('📦 压缩包路径:', self.browse_zip),
            ('📁 输出目录:', self.browse_output),
            ('📊 发票登记表:', self.browse_excel)
        ]):
            lbl = QLabel(label_text)
            lbl.setStyleSheet('font-weight: 500; color: #555;')
            entry = QLineEdit()
            entry.setPlaceholderText(f'请选择{label_text[1:]}...')
            btn = QPushButton('浏览')
            btn.setObjectName('btnBrowse')
            btn.clicked.connect(browse_method)
            self.file_entries.append(entry)
            self.file_browse_btns.append(btn)

            file_grid.addWidget(lbl, i, 0)
            file_grid.addWidget(entry, i, 1)
            file_grid.addWidget(btn, i, 2)

        file_group.setLayout(file_grid)
        content_layout.addWidget(file_group)

        # --- 邮件账号配置 ---
        mail_group = QGroupBox('📮 邮件账号配置 (至少2个QQ邮箱)')
        mail_vbox = QVBoxLayout()
        mail_header = QHBoxLayout()
        mail_header.addWidget(QLabel('QQ邮箱'))
        mail_header.addWidget(QLabel('授权码'))
        mail_header.addSpacing(60)
        mail_header.setSpacing(8)

        self.accounts_container = QVBoxLayout()
        self.accounts_container.setSpacing(2)

        self.add_account_btn = QPushButton('+ 添加账号')
        self.add_account_btn.setObjectName('btnAdd')
        self.add_account_btn.clicked.connect(self.add_account)

        # 批量粘贴区域（放在表头上方）
        self.batch_paste = QLineEdit()
        self.batch_paste.setPlaceholderText(
            '粘贴账号列表 格式: 邮箱 授权码 (空格/Tab/换行分隔), 如: 123@qq.com  abcdefgh'
        )
        batch_btn = QPushButton('解析导入')
        batch_btn.setObjectName('btnBrowse')
        batch_btn.setToolTip('自动识别"邮箱 授权码"格式并批量填充')
        batch_btn.clicked.connect(self.parse_batch_accounts)
        clear_all_btn = QPushButton('清空全部')
        clear_all_btn.setObjectName('btnDanger')
        clear_all_btn.setToolTip('清空所有已配置的邮箱账号')
        clear_all_btn.clicked.connect(self.clear_all_accounts)
        batch_row = QHBoxLayout()
        batch_row.addWidget(self.batch_paste)
        batch_row.addWidget(batch_btn)
        batch_row.addWidget(clear_all_btn)

        mail_vbox.addLayout(batch_row)
        mail_vbox.addLayout(mail_header)
        mail_vbox.addLayout(self.accounts_container)
        mail_vbox.addWidget(self.add_account_btn)
        mail_group.setLayout(mail_vbox)
        content_layout.addWidget(mail_group)

        # --- 进度 ---
        progress_group = QGroupBox('📊 处理进度')
        progress_layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        self.status_label = QLabel('准备就绪')
        self.status_label.setObjectName('statusLabel')

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        progress_group.setLayout(progress_layout)
        content_layout.addWidget(progress_group)

        # --- 操作按钮 ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(16)

        self.start_btn = QPushButton('🚀 开始处理')
        self.start_btn.setObjectName('btnPrimary')
        self.start_btn.setFixedHeight(44)
        self.start_btn.clicked.connect(self.start_processing)

        self.stop_btn = QPushButton('⏹ 停止')
        self.stop_btn.setObjectName('btnDanger')
        self.stop_btn.setFixedHeight(44)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)

        btn_layout.addStretch()
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()
        content_layout.addLayout(btn_layout)

        # --- 日志 ---
        log_group = QGroupBox('📜 运行日志')
        log_layout = QVBoxLayout()
        self.log_area = QTextEdit()
        self.log_area.setObjectName('logArea')
        self.log_area.setReadOnly(True)
        self.log_area.setMinimumHeight(180)
        log_layout.addWidget(self.log_area)
        log_group.setLayout(log_layout)
        content_layout.addWidget(log_group)

        content_layout.addStretch()
        scroll.setWidget(content)
        main_layout.addWidget(scroll)

        # 初始化2个账号
        self.add_account()
        self.add_account()

    # --- 账号管理 ---
    def add_account(self):
        row = EmailAccountRow(len(self.account_rows))
        row.removed.connect(self.remove_account)
        self.account_rows.append(row)
        self.accounts_container.addWidget(row)
        self.update_account_numbers()

    def remove_account(self, row):
        if len(self.account_rows) <= 2:
            QMessageBox.warning(self, '提示', '至少需要保留2个邮件账号！')
            return
        self.account_rows.remove(row)
        self.accounts_container.removeWidget(row)
        row.deleteLater()
        self.update_account_numbers()

    def parse_batch_accounts(self):
        """解析粘贴的账号列表(格式: 邮箱 授权码)并自动填充"""
        text = self.batch_paste.text().strip()
        if not text:
            return
        # 按换行、分号、逗号或连续的空白拆分
        tokens = re.split(r'[\n;；,，]+', text)
        new_accounts = []
        for token in tokens:
            token = token.strip()
            if not token or '@' not in token:
                continue
            parts = token.split()
            if len(parts) >= 2:
                email_part = parts[0].strip()
                auth_part = parts[-1].strip()  # 最后一个token作为授权码
                if '@' in email_part:
                    new_accounts.append((email_part, auth_part))
        if not new_accounts:
            QMessageBox.warning(self, '解析失败',
                                '未识别到有效账号，请使用格式：邮箱 授权码')
            return
        # 清空现有行
        for row in list(self.account_rows):
            self.accounts_container.removeWidget(row)
            row.deleteLater()
        self.account_rows.clear()
        # 填充新行
        for email, auth in new_accounts:
            self.add_account_with_data(email, auth)
        # 保证至少2行
        while len(self.account_rows) < 2:
            self.add_account()
        self.batch_paste.clear()
        QMessageBox.information(self, '导入成功',
                                f'成功导入 {len(new_accounts)} 个邮箱账号')

    def add_account_with_data(self, email='', auth=''):
        row = EmailAccountRow(len(self.account_rows), email=email, auth=auth)
        row.removed.connect(self.remove_account)
        self.account_rows.append(row)
        self.accounts_container.addWidget(row)
        self.update_account_numbers()

    def clear_all_accounts(self):
        """清空所有邮箱账号行，保留至少2个空行"""
        reply = QMessageBox.question(
            self, '确认清空', '确定要清空所有已配置的邮箱账号吗？',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        for row in list(self.account_rows):
            self.accounts_container.removeWidget(row)
            row.deleteLater()
        self.account_rows.clear()
        # 保留至少2个空行
        while len(self.account_rows) < 2:
            self.add_account()

    def update_account_numbers(self):
        for i, row in enumerate(self.account_rows):
            row.index = i
            row.num_label.setText(f'<b>#{i + 1}</b>')

    # --- 文件浏览 ---
    def browse_zip(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择压缩包', '', 'ZIP文件 (*.zip)')
        if path:
            self.file_entries[0].setText(path)

    def browse_output(self):
        path = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if path:
            self.file_entries[1].setText(path)

    def browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '选择发票登记表', '',
            'Excel文件 (*.xlsx *.xls)'
        )
        if path:
            self.file_entries[2].setText(path)

    # --- 日志 ---
    def append_log(self, msg, level='info'):
        colors = {
            'info': '#cdd6f4',
            'success': '#a6e3a1',
            'warn': '#f9e2af',
            'error': '#f38ba8',
            'title': '#89b4fa',
        }
        color = colors.get(level, '#cdd6f4')
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_area.append(
            f'<span style="color:#6c7086;">[{timestamp}]</span> '
            f'<span style="color:{color};">{msg}</span>'
        )
        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    # --- 处理流程 ---
    def start_processing(self):
        # 验证输入
        zip_path = self.file_entries[0].text().strip()
        output_dir = self.file_entries[1].text().strip()
        excel_path = self.file_entries[2].text().strip()

        if not all([zip_path, output_dir, excel_path]):
            QMessageBox.critical(self, '错误', '请先选择压缩包、输出目录和发票登记表！')
            return

        # 收集邮件账号
        accounts = []
        for row in self.account_rows:
            data = row.get_data()
            if data['email_from'] and data['auth_code']:
                accounts.append(data)

        if len(accounts) < 2:
            QMessageBox.critical(self, '错误',
                                 f'必须配置至少2个有效邮件账号！(当前{len(accounts)}个)')
            return

        # 清空日志
        self.log_area.clear()
        self.append_log('启动处理...', 'title')
        self.append_log(f'压缩包: {zip_path}')
        self.append_log(f'输出目录: {output_dir}')
        self.append_log(f'发票登记表: {excel_path}')
        self.append_log(f'邮件账号: {len(accounts)} 个')

        # 禁用UI
        self.is_processing = True
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        for entry in self.file_entries:
            entry.setEnabled(False)
        for btn in self.file_browse_btns:
            btn.setEnabled(False)
        self.add_account_btn.setEnabled(False)
        for row in self.account_rows:
            row.email_input.setEnabled(False)
            row.auth_input.setEnabled(False)
            row.remove_btn.setEnabled(False)

        # 启动线程
        self.thread = ProcessThread(zip_path, output_dir, excel_path, accounts)
        self.thread.log_signal.connect(self.append_log)
        self.thread.progress_signal.connect(self.progress_bar.setValue)
        self.thread.status_signal.connect(self.status_label.setText)
        self.thread.finished_signal.connect(self.on_finished)
        self.thread.start()

    def stop_processing(self):
        if hasattr(self, 'thread') and self.thread.isRunning():
            self.thread.terminate()
            self.thread.wait(3000)
        self.on_finished(False, '用户停止')

    # --- 无边框窗口拖拽 ---
    def title_bar_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPos()

    def title_bar_mouse_move(self, event):
        if hasattr(self, '_drag_pos') and event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self._drag_pos
            self.move(self.pos() + delta)
            self._drag_pos = event.globalPos()

    # --- 标题栏悬浮卡片 & 点击复制 ---
    def title_label_enter(self, event):
        """鼠标进入标题文字 → 在鼠标附近显示悬浮卡片"""
        if hasattr(self, 'hover_card'):
            from PyQt5.QtGui import QCursor
            mouse_pos = QCursor.pos()
            local_pos = self.mapFromGlobal(mouse_pos)
            self.hover_card.setParent(self)
            self.hover_card.adjustSize()
            self.hover_card.move(local_pos.x() + 15, local_pos.y() - 32)
            self.hover_card.show()
            self.hover_card.raise_()

    def title_label_leave(self, event):
        """鼠标离开标题文字 → 隐藏悬浮卡片"""
        if hasattr(self, 'hover_card'):
            self.hover_card.hide()

    def title_label_click(self, event):
        """点击标题文字 → 复制 xuandong__happy 到剪贴板 + Toast提示"""
        clipboard = QApplication.clipboard()
        clipboard.setText('xuandong__happy')
        # 显示 Toast
        self.toast.setText('✅ 已复制 xuandong__happy — 微信粘贴搜索🔍即可添加')
        self.toast.adjustSize()
        # 居中于窗口
        toast_x = (self.width() - self.toast.width()) // 2
        toast_y = self.height() - self.toast.height() - 30
        self.toast.move(toast_x, toast_y)
        self.toast.show()
        self.toast.raise_()
        # 1.5秒后自动消失
        self.toast_timer.start(1500)

    def toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def on_finished(self, success, message):
        self.is_processing = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        for entry in self.file_entries:
            entry.setEnabled(True)
        for btn in self.file_browse_btns:
            btn.setEnabled(True)
        self.add_account_btn.setEnabled(True)
        for row in self.account_rows:
            row.email_input.setEnabled(True)
            row.auth_input.setEnabled(True)
            row.remove_btn.setEnabled(True)

        if success:
            self.append_log(message, 'success')
        else:
            self.append_log(f'处理中断: {message}', 'error')


# ============ 入口 ============
def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = InvoiceMailerApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()