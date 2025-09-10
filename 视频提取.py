import os
import re
import sys
import json
import requests
import threading
from urllib.parse import urlparse, parse_qs
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QCheckBox, QGroupBox,
                             QTextEdit, QProgressBar, QMessageBox, QFileDialog, QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap
import yt_dlp
import time


class ParseThread(QThread):
    """解析视频信息的线程"""
    parse_finished = pyqtSignal(dict, str)  # 成功信号：视频信息, 错误信息
    parse_error = pyqtSignal(str)  # 错误信号：错误信息
    status_update = pyqtSignal(str)  # 状态更新信号

    def __init__(self, url):
        super().__init__()
        self.url = url

    def run(self):
        try:
            self.status_update.emit("正在解析URL...")
            downloader = BilibiliDownloader()
            bvid = downloader.extract_bvid(self.url)
            self.status_update.emit(f"提取到视频ID: {bvid}")

            self.status_update.emit("正在获取视频信息...")
            video_info = downloader.get_video_info(bvid)

            self.status_update.emit("解析完成")
            self.parse_finished.emit(video_info, "")

        except Exception as e:
            self.parse_error.emit(str(e))


class DownloadThread(QThread):
    """下载线程"""
    progress_signal = pyqtSignal(int)
    message_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str)  # 用于状态标签的信号
    speed_time_signal = pyqtSignal(str)  # 专门用于下载速度和时间的信号
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, bvid, page_index, download_options, download_path):
        super().__init__()
        self.bvid = bvid
        self.page_index = page_index
        self.download_options = download_options
        self.download_path = download_path
        self.is_cancelled = False
        self.last_progress_message_time = 0  # 用于控制进度消息的发送频率
        self.processing_stage = False  # 标记是否处于处理阶段

    def run(self):
        try:
            downloader = BilibiliDownloader()
            downloader.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Referer': 'https://www.bilibili.com/'
            })

            # 获取视频信息
            self.message_signal.emit("正在获取视频信息...")
            video_info = downloader.get_video_info(self.bvid)
            selected_page = video_info['pages'][self.page_index]
            filename = f"{self.bvid}_{selected_page['page']}_{selected_page['part']}"
            filename = re.sub(r'[\\/*?:"<>|]', "", filename)

            # 下载封面
            if self.download_options['cover']:
                self.message_signal.emit("正在下载封面...")
                downloader.download_cover(self.download_path, video_info)

            # 下载弹幕
            if self.download_options['danmaku']:
                self.message_signal.emit("正在下载弹幕...")
                downloader.download_danmaku(selected_page['cid'], self.download_path, filename)

            # 下载媒体文件
            if self.download_options['video'] or self.download_options['audio']:
                self.message_signal.emit("正在准备下载媒体文件...")

                # 构建yt-dlp选项 - 使用广泛兼容的视频编码
                ydl_opts = {
                    'outtmpl': os.path.join(self.download_path, f'{filename}.%(ext)s'),
                    'progress_hooks': [self.yt_dlp_progress_hook],
                    'quiet': True,
                    'noprogress': True,
                    'ffmpeg_location': './ffmpeg/bin'
                }

                # 设置格式和后处理
                if self.download_options['video'] and self.download_options['audio']:
                    # 下载视频+音频（合并为mp4，使用兼容编码）
                    ydl_opts['format'] = 'bestvideo+bestaudio/best'
                    ydl_opts['merge_output_format'] = 'mp4'
                    # 添加后处理器参数确保兼容性
                    ydl_opts['postprocessor_args'] = [
                        '-c:v', 'libx264', '-preset', 'medium', '-crf', '23',
                        '-c:a', 'aac', '-b:a', '128k',
                        '-movflags', '+faststart',
                        '-max_muxing_queue_size', '9999'  # 解决某些音频转换问题
                    ]
                elif self.download_options['video']:
                    # 只下载视频（带音频，使用兼容编码）
                    ydl_opts['format'] = 'bestvideo+bestaudio/best'
                    ydl_opts['merge_output_format'] = 'mp4'
                    ydl_opts['postprocessor_args'] = [
                        '-c:v', 'libx264', '-preset', 'medium', '-crf', '23',
                        '-c:a', 'aac', '-b:a', '128k',
                        '-movflags', '+faststart',
                        '-max_muxing_queue_size', '9999'
                    ]
                elif self.download_options['audio']:
                    # 只下载音频（转换为mp3）
                    ydl_opts['format'] = 'bestaudio/best'
                    ydl_opts['postprocessors'] = [{
                        'key': 'FFmpegExtractAudio',
                        'preferredcodec': 'mp3',
                        'preferredquality': '192',
                    }]
                    # 添加音频转换参数
                    ydl_opts['postprocessor_args'] = [
                        '-max_muxing_queue_size', '9999'
                    ]

                try:
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        ydl.download([f'https://www.bilibili.com/video/{self.bvid}?p={self.page_index + 1}'])

                    # 如果处于处理阶段，等待处理完成
                    if self.processing_stage:
                        self.simulate_processing()

                    self.message_signal.emit("下载完成!")
                    self.finished_signal.emit(True, "下载完成")
                except Exception as e:
                    error_msg = str(e)
                    # 处理音频转换失败的错误
                    if "audio conversion failed" in error_msg:
                        error_msg += "\n音频转换失败，尝试使用备用方案..."
                        # 尝试使用不同的音频编码
                        try:
                            self.message_signal.emit("音频转换失败，尝试备用方案...")
                            # 重新尝试使用不同的参数
                            backup_ydl_opts = ydl_opts.copy()
                            if 'postprocessor_args' in backup_ydl_opts:
                                # 移除可能有问题的参数
                                backup_ydl_opts['postprocessor_args'] = [
                                    '-c:v', 'libx264', '-preset', 'medium', '-crf', '23',
                                    '-c:a', 'aac', '-b:a', '128k',
                                    '-movflags', '+faststart'
                                ]

                            with yt_dlp.YoutubeDL(backup_ydl_opts) as ydl:
                                ydl.download([f'https://www.bilibili.com/video/{self.bvid}?p={self.page_index + 1}'])

                            self.message_signal.emit("备用方案成功!")
                            self.finished_signal.emit(True, "下载完成")
                        except Exception as backup_e:
                            error_msg += f"\n备用方案也失败: {str(backup_e)}"
                            self.message_signal.emit(f"下载失败: {error_msg}")
                            self.finished_signal.emit(False, f"下载失败: {error_msg}")
                    else:
                        self.message_signal.emit(f"下载失败: {error_msg}")
                        self.finished_signal.emit(False, f"下载失败: {error_msg}")
            else:
                self.message_signal.emit("下载完成!")
                self.finished_signal.emit(True, "下载完成")

        except Exception as e:
            self.message_signal.emit(f"下载过程中出错: {str(e)}")
            self.finished_signal.emit(False, f"下载过程中出错: {str(e)}")

    def yt_dlp_progress_hook(self, d):
        if self.is_cancelled:
            raise Exception("下载已取消")

        current_time = time.time()

        if d['status'] == 'downloading':
            if not self.processing_stage:  # 只在下载阶段更新下载进度
                if 'total_bytes' in d and d['total_bytes'] > 0:
                    percent = int(float(d['downloaded_bytes']) / float(d['total_bytes']) * 100)
                    self.progress_signal.emit(percent)

                    # 控制进度消息的发送频率
                    if current_time - self.last_progress_message_time >= 0.5:
                        speed = d.get('_speed_str', '未知速度')
                        elapsed = d.get('_elapsed_str', '未知时间')
                        speed_time_message = f"速度: {speed} - 已用时间: {elapsed}"
                        self.speed_time_signal.emit(speed_time_message)
                        self.last_progress_message_time = current_time

                elif 'downloaded_bytes' in d and 'total_bytes_estimate' in d:
                    percent = int(float(d['downloaded_bytes']) / float(d['total_bytes_estimate']) * 100)
                    self.progress_signal.emit(percent)

                    if current_time - self.last_progress_message_time >= 0.5:
                        speed = d.get('_speed_str', '未知速度')
                        elapsed = d.get('_elapsed_str', '未知时间')
                        speed_time_message = f"速度: {speed} - 已用时间: {elapsed}"
                        self.speed_time_signal.emit(speed_time_message)
                        self.last_progress_message_time = current_time

        elif d['status'] == 'finished':
            # 下载完成，进入处理阶段
            self.processing_stage = True
            self.progress_signal.emit(0)  # 进度条清零，开始处理阶段
            self.speed_time_signal.emit("下载完成，正在处理文件...")

    def simulate_processing(self):
        """模拟处理过程，更新进度条"""
        self.message_signal.emit("正在处理文件...")

        # 模拟处理进度，从0%开始到100%
        for i in range(101):
            if self.is_cancelled:
                raise Exception("处理已取消")

            self.progress_signal.emit(i)

            # 更新状态信息
            if i % 10 == 0:  # 每10%更新一次状态
                self.speed_time_signal.emit(f"处理中: {i}%")

            # 短暂延迟，模拟处理时间
            time.sleep(0.05)

        # 处理完成
        self.speed_time_signal.emit("处理完成")
        self.processing_stage = False

    def cancel(self):
        self.is_cancelled = True


class BilibiliDownloader:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Referer': 'https://www.bilibili.com/'
        })
        self.video_info = {}

        # 设置超时时间
        self.timeout = 10

    def extract_bvid(self, url):
        """从URL中提取BVID"""
        parsed_url = urlparse(url)
        if 'bilibili.com' not in parsed_url.netloc:
            raise ValueError("不是有效的B站链接")

        # 尝试从路径中提取BVID
        path_parts = parsed_url.path.split('/')
        for part in path_parts:
            if part.startswith('BV'):
                return part

        # 尝试从查询参数中提取
        query_params = parse_qs(parsed_url.query)
        if 'bvid' in query_params:
            return query_params['bvid'][0]

        raise ValueError("无法从URL中提取视频ID")

    def get_video_info(self, bvid):
        """获取视频信息"""
        # 获取视频基本信息
        info_url = f"https://api.bilibili.com/x/web-interface/view?bvid={bvid}"
        try:
            response = self.session.get(info_url, timeout=self.timeout)
            if response.status_code != 200:
                raise Exception(f"获取视频信息失败，HTTP状态码: {response.status_code}")

            data = response.json()
            if data['code'] != 0:
                raise Exception(data.get('message', '未知错误'))

            info = data['data']
            self.video_info = {
                'title': info['title'],
                'bvid': bvid,
                'cover': info['pic'],
                'desc': info['desc'],
                'duration': info['duration'],
                'pages': []
            }

            # 处理多P视频
            for page in info['pages']:
                self.video_info['pages'].append({
                    'page': page['page'],
                    'part': page['part'],
                    'cid': page['cid']
                })

            return self.video_info

        except requests.exceptions.Timeout:
            raise Exception("请求超时，请检查网络连接")
        except requests.exceptions.ConnectionError:
            raise Exception("网络连接错误，请检查网络设置")
        except requests.exceptions.RequestException as e:
            raise Exception(f"网络请求错误: {str(e)}")
        except json.JSONDecodeError:
            raise Exception("解析响应数据失败")

    def download_cover(self, download_path, video_info):
        """下载封面"""
        if not video_info or 'cover' not in video_info:
            return

        cover_url = video_info['cover']
        try:
            response = self.session.get(cover_url, timeout=self.timeout)
            if response.status_code == 200:
                cover_path = os.path.join(download_path, f"{video_info['bvid']}_cover.jpg")
                with open(cover_path, 'wb') as f:
                    f.write(response.content)
        except:
            pass  # 封面下载失败不影响主要功能

    def download_danmaku(self, cid, download_path, filename):
        """下载弹幕"""
        danmaku_url = f"https://api.bilibili.com/x/v1/dm/list.so?oid={cid}"
        try:
            response = self.session.get(danmaku_url, timeout=self.timeout)
            if response.status_code == 200:
                danmaku_path = os.path.join(download_path, f"{filename}.xml")
                with open(danmaku_path, 'wb') as f:
                    f.write(response.content)
        except:
            pass  # 弹幕下载失败不影响主要功能


class BilibiliDownloaderUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.parse_thread = None
        self.download_thread = None
        self.video_info = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('B站视频下载工具')
        self.setGeometry(100, 100, 800, 650)

        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        layout = QVBoxLayout(central_widget)

        # URL输入区域
        url_group = QGroupBox("视频链接")
        url_layout = QHBoxLayout()
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("请输入B站视频链接，例如: https://www.bilibili.com/video/BV1xx411c7mD")
        self.url_input.returnPressed.connect(self.parse_url)
        self.parse_btn = QPushButton("解析")
        self.parse_btn.clicked.connect(self.parse_url)
        url_layout.addWidget(self.url_input)
        url_layout.addWidget(self.parse_btn)
        url_group.setLayout(url_layout)
        layout.addWidget(url_group)

        # 视频信息区域
        info_group = QGroupBox("视频信息")
        info_layout = QVBoxLayout()

        # 封面和基本信息水平布局
        info_h_layout = QHBoxLayout()

        # 封面
        self.cover_label = QLabel()
        self.cover_label.setFixedSize(160, 100)
        self.cover_label.setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;")
        self.cover_label.setAlignment(Qt.AlignCenter)
        self.cover_label.setText("等待解析...")
        info_h_layout.addWidget(self.cover_label)

        # 基本信息
        info_v_layout = QVBoxLayout()
        self.title_label = QLabel("标题: 等待解析...")
        self.duration_label = QLabel("时长: 等待解析...")
        self.pages_combo = QComboBox()
        self.pages_combo.setEnabled(False)
        self.pages_combo.currentIndexChanged.connect(self.on_page_changed)
        info_v_layout.addWidget(self.title_label)
        info_v_layout.addWidget(self.duration_label)
        info_v_layout.addWidget(QLabel("选择分P:"))
        info_v_layout.addWidget(self.pages_combo)
        info_h_layout.addLayout(info_v_layout)

        info_layout.addLayout(info_h_layout)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 下载选项区域
        options_group = QGroupBox("下载选项")
        options_layout = QHBoxLayout()
        self.video_check = QCheckBox("下载视频(带音频)")
        self.video_check.setChecked(True)
        self.audio_check = QCheckBox("下载音频(MP3格式)")
        self.audio_check.setChecked(False)
        self.danmaku_check = QCheckBox("下载弹幕")
        self.cover_check = QCheckBox("下载封面")
        options_layout.addWidget(self.video_check)
        options_layout.addWidget(self.audio_check)
        options_layout.addWidget(self.danmaku_check)
        options_layout.addWidget(self.cover_check)

        # 连接信号，确保视频和音频不能同时选择
        self.video_check.stateChanged.connect(self.on_video_check_changed)
        self.audio_check.stateChanged.connect(self.on_audio_check_changed)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # 下载路径选择
        path_group = QGroupBox("下载路径")
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setText(os.path.join(os.path.expanduser("~"), "Downloads"))
        self.browse_btn = QPushButton("浏览")
        self.browse_btn.clicked.connect(self.browse_path)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_btn)
        path_group.setLayout(path_layout)
        layout.addWidget(path_group)

        # 进度条 - 设置为蓝色
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #2196F3;  /* 蓝色 */
                width: 10px;
            }
        """)
        layout.addWidget(self.progress_bar)

        # 下载速度和已用时间显示
        self.speed_time_label = QLabel("速度: - - 已用时间: - -")
        self.speed_time_label.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(self.speed_time_label)

        # 状态信息
        self.status_label = QLabel("就绪")
        layout.addWidget(self.status_label)

        # 按钮区域
        button_layout = QHBoxLayout()
        self.download_btn = QPushButton("开始下载")
        self.download_btn.clicked.connect(self.start_download)
        self.download_btn.setEnabled(False)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.cancel_operation)
        self.cancel_btn.setEnabled(False)
        button_layout.addWidget(self.download_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        # 日志区域
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

    def on_video_check_changed(self, state):
        """视频复选框状态改变时的处理"""
        if state == Qt.Checked and self.audio_check.isChecked():
            self.audio_check.setChecked(False)

    def on_audio_check_changed(self, state):
        """音频复选框状态改变时的处理"""
        if state == Qt.Checked and self.video_check.isChecked():
            self.video_check.setChecked(False)

    def parse_url(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "警告", "请输入B站视频链接")
            return

        # 禁用UI控件，防止重复点击
        self.set_ui_enabled(False)
        self.log_text.append(f"正在解析链接: {url}")
        self.status_label.setText("正在解析...")
        self.speed_time_label.setText("速度: - - 已用时间: - -")
        self.progress_bar.setValue(0)  # 重置进度条

        # 创建解析线程
        self.parse_thread = ParseThread(url)
        self.parse_thread.parse_finished.connect(self.on_parse_finished)
        self.parse_thread.parse_error.connect(self.on_parse_error)
        self.parse_thread.status_update.connect(self.log_text.append)
        self.parse_thread.start()

    def on_parse_finished(self, video_info, error_msg):
        self.video_info = video_info
        self.display_video_info()
        self.download_btn.setEnabled(True)
        self.set_ui_enabled(True)
        self.status_label.setText("解析完成")
        self.log_text.append("解析成功，请选择下载选项")

    def on_parse_error(self, error_msg):
        self.log_text.append(f"解析失败: {error_msg}")
        QMessageBox.critical(self, "错误", f"解析失败: {error_msg}")
        self.set_ui_enabled(True)
        self.status_label.setText("解析失败")

    def set_ui_enabled(self, enabled):
        """设置UI控件的启用状态"""
        self.parse_btn.setEnabled(enabled)
        self.url_input.setEnabled(enabled)
        self.cancel_btn.setEnabled(not enabled)

    def display_video_info(self):
        if not self.video_info:
            return

        # 显示基本信息
        self.title_label.setText(f"标题: {self.video_info['title']}")
        minutes, seconds = divmod(self.video_info['duration'], 60)
        self.duration_label.setText(f"时长: {minutes}分{seconds}秒")

        # 填充分P选择
        self.pages_combo.clear()
        for page in self.video_info['pages']:
            self.pages_combo.addItem(f"P{page['page']} - {page['part']}", page)
        self.pages_combo.setEnabled(len(self.video_info['pages']) > 1)

        # 异步加载封面
        self.load_cover_async()

    def load_cover_async(self):
        """异步加载封面"""
        if not self.video_info or 'cover' not in self.video_info:
            return

        def load_cover():
            try:
                response = requests.get(self.video_info['cover'], timeout=10)
                if response.status_code == 200:
                    pixmap = QPixmap()
                    pixmap.loadFromData(response.content)
                    pixmap = pixmap.scaled(160, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self.cover_label.setPixmap(pixmap)
            except:
                self.cover_label.setText("封面加载失败")

        threading.Thread(target=load_cover, daemon=True).start()

    def on_page_changed(self, index):
        if not self.video_info or index < 0:
            return

        page = self.pages_combo.itemData(index)
        if page:
            self.log_text.append(f"已选择分P: {page['part']}")

    def browse_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择下载目录", self.path_input.text())
        if path:
            self.path_input.setText(path)

    def start_download(self):
        if not self.video_info:
            QMessageBox.warning(self, "警告", "请先解析视频链接")
            return

        # 获取下载选项
        download_options = {
            'video': self.video_check.isChecked(),
            'audio': self.audio_check.isChecked(),
            'danmaku': self.danmaku_check.isChecked(),
            'cover': self.cover_check.isChecked()
        }

        if not any(download_options.values()):
            QMessageBox.warning(self, "警告", "请至少选择一个下载选项")
            return

        # 获取选中的分P
        page_index = self.pages_combo.currentIndex()
        if page_index < 0:
            page_index = 0

        # 重置进度和速度显示
        self.progress_bar.setValue(0)
        self.speed_time_label.setText("速度: - - 已用时间: - -")

        # 创建下载线程
        self.download_thread = DownloadThread(
            self.video_info['bvid'],
            page_index,
            download_options,
            self.path_input.text()
        )

        # 连接信号
        self.download_thread.progress_signal.connect(self.update_progress)
        self.download_thread.message_signal.connect(self.update_status)
        self.download_thread.status_signal.connect(self.update_status_label)
        self.download_thread.speed_time_signal.connect(self.update_speed_time)
        self.download_thread.finished_signal.connect(self.download_finished)

        # 更新UI状态
        self.set_ui_enabled(False)
        self.download_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        # 启动下载线程
        self.download_thread.start()

    def cancel_operation(self):
        """取消当前操作（解析或下载）"""
        if self.parse_thread and self.parse_thread.isRunning():
            self.parse_thread.terminate()
            self.parse_thread.wait()
            self.log_text.append("解析已取消")
        elif self.download_thread and self.download_thread.isRunning():
            self.download_thread.cancel()
            self.download_thread.terminate()
            self.download_thread.wait()
            self.log_text.append("下载已取消")

        self.reset_ui_state()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_status(self, message):
        self.log_text.append(message)

    def update_status_label(self, message):
        self.status_label.setText(message)

    def update_speed_time(self, message):
        self.speed_time_label.setText(message)

    def download_finished(self, success, message):
        if success:
            QMessageBox.information(self, "成功", message)
        else:
            QMessageBox.critical(self, "错误", message)

        self.reset_ui_state()

    def reset_ui_state(self):
        self.set_ui_enabled(True)
        self.download_btn.setEnabled(self.video_info is not None)
        self.cancel_btn.setEnabled(False)
        self.status_label.setText("就绪")
        self.speed_time_label.setText("速度: - - 已用时间: - -")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = BilibiliDownloaderUI()
    window.show()
    sys.exit(app.exec_())
