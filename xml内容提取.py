import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import Tk, filedialog, messagebox, ttk
import threading
import csv


class XMLToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("XML信息提取工具")
        self.root.geometry("800x600")

        # 创建界面组件
        self.create_widgets()

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 标题
        title_label = ttk.Label(main_frame, text="XML信息提取工具", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # 选择文件夹按钮
        self.select_button = ttk.Button(main_frame, text="选择文件夹", command=self.select_folder)
        self.select_button.pack(pady=10)

        # 文件夹路径显示
        self.folder_label = ttk.Label(main_frame, text="未选择文件夹", wraplength=700)
        self.folder_label.pack(pady=5)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(pady=10, fill="x")

        # 开始提取按钮
        self.start_button = ttk.Button(main_frame, text="开始提取", command=self.start_extraction, state="disabled")
        self.start_button.pack(pady=10)

        # 状态标签
        self.status_label = ttk.Label(main_frame, text="准备就绪")
        self.status_label.pack(pady=5)

        # 结果文本框框架
        text_frame = ttk.LabelFrame(main_frame, text="处理日志")
        text_frame.pack(pady=10, fill="both", expand=True)

        # 文本框和滚动条
        text_container = ttk.Frame(text_frame)
        text_container.pack(fill="both", expand=True, padx=5, pady=5)

        # 垂直滚动条
        v_scrollbar = ttk.Scrollbar(text_container, orient="vertical")
        v_scrollbar.pack(side="right", fill="y")

        # 水平滚动条
        h_scrollbar = ttk.Scrollbar(text_container, orient="horizontal")
        h_scrollbar.pack(side="bottom", fill="x")

        # 文本框
        self.result_text = tk.Text(text_container, height=15, width=80,
                                   yscrollcommand=v_scrollbar.set,
                                   xscrollcommand=h_scrollbar.set,
                                   wrap="none")
        self.result_text.pack(side="left", fill="both", expand=True)

        # 配置滚动条
        v_scrollbar.config(command=self.result_text.yview)
        h_scrollbar.config(command=self.result_text.xview)

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含XML文件的文件夹")
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.config(text=folder_path)
            self.start_button.config(state="enabled")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"已选择文件夹: {folder_path}\n")

    def start_extraction(self):
        self.start_button.config(state="disabled")
        self.status_label.config(text="正在提取...")
        self.progress["value"] = 0
        self.result_text.delete(1.0, tk.END)

        # 在新线程中执行提取操作
        thread = threading.Thread(target=self.extract_xml_data)
        thread.daemon = True
        thread.start()

    def parse_xml_file(self, xml_file):
        """解析XML文件，处理ANSI编码"""
        try:
            # 使用二进制读取并自动检测编码
            with open(xml_file, 'rb') as f:
                content = f.read()

            # 尝试多种编码
            encodings = ['gbk', 'gb2312', 'utf-8', 'latin-1']
            for encoding in encodings:
                try:
                    xml_content = content.decode(encoding)
                    # 移除BOM（如果有）
                    if xml_content.startswith('\ufeff'):
                        xml_content = xml_content[1:]
                    # 解析XML
                    return ET.fromstring(xml_content)
                except (UnicodeDecodeError, ET.ParseError):
                    continue

            raise Exception("无法解析XML文件，所有编码尝试都失败")

        except Exception as e:
            raise Exception(f"解析XML文件失败: {str(e)}")

    def extract_xml_data(self):
        try:
            # 收集所有XML文件
            xml_files = []
            for root_dir, _, files in os.walk(self.folder_path):
                for file in files:
                    if file.lower().endswith('.xml'):
                        xml_files.append(os.path.join(root_dir, file))

            if not xml_files:
                self.root.after(0, lambda: messagebox.showwarning("警告", "未找到XML文件"))
                self.root.after(0, self.reset_ui)
                return

            total_files = len(xml_files)
            data = []

            self.root.after(0, lambda: self.result_text.insert(tk.END, f"找到 {total_files} 个XML文件\n"))

            for i, xml_file in enumerate(xml_files):
                try:
                    file_name = os.path.basename(xml_file)
                    self.root.after(0, lambda f=file_name: self.result_text.insert(tk.END, f"\n处理文件: {f}\n"))

                    # 解析XML文件
                    root = self.parse_xml_file(xml_file)

                    # 获取文件夹名称
                    folder_name = os.path.basename(os.path.dirname(xml_file))

                    # 获取describe标签内容
                    describe = ""
                    describe_elem = root.find(".//describe")
                    if describe_elem is not None and describe_elem.text:
                        describe = describe_elem.text.strip()

                    # 获取expage元素的name属性 - 修正提取逻辑
                    expage_name = ""
                    # 查找所有expage元素
                    for elem in root.findall(".//expage"):
                        if 'name' in elem.attrib:
                            expage_name = elem.attrib['name']
                            break  # 只取第一个找到的

                    # 如果没找到，尝试直接获取根元素的name属性
                    if not expage_name and 'name' in root.attrib:
                        expage_name = root.attrib['name']

                    # 获取.cpt文件名称
                    cpt_files = set()

                    # 查找所有包含.cpt的文本内容
                    for elem in root.iter():
                        if elem.text and '.cpt' in elem.text:
                            text = elem.text.strip()
                            # 处理路径中的.cpt文件
                            if text.endswith('.cpt'):
                                cpt_files.add(os.path.basename(text))  # 只取文件名
                            elif '/' in text or '\\' in text:
                                # 从路径中提取文件名
                                parts = text.replace('\\', '/').split('/')
                                for part in parts:
                                    if part.endswith('.cpt'):
                                        cpt_files.add(part)

                    cpt_names = ", ".join(sorted(cpt_files)) if cpt_files else "无"

                    # 添加到数据列表
                    data.append({
                        "文件夹": folder_name,
                        "描述": describe,
                        "Expage名称": expage_name,
                        "CPT文件": cpt_names
                    })

                    # 在日志中显示处理结果
                    self.root.after(0, lambda f=folder_name, d=describe, e=expage_name, c=cpt_names:
                    self.result_text.insert(tk.END,
                                            f"  文件夹: {f}\n"
                                            f"  描述: {d}\n"
                                            f"  Expage名称: {e}\n"
                                            f"  CPT文件: {c}\n"
                                            ))

                    # 更新进度
                    progress_value = (i + 1) / total_files * 100
                    self.root.after(0, lambda v=progress_value: self.progress.config(value=v))

                except Exception as e:
                    error_msg = f"处理文件 {xml_file} 时出错: {str(e)}\n"
                    self.root.after(0, lambda: self.result_text.insert(tk.END, error_msg))
                    # 添加错误信息到数据中
                    data.append({
                        "文件夹": os.path.basename(os.path.dirname(xml_file)),
                        "描述": "解析错误",
                        "Expage名称": "无",
                        "CPT文件": "无",
                        "错误信息": str(e)
                    })

            # 保存为CSV文件
            if data:
                output_file = os.path.join(self.folder_path, "xml_extraction_result.csv")
                try:
                    # 使用UTF-8-BOM编码，确保Excel正确显示中文
                    with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                        fieldnames = ["文件夹", "描述", "Expage名称", "CPT文件"]
                        if any("错误信息" in item for item in data):
                            fieldnames.append("错误信息")

                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        writer.writeheader()
                        for row in data:
                            # 确保所有值都是字符串
                            safe_row = {}
                            for key, value in row.items():
                                if value is None:
                                    safe_row[key] = ""
                                else:
                                    safe_row[key] = str(value)
                            writer.writerow(safe_row)

                    success_msg = f"\n✅ 提取完成！共处理 {len(data)} 个XML文件\n"
                    success_msg += f"✅ 结果已保存到: {output_file}\n"
                    success_msg += f"✅ 文件使用UTF-8-BOM编码，Excel可以正确显示中文\n"
                    self.root.after(0, lambda: self.result_text.insert(tk.END, success_msg))

                    # 显示数据预览
                    self.root.after(0, lambda: self.result_text.insert(tk.END, "\n📊 数据预览 (前3行):\n"))
                    for i, item in enumerate(data[:3]):
                        preview = f"{i + 1}. 文件夹: {item['文件夹']}\n"
                        preview += f"   描述: {item['描述']}\n"
                        preview += f"   Expage名称: {item['Expage名称']}\n"
                        preview += f"   CPT文件: {item['CPT文件']}\n"
                        if "错误信息" in item:
                            preview += f"   错误信息: {item['错误信息']}\n"
                        self.root.after(0, lambda p=preview: self.result_text.insert(tk.END, p))

                    # 滚动到最底部
                    self.root.after(0, lambda: self.result_text.see(tk.END))

                except Exception as e:
                    error_msg = f"保存CSV文件时出错: {str(e)}\n"
                    self.root.after(0, lambda: self.result_text.insert(tk.END, error_msg))

            else:
                self.root.after(0, lambda: self.result_text.insert(tk.END, "❌ 未提取到任何数据\n"))

        except Exception as e:
            error_msg = f"提取过程中发生错误: {str(e)}\n"
            self.root.after(0, lambda: self.result_text.insert(tk.END, error_msg))
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))

        finally:
            self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.status_label.config(text="提取完成")
        self.start_button.config(state="enabled")


if __name__ == "__main__":
    root = Tk()
    app = XMLToExcelConverter(root)
    root.mainloop()
