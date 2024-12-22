import pandas as pd
from collections import defaultdict
import re
from datetime import datetime
from openpyxl.styles import Alignment, Border, Side, Font
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

class ExcelProcessor:
    def __init__(self):
        self.input_file = None
        self.output_file = None
        self.callback = None
        self.total_records = 0
        self.processed_records = 0
    
    def set_callback(self, callback):
        """设置日志回调函数"""
        self.callback = callback
    
    def log(self, message, level="INFO"):
        """输出带时间戳和级别的日志"""
        if self.callback:
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] [{level}] {message}"
            self.callback(formatted_message)
    
    def parse_and_merge_attendance(self, result):
        """解析考勤结果并合并同类型同日期的时间"""
        # 使用字典存储同一天同类型的小时数
        merged = defaultdict(lambda: defaultdict(float))
        
        # 使用正则表达式提取所有时间信息
        pattern = r'([\u4e00-\u9fa5]+)(\d{2}-\d{2}).*?到.*?(\d{2}-\d{2}).*?(\d+\.?\d*)小时'
        single_day_pattern = r'([\u4e00-\u9fa5]+)(\d{2}-\d{2}).*?(\d+\.?\d*)小时'
        
        def format_type(type_):
            """格式化考勤类型"""
            if type_ == "调休":
                return "调休（加班补）"
            elif type_ == "年假":
                return "休年假"
            elif type_ == "事假":
                return "请事假"
            return type_
        
        # 先尝试匹配跨天的情况
        multi_day_matches = list(re.finditer(pattern, result))
        if multi_day_matches:
            for match in multi_day_matches:
                type_, start_date, end_date, hours = match.groups()
                # 跳过加班记录
                if "加班" in type_:
                    continue
                # 将日期格式从 "11-04" 转换为 "11.04"
                start_date = start_date.replace('-', '.')
                end_date = end_date.split('-')[1]  # 只取结束日期的天数
                
                # 检查是否是同一天的不同时间段
                start_day = start_date.split('.')[1]
                end_day = end_date
                if start_day == end_day:
                    # 如果是同一天，只使用一个日期
                    date_str = start_date
                else:
                    # 如果是跨天，使用日期范围
                    date_str = f"{start_date}-{end_date}"
                
                merged[date_str][format_type(type_)] += float(hours)
        else:
            # 处理单天的情况
            for match in re.finditer(single_day_pattern, result):
                type_, date, hours = match.groups()
                # 跳过加班记录
                if "加班" in type_:
                    continue
                date = date.replace('-', '.')
                merged[date][format_type(type_)] += float(hours)
        
        # 格式化结果
        parts = []
        for date in sorted(merged.keys()):
            for type_, hours in merged[date].items():
                # 移除小时数中的 .0
                hours_str = str(hours).rstrip('0').rstrip('.')
                if "调休（加班补）" in type_:
                    # 调休的特殊格式
                    parts.append(f"{date}调休{hours_str}小时（加班补）")
                else:
                    # 其他类型的格式
                    parts.append(f"{date}{type_}{hours_str}小时")
        
        return "，".join(parts) if len(parts) > 1 else parts[0] if parts else ""

    def get_date_for_sorting(self, record):
        """从考勤记录中提取日期用于排序"""
        # 匹配日期格式 (11.04 或 11.04-05)
        match = re.search(r'(\d{2}\.\d{2})(?:-\d{2})?', record)
        if match:
            date_str = match.group(1)
            # 将日期转换为可比较的格式
            month, day = map(int, date_str.split('.'))
            return month * 100 + day  # 例如：11.04 转换为 1104
        return 0

    def process_excel(self, input_file, output_file):
        try:
            # 读取Excel文件
            self.log("开始读取Excel文件...")
            df = pd.read_excel(input_file, header=[2,3])
            self.total_records = len(df)
            self.log(f"成功读取文件: {Path(input_file).name}")
            self.log(f"总行数: {self.total_records}")
            
            # 找到姓名列
            self.log("正在识别表格结构...")
            name_column = None
            for col in df.columns:
                if isinstance(col, tuple):
                    if any("姓名" in str(x) for x in col):
                        name_column = col
                        break
                elif "姓名" in str(col):
                    name_column = col
                    break

            if name_column is None:
                raise ValueError("未找到姓名列，请检查Excel文件格式")
            
            self.log("开始处理考勤数据...")
            self.log("-" * 50)  # 分隔线
            
            # 使用字典存储每个人的考勤结果
            results_by_person = defaultdict(lambda: {"考勤": set(), "事假": None})
            self.processed_records = 0
            valid_records = 0
            person_details = []  # 存储每个人的处理详情
            
            # 遍历每一行
            for idx, row in df.iterrows():
                self.processed_records += 1
                name = row[name_column]
                if pd.isna(name):
                    continue
                
                attendance_found = False
                attendance_count = 0
                attendance_types = set()  # 记录考勤类型
                
                # 遍历所有列
                for col in df.columns:
                    # 检查考勤结果列
                    if isinstance(col, tuple) and "考勤结果" in str(col[0]):
                        result = row[col]
                        if pd.notna(result) and "休息" not in str(result) and "默认班次" not in str(result):
                            merged_result = self.parse_and_merge_attendance(str(result))
                            if merged_result:
                                attendance_found = True
                                attendance_count += 1
                                results_by_person[f"{name}_{idx}"]["考勤"].add(merged_result)
                                # 提取考勤类型
                                if "调休" in result:
                                    attendance_types.add("调休")
                                elif "年假" in result:
                                    attendance_types.add("年假")
                                elif "事假" in result:
                                    attendance_types.add("事假")
                    
                    # 检查事假列
                    elif isinstance(col, tuple) and "请假" in str(col[0]) and "事假(小时)" in str(col[1]):
                        sick_leave = row[col]
                        if pd.notna(sick_leave):
                            results_by_person[f"{name}_{idx}"]["事假"] = sick_leave
                
                if attendance_found:
                    valid_records += 1
                    # 记录该人员的处理详情
                    person_details.append({
                        "姓名": name,
                        "考勤记录数": attendance_count,
                        "考勤类型": ", ".join(attendance_types),
                        "事假小时": results_by_person[f"{name}_{idx}"]["事假"]
                    })
                    
                    if valid_records % 5 == 0:
                        progress = (self.processed_records / self.total_records) * 100
                        self.log(f"处理进度: {progress:.1f}% ({self.processed_records}/{self.total_records})")

            # 输出统计信息
            self.log("-" * 50)
            self.log("数据处理完成", "SUCCESS")
            self.log(f"总记录数: {self.total_records}")
            self.log(f"有效记录数: {valid_records}")
            self.log("-" * 50)
            
            # 输出每个人的处理详情
            self.log("人员处理详情:")
            for detail in person_details:
                self.log(f"姓名: {detail['姓名']}")
                self.log(f"  - 考勤记录数: {detail['考勤记录数']}")
                self.log(f"  - 考勤类型: {detail['考勤类型']}")
                if detail['事假小时']:
                    self.log(f"  - 事假时长: {detail['事假小时']}小时")
                self.log("-" * 30)
            
            # 生成结果表格
            self.log("正在生成结果表格...")
            result_data = []
            for i, (person_key, results) in enumerate(results_by_person.items(), 1):
                name = person_key.split('_')[0]
                attendance_records = sorted(results["考勤"], key=self.get_date_for_sorting)
                self.log(f"处理 [{name}] 的考勤记录: {len(attendance_records)} 条")
                # 创建一行数据，与原表格结构一致
                row_data = {
                    "序号": i,
                    "考勤月份": "2024年10月",
                    "职员代码": "(空)",
                    "姓名": name,
                    "组织单元名称": "研发中心",
                    "职位名称": "软件工程师",
                    "应出勤工时": "",
                    "入离职出勤天": "",
                    "病假小时": "",
                    "事假小时": results["事假"] if pd.notna(results["事假"]) else "",
                    "迟到次数": "",
                    "迟到分钟": "",
                    "旷工次数": "",
                    "旷工小时": "",
                    "平时加班": "",
                    "周末加班": "",
                    "合计加班时": "",
                    "平时加班费": "",
                    "周末加班费": "",
                    "合计加班费": "",
                    "备注": "；".join(attendance_records)
                }
                result_data.append(row_data)

            # 创建DataFrame
            result_df = pd.DataFrame(result_data)

            # 设置列顺序
            columns = [
                "序号", "考勤月份", "职员代码", "姓名", "组织单元名称", "职位名称", 
                "应出勤工时", "入离职出勤天", "病假小时", "事假小时", 
                "迟到次数", "迟到分钟", "旷工次数", "旷工小时", 
                "平时加班", "周末加班", "合计加班时", 
                "平时加班费", "周末加班费", "合计加班费", "备注"
            ]
            result_df = result_df[columns]

            # 保存结果
            self.log("正在保存Excel文件...")
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name='Sheet1', index=False)
                
                # 获取工作表
                worksheet = writer.sheets['Sheet1']
                
                # 设置所有列的基本宽度
                for col in worksheet.column_dimensions:
                    worksheet.column_dimensions[col].width = 12
                
                # 特别调整某些列的宽度
                worksheet.column_dimensions['B'].width = 15  # 考勤月份
                worksheet.column_dimensions['E'].width = 15  # 组织单元名称
                worksheet.column_dimensions['U'].width = 50  # 备注列
                
                # 定义样式
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                normal_font = Font(name='宋体', size=9)
                header_font = Font(name='宋体', size=9, bold=True)
                
                # 设置所有单元格的边框、字体和对齐方式
                for row in range(1, worksheet.max_row + 1):
                    for col in range(1, len(columns) + 1):
                        cell = worksheet.cell(row=row, column=col)
                        cell.border = thin_border
                        
                        # 设置字体
                        if row == 1:  # 表头行
                            cell.font = header_font
                        else:  # 数据行
                            cell.font = normal_font
                        
                        if col == len(columns):  # 备注列
                            cell.alignment = Alignment(wrapText=True, vertical='center')
                            if row > 1 and cell.value:
                                cell.value = cell.value.replace("；", "\n")
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center')

            self.log(f"文件已保存: {Path(output_file).name}", "SUCCESS")

            return True, "处理完成！"
        except Exception as e:
            self.log(f"发生错误: {str(e)}", "ERROR")
            return False, f"处理出错：{str(e)}"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("考勤数据处理工具V1.0.0")
        self.processor = ExcelProcessor()
        self.processor.set_callback(self.add_log)  # 设置日志回调函数
        
        # 设置窗口大小和位置
        window_width = 800  # 增加宽度
        window_height = 400  # 增加高度
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 禁止调整窗口大小
        self.root.resizable(False, False)
        
        # 创建界面元素
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 输入文件选择
        input_frame = tk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(input_frame, text="输入文件：", width=10).pack(side=tk.LEFT)
        self.input_path = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.input_path, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(input_frame, text="浏览", width=10, command=self.select_input_file).pack(side=tk.LEFT)
        
        # 输出文件选择
        output_frame = tk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(output_frame, text="输出文件：", width=10).pack(side=tk.LEFT)
        self.output_path = tk.StringVar()
        tk.Entry(output_frame, textvariable=self.output_path, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(output_frame, text="浏览", width=10, command=self.select_output_file).pack(side=tk.LEFT)
        
        # 处理按钮
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        tk.Button(button_frame, text="开始处理", width=20, command=self.process_file).pack()
        
        # 添加日志显示区域
        log_frame = tk.LabelFrame(main_frame, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建文本框和滚动条的容器
        text_container = tk.Frame(log_frame)
        text_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ���建垂直滚动条
        v_scrollbar = tk.Scrollbar(text_container)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建水平滚动条
        h_scrollbar = tk.Scrollbar(text_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建文本框
        self.log_text = tk.Text(text_container, 
                               height=10,  # 固定高度
                               wrap=tk.NONE,  # 禁用自动换行
                               xscrollcommand=h_scrollbar.set,
                               yscrollcommand=v_scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置滚动条
        v_scrollbar.config(command=self.log_text.yview)
        h_scrollbar.config(command=self.log_text.xview)
        
        # 设置文本框只读
        self.log_text.config(state='disabled')
    
    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.input_path.set(file_path)
            # 设置默认的输出路径
            input_path = Path(file_path)
            default_output = input_path.parent / f"考勤结果统计_{input_path.stem}.xlsx"
            self.output_path.set(str(default_output))
    
    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="保存文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def add_log(self, message):
        """添加日志信息到文本框"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 滚动到最新内容
        self.log_text.config(state='disabled')
    
    def process_file(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        
        if not input_file or not output_file:
            messagebox.showerror("错误", "请选择输入和输出文件！")
            return
        
        # 清空日志
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # 开始处理
        self.processor.set_callback(self.add_log)  # 确保回调函数已设置
        success, message = self.processor.process_excel(input_file, output_file)
        
        if success:
            messagebox.showinfo("成功", message)
        else:
            messagebox.showerror("错误", message)

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
