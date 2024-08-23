import tkinter as tk
from tkinter import Menu, filedialog, messagebox
import shutil
import subprocess
import threading
import queue
import os
import sys

def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# 使用示例
procedure_path = resource_path("procedure")
print(f"Procedure path: {procedure_path}")


def create_check_window(title, keylist):
    check_window = tk.Toplevel(root)
    check_window.title(title)
    check_window.geometry("+%d+%d" % (root.winfo_rootx() + 200, root.winfo_rooty() + 200))  # 调整位置

    vars = []
    for key in keylist:
        frame = tk.Frame(check_window)
        frame.pack(pady=5)
        tk.Label(frame, text=key).pack(side=tk.LEFT)
        var = tk.IntVar(value=1)  # 默认勾选
        vars.append(var)
        tk.Checkbutton(frame, variable=var).pack(side=tk.LEFT)

    def confirm_checked():
        selected_keys = [keylist[i] for i, var in enumerate(vars) if var.get() == 1]
        with open("para_library/transmit1", "w") as file:
            for key in selected_keys:
                file.write(f"{key}\n")
        output_message("已设置好关键词，请确认后提交")
        check_window.destroy()

    confirm_button = tk.Button(check_window, text="确定", command=confirm_checked)
    confirm_button.pack(pady=10)

def open_journal_window(event=None):
    journal_menu = Menu(root, tearoff=0)
    journals = [
        "Angew_communication",
        "Angew_article",
        "JACS",
        "Adv_Mater",
        "JMCA_article",
        "JMCA_communication",
        "JMCA_chem_comm"
    ]
    for journal in journals:
        journal_menu.add_command(label=journal, command=lambda j=journal: select_journal(j))
    journal_menu.post(event.x_root, event.y_root)

def select_journal(journal):
    journal_text.set(journal)
    with open("para_library/current_template", "w") as file:
        file.write(journal)
    copy_style_dictionary(journal)

def copy_style_dictionary(journal):
    source_path = f"tem_library/{journal}/style_dictionary"
    destination_path = "para_library/style_dictionary"
    try:
        shutil.copy(source_path, destination_path)
        output_message(f"已选择{journal}模板，请继续设置关键词")
    except FileNotFoundError:
        output_message(f"文件 {source_path} 未找到")

def confirm_selection():
    journal = journal_text.get()
    if not journal:
        output_message("请先选择模板")
        return

    try:
        with open(f"tem_library/{journal}/preinstall_keylist", "r") as file:
            keylist = file.read().splitlines()
            create_check_window(f"{journal} 预设关键字列表", keylist)
    except FileNotFoundError:
        output_message(f"文件 {journal}/preinstall_keylist 未找到")

def select_file():
    file_path = filedialog.askopenfilename(title="选择文件")
    if file_path:
        with open("para_library/paper_position", "w") as file:
            file.write(file_path)
        file_path_var.set(file_path)  # 更新Entry显示的文件路径
        output_message(f"选择的文件: {file_path}")

def output_message(msg):
    """在文本框中输出提示信息"""
    message_text.insert(tk.END, msg + "\n")
    message_text.see(tk.END)  # 滚动到最新消息

def run_script():
    def target():
        process = subprocess.Popen(["python", "run_scripts.py"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        while True:
            output = process.stdout.readline()
            if output == '' and process.poll() is not None:
                break
            if output:
                output_queue.put(output.strip())
        rc = process.poll()
        for line in process.stderr.readlines():
            output_queue.put(line.strip())
        output_queue.put(f"脚本执行完成，返回码: {rc}")

    def update_text():
        try:
            while True:
                msg = output_queue.get_nowait()
                output_message(msg)
        except queue.Empty:
            pass
        root.after(100, update_text)

    output_queue = queue.Queue()
    threading.Thread(target=target).start()
    root.after(100, update_text)

root = tk.Tk()
root.title("Main Window")
root.geometry("800x600")

# Main frame for file selection
file_frame = tk.Frame(root)
file_frame.pack(pady=10)

tk.Label(file_frame, text="请选择需要排版的原文件：").pack(side=tk.LEFT)

select_file_button = tk.Button(file_frame, text="选择文件", command=select_file)
select_file_button.pack(side=tk.LEFT, padx=10)

# Entry to display selected file path
file_path_var = tk.StringVar()
file_path_entry = tk.Entry(file_frame, textvariable=file_path_var, width=50)
file_path_entry.pack(side=tk.LEFT, padx=10)

# 选择预设的模板的提示与选择期刊按钮
template_frame = tk.Frame(root)
template_frame.pack(pady=10)

template_label = tk.Label(template_frame, text="请选择预设的模板：")
template_label.pack(side=tk.LEFT)

# Button to select journal
journal_button = tk.Button(template_frame, text="选择模板")
journal_button.pack(side=tk.LEFT, padx=10)
journal_button.bind("<Button-1>", open_journal_window)

# Text box to display selected journal
journal_text = tk.StringVar()
journal_entry = tk.Entry(template_frame, textvariable=journal_text, width=50)
journal_entry.pack(side=tk.LEFT)

# 选择所需的keyword的提示与选择keywords按钮
keyword_frame = tk.Frame(root)
keyword_frame.pack(pady=10)

keyword_label = tk.Label(keyword_frame, text="请设置模板所需的关键词：")
keyword_label.pack(side=tk.LEFT)

# Confirm button for keywords selection
confirm_button = tk.Button(keyword_frame, text="设置关键词", command=confirm_selection)
confirm_button.pack(side=tk.LEFT, padx=10)

# 新增的执行脚本按钮
run_script_button = tk.Button(keyword_frame, text="提交排版", command=run_script)
run_script_button.pack(side=tk.LEFT, padx=10)

# 输出提示消息的文本框
message_text = tk.Text(root, wrap="word", width=80, height=10)
message_text.pack(pady=20)

root.mainloop()
