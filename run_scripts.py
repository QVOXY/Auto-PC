import subprocess
import time

# 定义脚本路径
scripts = [
    "scripts_library/choose_tem.py",
    "scripts_library/execute_srs.py",
    "scripts_library/remove_empty_line.py",
    "scripts_library/write_to_tem.py",
    "scripts_library/style_operation.py",
    "scripts_library/delete_keyword.py"
]

# 依次运行每个脚本
for script in scripts:
    try:
        print(f"正在运行脚本: {script}")
        result = subprocess.run(["python", script], check=True)
        print(f"脚本 {script} 运行成功")
    except subprocess.CalledProcessError as e:
        print(f"脚本 {script} 运行失败: {e}")
        break
    # 在运行下一个脚本之前等待0.1秒
    time.sleep(0.1)
