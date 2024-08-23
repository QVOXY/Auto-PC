import subprocess
import time

# 读取para_library\transmit1里的内容存入一个列表名为keylist
with open(r'para_library\transmit1', 'r', encoding='utf-8') as file:
    keylist = file.readlines()

# 根据列表里元素的数量调用srs_library里的脚本
num_elements = len(keylist)

# 定义脚本名称列表
scripts = {
    4: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs4.py'],
    5: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs4.py'],
    6: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs4.py'],
    7: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs3_6.py', 'srs4.py'],
    8: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs3_6.py', 'srs3_7.py', 'srs4.py'],
    9: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs3_6.py', 'srs3_7.py', 'srs3_8.py', 'srs4.py'],
    10: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs3_6.py', 'srs3_7.py', 'srs3_8.py', 'srs3_9.py', 'srs4.py'],
    11: ['srs1.py', 'srs2.py', 'srs3_1.py', 'srs3_2.py', 'srs3_3.py', 'srs3_4.py', 'srs3_5.py', 'srs3_6.py', 'srs3_7.py', 'srs3_8.py', 'srs3_9.py', 'srs3_10.py', 'srs4.py']
}

# 根据元素数量调用相应的脚本
if num_elements in scripts:
    for script in scripts[num_elements]:
        subprocess.run(['python', f'srs_library\\{script}'])
        time.sleep(0.1)  # 暂停0.1秒
else:
    print(f"列表元素数量为{num_elements}，不符合调用条件。")

