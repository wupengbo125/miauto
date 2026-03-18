#!/bin/bash

# 打印执行前的信息
echo "🚀 启动 miauto 自动化脚本引擎"
echo "请确保已经安装所需依赖 (pip install -r requirements.txt)"
echo "----------------------------------------"

# 如果配置了环境变量，可以在这里注入
# export MY_APP_PWD="yoursecretpassword"

# 运行 wscript.py
# --script 必须指定，指向刚写的示范案例
# --delay 3 倒数 3 秒让界面焦点切出
# --sleep 0.5 全局默认两步间休息 0.5 秒防止连击过快
python wscript.py --script sample.txt --delay 3 --sleep 0.5
