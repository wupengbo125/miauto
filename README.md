# wscript - 极简 Windows 自动化执行器

抛弃复杂的脚本逻辑，只保留最纯粹直接的自动化实操功能。支持绝对坐标点击、固定区域 OCR 找字点击及常见键盘操作。

---

## 1. 抓取坐标 (`coords.py`)
写脚本前最常用的辅助工具，用于快速抓取屏幕坐标。无需管理员权限可直接运行。

```bash
# 默认模式：按 Ctrl+Z 复制完整语句，如 click "test",123:456
python coords.py

# ONLY 模式：按 Ctrl+Z 只复制坐标，如 123:456
python coords.py --only
```

---

## 2. 核心语法 (`*.txt` 脚本文件)

### 2.1 鼠标点击 (`click` / `dclick` / `rclick` / `move`)
目前支持三种点击模式：

**🟢 模式 A：直接点击（跳过 OCR，极速）**
直接向引擎提供一套绝对坐标即可执行操作。
```text
# 这里的 "登录" 只是给你当注释看的，引擎会直接点 1775:14
click "登录", 1775:14
```

**🟢 模式 B：区域 OCR 寻找点击**
在指定的左上角、右下角「框选范围」内寻找指定文字后点击。
```text
# 在两组坐标画出的矩形内寻找并点击 "导入文件"
click "导入文件", 1465:873, 1694:928

# 如果框内的匹配项不止一个，用 :N 按从上往下的顺序选第 N 个
click "板块":1, 851:428, 1694:836
```

**🟢 模式 C：区域全量点击 (`clickall`)**
在指定的框选范围内，找出**所有**匹配的文字，并自动依次逐个点击。（非常适合配合 `keydown ctrl` 实现多选）
```text
# 按住 Ctrl，在区域内寻找所有 "310" 并逐一点击，最后松开 Ctrl
keydown ctrl
clickall "310", 657:394, 1436:655
keyup ctrl
```

### 2.2 柔性跳过机制 (`?` 前缀)
如果不确定某个文字是否出现，可以在执行命令前加上 `?`。
如果找不到目标文字：
- **普通动作**：立即抛出报错，中止整个脚本。
- **带 `?` 动作：抛出柔性中止，直接跳出并中断当前的 `for` 循环或 `function`，且不影响其外层代码继续执行！**

```text
for 100:
    # 如果找不到 "0308"，会直接跳出最内层的 for 循环，不会导致程序卡死报错
    ?click "0308":1, 657:394, 1436:655
    click "删除", 1518:448
```

### 2.3 赋值与系统内置函数
支持标准的变量赋值，包括使用内置获取目录文件列表的函数。
注意：**为了代码严谨性（一眼辨别谁是被调用的），后续使用该变量时必须前置带上 `$` 符号！**

```text
# 常规赋值
name = "318po"
# 读取特定的目录系统下，指定扩展名的纯文本文件名列表
names = get_files("./daily_stock", "txt")

# 使用时必须加上 $
paste $name
```

### 2.4 文本输入与粘贴 (`paste`)
抛弃易错的键盘输入，一律通过操作系统的剪贴板进行，完美支持中文。
```text
paste "18592012523"        # 粘贴固定文本
paste $name                # 粘贴函数变量
paste env.ths_password     # 读取操作系统的系统环境变量，并粘贴
```

### 2.5 键盘与等待 (`press` / `sleep` / `keydown` / `keyup`)
```text
press s              # 按一下 s
press enter          # 按下回车键
keydown ctrl         # 长按 Ctrl 不松
keyup ctrl           # 松开 Ctrl
sleep 1.5            # 强行等待 1.5 秒
```

### 2.6 循环结构 (`for`)
```text
# 1. 计数制循环（例如重复执行100次）
for 100:
    ?click "0308":1, 657:394, 1436:655
    press enter

# 2. 列表遍历循环（必须通过 $ 提取被调用的列表变量）
names = get_files("./daily_stock", "txt")
for name in $names:
    click "新建板块", 1499:414
    paste $name
```

### 2.7 函数提取 (`function`)
```text
# 定义
function delete_board(name):
    # 将收到的参数以 $name 变量引用使用
    ?clickall $name, 657:394, 1436:655
    click "删除板块", 1518:448
    press enter

# 调用
delete_board(name="318po")
delete_board("311前高")
# 如果也是拿外部的变量传参进行套娃：
delete_board($my_name)
```

---

## 3. 运行脚本 (`wscript.py`)

配置好脚本后即可启动：

```bash
# --delay 2 : 开启 2 秒倒计时，给你留时间切回游戏或 APP 界面
# --sleep 0.3 : 每次点击/按键执行完后，默认硬等 0.3 秒，让界面刷新跟得上
# --log w.log : 找不到的错可以直接写入日志记录
python wscript.py --script your_script.txt --delay 2 --sleep 0.3 --log w.log
```
