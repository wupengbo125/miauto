"""
wscript.py - 自定义 Windows 自动化脚本解释器
用法: python wscript.py --script your_script.txt

支持的指令:
  click  "文字", x:y [,radius]   - 在坐标 (x,y) ±radius(默认100) 范围内找文字并左键点击
  rclick "文字", x:y [,radius]   - 同上，右键点击
  dclick "文字", x:y [,radius]   - 同上，双击
  move   "文字", x:y [,radius]   - 同上，只移动鼠标不点击
  press  键名                     - 按下并释放一个键 (如 enter, delete, s, tab 等)
  keydown 键名                    - 按住某键不放
  keyup   键名                    - 释放某键
  input   文字 / $变量            - 输入键盘文字
  sleep   秒数                    - 睡眠 N 秒 (支持小数)
  function 函数名(参数=默认值):   - 定义函数块 (以缩进的行为函数体)
  函数名("参数")                  - 调用已定义的函数
  # 注释                          - 以 # 开头的行为注释
"""

import argparse
import re
import sys
import time
from datetime import datetime
import numpy as np
import pyautogui
import pyperclip
from rapidocr_onnxruntime import RapidOCR
import os
# ─── 自定义异常：脚本执行错误，立即中止 ─────────────────────────────────────
class ScriptError(Exception):
    pass

class SoftAbort(Exception):
    pass

# ─── 全局配置（由 main 赋值） ──────────────────────────────────────────────────
_log_file    = None
_sleep_after = 0.5   # 每个操作完成后的默认等待时间（秒），--sleep 参数可覆盖

def log_warn(msg: str):
    """把警告信息写入日志文件（如果启用了 --log）"""
    if _log_file:
        ts = datetime.now().strftime("%H:%M:%S")
        _log_file.write(f"[{ts}] {msg}\n")
        _log_file.flush()

# ─── 全局 OCR 引擎 ────────────────────────────────────────────────────────────
print("[wscript] 正在初始化 OCR 引擎...")
ocr_engine = RapidOCR()
print("[wscript] OCR 引擎就绪。")

# ─── OCR 截屏并在区域内找文字，返回中心坐标 ──────────────────────────────────
def _ocr_region(text: str, region: tuple, index: int = None, return_all: bool = False):
    """对指定区域截图跑 OCR，若找到 text 返回屏幕逻辑坐标（兼容 DPI 缩放），否则返回 None。
    若 return_all=True，则返回所有匹配坐标的列表。
    若找到多个匹配且 index 为 None 且 return_all=False，则抛出 ScriptError；
    若指定 index（1起始，按 y 坐标从上到下排序）则返回第 index 个。
    """
    left, top, width, height = region
    screenshot = pyautogui.screenshot(region=region)
    img = np.array(screenshot)

    # 物理像素尺寸 vs 逻辑区域尺寸 → 算出缩放比例
    phys_h, phys_w = img.shape[:2]
    scale_x = phys_w / width
    scale_y = phys_h / height

    result, _ = ocr_engine(img)
    if not result:
        return None

    # 收集所有包含 text 的匹配
    matches = []
    for item in result:
        box, recognized, *_ = item
        if text in recognized:
            cx = (box[0][0] + box[2][0]) / 2 / scale_x + left
            cy = (box[0][1] + box[2][1]) / 2 / scale_y + top
            matches.append((cx, cy, recognized))

    if len(matches) == 0:
        return [] if return_all else None

    # 按 y 坐标从小到大排序（从上到下）
    matches.sort(key=lambda m: m[1])

    if return_all:
        return [(x, y) for x, y, _ in matches]

    if len(matches) > 1:
        if index is not None:
            if index < 1 or index > len(matches):
                raise ScriptError(
                    f"找到 {len(matches)} 处「{text}」，但 :{index} 超出范围（1~{len(matches)}）"
                )
            cx, cy, _ = matches[index - 1]
            print(f"    (多个匹配，按 y 排序取第 {index} 个: ({cx:.0f},{cy:.0f}))")
            return (cx, cy)
        else:
            detail = "  |  ".join(f"「{r}」at ({x:.0f},{y:.0f})" for x, y, r in matches)
            raise ScriptError(
                f"OCR 在区域内找到 {len(matches)} 处含「{text}」的文字，无法确定点哪个，请缩小 radius 或用 :{1}/{2}/... 指定序号！\n"
                f"    找到的结果: {detail}"
            )

    cx, cy, _ = matches[0]
    return (cx, cy)


# ─── 单条指令的动作执行 ────────────────────────────────────────────────────────
def do_find_and_act(action: str, text: str, x1: float, y1: float, x2: float = None, y2: float = None, index: int = None, allow_fail: bool = False):
    """找到文字后根据 action 执行操作。如果没有提供 x2, y2，直接点坐标，跳过 OCR。"""
    if x2 is None or y2 is None:
        # 直接点，不做 OCR
        print(f"  🎯 直接点击坐标 ({x1}, {y1})，执行 {action}（跳过 OCR）")
        cx, cy = x1, y1
    else:
        # 在指定区域找
        left = min(x1, x2)
        top = min(y1, y2)
        width = max(x1, x2) - left
        height = max(y1, y2) - top
        region = (int(left), int(top), int(width), int(height))
        
        if action == "clickall":
            points = _ocr_region(text, region, return_all=True)
            if not points:
                msg = f"【clickall】在区域 {region} 内找不到「{text}」"
                if allow_fail:
                    print(f"  [?] {msg} → 允许失败，立刻终止当前代码块 (函数)...")
                    return False
                print(f"  ❌ {msg} → 脚本中止")
                log_warn(f"{msg} → 脚本中止")
                raise ScriptError(msg)
            
            print(f"  ✅ 找到 {len(points)} 个「{text}」，依次点击：")
            for cx, cy in points:
                print(f"     → 点击 ({cx:.1f}, {cy:.1f})")
                pyautogui.moveTo(cx, cy, duration=0.2)
                time.sleep(0.3)
                pyautogui.click()
                time.sleep(0.1)
            return True

        pos = _ocr_region(text, region, index=index)
        if pos is None:
            msg = f"【{action}】在区域 {region} 内找不到「{text}」"
            if allow_fail:
                print(f"  [?] {msg} → 允许失败，立刻终止当前代码块 (函数)...")
                return False
            print(f"  ❌ {msg} → 脚本中止")
            log_warn(f"{msg} → 脚本中止")
            raise ScriptError(msg)
        
        cx, cy = pos
        print(f"  ✅ 找到「{text}」at ({cx:.1f}, {cy:.1f})，执行 {action}")
        # 写入日志，格式可直接复制回脚本
        log_warn(f"区域内找到「{text}」at ({cx:.1f},{cy:.1f}) → 建议更新为固定坐标: click \"{text}\",{cx:.0f}:{cy:.0f}")

    pyautogui.moveTo(cx, cy, duration=0.2)
    time.sleep(0.3)  # 等鼠标稳定后再点，避免点早了
    if action == "click":
        pyautogui.click()
    elif action == "rclick":
        pyautogui.rightClick()
    elif action == "dclick":
        pyautogui.doubleClick()
    elif action == "move":
        pass  # 只移动，不点

    return True


# ─── 解析一行里的操作对象（文字或变量） ──────────────────────────────────────
def resolve_value(token: str, variables: dict):
    """解析值：
    - $var      → 从脚本变量查
    - env.NAME  → 从系统环境变量查（大小写不敏感，先原样找，找不到再试大写）
    - "text"    → 直接返回字符串
    """
    token = token.strip()
    if token.startswith("$"):
        var_name = token[1:]
        return variables.get(var_name, "")
    if token.startswith("env."):
        env_name = token[4:]
        # 先原样找，再试大写
        val = os.environ.get(env_name) or os.environ.get(env_name.upper())
        if val is None:
            raise ScriptError(f"环境变量 {env_name!r} 未设置，请先 set {env_name.upper()}=... 或在系统环境变量里配置")
        return val
    return token.strip('"').strip("'")


# ─── 解析 click/rclick/dclick/move 行 ────────────────────────────────────────
# 格式: action "text"[:N]|$var[:N] , x1:y1 [, x2:y2]
# :N 为可选序号，找到多个时按 y 从上到下取第 N 个
_ACTION_RE = re.compile(
    r'^(\??)(click|rclick|dclick|move|clickall)\s+'  # ? and action (groups 1, 2)
    r'(".*?"|\'.*?\'|\$\w+)(?::(\d+))?'  # text or $var, optional :N index (groups 3,4)
    r'\s*,\s*'
    r'([\d.]+)\s*:\s*([\d.]+)'         # x1:y1  (groups 5,6)
    r'(?:\s*,\s*([\d.]+)\s*:\s*([\d.]+))?' # optional x2:y2 (groups 7,8)
    r'\s*$'
)


# ─── 解析函数定义 ─────────────────────────────────────────────────────────────
_FUNC_DEF_RE = re.compile(r'^function\s+(\w+)\s*\((.*?)\)\s*:\s*$')

# 解析函数调用：func_name("arg") 或 func_name("arg1","arg2")
_FUNC_CALL_RE = re.compile(r'^(\w+)\((.*)?\)\s*$')


# ─── 执行一组 lines ───────────────────────────────────────────────────────────
def execute_lines(lines: list[str], functions: dict, variables: dict):
    i = 0
    while i < len(lines):
        raw = lines[i]
        # 去掉行内注释（# 后的内容），但保留字符串内的 # 不处理
        line = raw.strip()
        # 简单策略：只去掉不在引号内的 # 及其后内容
        stripped = re.sub(r'\s*#(?=(?:[^"]*"[^"]*")*[^"]*$).*$', '', line).strip()
        line = stripped
        i += 1

        # 空行 / 注释 跳过
        if not line or line.startswith("#"):
            continue

        # ── 赋值语句: name = "305" 或 name = $other_var 或 name=get_files ────────
        if "=" in line and not line.startswith("click") and not line.startswith("rclick") and not line.startswith("dclick") and not line.startswith("move") and not line.startswith("clickall") and "==" not in line:
            parts = line.split("=", 1)
            var_name = parts[0].strip()
            val_str  = parts[1].strip()
            
            if val_str.startswith("get_files(") and val_str.endswith(")"):
                args_str = val_str[10:-1]
                args = [resolve_value(a.strip(), variables) for a in args_str.split(",")]
                path = args[0]
                ext = args[1] if len(args) > 1 else ""
                import os
                var_val = []
                if os.path.exists(path) and os.path.isdir(path):
                    for f in os.listdir(path):
                        if ext and f.endswith(ext):
                            # strip extension
                            stripped = f[:-(len(ext)+1)] if f.endswith("."+ext) else f[:-len(ext)]
                            var_val.append(stripped)
                        elif not ext:
                            var_val.append(f)
                variables[var_name] = var_val
                continue
                
            var_val = resolve_value(val_str, variables)
            variables[var_name] = var_val
            continue

        # ── click / rclick / dclick / move ───────────────────────────────────
        m = _ACTION_RE.match(line)
        if m:
            allow_fail = bool(m.group(1))
            action  = m.group(2)
            text    = resolve_value(m.group(3), variables)
            index   = int(m.group(4)) if m.group(4) else None  # :N 序号
            x1      = float(m.group(5))
            y1      = float(m.group(6))
            x2      = float(m.group(7)) if m.group(7) else None
            y2      = float(m.group(8)) if m.group(8) else None
            if not do_find_and_act(action, text, x1, y1, x2, y2, index=index, allow_fail=allow_fail):
                raise SoftAbort()
            time.sleep(_sleep_after)
            continue

        # ── press ─────────────────────────────────────────────────────────────
        if line.startswith("press "):
            key = line[6:].strip()
            print(f"  ⌨️  press [{key}]")
            pyautogui.press(key)
            time.sleep(_sleep_after)
            continue

        # ── paste（写入剪贴板后 Ctrl+V，支持中文）────────────────────────────
        if line.startswith("paste "):
            val = resolve_value(line[6:].strip(), variables)
            print(f"  📋 paste 「{val}」")
            pyperclip.copy(val)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(_sleep_after)
            continue

        # ── keydown ───────────────────────────────────────────────────────────
        if line.startswith("keydown "):
            key = line[8:].strip()
            print(f"  ⌨️  keydown [{key}]")
            pyautogui.keyDown(key)
            time.sleep(_sleep_after)
            continue

        # ── keyup ─────────────────────────────────────────────────────────────
        if line.startswith("keyup "):
            key = line[6:].strip()
            print(f"  ⌨️  keyup [{key}]")
            pyautogui.keyUp(key)
            time.sleep(_sleep_after)
            continue

        # ── sleep（只执行指定的等待，不额外叠加全局 sleep）────────────────────
        if line.startswith("sleep "):
            secs = float(line[6:].strip())
            print(f"  💤 sleep {secs}s")
            time.sleep(secs)
            continue

        # ── input ─────────────────────────────────────────────────────────────
        if line.startswith("input "):
            val = resolve_value(line[6:].strip(), variables)
            print(f"  📝 input 「{val}」")
            pyautogui.typewrite(val, interval=0.05)
            time.sleep(_sleep_after)
            continue

        # ── for 循环 ───────────────────────────────────────────────────────────
        if line.startswith("for ") and line.endswith(":"):
            loop_header = line[4:-1].strip()
            
            # 收集缩进的循环体
            body = []
            while i < len(lines) and (lines[i].startswith("    ") or lines[i].startswith("\t")):
                body.append(lines[i][4:] if lines[i].startswith("    ") else lines[i][1:])
                i += 1
            
            if " in " in loop_header:
                item_var, list_var = [s.strip() for s in loop_header.split(" in ", 1)]
                if not list_var.startswith("$"):
                    raise ScriptError(f"for ... in ... 循环中，引用的列表变量必须以 $ 开头（例如: {item_var} in ${list_var}）")
                
                list_var_name = list_var[1:]
                lst = variables.get(list_var_name, [])
                if not isinstance(lst, list):
                    raise ScriptError(f"for ... in ... 循环中，变量 {list_var} 解析出的不是列表")
                
                for loop_i, item_val in enumerate(lst):
                    variables[item_var] = item_val
                    try:
                        execute_lines(body, functions, variables)
                    except SoftAbort:
                        print(f"  [?] 检测到可选中止 (位于 {item_var}={item_val} 次)，已跳出 for 循环！")
                        break
            else:
                try:
                    count = int(loop_header)
                except ValueError:
                    raise ScriptError(f"for 循环次数无效，请使用正整数: {line}")
                
                for loop_i in range(count):
                    try:
                        execute_lines(body, functions, variables)
                    except SoftAbort:
                        print(f"  [?] 检测到可选中止 (位于循环第 {loop_i + 1} 次)，已跳出 for 循环！")
                        break
            continue

        # ── function 定义（跳过函数体，函数体在解析阶段已收集） ──────────────
        if _FUNC_DEF_RE.match(line):
            # 跳过缩进的函数体行（已在预处理阶段收集）
            while i < len(lines) and (lines[i].startswith("    ") or lines[i].startswith("\t")):
                i += 1
            continue

        # ── 函数调用 ──────────────────────────────────────────────────────────
        m2 = _FUNC_CALL_RE.match(line)
        if m2:
            func_name = m2.group(1)
            args_raw  = m2.group(2).strip() if m2.group(2) else ""
            if func_name in functions:
                func_def = functions[func_name]
                # 解析实参：支持具名参数 key="value" 和位置参数 "value"
                named_args = {}
                pos_args   = []
                if args_raw:
                    for part in args_raw.split(","):
                        part = part.strip()
                        if "=" in part:
                            k, v = part.split("=", 1)
                            named_args[k.strip()] = resolve_value(v.strip(), variables)
                        else:
                            pos_args.append(resolve_value(part.strip(), variables))
                # 把形参映射到局部变量（具名优先，其次按位置，最后用默认值）
                local_vars = dict(variables)
                pos_idx = 0
                for param_name in func_def["params"]:
                    if param_name in named_args:
                        local_vars[param_name] = named_args[param_name]
                    elif pos_idx < len(pos_args):
                        local_vars[param_name] = pos_args[pos_idx]
                        pos_idx += 1
                    else:
                        local_vars[param_name] = func_def["defaults"].get(param_name, "")
                print(f"  📌 调用函数 {func_name}({', '.join(f'{k}={local_vars[k]}' for k in func_def['params'])})")
                execute_lines(func_def["body"], functions, local_vars)
                continue
            else:
                print(f"  ⚠️  未知函数「{func_name}」，跳过")
                continue

        msg = f"无法解析的行: {line!r}"
        print(f"  ❌ {msg} → 脚本中止")
        raise ScriptError(msg)


# ─── 预处理：收集所有 function 定义 ──────────────────────────────────────────
def preprocess(lines: list[str]) -> dict:
    """扫描所有行，把 function 定义提取出来，返回函数字典"""
    functions = {}
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _FUNC_DEF_RE.match(line)
        if m:
            func_name  = m.group(1)
            params_raw = m.group(2).strip()
            # 解析参数和默认值：name="" 或 name
            params   = []
            defaults = {}
            if params_raw:
                for part in params_raw.split(","):
                    part = part.strip()
                    if "=" in part:
                        pname, pdefault = part.split("=", 1)
                        pname    = pname.strip()
                        pdefault = pdefault.strip().strip('"').strip("'")
                        params.append(pname)
                        defaults[pname] = pdefault
                    else:
                        params.append(part)
            # 收集缩进的函数体
            body = []
            i += 1
            while i < len(lines) and (lines[i].startswith("    ") or lines[i].startswith("\t")):
                body.append(lines[i][4:] if lines[i].startswith("    ") else lines[i][1:])
                i += 1
            functions[func_name] = {"params": params, "defaults": defaults, "body": body}
            print(f"[wscript] 已加载函数: {func_name}({', '.join(params)})")
        else:
            i += 1
    return functions


# ─── 主入口 ──────────────────────────────────────────────────────────────────
def main():
    global _log_file, _sleep_after

    parser = argparse.ArgumentParser(description="wscript - 自定义 Windows 自动化脚本解释器")
    parser.add_argument("--script", required=True, help="要执行的 .txt 脚本文件路径")
    parser.add_argument("--delay", type=int, default=5, help="开始执行前的倒计时秒数（默认 5）")
    parser.add_argument("--log", default=None, help="把首次找不到的警告写入指定日志文件（如 w.log）")
    parser.add_argument("--sleep", type=float, default=0.5, help="每个操作后的全局默认等待秒数（默认 0.5）")
    args = parser.parse_args()

    _sleep_after = args.sleep
    print(f"[wscript] 全局操作间隔: {_sleep_after}s")

    # 打开日志文件
    if args.log:
        _log_file = open(args.log, "a", encoding="utf-8")
        _log_file.write(f"\n{'='*60}\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 脚本开始: {args.script}\n")
        print(f"[wscript] 日志将写入: {args.log}")

    # 读脚本
    try:
        with open(args.script, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"错误：找不到脚本文件 {args.script}")
        sys.exit(1)

    # 预处理：收集函数定义
    functions = preprocess(lines)

    # 倒计时
    if args.delay > 0:
        print(f"\n脚本将在 {args.delay} 秒后开始执行，请切换到目标窗口...")
        for i in range(args.delay, 0, -1):
            print(f"  倒计时 {i}...")
            time.sleep(1)

    print("\n▶️  开始执行脚本...")
    try:
        execute_lines(lines, functions, variables={})
        print("\n🎉  脚本执行完毕！")
    except SoftAbort:
        print("\n  [?] 脚本遭遇可选中止，已静默结束。")
    except ScriptError as e:
        print(f"\n💥  脚本中止: {e}")
        if _log_file:
            _log_file.write(f"[{datetime.now().strftime('%H:%M:%S')}] 脚本中止: {e}\n")
        sys.exit(1)
    finally:
        if _log_file:
            _log_file.write(f"[{datetime.now().strftime('%H:%M:%S')}] 脚本结束\n")
            _log_file.close()


if __name__ == "__main__":
    main()
