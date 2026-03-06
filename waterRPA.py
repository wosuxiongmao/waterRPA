import pyautogui
import time
import xlrd
import pyperclip
import re
try:
    import keyboard
except ImportError:
    keyboard = None


# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

COMMAND_ALIASES = {
    '1': 'left_click',
    'left_click': 'left_click',
    'left': 'left_click',
    'click': 'left_click',
    '左键': 'left_click',
    '左键单击': 'left_click',
    '2': 'double_click',
    'double_click': 'double_click',
    'double': 'double_click',
    'dblclick': 'double_click',
    '双击': 'double_click',
    '左键双击': 'double_click',
    '3': 'right_click',
    'right_click': 'right_click',
    'right': 'right_click',
    'rclick': 'right_click',
    '右键': 'right_click',
    '4': 'input',
    'input': 'input',
    'type': 'input',
    '输入': 'input',
    '5': 'wait',
    'wait': 'wait',
    'sleep': 'wait',
    '等待': 'wait',
    '6': 'scroll',
    'scroll': 'scroll',
    'wheel': 'scroll',
    '滚轮': 'scroll',
    '7': 'if_exists',
    'if': 'if_exists',
    'if_exists': 'if_exists',
    '如果': 'if_exists',
    'ifnot': 'if_not_exists',
    'if_not': 'if_not_exists',
    'if_not_exists': 'if_not_exists',
    '如果不存在': 'if_not_exists',
    '8': 'else',
    'else': 'else',
    '否则': 'else',
    '9': 'endif',
    'endif': 'endif',
    'end_if': 'endif',
    '结束如果': 'endif',
    '10': 'while_exists',
    'while': 'while_exists',
    'while_exists': 'while_exists',
    '循环': 'while_exists',
    '12': 'while_not_exists',
    'whilenot': 'while_not_exists',
    'while_not': 'while_not_exists',
    'while_not_exists': 'while_not_exists',
    '循环直到不存在': 'while_not_exists',
    '11': 'endwhile',
    'endwhile': 'endwhile',
    'end_while': 'endwhile',
    '结束循环': 'endwhile',
    '13': 'for_loop',
    'for': 'for_loop',
    'for_loop': 'for_loop',
    '循环次数': 'for_loop',
    '14': 'endfor',
    'endfor': 'endfor',
    'end_for': 'endfor',
    '结束for': 'endfor',
    '结束循环次数': 'endfor',
    '15': 'move_mouse',
    'move': 'move_mouse',
    'move_mouse': 'move_mouse',
    'mousemove': 'move_mouse',
    '移动': 'move_mouse',
    '鼠标移动': 'move_mouse',
    '16': 'stop_if_exists',
    'stop_if_exists': 'stop_if_exists',
    'break_if_exists': 'stop_if_exists',
    '存在则结束': 'stop_if_exists',
    '17': 'stop',
    'stop': 'stop',
    'break': 'stop',
    '结束': 'stop',
    '停止': 'stop'
}

PAUSE_HOTKEY = 'f10'
RESUME_HOTKEY = 'f9'
STOP_HOTKEY = 'f8'
isPaused = False
isStopped = False
hotkeysReady = False


def pauseRun():
    global isPaused
    if isStopped:
        return
    if not isPaused:
        isPaused = True
        print("已暂停，按", RESUME_HOTKEY, "继续")


def resumeRun():
    global isPaused
    if isStopped:
        return
    if isPaused:
        isPaused = False
        print("已继续")


def stopRun():
    global isPaused
    global isStopped
    isStopped = True
    isPaused = False
    print("已停止，准备退出")


def setupHotkeys():
    global hotkeysReady
    if hotkeysReady:
        return
    if keyboard is None:
        print("未安装keyboard库，热键功能不可用")
        return
    keyboard.add_hotkey(PAUSE_HOTKEY, pauseRun, suppress=True)
    keyboard.add_hotkey(RESUME_HOTKEY, resumeRun, suppress=True)
    keyboard.add_hotkey(STOP_HOTKEY, stopRun, suppress=True)
    hotkeysReady = True
    print("热键已启用:", PAUSE_HOTKEY, "暂停;", RESUME_HOTKEY, "继续;", STOP_HOTKEY, "停止")


def flowControlPoint():
    while isPaused and (not isStopped):
        time.sleep(0.1)
    return isStopped


def normalizeCmdToken(token):
    token = str(token).strip()
    if token.endswith('.0'):
        intPart = token[:-2]
        if intPart.lstrip('-').isdigit():
            return intPart
    return token


def parseCmd(rawCmd):
    cmdText = str(rawCmd).strip()
    if cmdText == "":
        return "", []
    if ',' in cmdText:
        parts = [part.strip() for part in cmdText.split(',')]
        firstParts = parts[0].split(None, 1)
        cmdHead = normalizeCmdToken(firstParts[0]).lower()
        cmdArgs = []
        if len(firstParts) > 1:
            cmdArgs.append(firstParts[1].strip())
        cmdArgs.extend(parts[1:])
    else:
        parts = cmdText.split(None, 1)
        cmdHead = normalizeCmdToken(parts[0]).lower()
        cmdArgs = [parts[1].strip()] if len(parts) > 1 else []
    return cmdHead, cmdArgs


def resolveCommand(cmdHead):
    return COMMAND_ALIASES.get(cmdHead, cmdHead)


def renderTemplate(text, variables):
    if variables is None:
        return text

    def replacer(match):
        key = match.group(1)
        if key in variables:
            return str(variables[key])
        return match.group(0)

    return re.sub(r'\{([A-Za-z_][A-Za-z0-9_]*)\}', replacer, text)


def toNumber(valueText, defaultValue):
    try:
        return float(valueText)
    except (TypeError, ValueError):
        return defaultValue


def cellText(sheet, rowIndex, colIndex, variables=None):
    cell = sheet.row(rowIndex)[colIndex]
    if cell.ctype == 0:
        return ""
    return renderTemplate(str(cell.value).strip(), variables)


def cellInt(sheet, rowIndex, colIndex, defaultValue, variables=None):
    cell = sheet.row(rowIndex)[colIndex]
    if cell.ctype == 2:
        value = int(cell.value)
        if value != 0:
            return value
    if cell.ctype == 1:
        valueText = renderTemplate(str(cell.value).strip(), variables)
        parsed = toNumber(valueText, defaultValue)
        if parsed != 0:
            return int(parsed)
    return defaultValue


def parseOffset(cmdArgs, variables=None):
    offsetX = 0
    offsetY = 0
    if len(cmdArgs) > 1:
        try:
            offsetX = int(float(renderTemplate(cmdArgs[0], variables)))
            offsetY = int(float(renderTemplate(cmdArgs[1], variables)))
        except ValueError:
            print("偏移坐标格式有误，按(0,0)处理")
    return offsetX, offsetY


def locateImage(img, confidence=0.9):
    try:
        return pyautogui.locateCenterOnScreen(img, confidence=confidence), False
    except pyautogui.ImageNotFoundException:
        return None, False
    except OSError:
        print("图片文件不可用:", img)
        return None, True
    except Exception as e:
        print("图片识别异常，按未找到处理:", img, e)
        return None, False


def imageExists(img, checkTimes=1, interval=0.1, confidence=0.9):
    if not img:
        return False
    checkTimes = max(1, int(checkTimes))
    for _ in range(checkTimes):
        if flowControlPoint():
            return False
        location, invalidImage = locateImage(img, confidence=confidence)
        if invalidImage:
            return False
        if location is not None:
            return True
        time.sleep(interval)
    return False


def evaluateImageCondition(sheet, rowIndex, cmd, cmdArgs, variables=None):
    targetImg = cellText(sheet, rowIndex, 1, variables)
    if targetImg == "" and len(cmdArgs) > 0:
        targetImg = renderTemplate(cmdArgs[0], variables)
    checkTimes = cellInt(sheet, rowIndex, 2, 1, variables)
    exists = imageExists(targetImg, checkTimes=checkTimes)
    if cmd == 'if_exists' or cmd == 'while_exists':
        return targetImg, exists
    if cmd == 'if_not_exists' or cmd == 'while_not_exists':
        return targetImg, (not exists)
    return targetImg, False


def buildWhileMap(sheet):
    whileStartToEnd = {}
    whileEndToStart = {}
    stack = []
    i = 1
    while i < sheet.nrows:
        cmdHead, _ = parseCmd(sheet.row(i)[0].value)
        if cmdHead != "":
            cmd = resolveCommand(cmdHead)
            if cmd == 'while_exists' or cmd == 'while_not_exists':
                stack.append(i)
            elif cmd == 'endwhile':
                if len(stack) == 0:
                    print('第', i + 1, '行 endwhile 找不到匹配的 while')
                else:
                    start = stack.pop()
                    whileStartToEnd[start] = i
                    whileEndToStart[i] = start
        i += 1

    for start in stack:
        print('第', start + 1, '行 while 找不到匹配的 endwhile')

    return whileStartToEnd, whileEndToStart


def buildForMap(sheet):
    forStartToEnd = {}
    forEndToStart = {}
    stack = []
    rowIndex = 1
    while rowIndex < sheet.nrows:
        cmdHead, _ = parseCmd(sheet.row(rowIndex)[0].value)
        if cmdHead != "":
            cmd = resolveCommand(cmdHead)
            if cmd == 'for_loop':
                stack.append(rowIndex)
            elif cmd == 'endfor':
                if len(stack) == 0:
                    print('第', rowIndex + 1, '行 endfor 找不到匹配的 for')
                else:
                    start = stack.pop()
                    forStartToEnd[start] = rowIndex
                    forEndToStart[rowIndex] = start
        rowIndex += 1

    for start in stack:
        print('第', start + 1, '行 for 找不到匹配的 endfor')

    return forStartToEnd, forEndToStart


def parseForConfig(cmdArgs, variables):
    tokens = []
    for arg in cmdArgs:
        rendered = renderTemplate(arg, variables).replace(',', ' ')
        splitTokens = [part for part in rendered.split() if part != ""]
        tokens.extend(splitTokens)

    if len(tokens) < 3:
        return None

    varName = tokens[0]
    if not re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', varName):
        return None

    startValue = toNumber(tokens[1], None)
    endValue = toNumber(tokens[2], None)
    if startValue is None or endValue is None:
        return None

    if not float(startValue).is_integer() or not float(endValue).is_integer():
        return None

    stepValue = 1.0
    if len(tokens) > 3:
        stepValue = toNumber(tokens[3], None)
        if stepValue is None:
            return None
        if not float(stepValue).is_integer():
            return None

    if stepValue == 0:
        return None

    return {
        "var_name": varName,
        "start": int(startValue),
        "end": int(endValue),
        "step": int(stepValue)
    }


def inForRange(currentValue, endValue, stepValue):
    if stepValue > 0:
        return currentValue <= endValue
    return currentValue >= endValue


def mouseClick(clickTimes, lOrR, img, reTry, offsetX=0, offsetY=0):
    reTry = int(reTry)
    if reTry == 1:
        while True:
            if flowControlPoint():
                return False
            location, invalidImage = locateImage(img, confidence=0.9)
            if invalidImage:
                return False
            if location is not None:
                pyautogui.click(location.x + offsetX, location.y + offsetY, clicks=clickTimes, interval=0.2,
                                duration=0.2, button=lOrR)
                return True
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            if flowControlPoint():
                return False
            location, invalidImage = locateImage(img, confidence=0.9)
            if invalidImage:
                return False
            if location is not None:
                pyautogui.click(location.x + offsetX, location.y + offsetY, clicks=clickTimes, interval=0.2,
                                duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            if flowControlPoint():
                return False
            location, invalidImage = locateImage(img, confidence=0.9)
            if invalidImage:
                return False
            if location is not None:
                pyautogui.click(location.x + offsetX, location.y + offsetY, clicks=clickTimes, interval=0.2,
                                duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)
    return False


# 任务
def mainWork(img):
    rowIndex = 1
    ifStack = []
    variableContext = {}
    forState = {}
    whileStartToEnd, whileEndToStart = buildWhileMap(sheet1)
    forStartToEnd, forEndToStart = buildForMap(sheet1)

    while rowIndex < sheet1.nrows:
        if flowControlPoint():
            return False
        cmdHead, cmdArgs = parseCmd(sheet1.row(rowIndex)[0].value)
        if cmdHead == "":
            rowIndex += 1
            continue

        cmdArgs = [renderTemplate(arg, variableContext) for arg in cmdArgs]
        cmd = resolveCommand(cmdHead)

        parentExecute = True
        for block in ifStack:
            if not block["execute"]:
                parentExecute = False
                break

        if cmd == 'if_exists' or cmd == 'if_not_exists':
            targetImg, condition = evaluateImageCondition(sheet1, rowIndex, cmd, cmdArgs, variableContext)
            ifStack.append({
                "parentExecute": parentExecute,
                "condition": condition,
                "execute": parentExecute and condition,
                "hasElse": False
            })
            print("IF判断", targetImg if targetImg else "(空图片)", "结果:", condition)
            rowIndex += 1
            continue

        if cmd == 'else':
            if len(ifStack) == 0:
                print('第', rowIndex + 1, '行 else 找不到匹配的 if')
            else:
                current = ifStack[-1]
                if current["hasElse"]:
                    print('第', rowIndex + 1, '行 else 重复，已忽略')
                else:
                    current["hasElse"] = True
                    current["execute"] = current["parentExecute"] and (not current["condition"])
                    print("ELSE分支执行:", current["execute"])
            rowIndex += 1
            continue

        if cmd == 'endif':
            if len(ifStack) == 0:
                print('第', rowIndex + 1, '行 endif 找不到匹配的 if')
            else:
                ifStack.pop()
                print("END IF")
            rowIndex += 1
            continue

        if cmd == 'while_exists' or cmd == 'while_not_exists':
            endRow = whileStartToEnd.get(rowIndex)
            if endRow is None:
                print('第', rowIndex + 1, '行 while 找不到匹配的 endwhile')
                rowIndex += 1
                continue

            if not parentExecute:
                rowIndex = endRow + 1
                continue

            targetImg, condition = evaluateImageCondition(sheet1, rowIndex, cmd, cmdArgs, variableContext)
            if condition:
                print("WHILE判断", targetImg if targetImg else "(空图片)", "结果:", condition)
                rowIndex += 1
            else:
                print("WHILE判断", targetImg if targetImg else "(空图片)", "结果:", condition, "跳过循环")
                rowIndex = endRow + 1
            continue

        if cmd == 'endwhile':
            startRow = whileEndToStart.get(rowIndex)
            if startRow is None:
                print('第', rowIndex + 1, '行 endwhile 找不到匹配的 while')
                rowIndex += 1
                continue

            if not parentExecute:
                rowIndex += 1
                continue

            startHead, startArgs = parseCmd(sheet1.row(startRow)[0].value)
            startArgs = [renderTemplate(arg, variableContext) for arg in startArgs]
            startCmd = resolveCommand(startHead)
            targetImg, condition = evaluateImageCondition(sheet1, startRow, startCmd, startArgs, variableContext)
            if condition:
                rowIndex = startRow + 1
                print("WHILE继续", targetImg if targetImg else "(空图片)")
            else:
                rowIndex += 1
                print("WHILE结束", targetImg if targetImg else "(空图片)")
            continue

        if cmd == 'for_loop':
            endRow = forStartToEnd.get(rowIndex)
            if endRow is None:
                print('第', rowIndex + 1, '行 for 找不到匹配的 endfor')
                rowIndex += 1
                continue

            if not parentExecute:
                rowIndex = endRow + 1
                continue

            fullForArgs = list(cmdArgs)
            inlineSpec = cellText(sheet1, rowIndex, 1, variableContext)
            if inlineSpec != "":
                fullForArgs.append(inlineSpec)

            config = parseForConfig(fullForArgs, variableContext)
            if config is None:
                print('第', rowIndex + 1, '行 for 参数格式错误，示例: for i,1,5,1')
                rowIndex = endRow + 1
                continue

            currentValue = config["start"]
            if not inForRange(currentValue, config["end"], config["step"]):
                print("FOR跳过", config["var_name"], "初始值不在范围内")
                rowIndex = endRow + 1
                continue

            varName = config["var_name"]
            hadPrevValue = varName in variableContext
            prevValue = variableContext.get(varName)
            variableContext[varName] = currentValue
            forState[rowIndex] = {
                "var_name": varName,
                "current": currentValue,
                "end": config["end"],
                "step": config["step"],
                "had_prev_value": hadPrevValue,
                "prev_value": prevValue
            }
            print("FOR开始", varName, "=", currentValue, "到", config["end"], "步长", config["step"])
            rowIndex += 1
            continue

        if cmd == 'endfor':
            startRow = forEndToStart.get(rowIndex)
            if startRow is None:
                print('第', rowIndex + 1, '行 endfor 找不到匹配的 for')
                rowIndex += 1
                continue

            if not parentExecute:
                rowIndex += 1
                continue

            state = forState.get(startRow)
            if state is None:
                rowIndex += 1
                continue

            nextValue = state["current"] + state["step"]
            if inForRange(nextValue, state["end"], state["step"]):
                state["current"] = nextValue
                variableContext[state["var_name"]] = nextValue
                print("FOR继续", state["var_name"], "=", nextValue)
                rowIndex = startRow + 1
            else:
                print("FOR结束", state["var_name"])
                if state["had_prev_value"]:
                    variableContext[state["var_name"]] = state["prev_value"]
                elif state["var_name"] in variableContext:
                    del variableContext[state["var_name"]]
                del forState[startRow]
                rowIndex += 1
            continue

        if not parentExecute:
            rowIndex += 1
            continue

        if cmd == 'stop':
            print("收到结束指令，停止循环")
            return False
        elif cmd == 'stop_if_exists':
            targetImg = cellText(sheet1, rowIndex, 1, variableContext)
            if targetImg == "" and len(cmdArgs) > 0:
                targetImg = cmdArgs[0]
            if targetImg == "":
                print("stop_if_exists 缺少图片，已忽略")
                rowIndex += 1
                continue
            checkTimes = cellInt(sheet1, rowIndex, 2, 1, variableContext)
            if imageExists(targetImg, checkTimes=checkTimes):
                print("结束指令命中，停止循环:", targetImg)
                return False
            print("结束指令未命中，继续执行:", targetImg)
        elif cmd == 'left_click':
            img = cellText(sheet1, rowIndex, 1, variableContext)
            reTry = cellInt(sheet1, rowIndex, 2, 1, variableContext)
            offsetX, offsetY = parseOffset(cmdArgs, variableContext)
            coordMatch = re.match(r'^\s*(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)\s*$', img)
            if coordMatch is not None:
                targetX = int(float(coordMatch.group(1)))
                targetY = int(float(coordMatch.group(2)))
                pyautogui.click(targetX, targetY, clicks=1, interval=0.2, duration=0.2, button="left")
                print("左键坐标", targetX, targetY)
            elif img != "":
                mouseClick(1, "left", img, reTry, offsetX, offsetY)
                print("单击左键", img)
            elif len(cmdArgs) > 1:
                pyautogui.click(offsetX, offsetY, clicks=1, interval=0.2, duration=0.2, button="left")
                print("左键坐标", offsetX, offsetY)
            else:
                pyautogui.click(clicks=1, interval=0.2, duration=0.2, button="left")
                print("左键当前鼠标位置")
        elif cmd == 'double_click':
            img = cellText(sheet1, rowIndex, 1, variableContext)
            reTry = cellInt(sheet1, rowIndex, 2, 1, variableContext)
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        elif cmd == 'right_click':
            img = cellText(sheet1, rowIndex, 1, variableContext)
            reTry = cellInt(sheet1, rowIndex, 2, 1, variableContext)
            offsetX, offsetY = parseOffset(cmdArgs, variableContext)
            coordMatch = re.match(r'^\s*(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)\s*$', img)

            if coordMatch is not None:
                targetX = int(float(coordMatch.group(1)))
                targetY = int(float(coordMatch.group(2)))
                pyautogui.click(targetX, targetY, clicks=1, interval=0.2, duration=0.2, button="right")
                print("右键坐标", targetX, targetY)
            elif img != "":
                mouseClick(1, "right", img, reTry, offsetX, offsetY)
                print("右键", img)
            elif len(cmdArgs) > 1:
                pyautogui.click(offsetX, offsetY, clicks=1, interval=0.2, duration=0.2, button="right")
                print("右键坐标", offsetX, offsetY)
            else:
                pyautogui.click(clicks=1, interval=0.2, duration=0.2, button="right")
                print("右键当前鼠标位置")
        elif cmd == 'move_mouse':
            img = cellText(sheet1, rowIndex, 1, variableContext)
            offsetX, offsetY = parseOffset(cmdArgs, variableContext)
            coordMatch = re.match(r'^\s*(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)\s*$', img)
            if coordMatch is not None:
                targetX = int(float(coordMatch.group(1)))
                targetY = int(float(coordMatch.group(2)))
                pyautogui.moveTo(targetX, targetY, duration=0.2)
                print("移动鼠标到坐标", targetX, targetY)
            elif img != "":
                location, invalidImage = locateImage(img, confidence=0.9)
                if invalidImage:
                    rowIndex += 1
                    continue
                if location is not None:
                    pyautogui.moveTo(location.x + offsetX, location.y + offsetY, duration=0.2)
                    print("移动鼠标到图片", img, "偏移", offsetX, offsetY)
                else:
                    print("未找到图片，未移动鼠标", img)
            elif len(cmdArgs) > 1:
                pyautogui.moveTo(offsetX, offsetY, duration=0.2)
                print("移动鼠标到坐标", offsetX, offsetY)
            else:
                print("移动鼠标指令缺少坐标或图片")
        elif cmd == 'input':
            inputValue = cellText(sheet1, rowIndex, 1, variableContext)
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", inputValue)
        elif cmd == 'wait':
            waitText = cellText(sheet1, rowIndex, 1, variableContext)
            waitCandidates = []
            if waitText != "":
                waitCandidates.append(waitText)
            waitCandidates.extend(cmdArgs)

            waitTokens = []
            for candidate in waitCandidates:
                parts = re.split(r'[\s,]+', str(candidate).strip())
                for part in parts:
                    if part != "":
                        waitTokens.append(part)

            waited = False
            for token in waitTokens:
                try:
                    waitTime = float(token)
                    if waitTime < 0:
                        raise ValueError
                    time.sleep(waitTime)
                    print("等待", waitTime, "秒")
                    waited = True
                    break
                except (TypeError, ValueError):
                    print("等待参数无效，尝试下一个:", token)

            if not waited:
                print("等待参数都无效，已跳过")
        elif cmd == 'scroll':
            scroll = int(float(cellText(sheet1, rowIndex, 1, variableContext)))
            pyautogui.scroll(scroll)
            print("滚轮滑动", scroll, "距离")
        else:
            print('第', rowIndex + 1, "行未知指令:", cmdHead)
        rowIndex += 1
    return True


if __name__ == '__main__':
    file = 'cmd.xls'
    pyautogui.FAILSAFE = True
    setupHotkeys()
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('RPA启动~')
   
    while True:
        if flowControlPoint():
            break
        if not mainWork(sheet1):
            break
        time.sleep(1)
    if keyboard is not None and hotkeysReady:
        keyboard.unhook_all_hotkeys()
#         print("等待0.1秒")
# else:
#     print('输入有误或者已经退出!')
