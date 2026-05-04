#Requires AutoHotkey v2.0
#SingleInstance Force
#MaxThreadsPerHotkey 2

; ============================
; 管理员权限自提升
; ============================
; 如果游戏以管理员身份运行，脚本不是管理员时 PostMessage 会被静默拦截。
; 第一次运行会弹 UAC，确认即可。
if !A_IsAdmin {
    try Run '*RunAs "' A_AhkPath '" "' A_ScriptFullPath '"'
    ExitApp()
}

; ============================
; 配置
; ============================
global TARGET_EXE := "NRC-Win64-Shipping.exe"

; 即使窗口最小化或被隐藏，也允许 WinExist / ControlClick / PostMessage 找到它
DetectHiddenWindows true

; ============================
; 状态
; ============================
global isRunning := false
global gameHwnd := 0
global lastClickX := 0   ; 记录最后一次点击坐标，给 Click down/up 沿用
global lastClickY := 0

; ============================
; 热键: Ctrl+= 切换启动/停止 (全局生效)
; ============================
^=:: {
    global isRunning, gameHwnd, lastClickX, lastClickY
    isRunning := !isRunning ; 切换状态

    if (isRunning) {
        ; 启动前找游戏窗口
        hwnd := FindGameWindow()
        if (!hwnd) {
            isRunning := false
            ToolTip("未找到游戏窗口 (ahk_exe " TARGET_EXE ")")
            SetTimer(RemoveToolTip, -3000)
            return
        }
        gameHwnd := hwnd
        lastClickX := 0
        lastClickY := 0
        ToolTip("Script Started (Ctrl+= 停止)`n目标 hwnd: " hwnd)
        SetTimer(RemoveToolTip, -2000)
    } else {
        ToolTip("Script Stopped")
        SetTimer(RemoveToolTip, -2000)
        return ; 停止时打断后面的动作
    }

    ; ============================
    ; 主循环
    ; ============================
    while (isRunning) {
        ; 每轮开始前确认游戏窗口还在
        if (!WinExist("ahk_id " gameHwnd)) {
            ToolTip("游戏窗口已不存在，停止运行")
            SetTimer(RemoveToolTip, -3000)
            isRunning := false
            break
        }

        GetClientSize(gameHwnd, &w, &h)

        ; 1. 按 m 键
        if (!isRunning)
            break
        ToolTip("Action: Pressing M Key")
        BgKeyTap(gameHwnd, "m", 80, 120)
        if (!SafeSleep(1300, 1500))
            break

        ; 2. 点击屏幕正中央 (±5px 偏移)
        if (!isRunning)
            break
        ToolTip("Action: Clicking Screen Center")
        offsetX := Random(-5, 5)
        offsetY := Random(-5, 5)
        targetX := Round(w / 2) + offsetX
        targetY := Round(h / 2) + offsetY
        BgClick(gameHwnd, targetX, targetY)
        if (!SafeSleep(500, 700))
            break

        ; 3. 点击右下角区域 client(1070-1478, 833-864) 转换自1600x900
        if (!isRunning)
            break
        ToolTip("Action: Clicking Client (66.8% to 92.3%, 92.5% to 96%)")
        randPctX2 := Random(0.66875, 0.92375)
        randPctY2 := Random(0.9255, 0.9600)
        X2 := Round(randPctX2 * w)
        Y2 := Round(randPctY2 * h)
        BgClick(gameHwnd, X2, Y2)

        ; 4. 等待 10s
        ToolTip("Action: Waiting for 10s...")
        if (!SafeSleep(10000, 10000))
            break

        if (!isRunning)
            break

        ; 5. 依次遍历数字 1, 3, 4, 5, 6
        Loop Parse, "13456" {
            if (!isRunning)
                break

            ToolTip("Action: Pressing Key " A_LoopField)
            BgKeyTap(gameHwnd, A_LoopField, 80, 120)

            if (!SafeSleep(1000, 1200))
                break

            ToolTip("Action: Left Calling (Key " A_LoopField ")")
            BgClickDown(gameHwnd)
            RandomSleep(80, 120)
            BgClickUp(gameHwnd)

            if (!SafeSleep(1000, 1200))
                break
        }

        ; (2) 数字 2 键
        ToolTip("Action: Pressing 2 Key")
        BgKeyTap(gameHwnd, "2", 180, 240)
        if (!SafeSleep(3800, 4200))
            break

        ; 6. 鞠躬循环 8-12 次
        loopCount := Random(8, 12)
        ToolTip("Action: Loop Action " loopCount " times")
        Loop loopCount {
            if (!isRunning)
                break

            ; --- 反检测：1. 模拟玩家发呆/看手机 (约 3% 概率延迟 10-20秒) ---
            afkChance := Random(1, 100)
            if (afkChance <= 3) {
                ToolTip("Action: Simulating Human Distraction...")
                if (!SafeSleep(10000, 20000))
                    break
            }

            ; --- 反检测：2. 独立的随机鼠标游离抖动 (约 8% 概率) ---
            jitterChance := Random(1, 100)
            if (jitterChance <= 8) {
                BgMouseJitter(gameHwnd)
            }

            ; (1) Tab键
            ToolTip("Action: Pressing Tab Key")
            BgKeyTap(gameHwnd, "Tab", 80, 120)
            if (!SafeSleep(900, 1100))
                break

            ; (2) 数字2键
            ToolTip("Action: Pressing 2 Key")
            BgKeyTap(gameHwnd, "2", 180, 240)
            if (!SafeSleep(3800, 4200))
                break

            ; (3) Esc键
            ToolTip("Action: Pressing Esc Key")
            BgKeyTap(gameHwnd, "Esc", 80, 120)
            if (!SafeSleep(450, 550))
                break

            ; (4) R键 (反检测：概率性多按1~2次)
            ToolTip("Action: Pressing R Key")
            randR := Random(1, 100)
            rPresses := (randR <= 15) ? 2 : ((randR <= 20) ? 3 : 1) ; 15%按2次，5%按3次
            Loop rPresses {
                BgKeyTap(gameHwnd, "r", 80, 120)
                if (A_Index < rPresses)
                    RandomSleep(150, 300)
            }
            if (!SafeSleep(9500, 10500))
                break

            ; (5) X键 (反检测：概率性多按1~2次)
            ToolTip("Action: Pressing X Key")
            randX := Random(1, 100)
            xPresses := (randX <= 15) ? 2 : ((randX <= 20) ? 3 : 1) ; 15%按2次，5%按3次
            Loop xPresses {
                BgKeyTap(gameHwnd, "x", 80, 120)
                if (A_Index < xPresses)
                    RandomSleep(150, 300)
            }
            if (!SafeSleep(40, 60))
                break
        }

        ; --- 反检测：大循环结束后的随机喘息时间 (约 5% 概率停顿 5-10秒) ---
        restChance := Random(1, 100)
        if (restChance <= 5) {
            ToolTip("Action: Taking a short break...")
            if (!SafeSleep(5000, 10000))
                break
        } else {
            ; 正常情况下的短停顿
            if (!SafeSleep(800, 1200))
                break
        }
    }
}

; ============================
; 查找游戏窗口
; ============================
FindGameWindow() {
    hwnds := WinGetList("ahk_exe " TARGET_EXE)
    if (hwnds.Length = 0)
        return 0
    ; 如果有多个窗口，取第一个；可按需扩展为弹列表选择
    return hwnds[1]
}

; ============================
; 获取窗口客户区大小
; ============================
GetClientSize(hwnd, &w, &h) {
    rect := Buffer(16, 0)
    DllCall("GetClientRect", "ptr", hwnd, "ptr", rect)
    w := NumGet(rect, 8, "int")
    h := NumGet(rect, 12, "int")
}

; ============================
; 后台按键 - 完整一次按键 (down + 保持 + up)
; ============================
BgKeyTap(hwnd, keyName, holdMin := 80, holdMax := 120) {
    BgKeyDown(hwnd, keyName)
    Sleep(Random(holdMin, holdMax))
    BgKeyUp(hwnd, keyName)
}

BgKeyDown(hwnd, keyName) {
    vk := GetKeyVK(keyName)
    sc := GetKeySC(keyName)
    if (!vk || !sc)
        return false
    lParam := 1 | (sc << 16)
    PostMessage(0x0100, vk, lParam, , "ahk_id " hwnd) ; WM_KEYDOWN
    return true
}

BgKeyUp(hwnd, keyName) {
    vk := GetKeyVK(keyName)
    sc := GetKeySC(keyName)
    if (!vk || !sc)
        return false
    lParam := 1 | (sc << 16) | 0xC0000000
    PostMessage(0x0101, vk, lParam, , "ahk_id " hwnd) ; WM_KEYUP
    return true
}

; ============================
; 后台鼠标点击 (ControlClick + NA 选项不激活窗口)
; ============================
BgClick(hwnd, x, y) {
    global lastClickX, lastClickY
    lastClickX := x
    lastClickY := y
    ; "NA" 表示不激活目标窗口；坐标默认就是相对窗口 client 区域
    ControlClick("X" x " Y" y, "ahk_id " hwnd, , "Left", 1, "NA")
}

; 在最后一次点击的位置按下/抬起左键 (对应原脚本的 Click,down / Click,up)
BgClickDown(hwnd) {
    global lastClickX, lastClickY
    ControlClick("X" lastClickX " Y" lastClickY, "ahk_id " hwnd, , "Left", 1, "NA D")
}

BgClickUp(hwnd) {
    global lastClickX, lastClickY
    ControlClick("X" lastClickX " Y" lastClickY, "ahk_id " hwnd, , "Left", 1, "NA U")
}

; ============================
; 后台鼠标抖动 (PostMessage WM_MOUSEMOVE)
; 只是给窗口发"鼠标移动到 (x,y)"的消息，不会移动物理光标
; ============================
BgMouseJitter(hwnd) {
    global lastClickX, lastClickY
    if (lastClickX = 0 && lastClickY = 0)
        return
    jX := Random(-20, 20)
    jY := Random(-20, 20)
    newX := lastClickX + jX
    newY := lastClickY + jY
    ; lParam: 低16位 = X, 高16位 = Y (Windows 鼠标消息 lParam 编码)
    lParam := ((newY & 0xFFFF) << 16) | (newX & 0xFFFF)
    PostMessage(0x0200, 0, lParam, , "ahk_id " hwnd) ; WM_MOUSEMOVE
}

; ============================
; 增强：安全无延迟等待函数 (带随机延迟)
; ============================
; 将长时间的Sleep切碎为一个个50ms，每次循环都检查是否被中断。
SafeSleep(minTime, maxTime) {
    global isRunning
    targetTime := Random(minTime, maxTime)
    ticks := targetTime // 50
    Loop ticks {
        if (!isRunning)
            return false ; 被中断了，返回false
        Sleep(50)
    }
    return true ; 正常走完，返回true
}

; ============================
; 随机延迟函数（用于按键按压时长）
; ============================
RandomSleep(min, max) {
    Sleep(Random(min, max))
}

; ============================
; 清理屏幕提示的函数
; ============================
RemoveToolTip(*) {
    ToolTip()
}
