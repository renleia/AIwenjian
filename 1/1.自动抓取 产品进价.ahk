#SingleInstance Force
SetTitleMatchMode, 2
SendMode Input
CoordMode, Mouse, Screen

; 启动 GDI+
#Include Gdip_All.ahk
if !pToken := Gdip_Startup()
{
    MsgBox, GDI+ 初始化失败！
    ExitApp
}

; Excel 连接
excel := ComObjActive("Excel.Application")
sheet := excel.ActiveSheet
row := 1

Loop
{
    barcode := sheet.Cells(row, 1).Value
    if (barcode = "")
        break

    ; 去除 .000000
    barcode := RegExReplace(barcode, "\.0+$")

    IfWinExist, 欧华零售业管理系统 - 网络版V4.4.0.3
    {
        WinActivate
        WinWaitActive
        Sleep, 300

        ; 步骤1：点击输入框并输入条码
        MouseClick, left, 464, 86
        Sleep, 150
        SendInput, ^a
        Sleep, 100
        SendInput, {Backspace}
        Sleep, 100
        SendInput, %barcode%
        Sleep, 150

        ; 步骤2：点击搜索按钮
        MouseClick, left, 899, 87
        Sleep, 500

        ; 步骤3：截图并保存为唯一文件
        timestamp := A_Now
        ocrImage := A_ScriptDir . "\ocr_" . row . "_" . barcode . "_" . timestamp . ".png"
        CaptureRegion(866, 133, 920, 150, ocrImage)
        Sleep, 300

        ; 步骤4：OCR 识别
        ocrResultPath := A_ScriptDir . "\ocr_result_" . row . ".txt"
        RunWait, %ComSpec% /c tesseract "%ocrImage%" "%ocrResultPath%" -l chi_sim --psm 6, , Hide
        Sleep, 300

        ; 读取识别结果
        resultText := ReadFile(ocrResultPath . ".txt")
        resultText := Trim(resultText)

        ; 写入 Excel B列
        sheet.Cells(row, 4).Value := resultText

        row++
        Sleep, 500
    }
    else
    {
        MsgBox, 找不到目标窗口，请确认软件已打开
        ExitApp
    }
}
MsgBox, 所有条码处理完成
Gdip_Shutdown(pToken)
ExitApp

; -------- 函数：截图保存 --------
CaptureRegion(x1, y1, x2, y2, outputFile)
{
    width := x2 - x1
    height := y2 - y1

    pBitmap := Gdip_BitmapFromScreen(x1 "|" y1 "|" width "|" height)
    if (!pBitmap)
    {
        MsgBox, 截图失败！
        return
    }

    Gdip_SaveBitmapToFile(pBitmap, outputFile)
    Gdip_DisposeImage(pBitmap)
}

; -------- 函数：读取文件内容 --------
ReadFile(filePath)
{
    FileRead, content, %filePath%
    return content
}