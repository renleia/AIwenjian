#NoEnv
SendMode Input


; 配置参数
productNumber := "12345678"
excelPath := "C:\Sales_Data.xlsx"

; 打开目标软件
Run, D:\RetailSystemV4\client\OhRetailSystem_V4.exe, 软件窗口标题
Sleep, 1000

; 输入产品号（通过坐标点击输入框）
Click, 200, 150  ; 输入框坐标
Sleep, 2500
Send, %productNumber%

; 点击搜索按钮（绝对坐标）
Click, 187, 271
Sleep, 1000

; 点击仓库（绝对坐标）
Click, 586, 333
Sleep, 12000

; 点击仓库（绝对坐标）
Click, 419, 197
Sleep, 2000
