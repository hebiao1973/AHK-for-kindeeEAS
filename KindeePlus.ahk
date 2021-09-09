;----------------------------------------------------------------
; 几个关于奥园金蝶账套的快捷脚本：
; 1、Ctrl + Alt + L  自动登录金蝶;
; 2、自定义凭证导航快捷键，即在逐笔查看凭证时，可用自定义的快捷键导航
; 3、当按下金蝶本身的审核凭证快捷键(Ctrl+q)后，除了完成审核该凭证外，另
;    在所弹出的对话框中自动键击Enter(即点击"确定"按钮)，并到下一张凭证;
; 4、月底审核完凭证后，可以让它自动过账、生成结转损益凭证(自动切换公司);
;----------------------------------------------------------------

#SingleInstance force      ;让脚本只能单一实列运行
#NoEnv                     ;不检查空变量是否为环境变量

;以下脚本自动登录金蝶
^!l::    ; 定义快捷键("l"是login的首字母)
run_KindeeApp:  ; 自动登录金蝶 
If  WinExist("金蝶EAS-Aoyuan")  ;如果已存在，
{
    WinActivate                 ;激活
    WinMaximize                 ;最大化
    Return
}
Else
{
    If WinExist("金蝶EAS系统登录")
    {
        WinActivate
        Click, 400, 321   ;点击登录界面的密码输入编辑框，使得它获得输入焦点
        SendRaw, Hekexin2000@  ;输入密码
        Sleep, 100
        Send, {Enter}       ;发送回车键，相当于点击“确定”按钮
        Loop
        {
            If A_Index > 180
            {
                MsgBox, 0, 彪哥提醒您, 等待超时！
                Return
            }
            IfWinActive, 账号重复登录
            {
                Send, {Tab}         
                Sleep, 100
                Send {Enter}    
            }
            IfWinActive, 金蝶EAS-Aoyuan
                Break
            Sleep, 100
        }
        Loop 
        {
            If A_Index > 60
            {
                MsgBox, 0, 彪哥提醒您, 等待超时！
                Return
            }
            ImageSearch, x, y, 60, 180, 250, 350, *30 财务会计.png
            If !ErrorLevel  
                Break 
            Else
                Sleep, 500 
        }
        MouseMove, x+40, y+10, 0 
        Click
        Return      
    }
    Else
    {
        Run, D:\EAS\eas\client\bin\client.bat
        WinWait, 金蝶EAS系统登录, , 60
        If ErrorLevel
        {
            MsgBox, 0, 彪哥提醒您, 等待超时！
            Return
        }
        Else
            Goto, run_KindeeApp
    }
}
Return

;以下脚本自定义了凭证的前进和后退快捷键：
#IfWinActive 金蝶EAS-Aoyuan           ;让以下的快捷键只在金蝶EAS中有效

^Numpad0::   ;无论小键盘是否处于数字输入状态，按下Ctrl+0(小键盘)都会触发
^NumpadIns::
Click, 470, 90
Return
^NumpadDot::  ;无论小键盘是否处于数字输入状态，按下Ctrl+.(小键盘)都会触发
^NumpadDel::
Click, 436, 90
Return

;以下脚本让审核凭证快捷键Ctrl+u除了完成本职工作外干点更多的事情：
^q::
Send, ^u
WinWaitActive, 信息提示, , 2
Send, {Enter}    ;在所弹出的对话框中自动键击Enter(即点击"确定"按钮)
WinWaitClose, 信息提示
Click, 470, 90   ;点击工具栏图标到下一张凭证
Return

;以下代码自动切换公司
^!c::                        ;Ctrl+Alt+c  c代表change
WinMaximize, 金蝶EAS-Aoyuan
Loop, READ, 公司编码.txt
{
    Click, right, 1514, 1026   ;右键点击金蝶EAS右下的公司框
    Sleep, 100
    Click, 1472, 989           ;选择"更多……"
    Send, {Enter}              ;在弹出的对话框键击回车（相当于点击"确定"按钮）
    WinWaitActive, 组织切换, , 30
    If ErrorLevel = 1
    {
        MsgBox, 0, 彪哥提醒您, 等待超时
        Return
    }
    Sleep, 300
    Clipboard := A_LoopReadLine       ;用剪贴板，Send会受输入法拦截
    Send, ^v                ;粘贴
    Sleep, 100
    Click, 499, 88          ;去掉模糊查询的勾选
    Sleep, 100
    Click, 418, 90          ;点击快速查询按钮
    Sleep, 200             
    Click, 482, 182, 2      ;双击选择要切换进入的账套
    ;To-do： 此处调用你要对账套进项的操作子程序或者函数
}
Return

;-------------------------------------------------
;以下代码使金蝶EAS处于总账界面，方便下一步操作
^!Numpad1::
^!NumpadEnd::
返回总账:
WinMaximize, 金蝶EAS-Aoyuan
ImageSearch, x, y, 80, 200, 200, 260, *50 总帐.png
If ErrorLevel
{
    Sleep, 200
    Click, 90, 125          ;点击EAS左上“应用中心”页面
    Sleep, 300
    Click, 160, 185
    Sleep, 300
    Click, 165, 300     ;点击左边“财务会计” 
    Sleep, 100
    Click, 165, 300
    Sleep, 300
}
Return

;-------------------------------------------------
;以下代码进入凭证序时簿
^!Numpad2::
^!NumpadDown::
凭证序时簿:
ImageSearch, x, y, 250, 80, 800, 150, *50 凭证序时薄.png
If ErrorLevel
{
    ImageSearch, x, y, 250, 80, 800, 150, *50 凭证序时薄2.png
    If !ErrorLevel            ;如果找到，表示在EAS选项卡上已有凭证序时簿，未激活，激活即可
    {
        MouseMove, x+40, y+10, 0
        Click
        Sleep, 200
        Send, {F5}            ;刷新
        Return
    }
}
Else
{
    Send, {F5}            ;找到凭证序时簿，且处于被打开状态，刷新后结束
    Return    
}
Gosub, 返回总账
Click, 434, 294       ;点击选择“凭证处理”
Sleep, 500
Click, 1090, 254, 2   ;双击进入"凭证查询"
Loop
{
    If A_Index > 60
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
    Sleep, 500
    IfWinActive, 凭证-条件查询
    {
        Click, 688, 620
        Return
    }    
    ImageSearch, x, y, 250, 80, 800, 150, *50 凭证序时薄.png
}
Until !ErrorLevel
Return
;-------------------------------------------------
;以下代码审核凭证、记账所有凭证
^!Numpad3::
^!NumpadPgdn::
Gosub, 凭证序时簿 
Click, 120, 210    ;点击选择第一张凭证
Sleep, 300
Send, ^a           ;Ctrl + a (选择所有凭证)
Sleep, 500
Send, ^u           ;Ctrl + u (审核凭证)
WinWaitClose, 进度提示

Sleep, 300
Click, 120, 210    ;点击选择第一张凭证
Sleep, 500
Send, ^a           ;Ctrl + a (选择所有凭证), 
Sleep, 500    
Send, !p           ;Alt + p (记账)
WinWaitClose, 进度提示
Return

;-------------------------------------------------
;以下代码结转损益，并提交结转损益凭证
^!Numpad4::
^!NumpadLeft::
;SetMouseDelay, 30
ImageSearch, x, y, 200, 80, 1000, 150, *50 结转损益.png
If ErrorLevel
    ImageSearch, x, y, 200, 80, 1000, 150, *50 结转损益2.png
If ErrorLevel
{
    Gosub, 返回总账
    Sleep, 100
    Click, 434, 525         ;点击选择“期末处理”
    Sleep, 100
    Click, 1093, 293, 2     ;双击选择“结转损益”
    WinWaitActive, 结转损益 - 条件查询, ,30
    If ErrorLevel
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
    Else
    {
        Sleep, 100
        Click, 653, 409         ;点击"确认"按钮，选择默认方案
    }
}
Else
{
    MouseMove, x+40, y+10, 0    ;移动鼠标到所找到的图标上
    Click
}
Sleep, 500
Click, 116, 228        ;勾选
Sleep, 100            
Click, 529, 90         ;点击“生成凭证”
WinWaitActive, 凭证生成信息, , 30
If ErrorLevel
{
    MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
    Return
}
Else
{
    Sleep, 500
    ImageSearch, , , 80, 90, 120, 130, *50 成功.png
    If ErrorLevel
    {
        MsgBox, 4112, 彪哥提醒您, 生成结转凭证失败
        Return
    }
    Else 
        Click, 826, 378      ;点击"确认"按钮,关闭当前窗口
    
}
Gosub, 返回总账
Gosub, 凭证序时簿
Sleep, 300
Click, 120, 210
Sleep, 100
Send, ^s
Return

;以下代码自动生成5张报表，月报3份、上市报表2份
^!Numpad5::
^!NumpadClear::
ImageSearch, x, y, 200, 80, 1000, 150,*50 报表制作.png
If ErrorLevel
{
    ImageSearch, x, y, 200, 80, 1000, 150, *50 报表制作2.png
    If !ErrorLevel            ;如果找到，表示在EAS选项卡上已有凭证序时簿，未激活，激活即可
    {
        MouseMove, x+30, y+12, 0
        Click
        Sleep, 100
    }
    Else                      ;仍然找不到，表示在EAS没有打开报表制作，就打开
    {
        Gosub, 返回总账
        Click, 162, 504             ;点击左侧栏中的 "报表管理"
        Click, 435, 215             ;点击中间栏中的 "报表编制"
        Click, 1088, 255, 2         ;点击右侧栏中的 "报表制作"
        WinWaitActive, 报表 - 条件查询 ahk_class SunAwtDialog
        Click, 672, 482
        WinWaitClose, 报表 - 条件查询 ahk_class SunAwtDialog  
    }
}

Loop, 5
{
    If (A_Index < 4)
    {
        y1 := 165
        y2 := 120 + 20*(A_Index-1)
    }
    Else
    {
        y1 := 185
        y2 := 120 + 20*(A_Index-4)
    }
    Sleep, 300
    Send, ^n
    WinWaitActive, 报表 -新建 ahk_class SunAwtDialog, , 10
    If ErrorLevel
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
    Click, 344, 306
    WinWaitActive, 选择模板 ahk_class SunAwtDialog
    If ErrorLevel
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
    Click, 35, 146
    Sleep, 50
    Click, 120, %y1%
    Sleep, 50
    Click, 500, %y2%
    Sleep, 50
    Click, 592, 542
    WinWaitClose, 选择模板 ahk_class SunAwtDialog
    Click, 520, 630
    WinWaitClose, 报表 -新建 ahk_class SunAwtDialog
    WinWaitActive, 新建 ahk_class SunAwtFrame, , 10
    If ErrorLevel
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
    Send, {F9}
    Sleep, 1000
    CoordMode, Pixel, Screen
    Loop
    {
        ImageSearch, , , 740, 435, 835, 475, *50 计算工作薄.png
        If !ErrorLevel
            Sleep, 1000
    }
    Until ErrorLevel
    CoordMode, Pixel, Window
    Sleep, 500
    Send, ^s
    Sleep, 1000
    Send, ^w
    Sleep, 500
    If WinActive("信息提示")
    {
        Send, {Enter}
        Sleep, 100
    }
    WinWaitActive, 金蝶EAS-Aoyuan, , 30
    If ErrorLevel
    {
        MsgBox, 0, 彪哥提醒您, 等待超时，脚本已停止运行
        Return
    }
}
Return
#IfWinActive

^!b::
Send ^!l
WinWaitActive, 金蝶EAS-Aoyuan, , 60
If ErrorLevel
{
    MsgBox, 0, 彪哥提醒您, 等待超时！
    Return
}
Sleep, 6000
SetMouseDelay, 100
Click, 88, 125
Click, 165, 188
Click, 165, 228
Click, 165, 333
Click, 434, 334
Click, 1076, 216, 2
WinWaitActive, 合并范围选择,, 30
If ErrorLevel = 1
{
    MsgBox 超时
    Return
}
Sleep, 2000
Click, 342, 195
WinWaitActive, 合并范围,, 30
If ErrorLevel = 1
{
    MsgBox 超时
    Return
}
Sleep, 2000
Click, 493, 183, 2
Click, 210, 306
Return

#IfWinActive, 行政组织 ahk_class SunAwtDialog      ;快速选择组织单位
^k::
    Clipboard := "HXAYGSCB"
    Click, 186, 88
    Send, ^v
    Click, 419, 88
    Return
^z::
    Clipboard := "广州资金管理部"
    Click, 186, 88
    Send, ^v
    Click, 419, 88
    Click, 435, 184, 2
    ; Click, 728, 800
    Return
#IfWinActive


#IfWinActive, 金蝶EAS-Aoyuan ahk_class SunAwtFrame ahk_exe javaw.exe       
;快速打开“标准凭证引入”界面，快速输入要引入的目标文件
^k::   
    Clipboard := "GL0304"   
    MouseClickDrag, Left, 1700, 55, 1400, 55 
    Sleep, 50
    Send, ^v
    Sleep, 50
    Click, 1510, 90, 2   ;双击
    WinWaitActive, 标准凭证引入_第1步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 443, 407     ;点击“下一步”
    WinWaitActive, 标准凭证引入_第2步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 565, 168     ;点击“选择”
    WinWaitActive, 选择凭证库 ahk_class SunAwtDialog
    Sleep, 50
    Clipboard := "C:\Users\hebiao\Desktop\社保公积金导入\凭证导入.xlsx"
    Send, ^v        ;粘贴
    Sleep, 50
    Click, 660, 393
    WinWaitActive, 标准凭证引入_第2步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 210, 204     ;勾选“引入现金流量”
    Sleep, 50
    Click, 443, 407     ;点击“下一步”
    WinWaitActive, 标准凭证引入_第3步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 443, 407     ;点击“下一步”
    WinWaitActive, 标准凭证引入_第4步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 443, 407     ;点击“下一步”
    WinWaitActive, 标准凭证引入_第5步，共6步 ahk_class SunAwtDialog
    Sleep, 50
    Click, 270, 141     ;点击“检查凭证错误”

    Return
#IfWinActive

~Escape::
Pause
Return