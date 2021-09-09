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
    Click, 728, 800
    Return
#IfWinActive