;-----------------------------------------------------------
; 自动登陆OA
; 作者：何彪
; 时间：2021.6.9
; 
;-----------------------------------------------------------

; code UTF-8

^!a::
URL := "http://ayoa.aoyuan.net:8080/login.jsp"
URL_mysoft := "https://um.aoyuan.net/auth?redirecturl=https://mysofterp.aoyuan.net:8060/Product/Interface/SSO/Login.aspx?UserCode=xIIr4OhSJJ4%3d&targeturl=https://mysofterp.aoyuan.net:8060&appid=CBC19F6E257C46C9A2325ACE8C1D265D"
Login_Name := "hebiao"
My_Password := "Qwer1234!"


If !WinExist("ahk_exe iexplore.exe")
{
    run, "C:\Program Files (x86)\Internet Explorer\iexplore.exe" %URL%
    Sleep, 3000
}

If !IsObject(ie)
{
    ie := ComObjCreate("InternetExplorer.Application")
    ie.Navigate(URL)
    ie.Visible := true
    While ie.readyState != 4 || ie.document.readyState != "complete"|| ie.busy
        Sleep, 1000
    ie.document.forms[0].j_username.value := Login_Name
    ie.document.forms[0].j_password.value := My_Password
    ie.document.getElementsByName("btn_submit")[0].Click()
    Winwait, ��ҳ����԰���ţ� - Internet Explorer
    ie.Navigate(URL_mysoft)

    Winwait, ��ҵ���ʻؿ� - Internet Explorer,, 5
    WinClose     ;�رյ�������1,�ȴ�2��

    Winwait, ��Ŀ�ؿ� - Internet Explorer,,5
    WinClose       ;�رյ�������2���ȴ�2��

    While ie.readyState != 4 || ie.document.readyState != "complete" || ie.busy
        Sleep, 1000
}


Escape::
Pause
Return
