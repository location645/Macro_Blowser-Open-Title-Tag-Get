Attribute VB_Name = "Module8"
Sub IE()

Dim I As InternetExplorer
Set I = CreateObject("InternetExplorer.Application")

I.Visible = True
I.navigate "https://ja.wikipedia.org/wiki/%E5%B3%B6%E7%94%B0%E9%99%BD%E5%AD%90"


Do While I.Busy = True Or I.readyState < READYSTATE_COMPLETE '“Ç‚Ýž‚Ý‘Ò‚¿
 
    DoEvents
 
Loop
 



Dim M As HTMLDocument

Set M = I.document






Debug.Print M.Table





End Sub
