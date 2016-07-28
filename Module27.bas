Attribute VB_Name = "Module27"
Sub Test()


'IEを開く

Dim A As InternetExplorer
Set A = CreateObject("InternetExplorer.Application")

A.Visible = True



'セルからのURL入力

Dim IE As String

IE = Range("B1")


'確認用Box
MsgBox IE



' [変数IE]に入っているURL を　　[変数A]のオブジェクトで開く
A.navigate IE



'IEが開くまでの待機構文

Do While A.Busy = True Or A.readyState < READYSTATE_COMPLETE

    DoEvents
    
Loop


'
Dim Doc As HTMLDocument
Set Doc = A.document

'入力フォーム用
Dim User As String, Pass As String
User = Range("B2")
Pass = Range("B3")


Doc.getElementById("id").Value = User 'id="user_login"にユーザー名を入力
Doc.getElementById("password").Value = Pass  'id="user_pass"にパスワードを入力
Doc.getElementById("form1").submit 'フォームの内容を送信


End Sub

