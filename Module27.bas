Attribute VB_Name = "Module27"
Sub Test()


'IE���J���悤�̂��(open �Ƃ����ƕϐ��Ɏg���Ȃ�)

Dim A As InternetExplorer
Set A = CreateObject("InternetExplorer.Application")

A.Visible = True



'�Z�������URL����

Dim IE As String

IE = Range("B1")


'�m�F�pBox
MsgBox IE



' [�ϐ�IE]�ɓ����Ă���URL ���@�@[�ϐ�A]�̃I�u�W�F�N�g�ŊJ��
A.navigate IE



'IE���J���܂ł̑ҋ@�\��

Do While A.Busy = True Or A.readyState < READYSTATE_COMPLETE

    DoEvents
    
Loop


'
Dim Doc As HTMLDocument
Set Doc = A.document

'���̓t�H�[���p
Dim User As String, Pass As String
User = Range("B2")
Pass = Range("B3")


Doc.getElementById("id").Value = User 'id="user_login"�Ƀ��[�U�[�������
Doc.getElementById("password").Value = Pass  'id="user_pass"�Ƀp�X���[�h�����
Doc.getElementById("form1").submit '�t�H�[���̓��e�𑗐M


End Sub

