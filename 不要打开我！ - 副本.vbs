Dim Msg1,Msg2,Msg3 '������Ϣ�ı�
Dim Response '���巵�ض���
Dim Msg()
redim Msg(5)
Msg(0) = "��ȰͲ���"
Msg(1) = "�����ǣ�"
Msg(2) = "��Ĳ���"
Msg(3) = "..."
Msg(4) = "$-+��??dD?????????�e��"
Style = vbYesNo + vbInformation
Title = "��ȰͲ�"
for i=0 to 4
Response = MsgBox(Msg(i), Style, Title)
If Response = vbYes Then ' �û�ѡ��[��].
MyString = "Yes" ' Preform some action.
NextResponse = MsgBox("�Ҿ�֪�����ͬ��ģ�������",vbOKOnly ,Title)
exit for
Else ' �û�ѡ��[��],
MyStrig = "No" ' Perform some action.
If i =4 Then
i=-1
End If
End If
Next