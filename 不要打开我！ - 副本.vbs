Dim Msg1,Msg2,Msg3 '定义消息文本
Dim Response '定义返回对象
Dim Msg()
redim Msg(5)
Msg(0) = "歪比巴卜！"
Msg(1) = "球球惹！"
Msg(2) = "真的不吗？"
Msg(3) = "..."
Msg(4) = "$-+℃??dD?????????〆；"
Style = vbYesNo + vbInformation
Title = "歪比巴卜"
for i=0 to 4
Response = MsgBox(Msg(i), Style, Title)
If Response = vbYes Then ' 用户选择[是].
MyString = "Yes" ' Preform some action.
NextResponse = MsgBox("我就知道你会同意的，哈哈！",vbOKOnly ,Title)
exit for
Else ' 用户选择[否],
MyStrig = "No" ' Perform some action.
If i =4 Then
i=-1
End If
End If
Next