Dim Msg1,Msg2,Msg3 '定义消息文本
Dim Response '定义返回对象
Dim Msg()
redim Msg(5)
Msg(0) = "但是你还是打开了！发送电子邮件到xiao-miao@hotmail.com并附上这个文件（发送邮件 那是你自己的选择。当然 你也可以选择不发或者发送但不带附件）"
Msg(1) = "恭喜你 ):"
Msg(2) = "你感受过“改变”吗，对于你自己"
Msg(3) = ""
Msg(4) = "$-+℃??dD?????????〆；"
Style = vbYesNo + vbInformation
Title = ""
for i=0 to 4
Response = MsgBox(Msg(i), Style, Title)
If Response = vbYes Then ' 用户选择[是].
MyString = "Yes" ' Preform some action.
NextResponse = MsgBox("别忘了电子邮件！",vbOKOnly ,Title)
exit for
Else ' 用户选择[否],
MyStrig = "No" ' Perform some action.
If i =4 Then
i=-1
End If
End If
Next