<% @ Language = "VBScript" %>
<% ' Option Explicit %>
<%
'####################################
'#									#
'#		 阿江ASP探针 V1.70			#
'#									#
'#  阿江守候 http://www.ajiang.net  #
'#	 电子邮件 info@ajiang.net		#
'#									#
'#    转载本程序时请保留这些信息    #
'#								    #
'####################################

'不使用输出缓冲区，直接将运行结果显示在客户端
Response.Buffer = False

'声明待检测数组
Dim ObjTotest(26
Dim 4)

ObjTotest(0,0)  = "MSWC.AdRotator"
ObjTotest(1,0)  = "MSWC.BrowserType"
ObjTotest(2,0)  = "MSWC.NextLink"
ObjTotest(3,0)  = "MSWC.Tools"
ObjTotest(4,0)  = "MSWC.Status"
ObjTotest(5,0)  = "MSWC.Counters"
ObjTotest(6,0)  = "IISSample.ContentRotator"
ObjTotest(7,0)  = "IISSample.PageCounter"
ObjTotest(8,0)  = "MSWC.PermissionChecker"
ObjTotest(9,0)  = "Scripting.FileSystemObject"
ObjTotest(9,1)  = "(FSO 文本文件读写)"
ObjTotest(10,0) = "adodb.connection"
ObjTotest(10,1) = "(ADO 数据对象)"

ObjTotest(11,0) = "SoftArtisans.FileUp"
ObjTotest(11,1) = "(SA-FileUp 文件上传)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
ObjTotest(12,1) = "(SoftArtisans 文件管理)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
ObjTotest(13,1) = "(刘云峰的文件上传组件)"
ObjTotest(14,0) = "Persits.Upload.1"
ObjTotest(14,1) = "(ASPUpload 文件上传)"
ObjTotest(15,0) = "w3.upload"
ObjTotest(15,1) = "(Dimac 文件上传)"

ObjTotest(16,0) = "JMail.SmtpMail"
ObjTotest(16,1) = "(Dimac JMail 邮件收发) <a href='http://www.ajiang.net'>中文手册下载</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
ObjTotest(17,1) = "(虚拟 SMTP 发信)"
ObjTotest(18,0) = "Persits.MailSender"
ObjTotest(18,1) = "(ASPemail 发信)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
ObjTotest(19,1) = "(ASPmail 发信)"
ObjTotest(20,0) = "DkQmail.Qmail"
ObjTotest(20,1) = "(dkQmail 发信)"
ObjTotest(21,0) = "Geocel.Mailer"
ObjTotest(21,1) = "(Geocel 发信)"
ObjTotest(22,0) = "IISmail.Iismail.1"
ObjTotest(22,1) = "(IISmail 发信)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
ObjTotest(23,1) = "(SmtpMail 发信)"

ObjTotest(24,0) = "SoftArtisans.ImageGen"
ObjTotest(24,1) = "(SA 的图像读写组件)"
ObjTotest(25,0) = "W3Image.Image"
ObjTotest(25,1) = "(Dimac 的图像读写组件)"

Public IsObj,VerObj,TestObj

'检查预查组件支持情况及版本

Dim i

For i = 0 To 25
    On Error Resume Next
    IsObj       = False
    VerObj      = ""
    'dim TestObj
    TestObj     = ""
    Set TestObj = Server.CreateObject(ObjTotest(i,0))
    If - 2147221005 <> Err Then		'感谢网友iAmFisher的宝贵建议
    IsObj       = True
    VerObj      = TestObj.version
    If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
End If

ObjTotest(i,2) = IsObj
ObjTotest(i,3) = VerObj
Next

'检查组件是否被支持及组件版本的子程序
Sub ObjTest(strObj)
On Error Resume Next
IsObj       = False
VerObj      = ""
TestObj     = ""
Set TestObj = Server.CreateObject (strObj)
If - 2147221005 <> Err Then		'感谢网友iAmFisher的宝贵建议
IsObj       = True
VerObj      = TestObj.version
If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
End If

End Sub %>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>ASP探针V1.70－阿江http://www.ajiang.net</TITLE>
<style>
<!--
BODY
{
	FONT-FAMILY: 宋体;
	FONT-SIZE: 9pt
}
TD
{
	FONT-SIZE: 9pt
}
A
{
	COLOR: #000000;
	TEXT-DECORATION: none
}
A:hover
{
	COLOR: #3F8805;
	TEXT-DECORATION: underline
}
.input
{
	BORDER: #111111 1px solid;
	FONT-SIZE: 9pt;
	BACKGROUND-color: #F8FFF0
}
.backs
{
	BACKGROUND-COLOR: #3F8805;
	COLOR: #ffffff;

}
.backq
{
	BACKGROUND-COLOR: #EEFEE0
}
.backc
{
	BACKGROUND-COLOR: #3F8805;
	BORDER: medium none;
	COLOR: #ffffff;
	HEIGHT: 18px;
	font-size: 9pt
}
.fonts
{
	COLOR: #3F8805
}
-->
</STYLE>
</HEAD>
<BODY>
<a href="mailto:info@ajiang.net">阿江</a>改写的ASP探针-<font class=fonts>V1.70</font><br><br>
<font class=fonts>是否支持ASP</font>
<br>出现以下情况即表示您的空间不支持ASP：
<br>1、访问本文件时提示下载。
<br>2、访问本文件时看到类似“&lt;%@ Language="VBScript" %&gt;”的文字。
<br><br>

<font class=fonts>服务器的有关参数</font>
<table border=0 width=450 cellspacing=0 cellpadding=0 bgcolor="#3F8805">
<tr><td>

	<table border=0 width=450 cellspacing=1 cellpadding=0>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器名</td><td>&nbsp;<% = Request.ServerVariables("SERVER_NAME") %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器IP</td><td>&nbsp;<% = Request.ServerVariables("LOCAL_ADDR") %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器端口</td><td>&nbsp;<% = Request.ServerVariables("SERVER_PORT") %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器时间</td><td>&nbsp;<% = Now %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;IIS版本</td><td>&nbsp;<% = Request.ServerVariables("SERVER_SOFTWARE") %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;脚本超时时间</td><td>&nbsp;<% = Server.ScriptTimeout %> 秒</td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;本文件路径</td><td>&nbsp;<% = Request.ServerVariables("PATH_TRANSLATED") %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器CPU数量</td><td>&nbsp;<% = Request.ServerVariables("NUMBER_OF_PROCESSORS") %> 个</td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器解译引擎</td><td>&nbsp;<% = ScriptEngine & "/" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %></td>
	  </tr>
	  <tr bgcolor="#EEFEE0" height=18>
		<td align=left>&nbsp;服务器操作系统</td><td>&nbsp;<% = Request.ServerVariables("OS") %></td>
	  </tr>
	</table>

</td></tr>
</table>
<br>
<font class=fonts>组件支持情况</font>
<%
Dim strClass
strClass = Trim(Request.Form("classname"))

If "" <> strClass Then
Response.Write "<br>您指定的组件的检查结果："
Dim Verobj1
ObjTest(strClass)

If Not IsObj Then
Response.Write "<br><font color=red>很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
Else

If VerObj = "" Or IsNull(VerObj) Then
    Verobj1 = "无法取得该组件版本"
Else
    Verobj1 = "该组件版本是：" & VerObj
End If

Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
End If

Response.Write "<br>"
End If %>


<br>■ IIS自带的ASP组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
	<tr height=18 class=backs align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<% For i = 0 To 10 %>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<% = ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1) %></font></td>
		<td align=left>&nbsp;<%

If Not ObjTotest(i,2) Then
Response.Write "<font color=red><b>×</b></font>"
Else
Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & Left(ObjTotest(i,3),11) & "</a>"
End If %></td>
	</tr>
	<% Next %>
</table>

<br>■ 常见的文件上传和管理组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
	<tr height=18 class=backs align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<% For i = 11 To 15 %>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<% = ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1) %></font></td>
		<td align=left>&nbsp;<%

If Not ObjTotest(i,2) Then
Response.Write "<font color=red><b>×</b></font>"
Else
Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & Left(ObjTotest(i,3),11) & "</a>"
End If %></td>
	</tr>
	<% Next %>
</table>

<br>■ 常见的收发邮件组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
	<tr height=18 class=backs align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<% For i = 16 To 23 %>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<% = ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1) %></font></td>
		<td align=left>&nbsp;<%

If Not ObjTotest(i,2) Then
Response.Write "<font color=red><b>×</b></font>"
Else
Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & Left(ObjTotest(i,3),11) & "</a>"
End If %></td>
	</tr>
	<% Next %>
</table>

<br>■ 图像处理组件
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
	<tr height=18 class=backs align=center><td width=320>组 件 名 称</td><td width=130>支持及版本</td></tr>
	<% For i = 24 To 25 %>
	<tr height="18" class=backq>
		<td align=left>&nbsp;<% = ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1) %></font></td>
		<td align=left>&nbsp;<%

If Not ObjTotest(i,2) Then
Response.Write "<font color=red><b>×</b></font>"
Else
Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & Left(ObjTotest(i,3),11) & "</a>"
End If %></td>
	</tr>
	<% Next %>
</table>

<br>■ 其他组件支持情况检测<br>
在下面的输入框中输入你要检测的组件的ProgId或ClassId。
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
<FORM action=<% = Request.ServerVariables("SCRIPT_NAME") %> method=post id=form1 name=form1>
	<tr height="18" class=backq>
		<td align=center height=30><input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value=" 确 定 " class=backc id=submit1 name=submit1>
<INPUT type=reset value=" 重 填 " class=backc id=reset1 name=reset1> 
</td>
	  </tr>
</FORM>
</table>

<% If ObjTest("Scripting.FileSystemObject") Then

Set fsoobj = Server.CreateObject("Scripting.FileSystemObject") %>

<br><font class=fonts>磁盘相关测试</font>

<br>■ 服务器磁盘信息

<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
  <tr height="18" align=center class="backs">
	<td width="100">盘符和磁盘类型</td>
	<td width="50">就绪</td>
	<td width="80">卷标</td>
	<td width="60">文件系统</td>
	<td width="80">可用空间</td>
	<td width="80">总空间</td>
  </tr>
<%

' 测试磁盘信息的想法来自“COCOON ASP 探针”

Set drvObj = fsoobj.Drives

For Each d in drvObj %>
  <tr height="18" align=center>
	<td align="right"><% = cdrivetype(d.DriveType) & " " & d.DriveLetter %>:</td>
<%
If d.DriveLetter = "A" Then	'为防止影响服务器，不检查软驱
Response.Write "<td></td><td></td><td></td><td></td><td></td>"
Else %>
	<td><% = cIsReady(d.isReady) %></td>
	<td><% = d.VolumeName %></td>
	<td><% = d.FileSystem %></td>
	<td align="right"><% = cSize(d.FreeSpace) %></td>
	<td align="right"><% = cSize(d.TotalSize) %></td>
<%
End If %>
  </tr>
<%
Next %>
</td></tr>
</table>

<br>■ 当前文件夹信息
<%
dPath      = Server.MapPath("./")
Set dDir   = fsoObj.GetFolder(dPath)
Set dDrive = fsoObj.GetDrive(dDir.Drive) %>
文件夹: <% = dPath %>
<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
  <tr height="18" align="center" class="backs">
	<td width="75">已用空间</td>
	<td width="75">可用空间</td>
	<td width="75">文件夹数</td>
	<td width="75">文件数</td>
	<td width="150">创建时间</td>
  </tr>
  <tr height="18" align="center">
	<td><% = cSize(dDir.Size) %></td>
	<td><% = cSize(dDrive.AvailableSpace) %></td>
	<td><% = dDir.SubFolders.Count %></td>
	<td><% = dDir.Files.Count %></td>
	<td><% = dDir.DateCreated %></td>
  </tr>
</td></tr>
</table>

<br>■ 磁盘文件操作速度测试<br>
<%

' 测试文件读写的想法来自“迷城浪子”

Response.Write "正在重复创建、写入和删除文本文件50次..."

Dim thetime3
Dim tempfile
Dim iserr

iserr    = False
t1       = timer
tempfile = Server.MapPath("./") & "\aspchecktest.txt"

For i = 1 To 50
Err.Clear

Set tempfileOBJ = FsoObj.CreateTextFile(tempfile,True)

If Err <> 0 Then
Response.Write "创建文件错误！<br><br>"
iserr = True
Err.Clear
Exit For
End If

tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"

If Err <> 0 Then
Response.Write "写入文件错误！<br><br>"
iserr = True
Err.Clear
Exit For
End If

tempfileOBJ.Close
Set tempfileOBJ = FsoObj.GetFile(tempfile)
tempfileOBJ.Delete

If Err <> 0 Then
Response.Write "删除文件错误！<br><br>"
iserr = True
Err.Clear
Exit For
End If

Set tempfileOBJ = Nothing
Next

t2              = timer

If iserr <> True Then
thetime3        = CStr(Int(( (t2 - t1)*10000 ) + 0.5)/10)
Response.Write "...已完成！<font color=red>" & thetime3 & "毫秒</font>。<br>" %>
<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
  <tr height=18 align=center class="backs">
	<td width=320>供 对 照 的 服 务 器</td>
	<td width=130>完成时间(毫秒)</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.ajiang.net">阿江的个人主机（DDR512M赛扬1.7G,希捷7200转/2M）</a></td><td>&nbsp;140～200</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.ajiang.net">阿江单位的电脑（SD256M赛扬660,希捷5400转）</a></td><td>&nbsp;350～600</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<font color=red>这台服务器: <% = Request.ServerVariables("SERVER_NAME") %></font>&nbsp;</td><td>&nbsp;<font color=red><% = thetime3 %></font></td>
  </tr>
</table>
<%
End If

Set fsoobj = Nothing

End If %>
<br>
<font class=fonts>ASP脚本解释和运算速度测试</font><br>
<%

'感谢网际同学录 http://www.5719.net 推荐使用timer函数
'因为只进行50万次计算，所以去掉了是否检测的选项而直接检测

Response.Write "整数运算测试，正在进行50万次加法运算..."
Dim t1
Dim t2
Dim lsabc
Dim thetime
Dim thetime2
t1      = timer

For i = 1 To 500000
lsabc   = 1 + 1
Next

t2      = timer
thetime = CStr(Int(( (t2 - t1)*10000 ) + 0.5)/10)
Response.Write "...已完成！<font color=red>" & thetime & "毫秒</font>。<br>"

Response.Write "浮点运算测试，正在进行20万次开方运算..."
t1       = timer

For i = 1 To 200000
lsabc    = 2^0.5
Next

t2       = timer
thetime2 = CStr(Int(( (t2 - t1)*10000 ) + 0.5)/10)
Response.Write "...已完成！<font color=red>" & thetime2 & "毫秒</font>。<br>" %>
<table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#3F8805" width="450">
  <tr height=18 align=center class="backs">
	<td width=320>供对照的服务器及完成时间(毫秒)</td>
    <td width=65>整数运算</td><td width=65>浮点运算</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.100u.com?come=aspcheck&keyword=虚拟主机"
	>百优科技 100u 主机, <font color=#888888>2003-11-1</font></a></td><td>&nbsp;181～233</td><td>&nbsp;156～218</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.west263.com/index.asp?ads=ajiang"
	>西部数码 west263 主机, <font color=#888888>2003-11-1</font></a></td><td>&nbsp;171～233</td><td>&nbsp;156～171</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.linkwww.com "
	>联网科技 linkwww 主机,  <font color=#888888>2003-11-1</font></a></td><td>&nbsp;181～203</td><td>&nbsp;171</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.9s5.com/"
	>就是我www.9s5.com全功能(ASP+PHP+JSP)主机,<font color=#888888>2003-11-1</font></a></td><td>&nbsp;171～187</td><td>&nbsp;156～171</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.dnsmy.com/"
	>永讯网络 Dnsmy 主机, <font color=#888888>2003-11-1</font></a></td><td>&nbsp;155～180</td><td>&nbsp;122～172</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<a href="http://www.senye.net/"
	>胜易网络 Senye.net 主机, <font color=#888888>2003-10-28</font></a></td><td>&nbsp;171～187</td><td>&nbsp;156～171</td>
  </tr>
  <tr height=18>
	<td align=left>&nbsp;<font color=red>这台服务器: <% = Request.ServerVariables("SERVER_NAME") %></font>&nbsp;</td><td>&nbsp;<font color=red><% = thetime %></font></td><td>&nbsp;<font color=red><% = thetime2 %></font></td>
  </tr>
</table>
<br>
<table border=0 width=450 cellspacing=0 cellpadding=0>
<tr><td align=center>
<b>[<a href="http://www.ajiang.net/products/aspcheck/serverlist.asp#notice">提醒·说明</a>]
&nbsp;[<a href="http://www.ajiang.net/products/aspcheck/serverlist.asp">更多空间商即时实测数据</a>]
&nbsp;[<a href="http://www.ajiang.net/products/aspcheck/">查看下载最新版</a>]</b>
</td></tr>
</table>
<br>
<table border=0 width=450 cellspacing=0 cellpadding=0>
<tr><td align=center>
欢迎访问 【阿江守候】 <a href="http://www.ajiang.net">http://www.ajiang.net</a>
<br>本程序由阿江(<a href="mailto:info@ajiang.net?subject=阿江探针">info@ajiang.net</a>)编写，转载时请保留这些信息
</td></tr>
</table>
</BODY>
</HTML>

<%

Function cdrivetype(tnum)

Select Case tnum
Case 0: cdrivetype = "未知"
Case 1: cdrivetype = "可移动磁盘"
Case 2: cdrivetype = "本地硬盘"
Case 3: cdrivetype = "网络磁盘"
Case 4: cdrivetype = "CD-ROM"
Case 5: cdrivetype = "RAM 磁盘"
End Select

End Function

Function cIsReady(trd)

Select Case trd
Case True: cIsReady = "<font class=fonts><b>√</b></font>"
Case False: cIsReady = "<font color='red'><b>×</b></font>"
End Select

End Function

Function cSize(tSize)

If tSize >= 1073741824 Then
cSize = Int((tSize/1073741824)*1000)/1000 & " GB"
ElseIf tSize >= 1048576 Then
cSize = Int((tSize/1048576)*1000)/1000 & " MB"
ElseIf tSize >= 1024 Then
cSize = Int((tSize/1024)*1000)/1000 & " KB"
Else
cSize = tSize & "B"
End If

End Function %>