<!-- #INCLUDE file="include/config.asp" -->
<!--#include file="INCLUDE/upload_config.asp"-->
<!--#include file="include/conn.asp"-->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ==================== %>
<script langiage='javascript' src='style/beyondest_config.js'></script>
<script langiage='javascript' src='STYLE/mouse_on_title.js'></script>
<html>
<head>
<title><% Response.Write web_var(web_config,1) %> - 文件上传</title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href='include/beyondest.css' type=text/css>
</head>
<body topmargin=0 leftmargin=0 bgcolor=<% Response.Write web_var(web_color,1) %>>
<table border=0 height='100%'cellspacing=0 cellpadding=0><tr><td height='100%'>
<%
Dim formname,upload_path,upload_type,upload_size,uup
uup    = "|article|down|forum|gallery|news|other|product|video|website|"
action = Trim(Request.querystring("action"))

If var_null(login_username) = "" Or var_null(login_password) = "" Then
    Response.Write web_var(web_error,2)
ElseIf post_chk() = "no" Then
    Response.Write web_var(web_error,1)
Else
    upload_path = web_var(web_upload,1)
    upload_type = web_var(web_upload,2)
    upload_size = web_var(web_upload,3)

    Select Case action
        Case "upfile"
            Call upload_chk()
        Case Else
            Call upload_main()
    End Select

End If

Call close_conn()

Sub upload_chk()
    Server.ScriptTimeOut = 5000
    Dim upload,up_text,up_path,uptemp,uppath,up_name,MyFso,upfile,upcount,upfilename,upfile_name,upfile_name2,upfilesize,upid
    Set upload           = new upload_classes
    Set MyFso            = CreateObject("Scripting.FileSystemObject")
    upcount              = 1
    up_name              = Trim(upload.form("up_name"))
    up_text              = Trim(upload.form("up_text"))
    up_path              = Trim(upload.form("up_path"))
    If Session("beyondest_online_admin") <> "beyondest_admin" And Len(up_name) > 2 Then up_name = ""
    If Len(up_name) < 3 Then up_name = up_name & upload_time(now_time)
    If Int(InStr(uup,"|" & up_path & "|")) = 0 Then up_path = "other"
    If Len(up_path) < 3 Then up_path = "other"
    uppath          = up_path
    If Right(upload_path,1) <> "/" Then upload_path = upload_path & "/"
    up_path         = Server.mappath(upload_path & up_path)

    If Not MyFso.folderExists(up_path) Then
        Set up_path = MyFso.CreateFolder(up_path)
    End If

    If Right(up_path,1) <> "/" Then up_path = up_path & "/"
    'if right(uppath,1)<>"/" then uppath=uppath&"/"

    Set upfile               = upload.file("file_name1")
    upfilesize               = upfile.FileSize
    upload_size              = upload_size*1024

    If upfilesize > 0 Then
        upfilename           = upfile.FileName

        If upfilesize > upload_size Then
            uptemp           = "<font class=red_2>上传失败</font>：文件太大！(不能超过" & Int(upload_size/1024) & "KB) " & go_back
        Else
            upfile_name      = Right(upfilename,(Len(upfilename) - InStr(upfilename,".")))
            upfile_name      = LCase(upfile_name)

            If InStr("," & upload_type & ",","," & upfile_name & ",") > 0 Then
                upfile_name2 = upfile_name
                upfile_name  = up_name & "." & upfile_name
                upfile.SaveAs up_path & upfile_name

                sql    = "select id from upload where url='" & uppath & "/" & upfile_name & "'"
                Set rs = conn.execute(sql)

                If rs.eof And rs.bof Then
                    rs.Close
                    sql = "insert into upload(iid,nsort,types,username,url,genre,sizes,tim) " & _
                    "values(0,'',0,'" & login_username & "','" & uppath & "/" & upfile_name & "','" & upfile_name2 & "'," & upfilesize & ",'" & now_time & "')"
                    conn.execute(sql)
                    sql    = "select top 1 id from upload order by id desc"
                    Set rs = conn.execute(sql)
                    upid   = Int(rs("id"))
                Else
                    conn.execute("update upload set username='" & login_username & "',sizes=" & upfilesize & ",tim='" & now_time & "' where id=" & rs("id"))
                End If

                rs.Close:Set rs = Nothing

                uptemp     = "<font class=red>上传成功</font>：<a href='" & upload_path & uppath & "/" & upfile_name & "' target=_blank>" & upfile_name & "</a> (" & upfilesize & "Byte)"

                If InStr(up_text,"pic") > 0 Then
                    uptemp = uptemp & "<script>parent.document.all." & up_text & ".value='" & uppath & "/" & upfile_name & "';"
                Else
                    uptemp = uptemp & "&nbsp;&nbsp;[ <a href='?uppath=" & uppath & "&upname=&uptext=" & up_text & "'>点击继续上传</a> ]<script>parent.document.all." & up_text & ".value+='"

                    Select Case LCase(upfile_name2)
                        Case "gif","jpg","bmp","png"
                            uptemp = uptemp & "[IMG]" & upload_path & uppath & "/" & upfile_name & "[/IMG]"
                        Case "swf"
                            uptemp = uptemp & "[FLASH=" & web_var(web_num,9) & "," & web_var(web_num,10) & "]" & upload_path & uppath & "/" & upfile_name & "[/FLASH]"
                        Case Else
                            uptemp = uptemp & "[DOWNLOAD]" & upload_path & uppath & "/" & upfile_name & "[/DOWNLOAD]"
                    End Select

                    uptemp         = uptemp & "\n';"
                End If

                If Int(upid) > 0 Then uptemp = uptemp & "parent.document.all.upid.value+='," & upid & "';"
                uptemp = uptemp & "</script>"
            Else
                uptemp = "<font class=red_2>上传失败</font>：文件类型只能为：" & Replace(upload_type,"|","、") & "等格式) " & go_back
            End If

        End If

    Else
        uptemp = "<font class=red_2>上传失败</font>：您可能没有选择想要上传的文件！" & go_back
    End If

    Set upfile = Nothing
    Response.Write uptemp
    Set MyFso  = Nothing
    Set upload = Nothing
End Sub

Sub upload_main()
    Dim uppath,upname,uptext
    uppath = Trim(Request.querystring("uppath"))
    upname = Trim(Request.querystring("upname"))
    uptext = Trim(Request.querystring("uptext"))
    If Session("beyondest_online_admin") <> "beyondest_admin" And Len(upname) > 2 Then upname = ""

    If Int(InStr(uup,"|" & uppath & "|")) = 0 Then Response.Write "参数出错！":Exit Sub

        If Len(uppath) < 1 Or Len(uptext) < 1 Then Response.Write "参数出错！":Exit Sub
            'if len(upname)<3 then upname=upname&upload_time(now_time) %>
<table border=0 cellspacing=0 cellpadding=2>
<form name=form1 action='?action=upfile' method=post enctype='multipart/form-data'>
<input type=hidden name=up_path value='<% Response.Write uppath %>'>
<input type=hidden name=up_name value='<% Response.Write upname %>'>
<input type=hidden name=up_text value='<% Response.Write uptext %>'>
<tr>
<td><input type=file name=file_name1 value='' size=35></td>
<td align=center height=30><input type=submit name=submit value='点击上传'> (<=<% Response.Write upload_size %>KB)</td>
</tr>
</form>
</table>
<% End Sub %>
</td></tr></table>
</body>
</html>
