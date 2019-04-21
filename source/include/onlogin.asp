<!-- #include file="config.asp" -->
<!-- #include file="config_frm.asp" -->
<!-- #include file="config_upload.asp" -->
<!-- #include file="config_put.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

If Session("beyondest_online_admin") <> "beyondest_admin" Then
    Response.redirect "admin_login.asp"
    Response.End
End If

If web_login = 0 Then web_login = 2
Dim color1
Dim color2
Dim color3
Dim table1
Dim mtr
color1 = web_var(web_color,1)
color2 = web_var(web_color,5)
color3 = web_var(web_color,6)
table1 = " bordercolorlight=#c0c0c0 bordercolordark=" & color1
mtr    = " onmouseover=""javascript:this.bgColor='" & color2 & "';"" onmouseout=""javascript:this.bgColor='" & color1 & "';"""

Function del_select(delid)
    Dim del_i
    Dim del_num
    Dim del_dim
    Dim del_sql
    Dim del_rs
    Dim del_username
    Dim fobj
    Dim picc

    If delid <> "" And Not IsNull(delid) Then
        delid   = Replace(delid," ","")
        del_dim = Split(delid,",")
        del_num = UBound(del_dim)

        For del_i = 0 To del_num
            'del_sql
            del_sql    = "select username from " & data_name & " where id=" & del_dim(del_i)
            Set del_rs = conn.execute(del_sql)

            If Not(del_rs.eof And del_rs.bof) Then
                Call user_integral("del",web_varn(web_num,15),del_rs("username"))
            End If

            del_rs.Close:Set del_rs = Nothing
            Call upload_del(data_name,del_dim(del_i))
            del_sql = "delete from " & data_name & " where id=" & del_dim(del_i)
            conn.execute(del_sql)
        Next

        Erase del_dim
        del_select = vbcrlf & "<script language=javascript>alert(""共删除了 " & del_num + 1 & " 条记录！"");</script>"
    End If

End Function

Function header(popedomnum,titmenu)

    If Session("beyondest_online_admines") <> web_var(web_config,3) Then

        If Session("beyondest_online_admines") <> "beyondest" And popedom_formated(Session("beyondest_online_popedom"),popedomnum,0) = 0 Then
            Response.redirect "admin.asp?action=main&error=popedom"
            Response.End
        End If

    End If

    header = VbCrLf & "<html><head><title>" & web_var(web_config,1) & " - 管理后台</title>" & _
    VbCrLf & "<meta http-equiv=Content-Type content=text/html; charset=gb2312>" & _
    VbCrLf & "<link rel=stylesheet href='include/beyondest.css' type=text/css>" & _
    VbCrLf & "<script langiage='javascript' src='style/open_win.js'></script>" & _
    VbCrLf & "<script langiage='javascript' src='style/mouse_on_title.js'></script>" & _
    VbCrLf & "</head>" & VbCrLf & "<body topmargin=0 leftmargin=0 bgcolor=" & color1 & "><center>" & _
    VbCrLf & "<table border=0 width=600 cellspacing=0 cellpadding=0>" & _
    vbcrlf & "<tr><td height=50 align=center>" & titmenu & "&nbsp;┋&nbsp;<a href='javascript:;' onclick=""javascript:document.location.reload()"">刷新</a></td></tr><tr><td align=center height=350>"
End Function

Function popedom_formated(popedom1,popedomnum,popedomtype)
    Dim poptemp:poptemp = 0

    If Len(popedom1) = 50 And popedomnum <> - 1 Then
        poptemp = Mid(popedom1,popedomnum,1)
    End If

    If popedomtype <> 0 Then

        If poptemp = 0 Then
            poptemp = 1
        Else
            poptemp = 0
        End If

    End If

    If poptemp <> 0 Then poptemp = 1
    If popedomnum =  - 1 Then poptemp = 1
    popedom_formated = poptemp
End Function %>