<!--#include file="include/config.asp"-->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ==================== %>
<html>
<head>
<title><% Response.Write web_var(web_config,1) %> - 调查列表</title>
<meta name="Description"  content="Beyondest">
<meta name="keywords" content="最全的Beyond资料,最好的Beyond网站,asp,Beyondest,笼民">
<meta name="author" content="Beyondest">
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel=stylesheet href="include/beyondest.css" type=text/css>
</head>
<body leftmargin=0 topmargin=0>
<body topmargin=0 leftmargin=0 bgcolor=<%
Dim ttt,vid
vid = Trim(Request.querystring("vid"))
Response.Write web_var(web_color,1)
ttt = web_var(web_config,7)

If ttt <> "" Then
    Response.Write " background='images/" & ttt & ".gif'"
End If %>>
<center><table border=0 width='100%' height='100%' cellspacing=0 cellpadding=0>
<tr><td width='100%' height='100%' align=center>
  <table border=0 width='100%' height='100%' cellspacing=0 cellpadding=0>
  <tr><td width=20 height=16><img src='IMAGES/VOTE/vote_r1_c1.gif' width=20 height=16 border=0></td><td background='IMAGES/VOTE/vote_r1_c2.gif'></td><td width=16><img src='IMAGES/VOTE/vote_r1_c6.gif' width=16 height=16 border=0></td></tr>
  <tr><td background='IMAGES/VOTE/vote_r2_c1.gif' valign=bottom><img src='IMAGES/VOTE/vote_r4_c1.gif' width=20 border=0 height=8></td><td bgcolor=#ffffff align=center>
<%

If Not(IsNumeric(vid)) Then
    Call vote_error()
Else

    Select Case action
        Case "save"
            Call vote_save()
        Case Else
            Call vote_view()
    End Select

End If

Call close_conn() %>
  </td><td background='IMAGES/VOTE/vote_r2_c6.gif'></td></tr>
  <tr><td height=16 background='IMAGES/VOTE/vote_r5_c4.gif' colspan=2><a href="javascript:window.close();"><img src='IMAGES/VOTE/vote_r5_c1.gif' width=84 height=18 border=0></a></td><td width=16><img src='IMAGES/VOTE/vote_r5_c6.gif' width=16 height=18 border=0></td></tr>
  </table>
</td></tr></table></center>
</body>
</html>
<%

Sub vote_save()
    Dim go_tim,vvid:go_tim = web_var(web_num,5)

    If Trim(Request.cookies("beyondest_online")("vote_vid")) = "v" & vid Then %><font class=red_2>您已经投过一票！不可以重复多投……</font><br><br>
<a href='votetype.asp?type=view&vid=<% Response.Write vid %>'>查看投票结果</a><br><br>
<font class=gray>（系统将在 <font class=red><% Response.Write go_tim %></font> 秒钟后自动进入）</font><br><br>
<meta http-equiv='refresh' content='<% Response.Write go_tim %>; url=votetype.asp?type=view&vid=<% Response.Write vid %>'>
<% Exit Sub
        End If

        sql    = "select id from vote where vid=" & vid
        Set rs = conn.execute(sql)

        If rs.eof And rs.bof Then
            rs.Close:Set rs = Nothing

            Call vote_error():Exit Sub
            End If

            rs.Close:Set rs = Nothing
            Dim vote_id,ddim,j:j = 0
            vote_id  = Trim(Request.form("vote_id"))
            vote_id  = Replace(vote_id," ","")

            If Len(vote_id) < 1 Then Call vote_error():Exit Sub
                ddim = Split(vote_id,",")

                For i = 0 To UBound(ddim)

                    If IsNumeric(ddim(i)) Then
                        sql = "update vote set counter=counter+1 where vid=" & vid & " and vtype<>0 and id=" & ddim(i)
                        conn.execute(sql)
                        j   = j + 1
                    End If

                Next

                Erase ddim

                If j = 0 Then Call vote_error():Exit Sub
                    Response.cookies("beyondest_online")("vote_vid") = "v" & vid
                    Call cookies_yes() %>！！！<font class=red>谢谢你的支持与参与</font>！！！<br><br>
<a href='votetype.asp?type=view&vid=<% Response.Write vid %>'>查看投票结果</a><br><br>
<font class=gray>（系统将在 <font class=red><% Response.Write go_tim %></font> 秒钟后自动进入）</font><br><br>
<meta http-equiv='refresh' content='<% Response.Write go_tim %>; url=votetype.asp?type=view&vid=<% Response.Write vid %>'>
<%
                End Sub

                Sub vote_view()
                    Dim rssum,dimc,dimn,num,t
                    rssum  = 0:num = 0
                    Set rs = Server.CreateObject("adodb.recordset")
                    sql    = "select id,vname,counter from vote where vid=" & vid & " order by id"
                    rs.open sql,conn,1,1

                    If rs.eof And rs.bof Then
                        rs.Close:Set rs = Nothing

                        Call vote_error():Exit Sub
                        End If

                        rssum = Int(rs.recordcount)
                        rssum = rssum - 1
                        ReDim dimc(rssum),dimn(rssum)

                        For i = 0 To rssum
                            If rs.eof Then Exit For
                            dimc(i) = rs("counter")
                            dimn(i) = rs("vname")
                            num     = dimc(i) + num
                            rs.movenext
                        Next

                        rs.Close:Set rs = Nothing %>
<table border=0 width='96%' cellpadding=0 cellspacing=2 align=center>
<tr><td align=center colspan=3 class=red_3 height=20><b><% Response.Write code_html(dimn(0),1,0) %></b></td></tr>
<tr><td align=center colspan=3 class=gray>目前共有 <font class=blue><% Response.Write num %></font> 人参与了投票</td></tr>
<tr>
<td></td>
<td></td>
<td width='15%'></td>
</tr>
<%

                        For i = 1 To rssum

                            If Int(dimc(i)) = 0 Then
                                t = "0%"
                            Else
                                t = FormatPercent(dimc(i)/num,1)
                            End If %>
<tr>
<td height=18><% Response.Write i %>、<% Response.Write code_html(dimn(i),1,0) %> <font class=gray>(<font class=blue><% Response.Write dimc(i) %></font>)</font></td>
<td align=right><img src='IMAGES/VOTE/BAR.GIF' width=<% Response.Write dimc(i) %> height=10></td>
<td align=right><% Response.Write t %></td>
</tr>
<% Next %>
</table>
<%

                        Erase dimc:Erase dimn
                    End Sub

                    Sub vote_error() %>
<font class=red>您可能没有选择相关的择票选项！</font><br><br><font class=red_2>或进行非法的提交了投票数据！</font>
<br><br>
<%
                        Response.Write closer
                    End Sub %>