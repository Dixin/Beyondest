<!-- #include file="INCLUDE/config_user.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Select Case action
    Case "article"
        tit    = "查看我发表的文章"
    Case "down"
        tit    = "查看我添加的软件"
    Case "gallery"
        tit    = "查看我上传的贴图"
    Case "website"
        tit    = "查看我推荐的网站"
    Case Else
        action = "news"
        tit    = "查看我发布的新闻"
End Select

Dim rssum,nummer,page,thepages,viewpage,pageurl,types,topic,tim
rssum   = 0:thepages = 0:viewpage = 1:nummer = web_var(web_num,1)
pageurl = "?action=" & action & "&"

Call web_head(2,0,0,0,0)
'------------------------------------left----------------------------------
Call left_user()
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong & table1 %>
<tr<% Response.Write table2 %> height=25><td class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif>&nbsp;<% Response.Write img_small(us) %>&nbsp;&nbsp;<b>查看我所发表的相关信息</b></td></tr>
<tr<% Response.Write table3 %>><td align=center height=30>
<% Response.Write img_small("jt12") %><a href='?action=news'<% If action = "news" Then Response.Write "class=red_3" %>>查看我所发布的新闻</a>　
<% Response.Write img_small("jt12") %><a href='?action=article'<% If action = "article" Then Response.Write "class=red_3" %>>查看我所发表的文章</a>　
<% Response.Write img_small("jt12") %><a href='?action=down'<% If action = "down" Then Response.Write "class=red_3" %>>查看我所添加的软件</a>　
<% Response.Write img_small("jt12") %><a href='?action=gallery'<% If action = "gallery" Then Response.Write "class=red_3" %>>查看我所上传的图片</a>
</td></tr>
</table>
<%

Select Case action
    Case "article"
        sql     = "select id,topic,tim,counter from article where username='" & login_username & "' and hidden=1 order by id desc"
    Case "down"
        sql     = "select id,name,tim,counter from down where username='" & login_username & "' and hidden=1 order by id desc"
    Case "gallery"
        types   = Trim(Request.querystring("types"))
        If types <> "logo" And types <> "baner" Then types = "paste"
        pageurl = pageurl & "types=" & types & "&"

        Select Case types
            Case "logo"
                nummer = nummer*2
            Case "baner"
                nummer = web_var(web_num,3)
        End Select

        sql            = "select * from gallery where hidden=1 and types='" & types & "' and username='" & login_username & "' order by id desc"
    Case Else
        sql            = "select id,topic,tim,counter from news where username='" & login_username & "' and hidden=1 order by id desc"
End Select

Set rs = Server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1

If Not(rs.eof And rs.bof) Then
    rssum = rs.recordcount
End If

Call format_pagecute()

Response.Write ukong & table1 %>

<%

If Int(viewpage) > 1 Then
    rs.move (viewpage - 1)*nummer
End If

Select Case action
    Case "article"
        Call putview_article()
    Case "down"
        Call putview_down()
    Case "gallery"
        Call putview_gallery()
    Case Else
        Call putview_news()
End Select

rs.Close:Set rs = Nothing %>
<tr><td align=center bgcolor=<% = web_var(web_color,6) %> height=30 colspan=2<% Response.Write table3 %>>
  <table border=0 width='98%' cellspacing=0 cellpadding=0>
<tr align=center valign=bottom><td width='30%' >
现在有<font class=red><% Response.Write rssum %></font>条记录┋
每页<font class=red><% Response.Write nummer %></font>个
  </td><td width='70%' bgcolor=<% = web_var(web_color,6) %>>
页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font> 分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,5,"#ff0000") %>
  </td></tr>
  </table>
</td></tr>  
</table>
<br>
<%
'---------------------------------center end-------------------------------
Call web_end(0)

Sub putview_news() %>
<tr align=center<% Response.Write table2 %> height=25>
<td width='6%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>新闻标题111</b></td>
</tr>
<%

    For i = 1 To nummer
        If rs.eof Then Exit For
        topic = rs("topic"):tim = rs("tim")
        Response.Write vbcrlf & "<tr" & table3 & "><td align=center>" & (viewpage - 1)*nummer + i & ".</td><td><a target=_blank href='news_view.asp?id=" & rs("id") & "' title='新闻标题：" & code_html(topic,1,0) & "<br>浏览次数：" & rs("counter") & "<br>发布时间：" & tim & "'>" & code_html(topic,1,35) & "</a>" & format_end(1,time_type(tim,3)) & "</td></tr>"
        rs.movenext
    Next

End Sub

Sub putview_article() %>
<tr align=center<% Response.Write table2 %> height=25>
<td width='6%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>文章标题</b></td>
</tr>
<%

    For i = 1 To nummer
        If rs.eof Then Exit For
        topic = rs("topic"):tim = rs("tim")
        Response.Write vbcrlf & "<tr" & table3 & "><td align=center>" & (viewpage - 1)*nummer + i & ".</td><td><a target=_blank href='article_view.asp?id=" & rs("id") & "' title='文章标题：" & code_html(topic,1,0) & "<br>发表时间：" & tim & "'>" & code_html(topic,1,35) & "</a>" & format_end(1,time_type(tim,3) & ",<font class=blue>" & rs("counter") & "</font>") & "</td></tr>"
        rs.movenext
    Next

End Sub

Sub putview_down() %>
<tr align=center<% Response.Write table2 %> height=25>
<td width='6%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>序号</b></td>
<td width='84%' class=end background=images/<% = web_var(web_config,5) %>/bar_3_bg.gif><b>软件名称</b></td>
</tr>
<%

    For i = 1 To nummer
        If rs.eof Then Exit For
        topic = rs("name"):tim = rs("tim")
        Response.Write vbcrlf & "<tr" & table3 & "><td align=center>" & (viewpage - 1)*nummer + i & ".</td><td><a target=_blank href='article_view.asp?id=" & rs("id") & "' title='软件名称：" & code_html(topic,1,0) & "<br>添加时间：" & tim & "'>" & code_html(topic,1,35) & "</a>" & format_end(1,time_type(tim,3) & ",<font class=blue>" & rs("counter") & "</font>") & "</td></tr>"
        rs.movenext
    Next

End Sub

Sub putview_gallery()
    Dim j,k,kn,pic,name,nnum:nnum = 1
    Response.Write vbcrlf & "<tr" & table3 & "><td align=center>" & _
    vbcrlf & "<table border=0>" & _
    vbcrlf & "<tr><td width=100>" & img_small("jt1") & "<a href='?action=" & action & "&types=paste'"
    If types = "paste" Then Response.Write " class=red_3"
    Response.Write vbcrlf & ">精彩贴图</a></td>" & _
    vbcrlf & "<td width=100>" & img_small("jt1") & "<a href='?action=" & action & "&types=logo'"
    If types = "logo" Then Response.Write " class=red_3"
    Response.Write vbcrlf & ">精彩LOGO</a></td>" & _
    vbcrlf & "<td width=100>" & img_small("jt1") & "<a href='?action=" & action & "&types=baner'"
    If types = "baner" Then Response.Write " class=red_3"
    Response.Write vbcrlf & ">精彩BANNER</a></td></tr>" & _
    vbcrlf & "</table></td></tr><tr" & table3 & "><td align=center><table border=0 width='100%'>"

    Select Case types
        Case "logo"
            kn    = 5:nummer = 30

            If nummer Mod kn > 0 Then
                k = nummer\kn + 1
            Else
                k = nummer\kn
            End If

            If Int(viewpage) > 1 Then
                rs.move (viewpage - 1)*nummer
            End If

            For i = 1 To k
                'if rs.eof then exit for
                Response.Write "<tr align=center>"

                For j = 1 To kn
                    If rs.eof Or nnum > nummer Then Exit For
                    pic = rs("pic"):name = rs("name")
                    Response.Write "<td><table border=0><tr><td align=center><img src='" & web_var(web_upload,1) & pic & "' border=0 width=88 height=31></td></tr><tr><td align=center title='" & code_html(name,1,0) & "'>" & code_html(name,1,10) & "</td></tr></table></td>"
                    rs.movenext
                    nnum = nnum + 1
                Next

                Response.Write "</tr>"
            Next

        Case "baner"
            nummer = web_var(web_num,3)

            If Int(viewpage) > 1 Then
                rs.move (viewpage - 1)*nummer
            End If

            For i = 1 To nummer
                If rs.eof Then Exit For
                pic = rs("pic"):name = rs("name")
                Response.Write "<tr><td><table border=0 align=center><tr><td align=center><img src='" & web_var(web_upload,1) & pic & "' border=0 width=468 height=60></td></tr><tr><td align=center>" & code_html(name,1,0) & "</td></tr></table></td></tr>"
                rs.movenext
            Next

        Case Else
            kn    = 3:nummer = 12

            If nummer Mod kn > 0 Then
                k = nummer\kn + 1
            Else
                k = nummer\kn
            End If

            If Int(viewpage) > 1 Then
                rs.move (viewpage - 1)*nummer
            End If

            For i = 1 To k
                'if rs.eof then exit for
                Response.Write "<tr align=center>"

                For j = 1 To kn
                    If rs.eof Or nnum > nummer Then Exit For
                    pic = rs("pic"):name = rs("name")
                    Response.Write "<td><table border=0><tr><td align=center><a href='gallery.asp?action=view&c_id=" & rs("c_id") & "&s_id=" & rs("s_id") & "&id=" & rs("id") & "'><img src='" & web_var(web_down,5) & "/" & pic & "' border=0 width=" & web_var(web_num,7) & " height=" & web_var(web_num,7) & "></a></td></tr><tr><td align=center>" & rs("name") & "</td></tr></table></td>"
                    rs.movenext
                    nnum = nnum + 1
                Next

                Response.Write "</tr>"
            Next

    End Select

    Response.Write "</table></td></tr>"
End Sub %>