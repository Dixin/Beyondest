<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,page,rssum,thepages,viewpage,pageurl,sqladd,keyword,sea_type,sea_name,topic,topic2,sql1,sql2,linkurl,keywords,tims
pageurl = "?":sqladd = "":topic = "":sql1 = "":sql2 = "":linkurl = "":keywords = "":sea_name = "搜索"
nummer  = 20:viewpage = 1:thepages = 0
tit     = "站内搜索"

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Response.Write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong
Call web_search(0)
Response.Write ukong
Call sea_types()
Call sql_add()

If sqladd = "" Then
    Call search_error()
Else
    Call search_main()
End If

Response.Write ukong
'---------------------------------center end-------------------------------
Call web_end(0)

Sub search_error() %>
<table border=0 width='96%'>
<tr><td height=300 align=center>
  <table border=0>
  <tr><td colspan=2 height=30>您可能没有填写“搜索关键字”，请查看以下帮助说明：</td></tr>
  <tr><td width=10></td><td><% Response.Write img_small("jt1") %>在搜索时必须填写“搜索关键字”；</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt12") %>如要搜索多个关键字请用<font class=red>空格</font>将多个关键字隔开，如：<font class=blue>V6&nbsp;插件</font>；</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt0") %>“关键字”中不能含有单引号（'）；</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt0") %>“关键字”中含有的加号（+）将被视为空格处理；</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt13") %>“快速搜索”只在：新闻、文栏、下载里有效；</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt14") %>祝您在使用本站的“站内搜索”时轻松愉快。</td></tr>
  </table>
</td></tr>
</table>
<%
End Sub

Sub search_main()
    sql    = sql1 & sqladd & sql2
    tims   = timer()
    Set rs = Server.CreateObject("adodb.recordset")
    rs.open sql,conn,1,1

    If rs.eof And rs.bof Then
        rssum = 0
    Else
        rssum = rs.recordcount
    End If

    Call format_pagecute() %>
  <table border=0 width='96%' cellspacing=0 cellpadding=2>
  <tr><td height=1 colspan=4 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr align=center valign=bottom<% Response.Write table4 %>>
  <td width='6%'>序号</td>
  <td width='94%'>相关内容（您查询的关键字是：<% Response.Write keywords %>每页 <font class=red><% Response.Write nummer %></font> 条 <font class=blue><% Response.Write sea_name %></font> 查询结果）</td>
  </tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr><td height=5></td></tr>
<%

    If Int(viewpage) > 1 Then
        rs.move (viewpage - 1)*nummer
    End If

    For i = 1 To nummer
        If rs.eof Then Exit For %>
  <tr>
  <td align=center><% Response.Write (viewpage - 1)*nummer + i %>.</td>
  <td><a href='<%
        Response.Write linkurl & rs(0)
        If sea_type = "forum" Then Response.Write "&forum_id=" & rs(4) %>' target=_blank><% Response.Write code_html(rs(1),1,32) %></a>&nbsp;<font class=gray size=1><% Response.Write time_type(rs(3),3) %></font>&nbsp;<% Response.Write format_user_view(rs(2),1,"blue") %></td>
  </tr>
<%
        rs.movenext
    Next

    rs.Close:Set rs = Nothing %>
  <tr><td height=5></td></tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  <tr><td colspan=2<% Response.Write table4 %>>
    <table border=0 width='100%' cellspacing=0 cellpadding=0>
    <tr>
    <td>共&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;条结果&nbsp;
页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>&nbsp;
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,3,"#ff0000") %></td>
    <td align=right><font size=2 class=gray>查询用时：<font class=red_3><% Response.Write FormatNumber((timer() - tims)*1000,3) %></font> 毫秒</font></td>
    </tr>
    </table>
  </td></tr>
  <tr><td height=1 colspan=2 background='IMAGES/BG_DIAN.GIF'></td></tr>
  </table>
<%
End Sub

Sub sql_add()
    Dim ddim,dnum,i
    keyword          = code_form(Request.querystring("keyword"))

    If Len(keyword) < 1 Or Len(topic) < 1 Then sqladd = "":Exit Sub
        keyword      = Replace(keyword,"+"," ")
        pageurl      = pageurl & "keyword=" & Server.urlencode(keyword) & "&"
        ddim         = Split(keyword," ")
        dnum         = UBound(ddim)

        For i = 0 To dnum
            keywords = keywords & "<font class=red_3><b>" & ddim(i) & "</b></font>&nbsp;&nbsp;"
            sqladd   = sqladd & " and " & topic2 & " like '%" & ddim(i) & "%'"
        Next

        Erase ddim

        If sea_type = "forum" And sqladd <> "" Then
            sqladd = Right(sqladd,Len(sqladd) - 4)
        End If

    End Sub

    Sub sea_types()
        Dim celerity
        celerity = Trim(Request.querystring("celerity"))
        sea_type = Trim(Request.querystring("sea_type"))

        Select Case sea_type
            Case "news","article"
                topic = "topic":topic2 = topic
                If celerity = "yes" Then topic2 = "keyes"
                linkurl = sea_type & "_view.asp?id="
                sea_name = "新闻"
                If sea_type = "article" Then sea_name = "文栏"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "down"
                topic = "name":topic2 = topic
                If celerity = "yes" Then topic2 = "keyes"
                linkurl = sea_type & "_view.asp?id="
                sea_name = "软件"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "website"
                topic = "name":topic2 = topic
                linkurl = sea_type & ".asp?action=view&id="
                sea_name = "网站"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "paste","flash"
                topic = "name":topic2 = topic
                linkurl = "gallery.asp?action=" & sea_type & "&types=view&id="
                sea_name = "图片"
                If sea_type = "flash" Then sea_name = "Flash"
                sql1 = "select id," & topic & ",username,tim from gallery where hidden=1 and types='" & sea_type & "'"
                sql2 = " order by id desc"
            Case Else
                sea_type = "forum"
                topic = "topic":topic2 = topic
                linkurl = "forum_view.asp?view_id="
                sea_name = "论坛"
                sql1 = "select id," & topic & ",username,tim,forum_id from bbs_topic where"
                sql2 = " order by id desc"
        End Select

        pageurl = pageurl & "sea_type=" & sea_type & "&"
    End Sub %>