<!-- #include file="config.asp" -->
<!-- #include file="config_nsort.asp" -->
<!-- #include file="skin.asp" -->
<!-- #include file="jk_ubb.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim atb,ispic,topic,nid,sqladd
atb       = " target=_blank"
index_url = "news":n_sort = "news"
tit_fir   = format_menu(index_url)

Sub news_class_sort(t1,t2)
    Response.Write format_barc("<font class=" & sk_class & "><b>新闻分类</b></font>",class_sort(n_sort,index_url,t1,t2),1,1,8)
End Sub

Sub news_fpic(fsql,t_num,w_num,tt)
    Dim topic,temp1,rs,sql
    sql       = "select top 1 id,topic,comto,word,pic,tim from news where hidden=1 and istop=1 and ispic=1" & fsql & " order by id desc"
    Set rs    = conn.execute(sql)

    If Not(rs.eof And rs.bof) Then
        nid   = rs("id"):topic = rs("topic")
        temp1 = "<table border=0 width='100%' class=tf>" & _
        vbcrlf & "<tr align=center><td width='32%'>" & _
        vbcrlf & "  <table border=0>" & _
        vbcrlf & "  <tr><td align=center><a href='news_view.asp?id=" & nid & "' target=_blank class=red_3><img src='" & web_var(web_down,5) & "/" & rs("pic") & "' border=0 title='" & code_html(rs("topic"),1,0) & "' onload=""javascript:this.width=120""></a></td></tr>" & _
        vbcrlf & "  <tr><td height=25 align=center class=btd><a href='news_view.asp?id=" & nid & "' target=_blank class=red_3>" & code_html(rs("topic"),1,t_num) & "</a></td></tr>" & _
        vbcrlf & "  </table>" & _
        vbcrlf & "</td><td width='68%'>" & _
        vbcrlf & "  <table border=0 width='100%'>" & _
        vbcrlf & "  <tr><td vlign=middle height=150>" & Left(code_jk(rs("word")),w_num) & "……</td></tr>" & _
        vbcrlf & "  <tr><td height=25 class=gray align=right>（" & time_type(rs("tim"),tt) & "&nbsp;）&nbsp;</td></tr>" & _
        vbcrlf & "  </table>" & _
        vbcrlf & "</td></tr></table>"
    End If

    rs.Close
    Response.Write format_barc("<font class=" & sk_class & "><b>最新图片新闻</b></font>",temp1,1,1,1)
End Sub

Sub news_sea()
    Dim temp1,nid,nid2,rs,sql,rs2,sql2
    temp1 = vbcrlf & "<table border=0 cellspacing=0 cellpadding=0 align=center>" & _
    vbcrlf & "<script language=javascript><!--" & _
    vbcrlf & "function news_sea()" & _
    vbcrlf & "{" & _
    vbcrlf & "  if (news_sea_frm.keyword.value==""请输入关键字"")" & _
    vbcrlf & "  {" & _
    vbcrlf & "    alert(""请在搜索新闻前先输入要查询的 关键字 ！"");" & _
    vbcrlf & "    news_sea_frm.keyword.focus();" & _
    vbcrlf & "    return false;" & _
    vbcrlf & "  }" & _
    vbcrlf & "}" & _
    vbcrlf & "--></script>" & _
    vbcrlf & "<form name=news_sea_frm action='news_list.asp' method=get onsubmit=""return news_sea()"">" & _
    vbcrlf & "<input type=hidden name=action value='more'><tr><td height=3></td></tr>" & _
    vbcrlf & "<tr><td>" & _
    vbcrlf & "  <table border=0><tr><td colspan=2><input type=text name=keyword value='请输入关键字' onfocus=""if (value =='请输入关键字'){value =''}"" onblur=""if (value ==''){value='请输入关键字'}"" size=20 maxlength=20></td></tr><tr>" & _
    vbcrlf & "  <td><select name=sea_type sizs=1><option value='topic'>新闻标题</option><option value='username'>发布人</option></seelct></td>" & _
    vbcrlf & "  <td></td>" & _
    vbcrlf & "  </tr></table>" & _
    vbcrlf & "</td></tr><tr><td>" & _
    vbcrlf & "  <table border=0><tr>" & _
    vbcrlf & "  <td><select name=c_id sizs=1><option value=''>请选择新闻类别</option>"
    sql           = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order,c_id"
    Set rs        = conn.execute(sql)

    Do While Not rs.eof
        nid       = Int(rs(0))
        temp1     = temp1 & vbcrlf & "<option value='" & nid & "' class=bg_2"
        If cid = nid Then temp1 = temp1 & " selected"
        temp1     = temp1 & ">" & rs(1) & "</option>"
        sql2      = "select s_id,s_name from jk_sort where c_id=" & nid & " order by s_order,s_id"
        Set rs2   = conn.execute(sql2)

        Do While Not rs2.eof
            nid2  = rs2(0)
            temp1 = temp1 & vbcrlf & "<option value='" & nid & "&s_id=" & nid2 & "'"
            If sid = nid2 Then temp1 = temp1 & " selected"
            temp1 = temp1 & ">　" & rs2(1) & "</option>"
            rs2.movenext
        Loop

        rs2.Close:Set rs2 = Nothing
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    temp1 = temp1 & vbcrlf & "</select></td>" & _
    vbcrlf & "  <td><input type=image src='images/small/search_go.gif' border=0 width=40 height=25></td>" & _
    vbcrlf & "  </tr></table>" & _
    vbcrlf & "</td></tr>" & _
    vbcrlf & "</form><tr><td height=1></td></tr></table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>新闻搜索</b></font>",temp1,2,0,9)
End Sub

Sub news_scroll(sh,nsql,s_num,c_num,sbg)
    Dim cnum,temp1

    If sbg = 1 Then
        temp1 = vbcrlf & "<table border=0 width=176 cellspacing=2 cellpadding=2 height=25 background='images/" & web_var(web_config,5) & "/news_bg_scroll.gif'><tr><td width='13%'></td><td width='86%' valign=bottom>"
    Else
        temp1 = vbcrlf & "<table border=0 width='96%'>"
    End If

    temp1     = temp1 & "<marquee scrolldelay=120 scrollamount=4 onMouseOut=""if (document.all!=null){this.start()}"" onMouseOver=""if (document.all!=null){this.stop()}"">"
    sql       = "select top " & s_num & " id,topic,ispic from news where hidden=1" & nsql & " order by counter desc,id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        cnum  = c_num:ispic = "":topic = rs("topic")
        If rs("ispic") = True Then cnum = cnum - 2:ispic = sk_img
        temp1 = temp1 & "&nbsp;" & img_small(sh) & "<a href='news_view.asp?id=" & rs("id") & "'>" & code_html(topic,1,cnum) & "</a>" & ispic
        rs.movenext
    Loop

    rs.Close
    temp1     = temp1 & "</marquee>"

    If sbg = 1 Then
        temp1 = temp1 & "</td><td width='1%'></td></tr></table>"
    Else
        temp1 = temp1 & "</td></tr></table>"
    End If

    Response.Write temp1
End Sub

Sub news_new_hot(n_jt,nsql,nt,n_num,c_num,et,ct,tt)
    Dim htemp,tim,cnum,nhead
    If n_jt <> "" Then n_jt = img_small(n_jt)
    htemp     = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql       = "select top " & n_num & " id,username,topic,tim,counter,ispic from news where hidden=1" & nsql & " order by "

    If nt = "hot" Then
        sql   = sql & "counter desc,"
        nhead = "热门新闻"
    Else
        nhead = "近期更新"
    End If

    sql       = sql & "id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        cnum  = c_num:ispic = "":topic = rs("topic"):tim = rs("tim")
        If rs("ispic") = True Then cnum = cnum - 2:ispic = sk_img
        htemp = htemp & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & n_jt & "<a href='news_view.asp?id=" & rs("id") & "'" & atb & " title='新闻标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>浏览人次：" & rs("counter") & "<br>整理时间：" & tim & "'>" & code_html(topic,1,cnum) & "</a>" & ispic
        If tt > 0 Then htemp = htemp & format_end(et,time_type(tim,tt))
        htemp = htemp & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    htemp = ukong & htemp & vbcrlf & "</table>" & kong
    Response.Write kong & format_barc("<font class=" & sk_class & "><b>" & nhead & "</b></font>",htemp,2,0,12)
End Sub

Sub news_pic(nsql,n_num,c_num,pc)
    Dim temp1
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=0>"
    sql       = "select top " & n_num & " id,topic,pic from news where hidden=1 and ispic=1" & nsql & " order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic"):nid = rs("id")
        temp1 = temp1 & vbcrlf & "<tr><td align=center><a href='news_view.asp?id=" & nid & "'" & atb & "><img src='" & web_var(web_down,5) & "/" & rs("pic") & "' title='" & code_html(topic,1,0) & "' border=0 onload=""javascript:this.width=120""></a></td></tr>" & _
        vbcrlf & "<tr><td align=center height=25><a href='news_view.asp?id=" & nid & "'" & atb & " class=red_3><b>" & code_html(topic,1,c_num) & "</b></a></td></tr>"
        rs.movenext
    Loop

    temp1 = ukong & temp1 & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>图片新闻</b></font>",temp1,1,1,10)
End Sub

Sub news_picr(nsql,n_num,c_num,pc)
    Dim temp1
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=0>"
    sql       = "select top " & n_num & " id,topic,pic from news where hidden=1 and ispic=1" & nsql & " order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic"):nid = rs("id")
        temp1 = temp1 & vbcrlf & "<tr><td align=center><a href='news_view.asp?id=" & nid & "'" & atb & "><img src='" & web_var(web_down,5) & "/" & rs("pic") & "' title='" & code_html(topic,1,0) & "' border=0  onload=""javascript:this.width=120""></a></td></tr>" & _
        vbcrlf & "<tr><td align=center height=25><a href='news_view.asp?id=" & nid & "'" & atb & " class=red_3><b>" & code_html(topic,1,c_num) & "</b></a></td></tr>"
        rs.movenext
    Loop

    temp1 = ukong & temp1 & "</table>"
    Response.Write format_barc("<font class=" & sk_class & "><b>图片新闻</b></font>",temp1,2,0,10)
End Sub

Sub news_main(n_jt,n_num,c_num,et,ct,tt,pn,pl,pc)
    Dim ccid,ccname,sqla,crs,csql,nn,temp1,icon_num,tim,cnum:nn = 0
    csql          = "select c_id,c_name from jk_class where nsort='" & n_sort & "' order by c_order"
    Set crs       = conn.execute(csql)
    icon_num      = 1

    Do While Not crs.eof
        temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr><td height=1></td><td wdith=50></td></tr>"
        ccid      = crs("c_id"):ccname = crs("c_name"):sqla = " and c_id=" & ccid
        sql       = "select top " & n_num & " id,topic,tim,username,counter,ispic from news where hidden=1 and c_id=" & ccid & " order by id desc"
        Set rs    = conn.execute(sql)

        Do While Not rs.eof
            cnum  = c_num:ispic = "":topic = rs("topic"):tim = rs("tim")
            If rs("ispic") = True Then cnum = cnum - 2:ispic = sk_img
            temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & img_small(n_jt) & "<a href='news_view.asp?id=" & rs("id") & "'" & atb & " title='新闻标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>浏览人次：" & rs("counter") & "<br>整理时间：" & tim & "'>" & code_html(topic,1,cnum) & "</a>" & ispic & "</td><td>" & format_end(et,time_type(tim,tt)) & "</td></tr>"
            rs.movenext
        Loop

        rs.Close
        temp1 = temp1 & vbcrlf & "</table>"
        Response.Write vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr valign=top>"

        If nn = 0 Then
            Response.Write vbcrlf & "<td width=400>"
            Response.Write format_barc("<a href='news_list.asp?c_id=" & ccid & "'><b><font class=" & sk_class & ">" & ccname & "</font></b></a>",temp1,3,0,icon_num)
            Response.Write vbcrlf & "</td><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td bgcolor=" & web_var(web_color,1) & ">"
            Call news_pic(sqla,pn,pl,pc)
            nn = 1
        Else
            Response.Write vbcrlf & "<td bgcolor=" & web_var(web_color,1) & ">"
            Call news_pic(sqla,pn,pl,pc)
            Response.Write vbcrlf & "</td><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td width=400>"
            Response.Write format_barc("<a href='news_list.asp?c_id=" & ccid & "'><b><font class=" & sk_class & ">" & ccname & "</font></b></a>",temp1,3,0,icon_num)
            nn = 0
        End If

        Response.Write "</td></tr><tr><td colspan=3>" & gang & "</td></tr></table>"
        icon_num = icon_num + 1
        crs.movenext
    Loop

    crs.Close:Set crs = Nothing
End Sub

Sub news_more(n_jt,c_num,et,ct,tt,pn,pl,pc)
    Dim temp1,tim,cnum,sql2,mhead,cname,sname
    pageurl  = "?action=more&"
    keyword  = code_form(Request.querystring("keyword"))
    sea_type = Trim(Request.querystring("sea_type"))
    If sea_type <> "username" Then sea_type = "topic"
    Call cid_sid_sql(2,sea_type)

    temp1        = "<table border=0 width='100%'><tr><td height=1 width='5%'></td><td width='77%'></td><td wdith='18%'></td></tr>"
    sql          = "select id,topic,tim,username,counter,ispic from news where hidden=1" & sqladd

    If cid > 0 Then
        sql      = sql & " and c_id=" & cid

        If sid > 0 Then
            sql  = sql & " and s_id=" & sid
            sql2 = "select jk_class.c_name,jk_sort.s_name from jk_sort inner join jk_class on jk_sort.c_id=jk_class.c_id where jk_sort.c_id=" & cid & " and jk_sort.s_id=" & sid
        Else
            sql2 = "select c_name from jk_class where c_id=" & cid
        End If

    End If

    sql        = sql & " order by id desc"

    cname      = "搜索结果":sname = ""

    If Len(sql2) > 1 Then
        Set rs = conn.execute(sql2)

        If rs.eof And rs.bof Then
            rs.Close
            Call news_main("jt0",16,20,1,6,33,2,10,1)

            Exit Sub
            End If

            cname = rs("c_name")
            If sid > 0 Then sname = rs("s_name")
            rs.Close
        End If

        mhead  = "<a href='news_list.asp?c_id=" & cid & "'><b><font class=" & sk_class & ">" & cname & "</font></b></a>"
        If sid > 0 And sname <> "" Then mhead = mhead & "&nbsp;<font class=" & sk_class & ">→</font>&nbsp;<a href='news_list.asp?c_id=" & cid & "&s_id=" & sid & "'><b><font class=" & sk_class & ">" & sname & "</font></b></a>"
        Set rs = Server.CreateObject("adodb.recordset")
        rs.open sql,conn,1,1

        If rs.eof And rs.bof Then
            rssum = 0
        Else
            rssum = rs.recordcount
        End If

        Call format_pagecute()

        If Int(viewpage) > 1 Then
            rs.move (viewpage - 1)*nummer
        End If

        For i = 1 To nummer
            If rs.eof Then Exit For
            cnum  = c_num:ispic = "":topic = rs("topic"):tim = rs("tim")
            If rs("ispic") = True Then cnum = cnum - 2:ispic = sk_img
            temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & i + (viewpage - 1)*nummer & ".</td><td><a href='news_view.asp?id=" & rs("id") & "'" & atb & " title='新闻标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>浏览人次：" & rs("counter") & "<br>整理时间：" & tim & "'>" & code_html(topic,1,cnum) & "</a>" & ispic & "</td><td>" & format_end(et,time_type(tim,tt)) & "</td></tr>"
            rs.movenext
        Next

        rs.Close
        temp1 = temp1 & vbcrlf & "</table>"
        Response.Write format_barc(mhead,temp1,3,0,4) %>
<table border=0 width='100%' align=center>
<tr><td align=center><table border=0 width='100%'><tr><td height=1 background='images/bg_dian.gif'></td></tr></table></td></tr>
<tr><td>&nbsp;
本栏共有&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;条新闻&nbsp;
页次：<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>&nbsp;
分页：<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,8,"#ff0000") %>
</td></tr>
</table>
<%
    End Sub

    Sub news_list(n_jt,n_num,c_num,et,ct,tt,pn,pl,pc)
        Dim ssid,ssname,sqla,srs,icon_num,ssql,nn,temp1,tim,cnum:nn = 0
        ssql          = "select s_id,s_name from jk_sort where c_id=" & cid
        If sid <> 0 Then ssql = ssql & " and s_id=" & sid
        ssql          = ssql & " order by s_order"
        Set srs       = conn.execute(ssql)
        icon_num      = 1

        Do While Not srs.eof
            temp1     = "<table border=0 width='100%'><tr><td height=1></td><td wdith=50></td></tr>"
            ssid      = srs("s_id"):ssname = srs("s_name"):sqla = " and c_id=" & cid & " and s_id=" & ssid
            sql       = "select top " & n_num & " id,topic,tim,username,counter,ispic from news where hidden=1 and c_id=" & cid & " and s_id=" & ssid & " order by id desc"
            Set rs    = conn.execute(sql)

            Do While Not rs.eof
                cnum  = c_num:ispic = "":topic = rs("topic"):tim = rs("tim")
                If rs("ispic") = True Then cnum = cnum - 2:ispic = sk_img
                temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & img_small(n_jt) & "<a href='news_view.asp?id=" & rs("id") & "'" & atb & " title='新闻标题：" & code_html(topic,1,0) & "<br>发 布 人：" & rs("username") & "<br>浏览人次：" & rs("counter") & "<br>整理时间：" & tim & "'>" & code_html(topic,1,cnum) & "</a>" & ispic & "</td><td>" & format_end(et,time_type(tim,tt)) & "</td></tr>"
                rs.movenext
            Loop

            rs.Close
            temp1 = temp1 & vbcrlf & "</table>"
            Response.Write vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr valign=top>"

            If nn = 0 Then
                Response.Write vbcrlf & "<td width=400>"
                Response.Write format_barc("<a href='news_list.asp?c_id=" & cid & "&s_id=" & ssid & "'><b><font class=" & sk_class & ">" & ssname & "</font></b></a>",temp1,3,0,icon_num)
                Response.Write vbcrlf & "</td><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td bgcolor=" & web_var(web_color,1) & ">"
                Call news_pic(sqla,pn,pl,pc)
                nn = 1
            Else
                Response.Write vbcrlf & "<td bgcolor=" & web_var(web_color,1) & ">"
                Call news_pic(sqla,pn,pl,pc)
                Response.Write vbcrlf & "</td><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td width=400>"
                Response.Write format_barc("<a href='news_list.asp?c_id=" & cid & "&s_id=" & ssid & "'><b><font class=" & sk_class & ">" & ssname & "</font></b></a>",temp1,3,0,icon_num)
                nn = 0
            End If

            Response.Write vbcrlf & "</td></tr><tr><td colspan=3>" & gang & "</td></tr></table>"
            icon_num = icon_num + 1
            srs.movenext
        Loop

        srs.Close:Set srs = Nothing
    End Sub %>