<!-- #include file="config.asp" -->
<!-- #include file="skin.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

index_url = "main"
tit_fir   = format_menu(index_url)

Dim us
Dim table1
Dim table2
Dim table3
Dim table4
us     = "fk2"
table1 = format_table(1,3)
table2 = format_table(3,2)
table3 = format_table(3,1)
table4 = format_table(3,5)

Sub user_data_top(utt,ijt,sh,n_num)
    '��������	integral
    '��������	bbs_counter
    '�Ƹ�����	emoney
    Dim temp1
    temp1     = vbcrlf & "<table border=0 width='99%' cellspacing=2 cellpadding=0 align=center class=tf><tr height=0><td width='75%'></td><td width='25%'></td></tr>"
    sql       = "select top " & n_num & " username," & utt & " from user_data order by " & utt & " desc,id"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & img_small(ijt) & format_user_view(rs("username"),1,"") & "</td><td align=center class=red_3>" & rs(utt) & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Call left_btype(temp1,utt,sh,12)
End Sub

Sub vote_type(t1,t2,t3,t4)
    Call left_btype("<script language=javascript src='vote.asp?id=" & t1 & "&types=" & t2 & "&mcolor=" & t3 & "&bgcolor=" & t4 & "'></script>","vote",1,13)
End Sub

Sub left_type(ltv,pn,ltt)
    If ltt = 1 Then Response.Write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"
    Response.Write kong & format_barc("<img src='images/" & web_var(web_config,5) & "/left_" & pn & ".gif' border=0>",ltv,2,0,10)
    If ltt = 1 Then Response.Write "</td></tr></table>"
End Sub

Sub left_btype(ltv,pn,ltt,lbicon)
    If ltt = 1 Then Response.Write "<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center><tr><td align=center>"
    Response.Write kong & format_barc("<img src='images/" & web_var(web_config,5) & "/left_" & pn & ".gif' border=0>",ltv,2,0,lbicon)
    If ltt = 1 Then Response.Write "</td></tr></table>"
End Sub

Sub main_down_pic(nnhead,dsql,nt,n_num,c_num)
    Dim temp1
    Dim nid
    Dim name
    Dim nhead:nhead = nnhead
    temp1 = "<table border=0 width='100%' cellspacing=0 cellpadding=2><tr align=center valign=top>"
    sql   = "select top " & n_num & " id,name,tim,pic from down where hidden=1" & dsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "�ȵ�����"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "�����Ƽ�"
        Case Else
            sql = sql & " order by id desc"
            If nhead = "" Then nhead = "�����ϴ�"
    End Select

    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        name  = rs("name"):nid = rs("id")
        temp1 = temp1 & vbcrlf & "<td width='" & Int(100\n_num) & "%'><table border=0 cellspacing=0 cellpadding=2 width='100%' class=tf><tr><td align=center><a href='down_view.asp?id=" & nid & "'" & atb & "><img src='" & web_var(web_upload,1) & rs("pic") & "' border=0 width=" & web_var(web_down,1) & " height=" & web_var(web_down,2) & "></a></td></tr>" & _
        vbcrlf & "<tr><td align=center class=bw><a href='down_view.asp?id=" & nid & "'" & atb & " class=red_3><b>" & code_html(name,1,0) & "</b></a></td></tr></table></td>"
        rs.movenext
    Loop

    temp1 = temp1 & "</tr></table>"
    Response.Write kong & format_bar("<font class=" & sk_class & "><b>" & nhead & "</b></font>",temp1,sk_bar,0,0,"||","")
End Sub

Sub main_down(n_jt,nnhead,nmore,nsql,nt,n_num,n_m,c_num,et,tt)
    Dim temp1
    Dim nid
    Dim name
    Dim tim
    Dim counter
    Dim nhead:nhead = nnhead
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql   = "select top " & n_num + n_m & " id,name,username,tim,counter from down where hidden=1" & nsql

    Select Case nt
        Case "hot"
            sql = sql & " order by counter desc,id desc"
            If nhead = "" Then nhead = "��������"
        Case "good"
            sql = sql & " and types=5 order by id desc"
            If nhead = "" Then nhead = "�����Ƽ�"
        Case Else
            sql = sql & " order by id desc"
            If nhead = "" Then nhead = "���ָ���"
    End Select

    Set rs = conn.execute(sql)
    If n_m > 0 Then rs.move(n_m)

    Do While Not rs.eof
        name  = rs("name"):tim = rs("tim"):counter = rs("counter")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & " class=bw>" & n_jt & "<a href='down_view.asp?id=" & rs("id") & "'" & atb & " title='�������ƣ�" & code_html(name,1,0) & "<br>�� �� �ˣ�" & rs("username") & "<br>�����˴Σ�" & counter & "<br>����ʱ�䣺" & time_type(tim,88) & "'>" & code_html(name,1,c_num) & "</a>"
        If tt > 0 Then temp1 = temp1 & format_end(et,time_type(tim,tt) & "�������<font class=blue>" & counter & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & vbcrlf & "</table>"
    Response.Write format_barc("<a href=down.asp><font class=" & sk_class & "><b>" & nhead & "</b></font></a>",temp1 & kong,3,0,5)
End Sub

Sub main_news(n_jt,n_num,c_num,et,timt)
    Dim temp1
    Dim topic
    Dim oid
    Dim ooid
    Dim pic
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql       = "select top " & n_num & " id,topic,tim,counter from news where hidden=1 order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='news_view.asp?id=" & rs("id") & "'" & atb & " title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>"
        If et >  - 1 Then temp1 = temp1 & format_end(et,"<font class=gray>" & time_type(rs("tim"),timt) & "�������</font><font class=blue>" & rs("counter") & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    sql      = "select top 1 id,topic,pic from news where hidden=1 and istop=1 and ispic=1 order by id desc"
    Set rs   = conn.execute(sql)

    If Not(rs.eof) Then
        ooid = rs("id")
        If et = 1 Then oid = ooid
        pic  = "<table border=0><tr><td align=center valign=middle><a href='news_view.asp?id=" & ooid & "' target='_blank'><img src='" & url_true(web_var(web_down,5),rs("pic")) & "' border='0' width=" & web_var(web_num,8) & " height=" & web_var(web_num,7) & "></a></td></tr><tr><td align=center><a href='news_view.asp?id=" & ooid & "' target='_blank'>" & code_html(rs("topic"),1,c_num) & "</a></td></tr></table>"
    End If

    rs.Close
    temp1 = "<table width='100%'  border='0' cellspacing='0' cellpadding='0'><tr><td width='40%' align=center>" & pic & "</td><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td>" & temp1 & "</table></td></tr></table>"

    Response.Write format_barc("<a href='news.asp'><b><font class=end>�������</font></b></a>",temp1 & kong,3,0,7)
End Sub

Sub main_article(n_jt,n_num,c_num,et,timt)
    Dim temp1
    Dim topic
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql       = "select top " & n_num & " id,topic,tim,counter from article where hidden=1 order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='article_view.asp?id=" & rs("id") & "'" & atb & " title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>"
        If et >  - 1 Then temp1 = temp1 & format_end(et,"<font class=gray>" & time_type(rs("tim"),timt) & "�������</font><font class=blue>" & rs("counter") & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write format_barc("<a href='article.asp'><b><font class=end>��������</font></b></a>",temp1 & kong,3,0,2)
End Sub

Sub main_video(v_jt,v_num,vv_num,vet,vtimt)
    Dim temp1
    Dim topic
    If v_jt <> "" Then v_jt = img_small(v_jt)
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql       = "select top " & v_num & " * from gallery where types='film' and hidden=1 order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("name")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & v_jt & "<a href='gallery.asp?types=view&action=film&id=" & rs("id") & "'" & atb & " title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,vv_num) & "</a>"
        If vet >  - 1 Then temp1 = temp1 & format_end(vet,"<font class=gray>" & time_type(rs("tim"),vtimt) & "�������</font><font class=blue>" & rs("counter") & "</font>")
        temp1 = temp1 & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write format_barc("<a href='vouch.asp'><b><font class=end>������Ƶ</font></b></a>",temp1 & kong,3,0,4)
End Sub

Sub main_forum(n_jt,n_num,c_num,et,timt)
    Dim temp1
    Dim topic
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf>"
    sql       = "select top " & n_num & " forum_id,id,topic,tim,counter,re_counter from bbs_topic order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        temp1 = temp1 & vbcrlf & "<tr><td height=" & space_mod & ">" & n_jt & "<a href='forum_view.asp?forum_id=" & rs("forum_id") & "&view_id=" & rs("id") & "'" & atb & " title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=gray>" & time_type(rs("tim"),timt) & "��</font><font class=blue>" & rs("re_counter") & "</font>|<font class=blue>" & rs("counter") & "</font>") & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write format_barc("<a href='forum.asp'><b><font class=end>��̳����</font></b></a>",temp1 & kong,3,0,3)
End Sub

Sub main_update_view()
    Dim temp1
    Dim topic
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2 class=tf><tr><td></td><td  width='90%'>" & ukong & "</td><td></td></tr>"
    sql       = "select top 1 * from bbs_cast where sort='news' order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        temp1 = temp1 & vbcrlf & "<tr><td></td><td align=center background='images/" & web_var(web_config,5) & "/roll_bg.gif'><marquee scrollamount=2 direction=up width='100%' height='120' onMouseOver=this.stop() onMouseOut=this.start()><center><a href='update.asp?action=news&id=" & rs("id") & "'>" & code_html(topic,1,0) & "</center></a>" & code_jk(rs("word")) & "</marquee></td><td></td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write format_barc("<a href='update.asp'><b><font class=end>��վ����</font></b></a>",temp1,2,0,1)
End Sub

Sub main_update(n_jt,n_num,c_num,et,timt)
    Dim temp1
    Dim topic
    If n_jt <> "" Then n_jt = img_small(n_jt)
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=2>"
    sql       = "select top " & n_num & " id,topic,tim from bbs_cast where sort='news' order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        temp1 = temp1 & vbcrlf & "<tr align=center><td height=" & space_mod & ">" & n_jt & "<a href='update.asp?action=news&id=" & rs("id") & "' title='" & code_html(topic,1,0) & "'>" & code_html(topic,1,c_num) & "</a>" & format_end(et,"<font class=gray>" & time_type(rs("tim"),timt) & "</font>") & "</td></tr>"
        rs.movenext
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write kong & format_bar("<a href='update.asp'><b><font class=end>��վ����</font></b></a>",temp1,sk_bar,0,0,"|" & web_var(web_color,1) & "|80","") & "</td>"
End Sub

Sub main_shop()
    Dim sql
    Dim rs %><table border=0 width='98%' cellspacing=0 cellpadding=0><tr valign=top><%
    sql    = "select top 3 id,name,serial,brand,stock,smallimg,price_1,price_2,remark_1 from product where isgood=1 and hidden=1 order by tim desc,id desc"
    Set rs = conn.execute(sql)

    Do While Not rs.eof
        Response.Write vbcrlf & "<td width='33%' align=center>" & shop_view() & "</td>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing %></tr></table><%
End Sub

Sub links_main(lt,nummer)
    Dim temp1
    Dim nname
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=0>"
    sql       = "select * from links where sort='" & lt & "' and hidden=1 order by orders"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & vbcrlf & "<tr align=center height=40>"

        For i = 1 To nummer
            If rs.eof Then Exit For
            temp1     = temp1 & vbcrlf & "<td width='16%'><a href='" & rs("url") & "' target=_blank>"

            If lt = "txt" Then
                temp1 = temp1 & code_html(rs("nname"),1,0)
            Else
                temp1 = temp1 & "<img src='" & rs("pic") & "' border=0 width=88 height=31 title='" & code_html(rs("nname"),1,0) & "'>"
            End If

            temp1     = temp1 & "</a></td>"
            rs.movenext
        Next

        temp1 = temp1 & vbcrlf & "</tr>"
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write temp1
End Sub

Sub links_maini(lt,nummer)
    Dim temp1
    Dim nname
    temp1     = "<table border=0 width='100%' cellspacing=0 cellpadding=0>"
    sql       = "select * from links where sort='" & lt & "' and hidden=1 order by orders"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & vbcrlf & "<tr align=center height=40>"

        For i = 1 To nummer
            If rs.eof Then Exit For
            temp1     = temp1 & vbcrlf & "<td width='16%'><a href='" & rs("url") & "' target=_blank>"

            If lt = "txt" Then
                temp1 = temp1 & code_html(rs("nname"),1,0)
            Else
                temp1 = temp1 & "<img src='" & rs("pic") & "' border=0 width=88 height=31 title='" & code_html(rs("nname"),1,0) & "'>"
            End If

            temp1     = temp1 & "</a></td>"
            rs.movenext
        Next

        temp1 = temp1 & vbcrlf & "</tr>"
    Loop

    rs.Close
    temp1 = temp1 & "</table>"
    Response.Write format_barc("<a href=links.asp><font class=end><b>����վ��</b></font></a>",temp1,1,1,4)
End Sub

Sub web_search(wt)
    Dim temp11
    If wt = 1 Then Response.Write ukong
    temp11 = "<table border=0><form action='search.asp' method=get><tr><td>" & img_small("new") & "�ؼ��ʣ�</td><td>&nbsp;<input type=text name=keyword value='' size=20 maxlength=20></td><td>&nbsp;&nbsp;<select name=sea_type><option value='down'>����</option><option value='forum'>��̳</option><option value='news'>����</option><option value='article'>����</option><option value='paste'>��ֽ</option><option value='flash'>Flash</option><option value='website'>��վ</option></select>&nbsp;&nbsp;</td><td><input type=checkbox name=celerity value='yes' ></td><td>��������&nbsp;&nbsp;</td><td valign=top><input type=image src='images/small/web_sea.gif' border=0></td><td>&nbsp;&nbsp;<a href='search.asp?action=help' title='�๦������'>��������</a></td></tr></form></table>"
    Response.Write format_barc("<font class=end><b>վ������</b></font>",temp11,1,1,3)

End Sub

Sub news_fpic(d_num,t_num,w_num,nft)
    Dim temp1
    Dim topic
    Dim pic
    Dim tt
    Dim rs
    Dim sql
    Dim wnum
    Dim ispic
    Dim ooid
    wnum     = w_num
    sql      = "select top 1 id,topic,pic from news where hidden=1 and istop=1 and ispic=1 order by id desc"
    Set rs   = conn.execute(sql)

    If Not(rs.eof) Then
        ooid = rs("id")
        If nft = 1 Then oid = ooid
        pic  = "<table border=0><tr><td align=center><a href='news_view.asp?id=" & ooid & "'><img src='" & url_true(web_var(web_upload,1),rs("pic")) & "' border='0' width=" & web_var(web_num,7) & " height=" & web_var(web_num,8) & "></a></td></tr><tr><td align=center><a href='news_view.asp?id=" & ooid & "'>" & code_html(rs("topic"),1,d_num) & "</a></td></tr></table>"
    End If

    rs.Close
    sql       = "select top " & t_num & " id,topic,pic,ispic from news where hidden=1 and istop=1"
    If ooid <> 0 Then sql = sql & " and id<>" & ooid
    sql       = sql & " order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        topic = rs("topic")
        If rs("ispic") = True Then wnum = wnum - 2:ispic = sk_img
        temp1 = temp1 & "<tr><td height=" & space_mod & ">" & img_small("jt0") & "<a href='news_view.asp?id=" & rs("id") & "'>" & code_html(topic,1,wnum) & "</a>" & ispic & "</td></tr>"
        rs.movenext
    Loop

    rs.Close:Set rs = Nothing
    temp1 = "<table border=0 width='100%'><tr>" & _
    "<td width='30%' align=center>" & pic & "</td>" & _
    "<td width='70%'><table border=0>" & temp1 & "</table></td>" & _
    "</tr></table>"
    Response.Write kong & format_bar("<font class=" & sk_class & "><b>����ͼƬ����</b></font>",temp1,sk_bar,0,0,"||","")
End Sub

Sub main_best()
    Dim temp1
    Dim nid
    Dim n_jt
    Dim name
    Dim tim
    Dim counter
    Dim nhead
    n_jt  = img_small("jt0")
    temp1 = vbcrlf & "<table border=0 width='100%' cellspacing=0 cellpadding=0><tr valign=top>"

    sql       = "select top 1 id,name,counter,spic,tim,username,remark from gallery where hidden=1 and istop=1 and types='film' order by counter desc,id desc"
    nhead     = "�����Ƽ�"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & vbcrlf & "<td class=bw width='33%'>" & kong & "&nbsp;&nbsp;�Ƽ���Ƶ��<br>&nbsp;&nbsp;<a href='gallery.asp?types=view&action=film&id=" & rs("id") & "' title=''  target='_blank'><img src='" & url_true("images/video",rs("spic")) & "' border='0' width=150 height=150></a>"
        temp1 = temp1 & "<br>&nbsp;&nbsp;" & n_jt & "��Ƶ���ƣ�" & code_html(rs("name"),1,0) & "<br>&nbsp;&nbsp;" & n_jt & "�� �� �ˣ�" & rs("username") & "<br>&nbsp;&nbsp;" & n_jt & "���������" & rs("counter") & "<br>&nbsp;&nbsp;" & img_small("jt12") & Left(code_jk(rs("remark")),12) & "����"
        temp1 = temp1 & "</td>"
        rs.movenext
    Loop

    rs.Close
    temp1     = temp1 & "<td width=1 bgcolor=" & web_var(web_color,3) & "></td>"

    sql       = "select top 1 s_id,c_id,s_name,pic,intro from jk_sort where istop=1 order by s_id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & "<td align=left class=bw  >" & kong & "&nbsp;&nbsp;�Ƽ�ר����<br>&nbsp;&nbsp;<a href='down_list.asp?c_id=" & rs("c_id") & "&s_id=" & rs("s_id") & "' target='_blank'><img src=images/down/" & rs("pic") & ".jpg width=150 height=150 border=0></a>"
        temp1 = temp1 & "<br>&nbsp;&nbsp;" & n_jt & rs("s_name") & "<table width=180 border=0  cellspacing=0 cellpadding=0><tr><td>&nbsp;&nbsp;</td><td>" & img_small("jt12") & Left(rs("intro"),45) & "����</td></tr></table>"
        temp1 = temp1 & "</td>"
        rs.movenext
    Loop

    rs.Close
    temp1     = temp1 & "<td width=1 bgcolor=" & web_var(web_color,3) & "></td>"

    sql       = "select top 1 * from down where hidden=1 and types=5 order by id desc"
    Set rs    = conn.execute(sql)

    Do While Not rs.eof
        temp1 = temp1 & "<td align=left class=bw  width='33%'>" & kong & "&nbsp;&nbsp;�Ƽ����֣�<br>&nbsp;&nbsp;<a href='down_view.asp?id=" & rs("id") & "'  target='_blank'><img src=images/down/" & rs("pic") & " width=150 height=150 border=0></a>"
        temp1 = temp1 & "<br>&nbsp;&nbsp;" & n_jt & "�������ƣ�" & code_html(rs("name"),1,0) & "<br>&nbsp;&nbsp;" & n_jt & "�� �� �ˣ�" & rs("username") & "<br>&nbsp;&nbsp;" & n_jt & "���������" & rs("counter") & "<br>&nbsp;&nbsp;" & img_small("jt12") & Left(code_jk(rs("remark")),12) & "����"
        temp1 = temp1 & "</td>"
        rs.movenext
    Loop

    rs.Close

    temp1 = temp1 & vbcrlf & "</tr></table>"
    Response.Write format_barc("<a href=down.asp><font class=" & sk_class & "><b>" & nhead & "</b></font></a>",temp1 & kong,3,0,5)
End Sub %>