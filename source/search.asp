<!-- #include file="INCLUDE/config_other.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,page,rssum,thepages,viewpage,pageurl,sqladd,keyword,sea_type,sea_name,topic,topic2,sql1,sql2,linkurl,keywords,tims
pageurl = "?":sqladd = "":topic = "":sql1 = "":sql2 = "":linkurl = "":keywords = "":sea_name = "����"
nummer  = 20:viewpage = 1:thepages = 0
tit     = "վ������"

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
  <tr><td colspan=2 height=30>������û����д�������ؼ��֡�����鿴���°���˵����</td></tr>
  <tr><td width=10></td><td><% Response.Write img_small("jt1") %>������ʱ������д�������ؼ��֡���</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt12") %>��Ҫ��������ؼ�������<font class=red>�ո�</font>������ؼ��ָ������磺<font class=blue>V6&nbsp;���</font>��</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt0") %>���ؼ��֡��в��ܺ��е����ţ�'����</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt0") %>���ؼ��֡��к��еļӺţ�+��������Ϊ�ո���</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt13") %>������������ֻ�ڣ����š���������������Ч��</td></tr>
  <tr><td></td><td><% Response.Write img_small("jt14") %>ף����ʹ�ñ�վ�ġ�վ��������ʱ������졣</td></tr>
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
  <td width='6%'>���</td>
  <td width='94%'>������ݣ�����ѯ�Ĺؼ����ǣ�<% Response.Write keywords %>ÿҳ <font class=red><% Response.Write nummer %></font> �� <font class=blue><% Response.Write sea_name %></font> ��ѯ�����</td>
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
    <td>��&nbsp;<font class=red><% Response.Write rssum %></font>&nbsp;�����&nbsp;
ҳ�Σ�<font class=red><% Response.Write viewpage %></font>/<font class=red><% Response.Write thepages %></font>&nbsp;
��ҳ��<% Response.Write jk_pagecute(nummer,thepages,viewpage,pageurl,3,"#ff0000") %></td>
    <td align=right><font size=2 class=gray>��ѯ��ʱ��<font class=red_3><% Response.Write FormatNumber((timer() - tims)*1000,3) %></font> ����</font></td>
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
                sea_name = "����"
                If sea_type = "article" Then sea_name = "����"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "down"
                topic = "name":topic2 = topic
                If celerity = "yes" Then topic2 = "keyes"
                linkurl = sea_type & "_view.asp?id="
                sea_name = "���"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "website"
                topic = "name":topic2 = topic
                linkurl = sea_type & ".asp?action=view&id="
                sea_name = "��վ"
                sql1 = "select id," & topic & ",username,tim from " & sea_type & " where hidden=1"
                sql2 = " order by id desc"
            Case "paste","flash"
                topic = "name":topic2 = topic
                linkurl = "gallery.asp?action=" & sea_type & "&types=view&id="
                sea_name = "ͼƬ"
                If sea_type = "flash" Then sea_name = "Flash"
                sql1 = "select id," & topic & ",username,tim from gallery where hidden=1 and types='" & sea_type & "'"
                sql2 = " order by id desc"
            Case Else
                sea_type = "forum"
                topic = "topic":topic2 = topic
                linkurl = "forum_view.asp?view_id="
                sea_name = "��̳"
                sql1 = "select id," & topic & ",username,tim,forum_id from bbs_topic where"
                sql2 = " order by id desc"
        End Select

        pageurl = pageurl & "sea_type=" & sea_type & "&"
    End Sub %>