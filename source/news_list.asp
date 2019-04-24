<!-- #include file="INCLUDE/config_news.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer,page,rssum,thepages,viewpage,pageurl,keyword,sea_type,sqladd2
tit    = "新闻分类浏览"
Call cid_sid()
nummer = Int(web_var(web_num,2))

Call web_head(0,0,1,0,0)
'-----------------------------------center--------------------------------- %>
<table border=0 width='100%' cellspacing=0 cellpadding=0 align=center>
<tr><td align=center><%
Response.Write format_img("rnewslist.jpg") & gang
sqladd = ""

If cid = 0 And action <> "more" Then
    Call news_main("jt0",16,20,1,6,3,2,10,1)
Else
    If cid > 0 Then sqladd = " and c_id=" & cid
    If sid > 0 Then sqladd = sqladd & " and s_id=" & sid
    sqladd2 = sqladd

    If action = "more" Then
        Call news_more("jt0",35,1,6,3,5,10,1)
    Else

        If sid = 0 Then
            Call news_list("jt0",10,20,1,6,3,1,10,1)
        Else
            Call news_list("jt0",30,20,1,6,3,5,10,1)
        End If

    End If

End If %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
<tr><td align=center><% Call news_class_sort(cid,sid) %></td><td width=1 bgcolor=<% = web_var(web_color,3) %>></td></tr>
</table>
<%
'---------------------------------center end-------------------------------
Call web_center(1)
'------------------------------------right---------------------------------
Call format_login()
Call news_sea()
Call news_scroll("jt0","",3,15,1)
Call news_new_hot("jt0",sqladd2,"hot",10,12,1,6,0)
Call news_picr(sqladd2,1,10,6)
'----------------------------------right end-------------------------------
Call web_end(0) %>