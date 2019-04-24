<!-- #include file="INCLUDE/config_vouch.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="INCLUDE/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim types,tit2
types  = Trim(Request.querystring("types"))
n_sort = "gall"

If action = "view" Then If Not(IsNumeric(id)) Then action = "paste"
tit = "ÍøÕ¾ÌùÍ¼"

Select Case action
    Case "logo"
        tit  = "ÆäËû"
    Case "baner"
        tit  = "¾«²ÊÏà²á"
        tit2 = "Ïà²á"
    Case "film"
        tit  = "¾«²ÊÊÓÆµ"
        tit2 = "ÊÓÆµ"
        If types = "view" Then tit = "ä¯ÀÀÊÓÆµ"
    Case "flash"
        tit  = "FLASH"
        tit2 = "Flash"
        If types = "view" Then tit = "ä¯ÀÀFLASH"
    Case Else
        action = "paste"
        tit    = "×ÀÃæ±ÚÖ½"
        tit2   = "±ÚÖ½"
        If types = "view" Then tit = "ä¯ÀÀÍ¼Æ¬"
End Select

n_sort = action

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call vouch_left("jt12","jt1")
Call vouch_skin(tit2 & "·ÖÀà","<table border=0 width='100%' align=center><tr><td>" & nsort_left(n_sort,cid,sid,"?action=" & action & "&",0) & "</td></tr></table>","",1)

'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write "<table border=0 cellspacing=0 cellpadding=0 width='100%'><tr><td width=1 bgcolor=" & web_var(web_color,3) & "></td><td align=right>"

Select Case action
    Case "logo"
        Call gallery_main(action)
    Case "baner"
        Response.Write format_img("ralbum.jpg") & gang
        Call gallery_main(action)
    Case "film"
        Response.Write format_img("rmtv.jpg") & gang

        If types = "view" Then
            Call gallery_view()
        Else
            Call gallery_main(action)
        End If

    Case "flash"
        Response.Write format_img("rflash.jpg") & gang

        If types = "view" Then
            Call gallery_view()
        Else
            Call gallery_main(action)
        End If

    Case Else
        Response.Write format_img("rdesktop.jpg") & gang

        If types = "view" Then
            Call gallery_view()
        Else
            Call gallery_main(action)
        End If

End Select

Response.Write "</td></tr></table>"
'---------------------------------center end-------------------------------
Call web_end(0) %>