<!-- #include file="include/config_vouch.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
'                     Beyondest.Com V4.6 Demo版
' 
' http://beyondest.com
' ====================

Dim types
Dim tit2
types  = Trim(Request.querystring("types"))
n_sort = "gall"

If action = "view" Then If Not(IsNumeric(id)) Then action = "paste"
tit = "网站贴图"

Select Case action
    Case "logo"
        tit  = "其他"
    Case "baner"
        tit  = "精彩相册"
        tit2 = "相册"
    Case "film"
        tit  = "精彩视频"
        tit2 = "视频"
        If types = "view" Then tit = "浏览视频"
    Case "flash"
        tit  = "FLASH"
        tit2 = "Flash"
        If types = "view" Then tit = "浏览FLASH"
    Case Else
        action = "paste"
        tit    = "桌面壁纸"
        tit2   = "壁纸"
        If types = "view" Then tit = "浏览图片"
End Select

n_sort = action

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call vouch_left("jt12","jt1")
Call vouch_skin(tit2 & "分类","<table border=0 width='100%' align=center><tr><td>" & nsort_left(n_sort,cid,sid,"?action=" & action & "&",0) & "</td></tr></table>","",1)

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