<!-- #include file="include/config_vouch.asp" -->
<!-- #include file="include/config_review.asp" -->
<!-- #include file="include/jk_pagecute.asp" -->
<!-- #include file="include/jk_ubb.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V4.6 Demo版
' 
'           http://beyondest.com
' 
'*******************************************************************

dim types,tit2
types=trim(request.querystring("types"))
n_sort="gall"

if action="view" then if not(isnumeric(id)) then action="paste"
tit="网站贴图"

select case action
case "logo"
  tit="其他"
case "baner"
  tit="精彩相册"
  tit2="相册"
case "film"
  tit="精彩视频"
  tit2="视频"
  if types="view" then tit="浏览视频"
case "flash"
  tit="FLASH"
  tit2="Flash"
  if types="view" then tit="浏览FLASH"
case else
  action="paste"
  tit="桌面壁纸"
  tit2="壁纸"
  if types="view" then tit="浏览图片"
end select
n_sort=action

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call vouch_left("jt12","jt1")
call vouch_skin(tit2&"分类","<table border=0 width='100%' align=center><tr><td>"&nsort_left(n_sort,cid,sid,"?action="&action&"&",0)&"</td></tr></table>","",1)

'----------------------------------left end--------------------------------
call web_center(0)
'-----------------------------------center---------------------------------
response.write "<table border=0 cellspacing=0 cellpadding=0 width='100%'><tr><td width=1 bgcolor="&web_var(web_color,3)&"></td><td align=right>"
select case action
case "logo"
  call gallery_main(action)
case "baner"
  response.write format_img("ralbum.jpg")&gang
  call gallery_main(action)
case "film"
    response.write format_img("rmtv.jpg")&gang
  if types="view" then
    call gallery_view()
  else
    call gallery_main(action)
  end if
case "flash"
    response.write format_img("rflash.jpg")&gang
  if types="view" then
    call gallery_view()
  else
    call gallery_main(action)
  end if
case else
  response.write format_img("rdesktop.jpg")&gang
  if types="view" then
    call gallery_view()
  else
    call gallery_main(action)
  end if
end select

response.write "</td></tr></table>"
'---------------------------------center end-------------------------------
call web_end(0)



%>