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

dim types,tit2
types=trim(request.querystring("types"))
n_sort="gall"

if action="view" then if not(isnumeric(id)) then action="paste"
tit="��վ��ͼ"

select case action
case "logo"
  tit="����"
case "baner"
  tit="�������"
  tit2="���"
case "film"
  tit="������Ƶ"
  tit2="��Ƶ"
  if types="view" then tit="�����Ƶ"
case "flash"
  tit="FLASH"
  tit2="Flash"
  if types="view" then tit="���FLASH"
case else
  action="paste"
  tit="�����ֽ"
  tit2="��ֽ"
  if types="view" then tit="���ͼƬ"
end select
n_sort=action

call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
call format_login()
call vouch_left("jt12","jt1")
call vouch_skin(tit2&"����","<table border=0 width='100%' align=center><tr><td>"&nsort_left(n_sort,cid,sid,"?action="&action&"&",0)&"</td></tr></table>","",1)

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