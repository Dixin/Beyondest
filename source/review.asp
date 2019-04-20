<!-- #include file="include/config.asp" -->
<!-- #include file="include/skin.asp" -->
<!-- #include file="INCLUDE/config_review.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim rsort,rurl,re_id,rerr:rerr="":sql=""
tit="发表评论"
call web_head(0,2,0,0,0)

select case action
case "delete"
  call review_delete()
case "del"
  call review_del()
case else
  call review_main()
end select

call close_conn()

sub review_delete()
  call review_d()
  on error resume next
  conn.execute(sql)
  if err then
    err.clear
    call review_err("意外的错误！请与站长联系。\nhttp://www.beyondest.com/\n")
    exit sub
  end if
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""已成功删除了主题（n_sort："&rsort&"， id："&re_id&"）的所有评论！\n\n点击返回..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_del()
  call review_d()
  dim rid:rid=trim(request.querystring("rid"))
  if not(isnumeric(rid)) then
    rerr=rerr&"删除评论的 RID 出错！\n"
  end if
  if rerr<>"" then call review_err(rerr):exit sub
  sql=sql&" and rid="&rid
  on error resume next
  conn.execute(sql)
  if err then
    err.clear
    call review_err("意外的错误！请与站长联系。\nhttp://www.beyondest.com/\n")
    exit sub
  end if
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""已成功删除了一条主题（n_sort："&rsort&"， id："&re_id&"）评论（rid："&rid&"）！\n\n点击返回..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_d()
  if login_mode<>format_power2(1,1) then
    call close_conn()
    call review_err("您没有删除评论的权限！！！\n")
    response.end
  end if
  rsort=trim(request.querystring("rsort"))
  re_id=trim(request.querystring("re_id"))
  rurl=trim(request.querystring("rurl"))
  if review_rsort(rsort)<>"yes" then
    rerr=rerr&"删除评论的类型出错！\n"
  end if
  if not(isnumeric(re_id)) then
    rerr=rerr&"删除评论的 ID 出错！\n"
  end if
  if rerr<>"" then
    call close_conn()
    call review_err(rerr)
    response.end
  end if
  sql="delete from review where rsort='"&rsort&"' and re_id="&re_id
end sub

sub review_main()
  dim rusername,remail,rword
  rusername=code_form(trim(request.form("rusername")))
  remail=code_form(trim(request.form("remail")))
  rword=code_form(trim(request.form("rword")))
  rsort=trim(request.form("rsort"))
  re_id=trim(request.form("re_id"))
  rurl=trim(request.form("rurl"))
  if review_rsort(rsort)<>"yes" then
    rerr=rerr&"发表评论的类型出错！！！\n"
  end if
  if not(isnumeric(re_id)) then
    rerr=rerr&"发表评论的 ID 出错！！！\n"
  end if
  if symbol_name(rusername)<>"yes" then
    rerr=rerr&"请输入您的名称！（不得含有非法字符）\n"
  end if
  if len(remail)>0 then
    if email_ok(remail)<>"yes" or len(remail)>50 then
      rerr=rerr&"您的 E-mail 不得含有非法字符！\n"
    end if
  end if
  if len(rword)<1 then
    rerr=rerr&"您没有输入的评论内容！\n"
  elseif len(rword)>250 then
    rerr=rerr&"您输入的评论内容太长！(<=250字节)\n"
  end if
  if rerr<>"" then call review_err(rerr):exit sub

  on error resume next
  sql="insert into review(rsort,re_id,rusername,remail,rword,rtim,rtype) values('"&rsort&"',"&re_id&",'"&rusername&"','"&remail&"','"&rword&"','"&now_time&"',"
  if rusername=login_username then
    sql=sql&"1"
  else
    sql=sql&"0"
  end if
  sql=sql&")"
  conn.execute(sql)
  if err then
    err.clear
    call review_err("意外的错误！请与站长联系。\nhttp://www.beyondest.com/\n")
    exit sub
  end if

  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""您成功的发表了有关您的评论！\n\n谢谢您的参与！点击返回..."");"
  if len(rurl)<5 then
    response.write vbcrlf&"location.href='main.asp';"
  else
    response.write vbcrlf&"location.href='"&rurl&"';"
  end if
  response.write vbcrlf&"--></script>"
end sub

sub review_err(revar)
  response.write vbcrlf&"<script lanuage=javascript><!--" & _
		 vbcrlf&"alert(""您在发表评论时出现如下错误：\n\n"&revar&"\n点击返回..."");" & _
		 vbcrlf&"history.back(-1);" & _
		 vbcrlf&"--></script>"
end sub
%>