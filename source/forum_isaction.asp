<!-- #include file="INCLUDE/config_forum.asp" -->
<% if not(isnumeric(forumid)) then call cookies_type("view_id") %>
<!-- #include file="INCLUDE/config_upload.asp" -->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

call forum_first()
call web_head(2,2,0,0,0)
if action="istops" then
  if format_user_power(login_username,login_mode,"")<>"yes" then close_conn():call cookies_type("power")
else
  if format_user_power(login_username,login_mode,forumpower)<>"yes" then close_conn():call cookies_type("power")
end if

dim isaction,delid
isaction=trim(request.querystring("isaction"))

select case isaction
case "del"
  call is_del()
case "delete"
  call is_delete()
case else
  call is_action()
end select

call close_conn()

sub is_action()
  if not(isnumeric(viewid)) and (action<>"isgood" and action<>"islock" and action<>"istop" and action<>"istops") then
    call close_conn()
    call cookies_type("del_id")
  end if
  
  dim ismsg,ist,upss
  select case action
  case "isgood"
    ist="精华"
  case "islock"
    ist="锁定"
  case "istop"
    ist="固顶"
  case "istops"
    ist="总固顶"
  end select
  if trim(request.querystring("cancel"))="yes" then
    if action="istops" then  action="istop"
    upss=0
    ismsg="已成功的对主题（ID："&viewid&"）取消 "&ist&" ！"
  else
    if action="istops" then
      action="istop"
      upss=2
    else
      upss=1
    end if
    ismsg="已成功的将主题（ID："&viewid&"）设为 "&ist&" ！"
  end if

  sql="update bbs_topic set "&action&"="&upss&" where id="&viewid
  conn.execute(sql)

  response.write "<script language=javascript>" & _
		 vbcrlf & "alert("""&ismsg&"\n\n点击返回。"");" & _
		 vbcrlf & "location='forum_list.asp?forum_id="&forumid&"'" & _
		 vbcrlf & "</script>"
  'response.redirect "forum_list.asp?forum_id="&forumid
end sub

sub is_del()
  delid=trim(request.querystring("del_id"))
  if not(isnumeric(delid)) then
    call close_conn()
    call cookies_type("del_id")
  end if
  
  dim reid,username
  sql="select reply_id,username from bbs_data where forum_id="&forumid&" and id="&delid
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    call close_conn()
    call cookies_type("del_id")
  end if
  reid=rs("reply_id")
  username=rs("username")
  rs.close:set rs=nothing

  sql="delete from bbs_data where id="&delid
  conn.execute(sql)
  sql="update bbs_topic set re_counter=re_counter-1 where id="&reid
  conn.execute(sql)
  sql="update bbs_forum set forum_data_num=forum_data_num-1 where forum_id="&forumid
  conn.execute(sql)
  sql="update configs set num_data=num_data-1 where id=1"
  conn.execute(sql)
  sql="update user_data set bbs_counter=bbs_counter-1,integral=integral-2 where username='"&username&"'"
  conn.execute(sql)

  response.write "<script language=javascript>" & _
		 vbcrlf & "alert(""成功删除了一条回贴！\n\n点击返回。"");" & _
		 vbcrlf & "location='forum_list.asp?forum_id="&forumid&"'" & _
		 vbcrlf & "</script>"
end sub

sub is_delete()
  delid=trim(request("del_id"))
  if len(delid)<1 then
    call close_conn()
    call cookies_type("del_id")
  end if
  
  dim del_dim,del_num,i,del_true,iok,ifail
  iok=0:ifail=0
  delid=replace(delid," ","")
  del_dim=split(delid,",")
  del_num=UBound(del_dim)
  for i=0 to del_num
    del_true=forum_delete(del_dim(i))
    call upload_del(index_url,del_dim(i))
    if del_true="yes" then
      iok=iok+1
    else
      ifail=ifail+1
    end if
  next
  erase del_dim
  response.write "<script language=javascript>" & _
		 vbcrlf & "alert(""成功删除了 "&iok&" 条贴子及其回贴！\n删除失败 "&ifail&" 条！\n\n点击返回。"");" & _
		 vbcrlf & "location='forum_list.asp?forum_id="&forumid&"'" & _
		 vbcrlf & "</script>"
end sub

function forum_delete(did)
  dim username,numd,sqladd
  did=trim(did)
  numd=1:sqladd=""
  forum_delete="yes"
  sql="select username from bbs_topic where forum_id="&forumid&" and id="&did
  set rs=conn.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    forum_delete="no":exit function
  end if
  username=rs("username")
  rs.close
  
  sql="update user_data set bbs_counter=bbs_counter-1,integral=integral-3 where username='"&username&"'"
  conn.execute(sql)
  
  sql="select count(id) from bbs_data where forum_id="&forumid&" and reply_id="&did
  set rs=conn.execute(sql)
  numd=rs(0)
  rs.close:set rs=nothing
  
  sql="delete from bbs_data where reply_id="&did
  conn.execute(sql)
  sql="delete from bbs_topic where id="&did
  conn.execute(sql)
  sql="update bbs_forum set forum_topic_num=forum_topic_num-1,forum_data_num=forum_data_num-"&numd&" where forum_id="&forumid
  conn.execute(sql)
  sql="update configs set num_topic=num_topic-1,num_data=num_data-"&numd&" where id=1"
  conn.execute(sql)
end function
%>