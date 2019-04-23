<!-- #include file="include/config.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

dim forumid,viewid,html_temp,forumtype
html_temp=""
forumid=trim(request.querystring("forum_id"))
viewid=trim(request.querystring("view_id"))
if not(isnumeric(forumid)) or not(isnumeric(viewid)) then
  html_temp="<tr><td><font class=red_2>您的操作有误：ID 出錯（1）！</font></td></tr>"
end if

if var_null(login_username)="" or var_null(login_password)="" then
  html_temp="<tr><td>&nbsp;&nbsp;"&web_var(web_error,2)&"</td></tr>"
else
  sql="select forum_type from bbs_forum where forum_id="&forumid&" and forum_hidden=0"
  set rs=conn.execute(sql)
  if rs.eof and rs.eof then
    html_temp=html_temp&"<tr><td><font class=red_2>ForumID 出錯！</font></td></tr>"
  end if
  rs.close:set rs=nothing

  if html_temp="" then
    sql="select username,word from bbs_data where forum_id="&forumid&" and reply_id="&viewid&" order by id desc"
    set rs=conn.execute(sql)
    if rs.eof and rs.bof then
      html_temp=html_temp&"<tr><td><font class=red_2>您的操作有误：ID 出錯（2）！</font></td></tr>"
    else
      do while not rs.eof
        html_temp=html_temp&"<tr><td><img src=""images/small/fk_minus.gif"" border=0> "&code_html(rs("word"),1,45)&"&nbsp;<font class=gray>-</font>&nbsp;"&replace(format_user_view(rs("username"),1,0),"'","""")&"</td></tr>"
        rs.movenext
      loop
    end if
    rs.close:set rs=nothing
  end if
end if

call close_conn()

html_temp="<table border=0 width=99% align=right cellspacing=2 cellpadding=0>"&html_temp&"</table>"
%>
<script language=javascript>
parent.followTd<%=viewid%>.innerHTML='<%=html_temp%>';
</script>