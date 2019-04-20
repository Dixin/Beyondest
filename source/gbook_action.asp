<!--#include file="include/onlogin.asp"-->
<!-- #include file="include/conn.asp" -->
<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

dim html_temp,id,viewpage,affirm,reicon,reword,retim
id=trim(request.querystring("id"))
viewpage=trim(request.querystring("page"))
affirm=trim(request.form("affirm"))
tit="<a href='?action=reply&id="&id&"&page="&viewpage&"'>回复留言</a>&nbsp;┋&nbsp;<a href='?action=delete&id="&id&"&page="&viewpage&"'>删除留言</a>"

html_temp=header(-1,tit)

if action="delete" then
  html_temp=html_temp&"<script language=JavaScript><!--" & _
	  vbcrlf & "function confirm_delete()" & _
	  vbcrlf & "{" & _
	  vbcrlf & "if (confirm(""你确定要删除这条留言吗？"")){ return true; }" & _
	  vbcrlf & "return false;" & _
	  vbcrlf & "}" & _
	  vbcrlf & "//--></script>" & _
	  vbcrlf & "<table border=0 width=100% height=100% cellpadding=0 cellspacing=0>" & _
	  vbcrlf & "<tr><td width=100% align=center>" & _
	  vbcrlf & "  <table border=1 cellpadding=0 cellspacing=0 bordercolor=" & web_var(web_color,2) & " width=400>" & _
	  vbcrlf & "  <tr><td width=100% align=center>"
else
  html_temp=html_temp&"<script language=javascript><!--" & _
	  vbcrlf & "function check(reply_form)" & _
	  vbcrlf & "{" & _
	  vbcrlf & "if( reply_form.reply_word.innertext == """" )" & _
	  vbcrlf & "  { alert(""回复内容是必须要的。"");return false; }" & _
	  vbcrlf & "if (reply_form.reply_word.value.length > 10000)" & _
	  vbcrlf & "  { alert(""对不起，回复内容不能超过 10000 个字节！"");return false; }" & _
	  vbcrlf & "}" & _
	  vbcrlf & "function reset(reply_form)" & _
	  vbcrlf & "{" & _
	  vbcrlf & "  if (confirm(""该项操作要清除全部的内容，你确定要清除吗?"")){ return true; }" & _
	  vbcrlf & "  return false;" & _
	  vbcrlf & "}" & _
	  vbcrlf & "--></script>" & _
	  vbcrlf & "<table border=0 width=100% height=100% cellpadding=0 cellspacing=0>" & _
	  vbcrlf & "<tr><td width=100% align=center>" & _
	  vbcrlf & "<table border=1 cellpadding=0 cellspacing=0 bordercolor=" & web_var(web_color,2) & " width=506>" & _
	  vbcrlf & "<tr><td width=100% align=center>"
end if

if affirm="ok" then
  if action="delete" then
    conn.Execute("Delete from gb_data where ID = " & id)
    html_temp=html_temp & VbCrLf & "<br>已成功的删除ID为<font class=red>" & id & " </font>条留言！<br><br>" & VbCrLf & "<a href=gbook.asp?page=" & viewpage & ">返回留言簿</a><br><br>" & VbCrLf & "（系统将在 " & web_var(web_num,5) & " 秒钟后自动返回）<br><br>" & _
	      VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=gbook.asp?page=" & viewpage & "'>"
  else
    reicon=request.form("reply_icon")
    reword=request.form("reply_word")
    retim=now
    set rs=server.createobject("adodb.recordset")
    sql = "Select re_icon,re_word,re_tim from gb_data where ID=" & id 
    rs.open sql,conn,1,3
    if rs.eof and rs.bof then
      html_temp=html_temp & vbcrlf & "<p>操作错误，找不到序号为" & id & "的留言，因此不能进行回复操作！</p>" & closer
    else
      rs("re_icon")=reicon
      rs("re_word")=reword
      rs("re_tim")=retim
      rs.update
      rs.close:set rs=nothing
      html_temp=html_temp & vbcrlf & "<br>已成功的回复ID为<font class=red> " & id & " </font>条留言！" & VbCrLf & "<br><br><a href=gbook.asp?page=" & viewpage & ">返回留言簿</a><br><br>" & VbCrLf & "（系统将在 " & web_var(web_num,5) & " 秒钟后自动返回）<br><br>" & _
	        VbCrLf & "<meta http-equiv='refresh' content='" & web_var(web_num,5) & "; url=gbook.asp?page=" & viewpage & "'>"
    end if
  end if
else
  if action="delete" then
    html_temp=html_temp & VbCrLf & "<table border=0 width=100% cellpadding=0 cellspacing=0>" & _
	    vbcrlf & "<tr><td align=center bgcolor=" & web_var(web_color,2) & " height=20>" & _
	    vbcrlf & "<font class=end><b>删 除 留 言</b></font>" & _
	    vbcrlf & "</td></tr>" & _
	    vbcrlf & "<tr><form name=reply_form method=POST action='?action="&action&"&id="&id&"&page="&viewpage&"'>" & _
	    vbcrlf & "<input type=hidden name=affirm value='ok'>" & _
	    vbcrlf & "<td height=50 align=center>" & _
	    vbcrlf & "此操作将删除id为<font class=red> " & id & " </font>的留言！是否确定？！" & _
	    vbcrlf & "</td></tr>" & _
	    vbcrlf & "<tr><td align=center>" & _
	    vbcrlf & "<input type=submit value=' 确 定 ' onclick=""return confirm_delete()"">&nbsp;&nbsp;" & _
	    vbcrlf & "<input type=button value=' 取 消 ' onclick='javascript:history.back(1)'> " & _
	    vbcrlf & "</td></tr>"
  else
    html_temp=html_temp & vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width=550>" & _
	    vbcrlf & "<tr><form name=reply_form method=POST action='?action="&action&"&id="&id&"&page="&viewpage&"'>" & _
	    vbcrlf & "<input type=hidden name=affirm value='ok'>" & _
	    vbcrlf & "<td align=center colspan=2 bgcolor=" & web_var(web_color,2) & " height=20>" & _
	    vbcrlf & "<font class=end><b>回 复 留 言</b></font></td></tr>" & _
	    vbcrlf & "<tr><td height=15 colspan=2>  </td></tr>" & _
	    vbcrlf & "<tr><td align=center height=30 colspan=2>些操作将回复ID为<font class=red>  " & Request("id") & " </font>条留言</td></tr>" & _
	    vbcrlf & "<tr><td align=center width=80 height=10>　</td><td align=left width=440>  </td></tr>" & _
	    vbcrlf & "<tr><td align=center height=25>表情图标: </td><td align=left>" & _
	    vbcrlf & "<img border=0 src='images/icon/0.gif'>" & _
	    vbcrlf & "<input type=radio value=0 name=reply_icon checked class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/1.gif'>" & _
	    vbcrlf & "<input type=radio value=1 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/2.gif'>" & _
	    vbcrlf & "<input type=radio value=2 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/3.gif'>" & _
	    vbcrlf & "<input type=radio value=3 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/4.gif'>" & _
	    vbcrlf & "<input type=radio value=4 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/5.gif'>" & _
	    vbcrlf & "<input type=radio value=5 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/6.gif'>" & _
	    vbcrlf & "<input type=radio value=6 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/7.gif'>" & _
	    vbcrlf & "<input type=radio value=7 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/8.gif'>" & _
	    vbcrlf & "<input type=radio value=8 name=reply_icon class=bg_1>" & _
	    vbcrlf & "<img border=0 src='images/icon/9.gif'>" & _
	    vbcrlf & "<input type=radio value=9 name=reply_icon class=bg_1>" & _
	    vbcrlf & "</td></tr>" & _
	    vbcrlf & "<tr><td align=center valign=top><br>回复留言:<br><br>"&web_var(web_error,3)&"<br><br>留言<=10KB</td>" & _
	    vbcrlf & "<td align=left>" & _
	    vbcrlf & "<table border=0 cellpadding=0 cellspacing=0 width='100%'>" & _
	    vbcrlf & "<tr><td>" & _
	    vbcrlf & "<textarea rows=8 name=reply_word cols=70></textarea></td></tr></table>" & _
	    vbcrlf & "</td></tr>" & _
	    vbcrlf & "<tr><td align=center colspan=2></td></tr>" & _
	    vbcrlf & "<tr><td align=center colspan=2 height=50>" & _
	    vbcrlf & "<input type=submit value=' 点 击 回 复 ' onclick=""return check(this.form)"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & _
	    vbcrlf & "<input type=reset value=' 重 写 ' onclick=""return reset()"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type=button value=' 取 消 ' onclick='javascript:history.back(1)'>" & _
	    vbcrlf & "</td></form></tr>" & _
	    vbcrlf & "</table>"
  end if
end if
call close_conn()
response.write html_temp&ender()
%>