<%
'*******************************************************************

'

'                     Beyondest.Com V3.6 Demo版

' 




'           网址：http://www.beyondest.com

' 

'*******************************************************************

'call review_type("art",id,"article_view.asp?id="&id,1|2|3|4)

function review_rsort(rvar)
  if rvar<>"news" and rvar<>"art" and rvar<>"down" and rvar<>"gall" and rvar<>"web" and rvar<>"pro" and rvar<>"video" then
    review_rsort="no"
  else
    review_rsort="yes"
  end if
end function

sub review_type(rtsort,rtid,rturl,rtt)
  dim rsr,sqlr,rusername,remail,rtype
%>
<table border=0 width='100%' cellspacing=0 cellpadding=0>
<tr><td height=1 background='images/bg_dian.gif'></td></tr>
<tr><td height=25 valign=middle bgcolor=<%response.write web_var(web_color,5)%> class=red_3>&nbsp;&nbsp;<b>→&nbsp;<a onclick="document.all.review_word.style.display=document.all.review_word.style.display=='none'?'':'none';return false;" href="javsscript:;"><font class=red_3>评论内容</font></b> (点击查看)</a><%
  if login_mode=format_power2(1,1) then
    response.write "&nbsp;&nbsp;&nbsp;<a href='review.asp?action=delete&rsort="&rtsort&"&re_id="&rtid&"&rurl="&rturl&"'>删除本主题的所有评论</a>"
  end if
%></td></tr>
<tr><td height=1 background='images/bg_dian.gif'></td></tr>
<tr id=review_word<%if rtt<>1 then response.write " style=""display:none;"""%>><td align=center>
<table border=0 width='98%'>
<%
  sqlr="select top 10 rid,rusername,remail,rword,rtim,rtype from review where rsort='"&rtsort&"' and re_id="&rtid&" order by rid desc"
  set rsr=conn.execute(sqlr)
  if rsr.eof and rsr.bof then
    response.write "<tr><td class=gray>（没有相关评论）</td></tr>"
  else
    do while not rsr.eof
      rusername=rsr("rusername"):remail=rsr("remail"):rtype=int(rsr("rtype"))
      response.write vbcrlf&"<tr><td>姓名："
      if rtype=1 then
        response.write format_user_view(rusername,1,1)
      else
        response.write rusername
      end if
      if len(remail)>5 then
        response.write "&nbsp;&nbsp;E-mail：<a href='mailto:"&remail&"'>"&remail&"</a>"
      end if
      response.write "&nbsp;&nbsp;发表时间："&time_type(rsr("rtim"),88)
      if login_mode=format_power2(1,1) then
        response.write "&nbsp;&nbsp;&nbsp;<a href='review.asp?action=del&rsort="&rtsort&"&re_id="&rtid&"&rid="&rsr("rid")&"&rurl="&rturl&"'>删除这条评论</a>"
      end if
      response.write "</td></tr><tr><td>评论内容："&code_html(rsr("rword"),3,0)&"</td></tr><tr><td height=1 background='images/bg_dian.gif'></td></tr>"
      rsr.movenext
    loop
  end if
  rsr.close:set rsr=nothing
%>
</table>
</td></tr>
<tr><td height=2></td></tr>
<tr><td height=1 background='images/bg_dian.gif'></td></tr>
<tr><td height=25 valign=middle bgcolor=<%response.write web_var(web_color,5)%> class=red_3>&nbsp;&nbsp;<b>→&nbsp;<a onclick="document.all.review_add.style.display=document.all.review_add.style.display=='none'?'':'none';return false;" href="javsscript:;"><font class=red_3>发表我的评论</font></a></b></td></tr>
<tr><td height=1 background='images/bg_dian.gif'></td></tr>
<tr id=review_add><td align=center>
<table border=0 width='90%'>
<form action='review.asp' method=post>
<input type=hidden name=rsort value='<%response.write rtsort%>'>
<input type=hidden name=re_id value='<%response.write rtid%>'>
<input type=hidden name=rurl value='<%response.write rturl%>'>
<tr height=30><td>您的姓名：</td><td><input type=text name=rusername value='<%response.write login_username%>' size=16 maxlength=20>　　您的E-mail：<input type=text name=remail size=24 maxlength=20></td></tr>
<tr valign=top><td><br>评论内容：</td><td><textarea rows=5 cols=60 name=rword></textarea></td></tr>
<tr height=30><td>发表评论：</td><td><input type=submit value='点 击 发 表 评 论'>　　<input type=reset value='重新填写'></td></tr>
</form></table>
</td></tr>
<tr><td height=2></td></tr>
</table>
<%
end sub

sub font_word_js()
%>
<script language=JavaScript>
<!--
  function do_color(vobject,vvar)
  { document.getElementById(vobject).style.color=vvar; }
  function do_zooms(vobject,vvar)
  { document.getElementById(vobject).style.fontsize=vvar+'px'; }
-->
</script>
<%
end sub

sub font_word_action()
%>浏览选项：
<!--
<a href="javascript:;" onclick="javascript:do_zooms('font_word',16);">大</a>
<a href="javascript:;" onclick="javascript:do_zooms('font_word',14);">中</a>
<a href="javascript:;" onclick="javascript:do_zooms('font_word',12);">小</a>&nbsp;
-->
<select name=do_color_frm size=1 onchange="if(this.options[this.selectedIndex].value!=''){do_color('font_word',this.options[this.selectedIndex].value);}">
<option value=''>颜色</option>
<option value='#000000' style="color:#000000">默认</option>
<option value='#808080' style="color:#808080">灰度</option>
<option value='#808000' style="color:#808000">橄榄色</option>
<option value='#008000' style="color:#008000">绿色</option>
<option value='#0000FF' style="color:#0000FF">蓝色</option>
<option value='#800000' style="color:#800000">褐色</option>
<option value='#FF0000' style="color:#FF0000">红色</option>
</select>&nbsp;<%
end sub
' style="font-size:14px; line-height:150%;"
sub font_word_type(fvar)
%>
  <table border=0 cellpadding=0 cellspacing=0 width='100%'>
  <tr>
  <td width=22 height=1 background='images/main/view_line.gif'></td>
  <td bgcolor=#666666></td>
  <td width=1 rowspan=5 bgcolor=#666666></td>
  </tr>
  <tr><td width=22 height=5 background='images/main/view_b.gif'></td><td></td></tr>
  <tr>
  <td background='images/main/view_bg.gif'></td>
  <td align=center>
    <table border=0 width='98%' align=center class=tf>
    <tr><td width='100%' class=bw><font id="font_word" class=htd style="font-size:14px; font-family:宋体, Verdana, Arial, Helvetica, sans-serif;"><% response.write fvar %></font></td></tr>
    </table>
  </td>
  </tr>
  <tr><td width=22 height=5 background='images/main/view_b.gif'></td><td></td></tr>
  <tr>
  <td height=1 background='images/main/view_line.gif'></td>
  <td height=1 bgcolor=#666666></td>
  </tr>
  </table>
<%
end sub
%>