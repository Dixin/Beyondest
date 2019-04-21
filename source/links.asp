<!-- #include file="include/config_other.asp" -->
<!-- #include file="include/conn.asp" -->
<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

Dim nummer:nummer = 5
tit = "友情链接"

Call web_head(0,0,0,0,0)
'------------------------------------left----------------------------------
Call format_login()
Call links_left()
'response.write left_action("jt13",4)
'----------------------------------left end--------------------------------
Call web_center(0)
'-----------------------------------center---------------------------------
Response.Write ukong & table1
Response.Write vbcrlf & "<tr" & table2 & "><td class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>合作站点</b></td></tr><tr" & table3 & "><td align=center>" %>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><% Call links_main("fir",nummer) %></td></tr>
  </table>
<%
Response.Write vbcrlf & "</td></tr>"
Response.Write vbcrlf & "<tr" & table2 & "><td class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>友情链接</b></td></tr><tr" & table3 & "><td align=center>" %>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><% Call links_main("sec",nummer) %></td></tr>
  </table>
<%
Response.Write vbcrlf & "</td></tr>"
Response.Write vbcrlf & "<tr" & table2 & "><td class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>文字链接</b></td></tr><tr" & table3 & "><td align=center>" %>
  <table border=0 width='100%' cellspacing=0 cellpadding=0>
  <tr><td align=center><% Call links_main("txt",nummer) %></td></tr>
  </table>
<%
Response.Write vbcrlf & "</td></tr>"
Response.Write vbcrlf & "<tr" & table2 & "><td class=end background=images/" & web_var(web_config,5) & "/bar_3_bg.gif>&nbsp;" & img_small(us) & "&nbsp;&nbsp;<b>链接说明</b></td></tr><tr" & table3 & "><td align=center>" %>
<table border=0 width=450>
<tr><td>
<table border=0><tr><td class=htd>1、站点必须内容详实，制作马虎或名不副实者恕不链接。<br>
2、含有修改浏览者默认页或注册表的网页不予链接。<br>
3、含有任何反动内容或色情内容的不予链接。</td></tr></table>
</td></tr>
</table>
<%
Response.Write vbcrlf & "</td></tr></table><br>"
'---------------------------------center end-------------------------------
Call web_end(0)

Sub links_left()
    tit = vbcrlf & "<table border=0 width='100%' cellpadding=0 cellspacing=0>" & _
    vbcrlf & "<tr><td align=center height=50><a href='" & web_var(web_config,2) & "' target=_blank><img border=0 src='images/" & web_var(web_config,5) & "/logo.gif' width=120 heigh=40 alt='" & web_var(web_config,1) & "'></a></td></tr>" & _
    vbcrlf & "<tr><td align=center valign=top><textarea name=flink_main rows=5 cols=21  onfocus=this.select() onmouseover=this.focus()><a href='" & web_var(web_config,2) & "' target=_blank><img border=0 src='" & web_var(web_config,2) & "images/" & web_var(web_config,5) & "/logo.gif' width=88 heigh=31 alt='" & web_var(web_config,1) & "'></a></textarea></td></tr>" & _
    vbcrlf & "</table>"
    Call left_type(tit,"links",1)
End Sub %>