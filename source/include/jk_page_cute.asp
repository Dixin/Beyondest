<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

function pagecute_fun(viewpage,thepages,pagecuteurl)
  dim re_color,pf0,pf1,pf2,pf3,pf4,pf5
  re_color="#c0c0c0"
  pf0="���ǵ�һҳ"
  pf1="��һҳ"
  pf2="��һҳ"
  pf3="��һҳ"
  pf4="���һҳ"
  pf5="�������һҳ"
  pagecute_fun=VbCrLf & "<table border=0 cellspacing=0 cellpadding=0><tr><form action='"&pagecuteurl&"' method=post><td>"
  if cint(viewpage)=1 then
    pagecute_fun=pagecute_fun & VbCrLf & "<font color="&re_color&">"&pf0&"</font>&nbsp;"
  else
    pagecute_fun=pagecute_fun & VbCrLf & "<a href='"&pagecuteurl&"page=1' alt='"&pf1&"'>"&pf1&"</a>��<a href='"&pagecuteurl&"page="&cint(viewpage)-1&"' alt='"&pf2&"'>"&pf2&"</a>&nbsp;"
  end if

  if cint(viewpage)=cint(thepages) then
    pagecute_fun=pagecute_fun & VbCrLf & "<font color="&re_color&" alt='"&pf5&"'>"&pf5&"</font>"
  else
    pagecute_fun=pagecute_fun & VbCrLf & "<a href='"&pagecuteurl&"page="&cint(viewpage)+1&"' alt='"&pf3&"'>"&pf3&"</a>��<a href='"&pagecuteurl&"page="&cint(thepages)&"' alt='"&pf4&"'>"&pf4&"</a>"
  end if
  if cint(thepages)<>1 then
    pagecute_fun=pagecute_fun & VbCrLf & "&nbsp;<input type=text name=page value='"&viewpage&"' size=2>&nbsp;<input type=submit value='GO'>"
  end if
  pagecute_fun=pagecute_fun & VbCrLf & "</td></form></tr></table>"
end function
%>