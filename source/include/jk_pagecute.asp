<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================

function jk_pagecute(maxpage,thepages,viewpage,pageurl,pp,font_color)
  dim pn,pi,page_num,ppp,pl,pr:pi=1
  ppp=pp\2
  if pp mod 2 = 0 then ppp=ppp-1
  pl=viewpage-ppp
  pr=pl+pp-1
  if pl<1 then
    pr=pr-pl+1:pl=1
    if pr>thepages then pr=thepages
  end if
  if pr>int(thepages) then
    pl=pl+thepages-pr:pr=thepages
    if pl<1 then pl=1
  end if
  if pl>1 then
    jk_pagecute=jk_pagecute&" <a href='"& pageurl &"' title='第一页'>[|<]</a> " & _
		" <a href='"& pageurl &"page="&pl-1&"' title='上一页'>[<]</a> "
  end if
  for pi=pl to pr
    if cint(viewpage)=cint(pi) then
      jk_pagecute=jk_pagecute&" <font color=" & font_color & ">[" & pi & "]</font> "
    else
      jk_pagecute=jk_pagecute&" <a href='"& pageurl &"page="& pi &"' title='第 " & pi & " 页'>[" & pi & "]</a> "
    end if
  next
  if pr<thepages then
    jk_pagecute=jk_pagecute&" <a href='"& pageurl &"page="&pi&"' title='后一页'>[>]</a> " & _
		   " <a href='"& pageurl &"page="& thepages &"' title='最后一页'>[>|]</a> "
  end if
end function
%>