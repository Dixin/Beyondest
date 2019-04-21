<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Function code_jk(strers)
    Dim strer:strer = strers
    If strer = "" Or IsNull(strer) Then code_jk = "":Exit Function
    strer = health_var(strer,1)
    strer = Replace(strer,"<","&lt;")
    strer = Replace(strer,">","&gt;")
    strer = Replace(strer," ","&nbsp")		'空格
    strer = Replace(strer,Chr(9),"&nbsp")		'table
    strer = Replace(strer,"'","&#39;")		'单引号
    strer = Replace(strer,"""","&quot;")		'双引号
    Dim re
    Dim re_v
    re_v          = "[^\(\)\;\'\[]*"
    Set re        = new RegExp
    re.IgnoreCase = True
    re.Global     = True
    re.Pattern    = "(javascript:)"
    strer         = re.Replace(strer,"<i>javascript</i>:")
    re.Pattern    = "(javascript)"
    strer         = re.Replace(strer,"<i>&#106avascript</i>")
    re.Pattern    = "(jscript:)"
    strer         = re.Replace(strer,"<i>&#106script</i>:")
    re.Pattern    = "(js:)"
    strer         = re.Replace(strer,"<i>&#106s</i>:")
    re.Pattern    = "(value)"
    strer         = re.Replace(strer,"<i>&#118alue</i>")
    re.Pattern    = "(about:)"
    strer         = re.Replace(strer,"<i>about&#58</i>")
    re.Pattern    = "(file:)"
    strer         = re.Replace(strer,"<i>file&&#58</i>")
    re.Pattern    = "(document.)"
    strer         = re.Replace(strer,"<i>document</i>.")
    re.Pattern    = "(vbscript:)"
    strer         = re.Replace(strer,"<i>&#118bscript</i>:")
    re.Pattern    = "(vbs:)"
    strer         = re.Replace(strer,"<i>&#118bs</i>&#58")
    re.Pattern    = "(on(mouse|exit|error|click|key))"
    strer         = re.Replace(strer,"<i>&#111n$2</i>")

    re.Pattern    = "\[IMGS\](.[^\[]*(gif|jpg|jpeg|bmp|png))\[\/IMGS\]"
    strer         = re.Replace(strer,"<IMG SRC='$1' align=center border=0 onload=""javascript:if(this.width>max-width)this.width=max-width"">")
    re.Pattern    = "\[IMG\](.[^\[]*(gif|jpg|jpeg|bmp|png))\[\/IMG\]"
    strer         = re.Replace(strer,"<img src='images/small/image.gif' border=0 align=absMiddle width=16 height=16> <a href=$1 alt='按此在新窗口浏览图片' target=_blank>[ 相关贴图 ]</a><br><IMG SRC=$1 align=center border=0 onload=""javascript:if(this.width>max-width)this.width=max-width"">")

    re.Pattern    = "\[DIR=*([0-9]*),*([0-9]*)\](" & re_v & ")\[\/DIR]"
    strer         = re.Replace(strer,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
    re.Pattern    = "\[QT=*([0-9]*),*([0-9]*)\](" & re_v & ")\[\/QT]"
    strer         = re.Replace(strer,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#ededed kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
    re.Pattern    = "\[MP=*([0-9]*),*([0-9]*)\](" & re_v & ")\[\/MP]"
    strer         = re.Replace(strer,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
    re.Pattern    = "\[RM=*([0-9]*),*([0-9]*)\](" & re_v & ")\[\/RM]"
    strer         = re.Replace(strer,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
    re.Pattern    = "(\[FLASH=*([0-9]*),*([0-9]*)\])(" & re_v & "(.swf))(\[\/FLASH\])"
    strer         = re.Replace(strer,"<img src='images/small/flash.gif' border=0 align=absMiddle width=16 height=16> <a href='$4' TARGET=_blank>[ 全屏欣赏 ]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE='$4'><PARAM NAME=quality VALUE=high><embed src='$4' quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
    re.Pattern    = "(\[FLASHS=*([0-9]*),*([0-9]*)\])(" & re_v & "(.swf))(\[\/FLASHS\])"
    strer         = re.Replace(strer,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=max-width height=max-height><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=max-width height=max-height>$2</embed></OBJECT>")
    re.Pattern    = "(\[DOWNLOAD\])(" & re_v & ")(\[\/DOWNLOAD\])"
    strer         = re.Replace(strer,"<img src=images/small/download.gif border=0 align=absMiddle height=16 width=16> <a href=""$2"" TARGET=_blank>[ 点击下载 ]</a>")
    re.Pattern    = "(\[URL\])(" & re_v & ")(\[\/URL\])"
    strer         = re.Replace(strer,"<A HREF='$2' TARGET=_blank>$2</A>")
    re.Pattern    = "(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
    strer         = re.Replace(strer,"<A HREF='$2' TARGET=_blank>$3</A>")
    re.Pattern    = "(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
    strer         = re.Replace(strer,"<img align=absmiddle src=images/small/email.gif><A HREF=""mailto:$2"">$2</A>")
    re.Pattern    = "(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
    strer         = re.Replace(strer,"<img align=absmiddle src=images/small/email.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
    re.Pattern    = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "([^>='])(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"$1<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$2>$2</a>")
    re.Pattern    = "^(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "[^>='](ftp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "^(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "[^>='](rtsp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "^(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "[^>='](mms://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
    strer         = re.Replace(strer,"<img align=absmiddle src='images/small/url.gif'><a target=_blank href=$1>$1</a>")
    re.Pattern    = "\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
    strer         = re.Replace(strer,"<font color=$1>$2</font>")
    re.Pattern    = "\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
    strer         = re.Replace(strer,"<font face=$1>$2</font>")
    re.Pattern    = "\[size=([-3-7])\](.[^\[]*)\[\/size\]"
    strer         = re.Replace(strer,"<font size=$1>$2</font>")
    re.Pattern    = "\[align=(.[^\[]*)\](.[^\[]*)\[\/align\]"
    strer         = re.Replace(strer,"<div align=$1>$2</div>")
    re.Pattern    = "\[align=(.[^\[]*)\](.*)\[\/align\]"
    strer         = re.Replace(strer,"<div align=$1>$2</div>")
    re.Pattern    = "\[center\](.[^\[]*)\[\/center\]"
    strer         = re.Replace(strer,"<div align=center>$1</div>")
    re.Pattern    = "\[QUOTE\](.*)\[\/QUOTE\]"
    strer         = re.Replace(strer,"<table border=1 cellspacing=0 cellpadding=4 width='98%' bordercolorlight=" & web_var(web_color,4) & " bordercolordark=" & web_var(web_color,1) & " bgcolor=" & web_var(web_color,1) & " style=""TABLE-LAYOUT: fixed"" align=center><tr><td style=""WORD-WRAP: break-word"">$1</td></tr></table><br>")
    're.Pattern="\[HTML\](.[^\[]*)\[\/HTML\]"
    'strer=re.Replace(strer,"<table width='100%' border='0' cellspacing='0' cellpadding='6' class='"&abgcolor&"'><td><b>以下内容为程序代码:</b><br>$1</td></table>")
    re.Pattern = "\[CODE\](.*)\[\/CODE\]"
    strer      = re.Replace(strer,"<table border=1 cellspacing=0 cellpadding=4 width='98%' bordercolorlight=" & web_var(web_color,4) & " bordercolordark=" & web_var(web_color,1) & " bgcolor=" & web_var(web_color,1) & " style=""TABLE-LAYOUT: fixed"" align=center><tr><td style=""WORD-WRAP: break-word"">$1</td></tr></table><br>")
    re.Pattern = "\[fly\](.[^\[]*)\[\/fly\]"
    strer      = re.Replace(strer,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
    re.Pattern = "\[move\](.[^\[]*)\[\/move\]"
    strer      = re.Replace(strer,"<MARQUEE scrollamount=3>$1</marquee>")
    re.Pattern = "\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
    strer      = re.Replace(strer,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
    re.Pattern = "\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
    strer      = re.Replace(strer,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
    re.Pattern = "\[i\](.[^\[]*)\[\/i\]"
    strer      = re.Replace(strer,"<i>$1</i>")
    re.Pattern = "\[u\](.[^\[]*)(\[\/u\])"
    strer      = re.Replace(strer,"<u>$1</u>")
    re.Pattern = "\[b\](.[^\[]*)(\[\/b\])"
    strer      = re.Replace(strer,"<b>$1</b>")
    Set re     = Nothing
    Dim emci

    For emci = 1 To 16
        strer = Replace(strer,"[em" & emci & "]","<image src='images/icon/em" & emci & ".gif' border=0>&nbsp;")
    Next

    For emci = 1 To 13
        strer = Replace(strer,"[emb" & emci & "]","<image src='images/icon/emb" & emci & ".gif' border=0>&nbsp;")
    Next

    strer     = Replace(strer,"[br]","<br>")
    strer     = Replace(strer,"[BR]","<br>")
    strer     = Replace(strer,vbCrLf,"<br>")
    strer     = Replace(strer,"max-width",web_var(web_num,9))
    strer     = Replace(strer,"max-height",web_var(web_num,10))
    code_jk   = strer
End Function

Function code_jk2(strers)
    Dim strer:strer = strers
    If strer = "" Or IsNull(strer) Then code_jk2 = "":Exit Function
    strer = health_var(strer,1)
    strer = Replace(strer,"<","&lt;")
    strer = Replace(strer,">","&gt;")
    strer = Replace(strer," ","&nbsp")		'空格
    strer = Replace(strer,Chr(9),"&nbsp")		'table
    strer = Replace(strer,"'","&#39;")		'单引号
    strer = Replace(strer,"""","&quot;")		'双引号
    Dim re
    Dim re_v
    re_v          = "[^\(\)\;\'\[]*"
    Set re        = new RegExp
    re.IgnoreCase = True
    re.Global     = True
    re.Pattern    = "(javascript:)"
    strer         = re.Replace(strer,"<i>javascript</i>:")
    re.Pattern    = "(javascript)"
    strer         = re.Replace(strer,"<i>&#106avascript</i>")
    re.Pattern    = "(jscript:)"
    strer         = re.Replace(strer,"<i>&#106script</i>:")
    re.Pattern    = "(js:)"
    strer         = re.Replace(strer,"<i>&#106s</i>:")
    re.Pattern    = "(value)"
    strer         = re.Replace(strer,"<i>&#118alue</i>")
    re.Pattern    = "(about:)"
    strer         = re.Replace(strer,"<i>about&#58</i>")
    re.Pattern    = "(file:)"
    strer         = re.Replace(strer,"<i>file&&#58</i>")
    re.Pattern    = "(document.)"
    strer         = re.Replace(strer,"<i>document</i>.")
    re.Pattern    = "(vbscript:)"
    strer         = re.Replace(strer,"<i>&#118bscript</i>:")
    re.Pattern    = "(vbs:)"
    strer         = re.Replace(strer,"<i>&#118bs</i>&#58")
    re.Pattern    = "(on(mouse|exit|error|click|key))"
    strer         = re.Replace(strer,"<i>&#111n$2</i>")
    re.Pattern    = "\[IMG\](.[^\[]*(gif|jpg|jpeg|bmp|png))\[\/IMG\]"
    strer         = re.Replace(strer,"<IMG SRC='$1' align=center border=0 onload=""javascript:if(this.width>250)this.width=250"">")
    re.Pattern    = "(\[URL\])(.[^\[]*)(\[\/URL\])"
    strer         = re.Replace(strer,"<A HREF='$2' TARGET=_blank>$2</A>")
    re.Pattern    = "(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
    strer         = re.Replace(strer,"<A HREF='$2' TARGET=_blank>$3</A>")
    re.Pattern    = "(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
    strer         = re.Replace(strer,"<img align=absmiddle src=images/small/email.gif><A HREF=""mailto:$2"">$2</A>")
    re.Pattern    = "(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
    strer         = re.Replace(strer,"<img align=absmiddle src=images/small/email.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
    re.Pattern    = "\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
    strer         = re.Replace(strer,"<font color=$1>$2</font>")
    re.Pattern    = "\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
    strer         = re.Replace(strer,"<font face=$1>$2</font>")
    re.Pattern    = "\[align=(.[^\[]*)\](.[^\[]*)\[\/align\]"
    strer         = re.Replace(strer,"<div align=$1>$2</div>")
    re.Pattern    = "\[align=(.[^\[]*)\](.*)\[\/align\]"
    strer         = re.Replace(strer,"<div align=$1>$2</div>")
    re.Pattern    = "\[center\](.[^\[]*)\[\/center\]"
    strer         = re.Replace(strer,"<div align=center>$1</div>")
    re.Pattern    = "\[fly\](.[^\[]*)\[\/fly\]"
    strer         = re.Replace(strer,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
    re.Pattern    = "\[move\](.[^\[]*)\[\/move\]"
    strer         = re.Replace(strer,"<MARQUEE scrollamount=3>$1</marquee>")
    re.Pattern    = "\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
    strer         = re.Replace(strer,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
    re.Pattern    = "\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
    strer         = re.Replace(strer,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
    re.Pattern    = "\[i\](.[^\[]*)\[\/i\]"
    strer         = re.Replace(strer,"<i>$1</i>")
    re.Pattern    = "\[u\](.[^\[]*)(\[\/u\])"
    strer         = re.Replace(strer,"<u>$1</u>")
    re.Pattern    = "\[b\](.[^\[]*)(\[\/b\])"
    strer         = re.Replace(strer,"<b>$1</b>")
    re.Pattern    = "\[size=([1-4])\](.[^\[]*)\[\/size\]"
    strer         = re.Replace(strer,"<font size=$1>$2</font>")
    Set re        = Nothing
    Dim emci

    For emci = 1 To 16
        strer = Replace(strer,"[em" & emci & "]","<image src='images/icon/em" & emci & ".gif' border=0>&nbsp;")
    Next

    For emci = 1 To 13
        strer = Replace(strer,"[emb" & emci & "]","<image src='images/icon/emb" & emci & ".gif' border=0>&nbsp;")
    Next

    strer = Replace(strer,"[br]","<br>")
    strer = Replace(strer,"[BR]","<br>")
    strer = Replace(strer,vbCrLf,"<br>")
    code_jk2 = strer
End Function %>