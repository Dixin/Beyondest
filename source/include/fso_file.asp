<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

sub config_file()
  dim filetype,file_name
  filetype=filetype&"<"&""&"%" & vbcrlf & _
	"'*******************************************************************"&vbcrlf&_
	"'"&vbcrlf&_
	"'                     Beyondest.Com V3.6 Demo版"&vbcrlf&_
	"' "&vbcrlf&_
	"'           网址：http://www.beyondest.com"&vbcrlf&_
	"' "&vbcrlf&_
	"'*******************************************************************"&vbcrlf&_
	"web_config="""&web_config&""""&vbcrlf&_
	"web_cookies="""&web_cookies&""""&vbcrlf&_
	"web_login="&web_login&vbcrlf&_
	"web_setup="""&web_setup&""""&vbcrlf&_
	"web_num="""&web_num&""""&vbcrlf&_
	"web_menu="""&web_menu&""""&vbcrlf&_
	"web_color="""&web_color&""""&vbcrlf&_
	"web_upload="""&web_upload&""""&vbcrlf&_
	"web_safety="""&web_safety&""""&vbcrlf&_
	"web_error="""&web_error&""""&vbcrlf&_
	"web_news_art="""&web_news_art&""""&vbcrlf&_
	"web_down="""&web_down&""""&vbcrlf&_
	"web_shop="""&web_shop&""""&vbcrlf&_
	"web_stamp="""&web_stamp&""""&vbcrlf&_
	"user_power="""&user_power&""""&vbcrlf&_
	"user_grade="""&user_grade&""""&vbcrlf&_
	"forum_type="""&forum_type&""""&vbcrlf&_
	"%"&""&">"
  file_name="include/common.asp"
  call create_file(file_name,filetype)
  filetype="<"&""&"%" & vbcrlf & _
        "dim web_font_family,web_font_size"&vbcrlf&_
	"web_font_family="""&web_font_family&""""&vbcrlf&_
	"web_font_size="""&web_font_size&""""&vbcrlf&_
	"%"&""&">"
  file_name="include/common_other.asp"
  call create_file(file_name,filetype)
end sub

sub config_mouse_on_title()
  dim filetype,file_name
  file_name="style/mouse_on_title.js"
  filetype="<!--" & _
	   "//******************************************************************"&vbcrlf&_
	   "//"&vbcrlf&_
	   "//                     Beyondest.Com V3.6 Demo版"&vbcrlf&_
	   "//"&vbcrlf&_
	   "//           网址：http://www.beyondest.com"&vbcrlf&_
	   "//"&vbcrlf&_
	   "//******************************************************************"&vbcrlf&_
           vbcrlf&"//***********默认设置定义.*********************" & _
           vbcrlf&"tPopWait=50;		//停留tWait豪秒后显示提示。" & _
           vbcrlf&"tPopShow=6000;		//显示tShow豪秒后关闭提示" & _
           vbcrlf&"showPopStep=20;" & _
           vbcrlf&"popOpacity=95;" & _
           vbcrlf&"fontcolor="""&code_config(request.form("web_color_7"),2)&""";" & _
           vbcrlf&"bgcolor="""&code_config(request.form("web_color_5"),2)&""";" & _
           vbcrlf&"bordercolor="""&code_config(request.form("web_color_2"),2)&""";" & _
           vbcrlf&vbcrlf&"//***************内部变量定义*****************" & _
           vbcrlf&"sPop=null;curShow=null;tFadeOut=null;tFadeIn=null;tFadeWaiting=null;" & _
           vbcrlf&vbcrlf&"document.write(""<style type='text/css'id='defaultPopStyle'>"");" & _
           vbcrlf&"document.write("".cPopText {  background-color: "" + bgcolor + "";color:"" + fontcolor + ""; border: 1px "" + bordercolor + "" solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0)}"");" & _
           vbcrlf&"document.write(""</style>"");" & _
           vbcrlf&"document.write(""<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>"");" & _
           vbcrlf&vbcrlf&vbcrlf&"function showPopupText(){" & _
           vbcrlf&"var o=event.srcElement;" & _
           vbcrlf&"	MouseX=event.x;" & _
           vbcrlf&"	MouseY=event.y;" & _
           vbcrlf&"	if(o.alt!=null && o.alt!=""""){o.dypop=o.alt;o.alt=""""};" & _
           vbcrlf&"        if(o.title!=null && o.title!=""""){o.dypop=o.title;o.title=""""};" & _
           vbcrlf&"	if(o.dypop!=sPop) {" & _
           vbcrlf&"			sPop=o.dypop;" & _
           vbcrlf&"			clearTimeout(curShow);" & _
           vbcrlf&"			clearTimeout(tFadeOut);" & _
           vbcrlf&"			clearTimeout(tFadeIn);" & _
           vbcrlf&"			clearTimeout(tFadeWaiting);	" & _
           vbcrlf&"			if(sPop==null || sPop=="""") {" & _
           vbcrlf&"				dypopLayer.innerHTML="""";" & _
           vbcrlf&"				dypopLayer.style.filter=""Alpha()"";" & _
           vbcrlf&"				dypopLayer.filters.Alpha.opacity=0;	" & _
           vbcrlf&"				}" & _
           vbcrlf&"			else {" & _
           vbcrlf&"				if(o.dyclass!=null) popStyle=o.dyclass " & _
           vbcrlf&"					else popStyle=""cPopText"";" & _
           vbcrlf&"				curShow=setTimeout(""showIt()"",tPopWait);" & _
           vbcrlf&"			}" & _
           vbcrlf&"			" & _
           vbcrlf&"	}" & _
           vbcrlf&"}" & _
           vbcrlf&vbcrlf&"function showIt(){" & _
           vbcrlf&"		dypopLayer.className=popStyle;" & _
           vbcrlf&"		dypopLayer.innerHTML=sPop;" & _
           vbcrlf&"		popWidth=dypopLayer.clientWidth;" & _
           vbcrlf&"		popHeight=dypopLayer.clientHeight;" & _
           vbcrlf&"		if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24" & _
           vbcrlf&"			else popLeftAdjust=0;" & _
           vbcrlf&"		if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24" & _
           vbcrlf&"			else popTopAdjust=0;" & _
           vbcrlf&"		dypopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;" & _
           vbcrlf&"		dypopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;" & _
           vbcrlf&"		dypopLayer.style.filter=""Alpha(Opacity=0)"";" & _
           vbcrlf&"		fadeOut();" & _
           vbcrlf&"}" & _
           vbcrlf&vbcrlf&"function fadeOut(){" & _
           vbcrlf&"	if(dypopLayer.filters.Alpha.opacity<popOpacity) {" & _
           vbcrlf&"		dypopLayer.filters.Alpha.opacity+=showPopStep;" & _
           vbcrlf&"		tFadeOut=setTimeout(""fadeOut()"",1);" & _
           vbcrlf&"		}" & _
           vbcrlf&"		else {" & _
           vbcrlf&"			dypopLayer.filters.Alpha.opacity=popOpacity;" & _
           vbcrlf&"			tFadeWaiting=setTimeout(""fadeIn()"",tPopShow);" & _
           vbcrlf&"			}" & _
           vbcrlf&"}" & _
           vbcrlf&vbcrlf&"function fadeIn(){" & _
           vbcrlf&"	if(dypopLayer.filters.Alpha.opacity>0) {" & _
           vbcrlf&"		dypopLayer.filters.Alpha.opacity-=1;" & _
           vbcrlf&"		tFadeIn=setTimeout(""fadeIn()"",1);" & _
           vbcrlf&"		}" & _
           vbcrlf&"}" & _
           vbcrlf&"document.onmouseover=showPopupText;" & _
           vbcrlf&"-->"
  call create_file(file_name,filetype)
end sub
sub config_css()
  dim filetype,file_name
  file_name="include/beyondest.css"
  filetype="<!--" & _
	   vbcrlf&"body,p,td { font-family:"&web_font_family&"; font-size:"&web_font_size&"; color:"&code_config(request.form("web_color_7"),2)&" }" & _
	   vbcrlf&"body { cursor: url('images/beyondest.cur');" & _
	   vbcrlf&"scrollbar-face-color: #EEEEEE;" & _
	   vbcrlf&"scrollbar-highlight-color: #FFFFFF;" & _
	   vbcrlf&"scrollbar-shadow-color: #DEE3E7;" & _
	   vbcrlf&"scrollbar-3dlight-color: #D1D7DC;" & _
	   vbcrlf&"scrollbar-arrow-color:  "&code_config(request.form("web_color_2"),2)&";" & _
	   vbcrlf&"scrollbar-track-color: #ededed;" & _
	   vbcrlf&"scrollbar-darkshadow-color: "&code_config(request.form("web_color_3"),2)&"; }" & _
	   vbcrlf&"" & _
	   vbcrlf&"INPUT { BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px;" & _
	   vbcrlf&" BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px;" & _
	   vbcrlf&" PADDING-BOTTOM: 1px; PADDING-TOP: 1px; HEIGHT: 18px;" & _
	   vbcrlf&" BORDER-LEFT-COLOR: #c0c0c0; BORDER-BOTTOM-COLOR: #c0c0c0;" & _
	   vbcrlf&" BORDER-TOP-COLOR: #c0c0c0; BORDER-RIGHT-COLOR: #c0c0c0;" & _
	   vbcrlf&" background-color: "&code_config(request.form("web_color_1"),2)&"; CURSOR: HAND;" & _
	   vbcrlf&" FONT-SIZE: "&web_font_size&"; font-family: "&web_font_family&"; COLOR: "&code_config(request.form("web_color_7"),2)&";" & _
	   vbcrlf&"}" & _
	   vbcrlf&"textarea { border-width: 1; border-color: #c0c0c0; background-color: "&code_config(request.form("web_color_1"),2)&";" & _
	   vbcrlf&" font-family: "&web_font_family&"; font-size: "&web_font_size&"; CURSOR: HAND; COLOR: "&code_config(request.form("web_color_7"),2)&";" & _
	   vbcrlf&"}" & _
	   vbcrlf&"select { border-width: 1; border-color: #c0c0c0; background-color: "&code_config(request.form("web_color_1"),2)&";" & _
	   vbcrlf&" font-family: "&web_font_family&"; font-size:"&web_font_size&"; CURSOR: HAND; COLOR: "&code_config(request.form("web_color_7"),2)&";" & _
	   vbcrlf&"}" & _
	   vbcrlf&".bg_1 { background-color: "&code_config(request.form("web_color_1"),2)&" }" & _
	   vbcrlf&".bg_2 { background-color: "&code_config(request.form("web_color_6"),4)&" }" & _
	   vbcrlf&".bg_3 { background-color: "&code_config(request.form("web_color_5"),5)&" }" & _
	   vbcrlf&".timtd { font-family: Arial, 宋体, Verdana, Helvetica, sans-serif }" & _
	   vbcrlf&".htd { line-height:150% }" & _
	   vbcrlf&".btd { font-weight:bold }" & _
	   vbcrlf&".bw { WORD-WRAP: break-word }" & _
	   vbcrlf&".tf { TABLE-LAYOUT: fixed }" & _
	   vbcrlf&".red { color:"&code_config(request.form("web_color_10"),2)&" }" & _
	   vbcrlf&".red_2 { color:"&code_config(request.form("web_color_11"),2)&" }" & _
	   vbcrlf&".red_3 { color:"&code_config(request.form("web_color_12"),2)&" }" & _
	   vbcrlf&".end { color:"&code_config(request.form("web_color_1"),2)&" }" & _
	   vbcrlf&".blue { color: "&code_config(request.form("web_color_8"),2)&"; }" & _
	   vbcrlf&".gray { color: "&code_config(request.form("web_color_9"),2)&"; }" & _
	   vbcrlf&"a { COLOR: "&code_config(request.form("web_color_3"),2)&"; TEXT-DECORATION: none }" & _
	   vbcrlf&"a:hover { COLOR: "&code_config(request.form("web_color_10"),2)&"; TEXT-DECORATION: underline }" & _
	   vbcrlf&"a.menu { COLOR: "&code_config(request.form("web_color_1"),2)&"; TEXT-DECORATION: none }" & _
	   vbcrlf&"a.menu:hover { COLOR: "&code_config(request.form("web_color_12"),2)&"; TEXT-DECORATION: none }" & _
	   vbcrlf&"-->"
  call create_file(file_name,filetype)
end sub
sub create_file(file_name,filetype)
  dim filetemp,fileos,filepath
  set fileos=CreateObject("Scripting.FileSystemObject")
  filepath=server.mappath(file_name)
  set filetemp=fileos.createtextfile(filepath,true)
  filetemp.writeline( filetype )
  filetemp.close
  set filetemp=nothing
  set fileos=nothing
end sub
%>