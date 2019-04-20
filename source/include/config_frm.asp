<%
'*******************************************************************
'
'                     Beyondest.Com V3.6 Demo版
' 
'           网址：http://www.beyondest.com
' 
'*******************************************************************

sub frm_ubb(fn,tn,uw)
%><script language=javascript>
<!--
var defmode="advmode";		//默认模式，可选 normalmode, advmode, 或 helpmode
var ubb_w=450;
var ubb_h=350;
var ubb_name="UBB代码 - ";

if (defmode == "advmode")
{ helpmode = false; normalmode = false; advmode = true; }
else if (defmode == "helpmode")
{ helpmode = true; normalmode = false; advmode = false; }
else
{ helpmode = false; normalmode = true; advmode = false; }

function jk_ubb_mode(swtch)
{
  if (swtch == 1)
  {
    advmode = false; normalmode = false; helpmode = true;
    alert(ubb_name+"帮助信息\n\n点击相应的代码按钮即可获得相应的说明和提示");
  }
  else if (swtch == 0)
  {
    helpmode = false; normalmode = false; advmode = true;
    alert(ubb_name+"直接插入\n\n点击代码按钮后不出现提示即直接插入相应代码");
  }
  else if (swtch == 2)
  {
    helpmode = false; advmode = false; normalmode = true;
    alert(ubb_name+"提示插入\n\n点击代码按钮后出现向导窗口帮助您完成代码插入");
  }
}

function AddText(NewCode)
{
  if(document.all)
  { insertAtCaret(document.<%response.write fn&"."&tn%>, NewCode); setfocus(); } 
  else
  { document.<%response.write fn&"."&tn%>.value += NewCode; setfocus(); }
}

function storeCaret (textEl)
{ if(textEl.createTextRange){ textEl.caretPos = document.selection.createRange().duplicate();} }

function insertAtCaret (textEl, text)
{
  if (textEl.createTextRange && textEl.caretPos)
  {
    var caretPos = textEl.caretPos;
    caretPos.text += caretPos.text.charAt(caretPos.text.length - 2) == ' ' ? text + ' ' : text;
  }
  else if(textEl)
  { textEl.value += text; }
  else
  { textEl.value = text; }
}

function jk_ubb_email()
{
  if (helpmode)
  { alert(ubb_name +"插入邮件地址\n\n插入邮件地址连接！\n例如：\n[email]plinq@live.com[/email]\n[email=plinq@live.com]笼民[/email]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[email]" + range.text + "[/email]"; }
  else if (advmode)
  { AddTxt="[email][/email]"; AddText(AddTxt); }
  else
  { 
    txt2=prompt(ubb_name+"请输入链接显示的文字，如果留空则直接显示邮件地址！",""); 
    if (txt2!=null)
    {
      txt = prompt(ubb_name +"请输入邮件地址！例：plinq@live.com","");      
      if (txt!=null)
      {
        if (txt2=="")
        { AddTxt="[email]"+txt+"[/email]"; }
        else
        { AddTxt="[email="+txt+"]"+txt2+"[/email]"; } 
        AddText(AddTxt);
      }
    }
  }
}

function jk_ubb_size(size)
{
  if (helpmode)
  { alert(ubb_name+"设置字号\n\n将标签所包围的文字设置成指定字号！\n例如：[size=3]文字大小为 3[/size]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[size=" + size + "]" + range.text + "[/size]"; }
  else if (advmode)
  { AddTxt="[size="+size+"][/size]"; AddText(AddTxt); }
  else
  {           
    txt=prompt(ubb_name+"请输入要设置为字号 "+size+" 的文字！","文字"); 
    if (txt!=null) { AddTxt="[size="+size+"]"+txt; AddText(AddTxt); AddText("[/size]"); }  
  }
}

function jk_ubb_font(font)
{
  if (helpmode)
  { alert(ubb_name+"设定字体\n\n将标签所包围的文字设置成指定字体！\n例如：[face=仿宋]字体为仿宋[/face]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[face=" + font + "]" + range.text + "[/face]"; }
  else if (advmode)
  { AddTxt="[face="+font+"][/face]"; AddText(AddTxt); }
  else
  {      
    txt=prompt(ubb_name+"请输入要设置成 "+font+" 的文字！","文字");
    if (txt!=null) { AddTxt="[face="+font+"]"+txt; AddText(AddTxt); AddText("[/face]"); }  
  }
}


function jk_ubb_bold()
{
  if (helpmode)
  { alert(ubb_name+"插入粗体文本\n\n将标签所包围的文本变成粗体！\n例如：[b]Beyondest.com[/b]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[b]" + range.text + "[/b]"; }
  else if (advmode)
  { AddTxt="[b][/b]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"请输入要设置成粗体的文字！","文字");     
    if (txt!=null) { AddTxt="[b]"+txt; AddText(AddTxt); AddText("[/b]"); }       
  }
}

function jk_ubb_italicize()
{
  if (helpmode)
  { alert(ubb_name+"插入斜体文本\n\n将标签所包围的文本变成斜体！\n例如：[i]Beyondest.com[/i]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[i]" + range.text + "[/i]"; }
  else if (advmode)
  { AddTxt="[i][/i]"; AddText(AddTxt); }
  else
  {   
    txt=prompt(ubb_name+"请输入要设置成斜体的文字！","文字");     
    if (txt!=null) { AddTxt="[i]"+txt; AddText(AddTxt); AddText("[/i]"); }         
  }
}

function jk_ubb_quote()
{
  if (helpmode)
  { alert(ubb_name+"插入引用\n\n将标签所包围的文本作为引用特殊显示！\n例如：[quote]Beyondest.com[/quote]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[quote]" + range.text + "[/quote]"; }
  else if (advmode)
  { AddTxt="\r[quote]\r[/quote]"; AddText(AddTxt); }
  else
  {
    txt=prompt(ubb_name+"请输入要作为引用显示的文字！","文字");     
    if(txt!=null) { AddTxt="\r[quote]\r"+txt; AddText(AddTxt); AddText("\r[/quote]"); }         
  }
}

function jk_ubb_color(color)
{
  if (helpmode)
  { alert(ubb_name+"插入定义颜色文本\n\n将标签所包围的文本变为制定颜色！\n例如：[color=red]红颜色[/color]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[color=" + color + "]" + range.text + "[/color]"; }
  else if (advmode)
  { AddTxt="[color="+color+"][/color]"; AddText(AddTxt); }
  else
  {
    txt=prompt(ubb_name+"请输入要设置成颜色 "+color+" 的文字！","文字");
    if(txt!=null) { AddTxt="[color="+color+"]"+txt; AddText(AddTxt); AddText("[/color]"); }
  }
}

function jk_ubb_center()
{
  if (helpmode)
  { alert(ubb_name+"居中对齐\n\n将标签所包围的文本居中对齐显示！\n例如：[align=center]内容居中[/align]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[center]" + range.text + "[/center]"; }
  else if (advmode)
  { AddTxt="[align=center][/align]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"请输入要居中对齐的文字！","文字");     
    if (txt!=null) { AddTxt="\r[align=center]"+txt; AddText(AddTxt); AddText("[/align]"); }        
  }
}

function jk_ubb_link()
{
  if (helpmode)
  { alert(ubb_name+"插入超级链接\n\n插入一个超级连接！\n例如：\n[url]http://www.beyondest.com/[/url]\n[url=http://www.beyondest.com/]Beyondest.com[/url]"); }
  else if (advmode)
  { AddTxt="[url][/url]"; AddText(AddTxt); }
  else
  { 
    txt2=prompt(ubb_name+"请输入链接显示的文字，如果留空则直接显示链接！",""); 
    if (txt2!=null)
    {
      txt=prompt(ubb_name+"请输入 URL！例：http://www.beyondest.com/","http://");      
      if (txt!=null)
      {
        if (txt2=="")
        { AddTxt="[url]"+txt; AddText(AddTxt); AddText("[/url]"); }
        else
        { AddTxt="[url="+txt+"]"+txt2; AddText(AddTxt); AddText("[/url]"); }   
      } 
    }
  }
}

function jk_ubb_image()
{
  if (helpmode)
  { alert(ubb_name+"插入图像\n\n在文本中插入一幅图像！\n例如：[IMG]http://www.beyondest.com/images/logo.gif[/IMG]"); }
  else if (advmode)
  { AddTxt="[IMG][/IMG]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"请输入图像的 URL！例：http://www.beyondest.com/images/logo.gif","http://");    
    if(txt!=null) { AddTxt="\r[IMG]"+txt; AddText(AddTxt); AddText("[/IMG]");
    }       
  }
}

function jk_ubb_code()
{
  if (helpmode)
  { alert(ubb_name+"插入代码\n\n插入程序或脚本原始代码！\n例如：[code]Beyondest.com[/code]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[code]" + range.text + "[/code]"; }
  else if (advmode)
  { AddTxt="\r[code][/code]"; AddText(AddTxt); }
  else
  {   
    txt=prompt(ubb_name+"请输入要插入的代码！","");     
    if (txt!=null) { AddTxt="\r[code]"+txt; AddText(AddTxt); AddText("[/code]"); }
  }
}

function jk_ubb_flash()
{
  if (helpmode)
  { alert(ubb_name+"插入 Flash\n\n在文本中插入 Flash 动画！\n例如：[FLASH="+ubb_w+","+ubb_h+"]http://www.Beyondrest.com/images/banner.swf[/FLASH]"); }
  else if (advmode)
  { AddTxt="[FLASH="+ubb_w+","+ubb_h+"][/FLASH]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"请输入 Flash 动画的大小　",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"请输入 Flash 动画的地址！","http://");
      if(txt!=null) { AddTxt="\r[FLASH="+stxt+"]"+txt; AddText(AddTxt); AddText("[/FLASH]"); }
    }
  }
}

function jk_ubb_rm()
{
  if (helpmode)
  { alert(ubb_name+"插入 RM\n\n在文本中插入 Realplay 视频文件！\n例如：[RM="+ubb_w+","+ubb_h+"]http://www.Beyondest.com/images/test.ram[/RM]"); }
  else if (advmode)
  { AddTxt="[RM="+ubb_w+","+ubb_h+"][/RM]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"请输入 Realplay 视频文件的大小！",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"请输入 Realplay 视频文件的地址 rstp://等都支持","http://");
      if(txt!=null) { AddTxt="\r[RM="+stxt+"]"+txt; AddText(AddTxt); AddText("[/RM]"); }
    }
  }
}

function jk_ubb_mp()
{
  if (helpmode)
  { alert(ubb_name+"插入 MP\n\n在文本中插入 Windows Media Player 视频文件！\n例如：[MP="+ubb_w+","+ubb_h+"]http://www.beyondest.com/images/test.wmv[/MP]"); }
  else if (advmode)
  { AddTxt="[MP="+ubb_w+","+ubb_h+"][/MP]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"请输入 Windows Media Player 视频文件的大小！",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"请输入 Windows Media Player 视频文件的地址，各种地址头都支持！","http://");    
      if(txt!=null) { AddTxt="\r[MP="+stxt+"]"+txt; AddText(AddTxt); AddText("[/MP]"); }
    }
  }
}

function jk_ubb_underline()
{
  if (helpmode)
  { alert(ubb_name+"插入下划线\n\n给标签所包围的文本加上下划线！\n例如：[u]Beyondest.com[/u]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[u]" + range.text + "[/u]"; }
  else if (advmode)
  { AddTxt="[u][/u]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"请输入要加下划线的文字！","文字");
    if (txt!=null) { AddTxt="[u]"+txt; AddText(AddTxt); AddText("[/u]"); }         
  }
}

function setfocus() { document.<%response.write fn&"."&tn%>.focus(); }
-->
</script><table border=0 cellspacing=0 cellpadding=0><tr>
<td height=30><%response.write uw%><select onchange="javascript:jk_ubb_font(this.options[this.selectedIndex].value);" size=1 name=font>
<option value=宋体>宋体</option>
<option value=黑体>黑体</option>
<option value=arial>arial</option>
<option value="Book antiqua">Book antiqua</option>
<option value="Century Gothic">Century Gothic</option>
<option value="Courier New" selected>Courier New</option>
<option value=Georgia>Georgia</option>
<option value=Impact>Impact</option>
<option value=Tahoma>Tahoma</option>
<option value="Times New Roman">Times New Roman</option>
<option value=Verdana>Verdana</option>
</select></td>
<td>&nbsp;&nbsp;<select onchange="javascript:jk_ubb_size(this.options[this.selectedIndex].value);" size=1 name=size>
<option value=-2>-2</option>
<option value=-1>-1</option>
<option value=1>1</option>
<option value=2>2</option>
<option value=3 selected>3</option>
<option value=4>4</option>
<option value=5>5</option>
<option value=6>6</option>
<option value=7>7</option>
</select></td>
<td>&nbsp;&nbsp;<select onchange="javascript:jk_ubb_color(this.options[this.selectedIndex].value);" size=1 name=color>
<option style="COLOR: white" value=White selected>White</option>
<option style="COLOR: black" value=Black>Black</option>
<option style="COLOR: red" value=Red>Red</option>
<option style="COLOR: yellow" value=Yellow>Yellow</option>
<option style="COLOR: pink" value=Pink>Pink</option>
<option style="COLOR: green" value=Green>Green</option>
<option style="COLOR: orange" value=Orange>Orange</option>
<option style="COLOR: purple" value=Purple>Purple</option>
<option style="COLOR: blue" value=Blue>Blue</option>
<option style="COLOR: beige" value=Beige>Beige</option>
<option style="COLOR: brown" value=Brown>Brown</option>
<option style="COLOR: teal" value=Teal>Teal</option>
<option style="COLOR: navy" value=Navy>Navy</option>
<option style="COLOR: maroon" value=Maroon>Maroon</option>
<option style="COLOR: limegreen" value=LimeGreen>LimeGreen</option>
</select></td>
</tr>
</table>
<table border=0 cellspacing=0 cellpadding=0>
<tr><td height=30><%response.write uw%><a href="javascript:jk_ubb_bold();"><img alt='插入粗体文本' src='images/ubb/bold.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_italicize();"><img alt='插入斜体文本' src='images/ubb/italicize.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_underline();"><img alt='插入下划线' src='images/ubb/underline.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_center();"><img alt='居中对齐' src='images/ubb/center.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_link();"><img alt='插入超级链接' src='images/ubb/url.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_email();"><img alt='插入邮件地址' src='images/ubb/email.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_image();"><img alt='插入图像' src='images/ubb/image.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_flash();"><img alt='插入 flash' src='images/ubb/flash.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_code();"><img alt='插入代码' src='images/ubb/code.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_quote();"><img alt='插入引用' src='images/ubb/quote.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_rm();"><img alt='插入Realplay视频文件' src='images/ubb/rm.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_mp();"><img alt='插入Media Player播放文件' src='images/ubb/mp.gif' border=0></a></td></tr></table><%
end sub

sub frm_ubb_type()
%>辅助模式：<br><input onclick="javascript:jk_ubb_mode('2');" type=radio value=2 name=mode class=bg_1> 提示插入<br>
<input onclick="javascript:jk_ubb_mode('0');" type=radio value=0 name=mode class=bg_1 checked> 直接插入<br>
<input onclick="javascript:jk_ubb_mode('1');" type=radio value=1 name=mode class=bg_1> 帮助信息<%
end sub

sub frm_word_size(fn,tn,ws,wv)
%>
<a href="javascript:checklength(document.<%response.write fn%>);">[字数检查]</a>
<script language=JavaScript>
<!--
function checklength(theform)
{
  var postmaxchars=<%response.write ws%>;
  var postchars=theform.<%response.write tn%>.value.length;
  var post_1="您输入的<%response.write wv%>已经超过系统允许的最大值！\n请删减部分<%response.write wv%>！";
  var message="";
  if (postmaxchars != 0) { message = "系统允许："+postmaxchars+"KB（约"+postmaxchars*1024+"个字符）"; }
  if (postmaxchars*1024>=postchars)
  {
    var postc=postmaxchars*1024-postchars;
    post_1="您还可以输入："+postc+"个字符（约"+Math.round(postc/1024)+"KB）";
  }
  alert("您输入的<%response.write wv%>的统计信息如下：\n\n当前长度："+postchars+"个字符（约"+Math.round(postchars/1024)+"KB）\n"+message+"\n"+post_1);
}
-->
</script><%
end sub

sub frm_topic(fn,tn)
%><select onchange="document.<%response.write fn&"."&tn%>.focus(); document.<%response.write fn&"."&tn%>.value = this.options[this.selectedIndex].value + document.<%response.write fn&"."&tn%>.value;"> 
<option value='' selected>选择</option>
<option value=[原创]>[原创]</option>
<option value=[转帖]>[转帖]</option>
<option value=[Hack]>[Hack]</option>
<option value=[砖头]>[砖头]</option>
<option value=[乱弹]>[乱弹]</option>
<option value=[灌水]>[灌水]</option>
<option value=[讨论]>[讨论]</option>
<option value=[警告]>[警告]</option> 
<option value=[求助]>[求助]</option>
<option value=[推荐]>[推荐]</option> 
<option value=[公告]>[公告]</option>
<option value=[成人]>[成人]</option> 
<option value=[注意]>[注意]</option>
<option value=[贴图]>[贴图]</option> 
<option value=[建议]>[建议]</option>
<option value=[下载]>[下载]</option>
<option value=[分享]>[分享]</option>
</select><%
end sub
%>