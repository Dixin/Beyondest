<%
' ====================
' Beyondest.Com v3.6.1
' http://beyondest.com
' ====================
Sub frm_ubb(fn,tn,uw) %><script language=javascript>
<!--
var defmode="advmode";		//Ĭ��ģʽ����ѡ normalmode, advmode, �� helpmode
var ubb_w=450;
var ubb_h=350;
var ubb_name="UBB���� - ";

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
    alert(ubb_name+"������Ϣ\n\n�����Ӧ�Ĵ��밴ť���ɻ����Ӧ��˵������ʾ");
  }
  else if (swtch == 0)
  {
    helpmode = false; normalmode = false; advmode = true;
    alert(ubb_name+"ֱ�Ӳ���\n\n������밴ť�󲻳�����ʾ��ֱ�Ӳ�����Ӧ����");
  }
  else if (swtch == 2)
  {
    helpmode = false; advmode = false; normalmode = true;
    alert(ubb_name+"��ʾ����\n\n������밴ť������򵼴��ڰ�������ɴ������");
  }
}

function AddText(NewCode)
{
  if(document.all)
  { insertAtCaret(document.<% Response.Write fn & "." & tn %>, NewCode); setfocus(); } 
  else
  { document.<% Response.Write fn & "." & tn %>.value += NewCode; setfocus(); }
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
  { alert(ubb_name+"�����ʼ���ַ\n\n�����ʼ���ַ���ӣ�\n���磺\n[email]dixinyan@live.com[/email]\n[email=dixinyan@live.com]����[/email]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[email]" + range.text + "[/email]"; }
  else if (advmode)
  { AddTxt="[email][/email]"; AddText(AddTxt); }
  else
  { 
    txt2=prompt(ubb_name+"������������ʾ�����֣����������ֱ����ʾ�ʼ���ַ��",""); 
    if (txt2!=null)
    {
      txt=prompt(ubb_name+"�������ʼ���ַ������dixinyan@live.com","");      
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
  { alert(ubb_name+"�����ֺ�\n\n����ǩ����Χ���������ó�ָ���ֺţ�\n���磺[size=3]���ִ�СΪ 3[/size]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[size=" + size + "]" + range.text + "[/size]"; }
  else if (advmode)
  { AddTxt="[size="+size+"][/size]"; AddText(AddTxt); }
  else
  {           
    txt=prompt(ubb_name+"������Ҫ����Ϊ�ֺ� "+size+" �����֣�","����"); 
    if (txt!=null) { AddTxt="[size="+size+"]"+txt; AddText(AddTxt); AddText("[/size]"); }  
  }
}

function jk_ubb_font(font)
{
  if (helpmode)
  { alert(ubb_name+"�趨����\n\n����ǩ����Χ���������ó�ָ�����壡\n���磺[face=����]����Ϊ����[/face]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[face=" + font + "]" + range.text + "[/face]"; }
  else if (advmode)
  { AddTxt="[face="+font+"][/face]"; AddText(AddTxt); }
  else
  {      
    txt=prompt(ubb_name+"������Ҫ���ó� "+font+" �����֣�","����");
    if (txt!=null) { AddTxt="[face="+font+"]"+txt; AddText(AddTxt); AddText("[/face]"); }  
  }
}


function jk_ubb_bold()
{
  if (helpmode)
  { alert(ubb_name+"��������ı�\n\n����ǩ����Χ���ı���ɴ��壡\n���磺[b]Beyondest.com[/b]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[b]" + range.text + "[/b]"; }
  else if (advmode)
  { AddTxt="[b][/b]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"������Ҫ���óɴ�������֣�","����");     
    if (txt!=null) { AddTxt="[b]"+txt; AddText(AddTxt); AddText("[/b]"); }       
  }
}

function jk_ubb_italicize()
{
  if (helpmode)
  { alert(ubb_name+"����б���ı�\n\n����ǩ����Χ���ı����б�壡\n���磺[i]Beyondest.com[/i]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[i]" + range.text + "[/i]"; }
  else if (advmode)
  { AddTxt="[i][/i]"; AddText(AddTxt); }
  else
  {   
    txt=prompt(ubb_name+"������Ҫ���ó�б������֣�","����");     
    if (txt!=null) { AddTxt="[i]"+txt; AddText(AddTxt); AddText("[/i]"); }         
  }
}

function jk_ubb_quote()
{
  if (helpmode)
  { alert(ubb_name+"��������\n\n����ǩ����Χ���ı���Ϊ����������ʾ��\n���磺[quote]Beyondest.com[/quote]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[quote]" + range.text + "[/quote]"; }
  else if (advmode)
  { AddTxt="\r[quote]\r[/quote]"; AddText(AddTxt); }
  else
  {
    txt=prompt(ubb_name+"������Ҫ��Ϊ������ʾ�����֣�","����");     
    if(txt!=null) { AddTxt="\r[quote]\r"+txt; AddText(AddTxt); AddText("\r[/quote]"); }         
  }
}

function jk_ubb_color(color)
{
  if (helpmode)
  { alert(ubb_name+"���붨����ɫ�ı�\n\n����ǩ����Χ���ı���Ϊ�ƶ���ɫ��\n���磺[color=red]����ɫ[/color]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[color=" + color + "]" + range.text + "[/color]"; }
  else if (advmode)
  { AddTxt="[color="+color+"][/color]"; AddText(AddTxt); }
  else
  {
    txt=prompt(ubb_name+"������Ҫ���ó���ɫ "+color+" �����֣�","����");
    if(txt!=null) { AddTxt="[color="+color+"]"+txt; AddText(AddTxt); AddText("[/color]"); }
  }
}

function jk_ubb_center()
{
  if (helpmode)
  { alert(ubb_name+"���ж���\n\n����ǩ����Χ���ı����ж�����ʾ��\n���磺[align=center]���ݾ���[/align]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[center]" + range.text + "[/center]"; }
  else if (advmode)
  { AddTxt="[align=center][/align]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"������Ҫ���ж�������֣�","����");     
    if (txt!=null) { AddTxt="\r[align=center]"+txt; AddText(AddTxt); AddText("[/align]"); }        
  }
}

function jk_ubb_link()
{
  if (helpmode)
  { alert(ubb_name+"���볬������\n\n����һ���������ӣ�\n���磺\n[url]http://beyondest.com/[/url]\n[url=http://beyondest.com/]Beyondest.com[/url]"); }
  else if (advmode)
  { AddTxt="[url][/url]"; AddText(AddTxt); }
  else
  { 
    txt2=prompt(ubb_name+"������������ʾ�����֣����������ֱ����ʾ���ӣ�",""); 
    if (txt2!=null)
    {
      txt=prompt(ubb_name+"������ URL������http://beyondest.com/","http://");      
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
  { alert(ubb_name+"����ͼ��\n\n���ı��в���һ��ͼ��\n���磺[IMG]http://beyondest.com/images/logo.gif[/IMG]"); }
  else if (advmode)
  { AddTxt="[IMG][/IMG]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"������ͼ��� URL������http://beyondest.com/images/logo.gif","http://");    
    if(txt!=null) { AddTxt="\r[IMG]"+txt; AddText(AddTxt); AddText("[/IMG]");
    }       
  }
}

function jk_ubb_code()
{
  if (helpmode)
  { alert(ubb_name+"�������\n\n��������ű�ԭʼ���룡\n���磺[code]Beyondest.com[/code]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[code]" + range.text + "[/code]"; }
  else if (advmode)
  { AddTxt="\r[code][/code]"; AddText(AddTxt); }
  else
  {   
    txt=prompt(ubb_name+"������Ҫ����Ĵ��룡","");     
    if (txt!=null) { AddTxt="\r[code]"+txt; AddText(AddTxt); AddText("[/code]"); }
  }
}

function jk_ubb_flash()
{
  if (helpmode)
  { alert(ubb_name+"���� Flash\n\n���ı��в��� Flash ������\n���磺[FLASH="+ubb_w+","+ubb_h+"]http://www.Beyondrest.com/images/banner.swf[/FLASH]"); }
  else if (advmode)
  { AddTxt="[FLASH="+ubb_w+","+ubb_h+"][/FLASH]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"������ Flash �����Ĵ�С��",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"������ Flash �����ĵ�ַ��","http://");
      if(txt!=null) { AddTxt="\r[FLASH="+stxt+"]"+txt; AddText(AddTxt); AddText("[/FLASH]"); }
    }
  }
}

function jk_ubb_rm()
{
  if (helpmode)
  { alert(ubb_name+"���� RM\n\n���ı��в��� Realplay ��Ƶ�ļ���\n���磺[RM="+ubb_w+","+ubb_h+"]http://beyondest.com/images/test.ram[/RM]"); }
  else if (advmode)
  { AddTxt="[RM="+ubb_w+","+ubb_h+"][/RM]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"������ Realplay ��Ƶ�ļ��Ĵ�С��",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"������ Realplay ��Ƶ�ļ��ĵ�ַ rstp://�ȶ�֧��","http://");
      if(txt!=null) { AddTxt="\r[RM="+stxt+"]"+txt; AddText(AddTxt); AddText("[/RM]"); }
    }
  }
}

function jk_ubb_mp()
{
  if (helpmode)
  { alert(ubb_name+"���� MP\n\n���ı��в��� Windows Media Player ��Ƶ�ļ���\n���磺[MP="+ubb_w+","+ubb_h+"]http://beyondest.com/images/test.wmv[/MP]"); }
  else if (advmode)
  { AddTxt="[MP="+ubb_w+","+ubb_h+"][/MP]"; AddText(AddTxt); }
  else
  {
    stxt=prompt(ubb_name+"������ Windows Media Player ��Ƶ�ļ��Ĵ�С��",ubb_w+","+ubb_h);
    if (stxt!=null)
    {
      txt=prompt(ubb_name+"������ Windows Media Player ��Ƶ�ļ��ĵ�ַ�����ֵ�ַͷ��֧�֣�","http://");    
      if(txt!=null) { AddTxt="\r[MP="+stxt+"]"+txt; AddText(AddTxt); AddText("[/MP]"); }
    }
  }
}

function jk_ubb_underline()
{
  if (helpmode)
  { alert(ubb_name+"�����»���\n\n����ǩ����Χ���ı������»��ߣ�\n���磺[u]Beyondest.com[/u]"); }
  else if (document.selection && document.selection.type == "Text")
  { var range = document.selection.createRange(); range.text = "[u]" + range.text + "[/u]"; }
  else if (advmode)
  { AddTxt="[u][/u]"; AddText(AddTxt); }
  else
  {  
    txt=prompt(ubb_name+"������Ҫ���»��ߵ����֣�","����");
    if (txt!=null) { AddTxt="[u]"+txt; AddText(AddTxt); AddText("[/u]"); }         
  }
}

function setfocus() { document.<% Response.Write fn & "." & tn %>.focus(); }
-->
</script><table border=0 cellspacing=0 cellpadding=0><tr>
<td height=30><% Response.Write uw %><select onchange="javascript:jk_ubb_font(this.options[this.selectedIndex].value);" size=1 name=font>
<option value=����>����</option>
<option value=����>����</option>
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
<tr><td height=30><% Response.Write uw %><a href="javascript:jk_ubb_bold();"><img alt='��������ı�' src='images/ubb/bold.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_italicize();"><img alt='����б���ı�' src='images/ubb/italicize.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_underline();"><img alt='�����»���' src='images/ubb/underline.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_center();"><img alt='���ж���' src='images/ubb/center.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_link();"><img alt='���볬������' src='images/ubb/url.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_email();"><img alt='�����ʼ���ַ' src='images/ubb/email.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_image();"><img alt='����ͼ��' src='images/ubb/image.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_flash();"><img alt='���� flash' src='images/ubb/flash.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_code();"><img alt='�������' src='images/ubb/code.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_quote();"><img alt='��������' src='images/ubb/quote.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_rm();"><img alt='����Realplay��Ƶ�ļ�' src='images/ubb/rm.gif' border=0></a>&nbsp;
<a href="javascript:jk_ubb_mp();"><img alt='����Media Player�����ļ�' src='images/ubb/mp.gif' border=0></a></td></tr></table><%
End Sub

Sub frm_ubb_type() %>����ģʽ��<br><input onclick="javascript:jk_ubb_mode('2');" type=radio value=2 name=mode class=bg_1> ��ʾ����<br>
<input onclick="javascript:jk_ubb_mode('0');" type=radio value=0 name=mode class=bg_1 checked> ֱ�Ӳ���<br>
<input onclick="javascript:jk_ubb_mode('1');" type=radio value=1 name=mode class=bg_1> ������Ϣ<%
End Sub

Sub frm_word_size(fn,tn,ws,wv) %>
<a href="javascript:checklength(document.<% Response.Write fn %>);">[�������]</a>
<script language=JavaScript>
<!--
function checklength(theform)
{
  var postmaxchars=<% Response.Write ws %>;
  var postchars=theform.<% Response.Write tn %>.value.length;
  var post_1="�������<% Response.Write wv %>�Ѿ�����ϵͳ��������ֵ��\n��ɾ������<% Response.Write wv %>��";
  var message="";
  if (postmaxchars != 0) { message = "ϵͳ����"+postmaxchars+"KB��Լ"+postmaxchars*1024+"���ַ���"; }
  if (postmaxchars*1024>=postchars)
  {
    var postc=postmaxchars*1024-postchars;
    post_1="�����������룺"+postc+"���ַ���Լ"+Math.round(postc/1024)+"KB��";
  }
  alert("�������<% Response.Write wv %>��ͳ����Ϣ���£�\n\n��ǰ���ȣ�"+postchars+"���ַ���Լ"+Math.round(postchars/1024)+"KB��\n"+message+"\n"+post_1);
}
-->
</script><%
End Sub

Sub frm_topic(fn,tn) %><select onchange="document.<% Response.Write fn & "." & tn %>.focus(); document.<% Response.Write fn & "." & tn %>.value = this.options[this.selectedIndex].value + document.<% Response.Write fn & "." & tn %>.value;"> 
<option value='' selected>ѡ��</option>
<option value=[ԭ��]>[ԭ��]</option>
<option value=[ת��]>[ת��]</option>
<option value=[Hack]>[Hack]</option>
<option value=[שͷ]>[שͷ]</option>
<option value=[�ҵ�]>[�ҵ�]</option>
<option value=[��ˮ]>[��ˮ]</option>
<option value=[����]>[����]</option>
<option value=[����]>[����]</option> 
<option value=[����]>[����]</option>
<option value=[�Ƽ�]>[�Ƽ�]</option> 
<option value=[����]>[����]</option>
<option value=[����]>[����]</option> 
<option value=[ע��]>[ע��]</option>
<option value=[��ͼ]>[��ͼ]</option> 
<option value=[����]>[����]</option>
<option value=[����]>[����]</option>
<option value=[����]>[����]</option>
</select><%
End Sub %>