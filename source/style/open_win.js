// ====================
// Beyondest.Com v3.6.1
// http://beyondest.com
// ====================

function open_win(url,name,width,height,scroll)
{
var Left_size = (screen.width) ? (screen.width-width)/2 : 0;
var Top_size = (screen.height) ? (screen.height-height)/2 : 0;
var open_win=window.open(url,name,'width=' + width + ',height=' + height + ',left=' + Left_size + ',top=' + Top_size + ',toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=' + scroll + ',resizable=no' );
}