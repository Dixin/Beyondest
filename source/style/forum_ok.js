<!--
//******************************************************************
//
//                        Beyondest V3.6 Demo版
//
//                      网址：http://www.beyondest.com
//
//******************************************************************

//调用方法:onsubmit="frm_submitonce(this);"
function frm_submitonce(theform)
{	//if IE 4+ or NS 6+
  if (document.all||document.getElementById)
  {	//screen thru every element in the form, and hunt down "submit" and "reset"
    for (i=0;i<theform.length;i++)
    {
      var tempobj=theform.elements[i]
      if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
      //disable em
      tempobj.disabled=true
    }
  }
}

function frm_quicksubmit(eventobject)
{
  if (event.keyCode==13 && event.ctrlKey)
  //判断是不是CTRL+ENTER
  write_frm.wsubmit.click();
}
-->