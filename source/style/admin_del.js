<!--

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function suredel(del_num){
  var length = del_num;
  var ifcheck;
  var cf;
  
  if (length == 0 ){ return false; }
  if (length ==1 )
    {
      if (document.del_form.del_id.checked) 
        {
          ifcheck=true;
        }
    }
  if (length>1)
    {
      for (var i = 0; i < length; i++)
        { 
          if (document.del_form.del_id[i].checked)
	    {  
	      cf = window.confirm("记录被删除后将无法恢复！您确定吗？");
	      if (cf) 
		{ return true; } 
	      else
		{ return false; }
	    }
	  else
	    { ifcheck=false; }
        }
    }
	
  if (ifcheck)
    {
      cf = window.confirm("记录被删除后将无法恢复！您确定吗？");
      if (cf) 
        { return true; } 
      else
        { return false; }
    }
  else
    {
      window.alert("没有选择任何记录！");
      return false;
    }
}

function selectall(del_num)
{
    var length = del_num;
    document.del_form.del_all.checked = document.del_form.del_all.checked|0;
    if (length == 0 ){
          return;
    }
    if (length ==1 )
    {
       document.del_form.del_id.checked=document.del_form.del_all.checked ;
    }
    if (length>1)
    {
      for (var i = 0; i < length; i++)
       {
        document.del_form.del_id[i].checked=document.del_form.del_all.checked;         
       }
    }
}
function unselectall()
{
    if(document.del_form.del_all.checked){
	document.del_form.del_all.checked = document.del_form.del_all.checked&0;
    } 	
}

//-->