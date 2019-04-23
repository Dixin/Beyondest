<!--
// ====================
// Beyondest.Com v3.6.1
// http://beyondest.com
// ====================

function load_tree(f_id,v_id){
  var targetImg =eval("document.all.followImg" + v_id);
  var targetDiv =eval("document.all.follow" + v_id);
  if (targetImg.src.indexOf("nofollow")!=-1){return false;}
    if ("object"==typeof(targetImg)){
      if (targetDiv.style.display!='block'){
        targetDiv.style.display="block";
        targetImg.src="images/small/fk_minus.gif";
        if (targetImg.loaded=="no"){
          document.frames["hiddenframe"].location.replace("forum_loadtree.asp?forum_id="+f_id+"&view_id="+v_id);
        }
      }else{
      targetDiv.style.display="none";
      targetImg.src="images/small/fk_plus.gif";
    }
  }
}
-->