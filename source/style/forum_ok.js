<!--
// ====================
// Beyondest.Com v3.6.1
// http://beyondest.com
// ====================

//���÷���:onsubmit="frm_submitonce(this);"
function frm_submitonce(theform) {	//if IE 4+ or NS 6+
    if (document.all || document.getElementById) {	//screen thru every element in the form, and hunt down "submit" and "reset"
        for (i = 0; i < theform.length; i++) {
            var tempobj = theform.elements[i]
            if (tempobj.type.toLowerCase() == "submit" || tempobj.type.toLowerCase() == "reset")
                //disable em
                tempobj.disabled = true
        }
    }
}

function frm_quicksubmit(eventobject) {
    if (event.keyCode == 13 && event.ctrlKey)
        //�ж��ǲ���CTRL+ENTER
        write_frm.wsubmit.click();
}
-->