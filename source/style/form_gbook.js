<!--


function isCharsInBag(inputchar, bagchar) {
    var ii, cc;
    for (ii = 0; ii < inputchar.length; ii++) {
        cc = inputchar.charAt(ii);//�ַ���inputchar�е��ַ�
        if (bagchar.indexOf(cc) > -1) {
            return "no";
        }
        else {
            return "yes";
        }
    }
}
//-----------------------------------------------------------------------------
function check(write_frm) {
    var errorcharname = "><,[>{}?/+=|\\'\":;~!#$%()`@"
    var errorcharqq = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz><,[>{}?/+=|\\'\":;~!#$%()`"

    var wr_name = write_frm.wrname.value
    if (wr_name == "") {
        alert("�㻹û��ȫ����������Ϣ��\r\n��� ���� �Ǳ���Ҫ�ġ�");
        return false;
    }
    var errorname = isCharsInBag(wr_name, errorcharname)
    if (errorname == "no") {
        alert("��� ���� ���ú��������ַ���\r\n  ><,[{}?/+=|\\'\":;~!#$%()`");
        return false;
    }

    var wr_qq = write_frm.wrqq.value
    if (write_frm.wrqq.value !== "") {
        var errorqq = isCharsInBag(wr_qq, errorcharqq)
        if (errorqq == "no") {
            alert("��� QQ ֻ�������֣�\r\n���������롣");
            return false
        }
    }

    var wr_email = write_frm.wremail.value
    if (wr_email !== "") {
        var AtSym = wr_email.indexOf('@')
        var Period = wr_email.lastIndexOf('.')
        var Space = wr_email.indexOf(' ')
        var Length = wr_email.length - 1 // Array is from 0 to length-1 
        if ((AtSym < 1) || (Period <= AtSym + 1) || (Period == Length) || (Space != -1))
        // '@' cannot be in first position 
        // Must be atleast one valid char btwn '@' and '.' 
        // Must be atleast one valid char after '.' 
        // No empty spaces permitted 
        {
            alert('��� eMail��ַ ��ʽ���ԣ�\r\n���������롣')
            return false
        }
    }

    var wr_whe = write_frm.wrwhe.value
    if (wr_whe !== "") {
        var errorwhe = isCharsInBag(wr_whe, errorcharname)
        if (errorwhe == "no") {
            alert("��� ���� ���ú��������ַ���\r\n  ><,[{}?/+=|\\'\":;~!#$%()`");
            return false
        }
    }

    var wr_topic = write_frm.wrtopic.value
    if (wr_topic == "") {
        alert("�㻹û��ȫ����������Ϣ��\r\n��� ���� �Ǳ���Ҫ�ġ�");
        return false;
    }

    var wr_word = write_frm.wrword.value
    if (wr_word == "") {
        alert("�㻹û��ȫ����������Ϣ��\r\n��� �������� �Ǳ���Ҫ�ġ�");
        return false;
    }
    //document.write_frm.submit()
}
//-----------------------------------------------------------------------
function reset(write_frm) {
    if (confirm("�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?")) {
        return true;
    }
    return false;
}

function showimage() {
    document.images.wrimg.src = "images/face/" + document.write_frm.wrface.options[document.write_frm.wrface.selectedIndex].value + ".gif";
}
-->
