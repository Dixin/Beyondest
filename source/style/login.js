<!--
// ====================
// Beyondest.Com v3.6.1
// http://beyondest.com
// ====================

function login_true() {
    if (login_frm.username.value == "") {
        alert("请输入您在本站注册时的 用户名称 ！");
        login_frm.username.focus();
        return false;
    }
    if (login_frm.password.value == "") {
        alert("请输入您在本站注册时的 登陆密码 ！");
        login_frm.password.focus();
        return false;
    }
}
-->