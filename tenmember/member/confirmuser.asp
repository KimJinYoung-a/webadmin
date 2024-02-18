<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��Ʈ��� �������� ���� Ȯ�� ����
' History : 2018.08.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim param, empno, userid, loginType
param = requestCheckVar(request("pflag"),1)

userid = session("ssBctId")     '���̵�α���
empno = session("ssBctSn")      '����α��� (���̵�α��νÿ��� ����� ����)

'�α��� ���� ����
if userid<>"" then
    loginType = "id"
else
    loginType = "emp"
end if
%>
<script type="text/javascript" src="/js/jquery-2.2.2.min.js"></script>
<script type="text/javascript">
$(function(){
    $("#userpass").focus();
});

function submitForm() {
    var frm = document.frm;

    if(frm.password.value.length<1) {
        alert("��й�ȣ�� �Է����ּ���.");
        frm.password.focus();
        return false;
    }

    frm.submit();
}
function jsMemberPayView(){
 	var winPrint = window.open("/admin/member/tenbyten/tenbyten_pay_view.asp?sEN=<%=empno%>","prtWT","width=1300,height=600,scrollbars=yes,resizable=yes");
 	winPrint.focus();
}
function jsPopView(sPage){
    var winNew = window.open(sPage,"popNew","width=880, height=600,scrollbars=yes, resizable=yes");
    winNew.focus();
}

</script>
<div style="width:500px;text-align:center; margin-top:20px;">
    <p style="font-weight:bold; padding:10px 0;">������ �����ϰ� ��ȣ�ϱ� ���� ������ �ٽ� �� �� Ȯ���մϴ�.</p>
    <form name="frm" method="POST" action="/tenmember/member/doConfirmuser.asp">
    <input type="hidden" name="menupos" value="<%=menupos%>" />
    <input type="hidden" name="loginType" value="<%=loginType%>" />
    <table width="400" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
    <tr bgcolor="#FFFFFF">
        <td width="100" bgcolor="<%=adminColor("tabletop")%>" align="center">�α��� ���</td>
        <td><strong><%=chkIIF(loginType="id","���̵�","���")%></strong></td>
    </tr>
    <% if loginType="id" then %>
    <tr id="trId" bgcolor="#FFFFFF">
        <td bgcolor="<%=adminColor("tabletop")%>" align="center">���̵�</td>
        <td align="center">
            <input type="text" name="userid" value="<%=userid%>" maxlength="32" class="text_ro" readonly="readonly" style="width:100%;" />
        </td>
    </tr>
    <% else %>
    <tr id="trEmp" bgcolor="#FFFFFF">
        <td bgcolor="<%=adminColor("tabletop")%>" align="center">���</td>
        <td align="center">
            <input type="text" name="empno" value="<%=empno%>" maxlength="14" class="text_ro" readonly="readonly" style="width:100%;" />
        </td>
    </tr>
    <% end if %>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%=adminColor("tabletop")%>" maxlength="32" align="center">��й�ȣ</td>
        <td align="center">
            <input type="password" id="userpass" name="password" value="" class="text" style="width:100%;" onKeyPress="if(event.keyCode == 13){submitForm();}" />
        </td>
    </tr>
    </table>
    </form>
    <p style="padding:5px 0;"><input type="button" value="�α���" class="button_auth" onclick="submitForm()" /></p>
</p>
<% if session("ssAdminPOSITsn") = "13" then %>
<p style="padding:5px 0;"><a href="javascript:jsMemberPayView();"><font style="color:red;font-size:15px">�����޿� ����</font></a></p>
<p style="padding:5px 0;"><a href="javascript:jsPopView('/admin/approval/eapp/epop/regeappform.asp');"><font style="color:red;font-size:15px">���ڰ��� �ۼ�</font></a></p>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->