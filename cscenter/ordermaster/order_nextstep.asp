<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : �̻� ����
'			2020.01.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim ojumun, IsAppExists, IsTempOrderExists, ix, preipkumdiv, isfailorder, accountdiv, menupos, cancelyn, acctdivchg
    menupos = requestcheckvar(getNumeric(request("menupos")),10)

set ojumun = new COrderMaster
    ojumun.FRectOrderSerial = request("orderserial")
    ojumun.QuickSearchOrderMaster

IF ojumun.FTotalCount < 1 THEN
    response.write "�������� �ֹ����� �ƴմϴ�."
    dbget.close() : response.end
end if

cancelyn = ojumun.FOneItem.Fcancelyn
accountdiv = ojumun.FOneItem.Faccountdiv

IsAppExists = False
IsTempOrderExists = False
isfailorder = false

'// ���γ��� �ִ���
IsAppExists = ojumun.getAppLogExists()
IsTempOrderExists = ojumun.getTempOrderExists()

if C_CSPowerUser or C_ADMIN_AUTH then
    ' �ֹ����� or �ֹ���� �ϰ�쿡��
    if (ojumun.FOneItem.Fipkumdiv = "1") or (ojumun.FOneItem.Fipkumdiv = "0") then
        isfailorder=true
    end if
end if

%>
<script type='text/javascript'>

function SubmitForm() {
    var IsTempOrderExists = <%= LCase(IsTempOrderExists) %>;
    var IsAppExists = <%= LCase(IsAppExists) %>;
    var IsallAppExists = <%= LCase(IsAppExists or IsTempOrderExists) %>;
    var ipkumdatediff = <%= datediff("d",left(ojumun.FOneItem.Fregdate,10),date()) %>;

    if (frm.cancelyn.value!='N'){
        alert('��ҵ� ���� �Դϴ�.');
        return;
    }

    <%
    ' ���� �ֹ� ���� �ֹ����� ����
    if isfailorder then
    %>
        // üũ�ϸ� �ȵ�. �ӽ��ֹ������� �ȵ��� ���̽� ����.
        //if (IsTempOrderExists==false){
        //    alert('�ӽ� �ֹ������� �����ϴ�.');
        //    return;
        //}
        if (frm.checkappexists.checked){
            if (IsAppExists==false){
                if (ipkumdatediff>2){
                    alert('�������� 2���� �ʰ� �Ǿ�����, PG�翡�� �Ѿ�� ���γ����� ���� �ֹ� �Դϴ�.\n�������� ���� �Ͻðų�, PG����γ���üũ�� ������ �ּ���.');
                    return;
                }
            }
        }
        if (frm.acctdiv.value==''){
            alert('��������� ������ �ּ���.');
            return;
        }
        if (!(frm.acctdiv.value=='100' || frm.acctdiv.value=='20' || frm.acctdiv.value=='400' || frm.acctdiv.value=='7')){
            alert('��������� �ſ�ī��,�ǽð���ü,�ڵ�������,������ �� ���ð��� �մϴ�.');
            return;
        }
        if (frm.paygatetid.value==''){
            alert('PG�� TID�� �Է��� �ּ���.');
            return;
        }
        if (frm.paygatetid.value.length<10){
            alert('�������� PG�� TID�� �ƴմϴ�.');
            frm.paygatetid.focus();
            return;
        }

        if(frm.authcode.value!=''){
            if (!IsDouble(frm.authcode.value)){
                alert('���ι�ȣ�� ���ڸ� �����մϴ�.');
                frm.authcode.focus();
                return;
            }
            if (frm.authcode.value.length>10){
                alert('�������� ���ι�ȣ�� �ƴմϴ�.');
                frm.authcode.focus();
                return;
            }
        }

        if (confirm("�����ֹ����� ���� ���� �Ͻðڽ��ϱ�?") == true) {
            frm.submit();
        }

    <% else %>
        //if ((frm.ipkumdiv.value!='2')&&(frm.ipkumdiv.value!='5')&&(frm.ipkumdiv.value!='6')){
        if ((frm.ipkumdiv.value != '2') && (IsallAppExists == false)) {
            alert('�ֹ� ���������� ���� ���·� ������ ���� �մϴ�.');
            return;
        }

        if (confirm("���� ���·� ���� �Ͻðڽ��ϱ�?") == true) {
            frm.submit();
        }
    <% end if %>
}

</script>

<form name="frm" onsubmit="return false;" action="/cscenter/ordermaster/order_info_edit_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="ipkumdivnextstep">
<input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.Forderserial %>">
<input type="hidden" name="ipkumdiv" value="<%= ojumun.FOneItem.Fipkumdiv %>">
<input type="hidden" name="userid" value="<%= ojumun.FOneItem.Fuserid %>">
<input type="hidden" name="cancelyn" value="<%= ojumun.FOneItem.Fcancelyn %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr>
                <td width="160">
                    <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���� ���� ����</b>
                </td>
                <td align="right">
                    <input type="button" value="���� ���� ����" class="csbutton" onclick="SubmitForm();">
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
    <td width="150" >�������</td>
    <td width="150" >��������</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF">
    <td width="150"><font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font></td>
    <%
    preipkumdiv = ojumun.FOneItem.Fipkumdiv
    if (preipkumdiv="2") then
        ojumun.FOneItem.Fipkumdiv = "4"
    elseif (preipkumdiv="4") then
        ojumun.FOneItem.Fipkumdiv = "5"
    elseif (preipkumdiv="5") then
        ojumun.FOneItem.Fipkumdiv = "6"
    elseif (preipkumdiv="6") then
        ojumun.FOneItem.Fipkumdiv = "7"
    elseif IsAppExists or IsTempOrderExists then
        ojumun.FOneItem.Fipkumdiv = "4"
    end if
    %>
    <td width="150"><font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font></td>
    <%
    ojumun.FOneItem.Fipkumdiv = preipkumdiv
    %>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF">
    <td colspan="2"><%= JumunMethodName(accountdiv) %> <font color="<%= ojumun.FOneItem.CancelYnColor %>"><%= ojumun.FOneItem.CancelYnName %></font></td>
</tr>

<% if (ojumun.FOneItem.Fipkumdiv="2") and (ojumun.FOneItem.Fcancelyn="N") then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="emailok" checked >�Ա�Ȯ�θ��Ϲ߼�</td>
        <td ><%= ojumun.FOneItem.Fbuyemail %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="smsok" checked >�Ա�Ȯ��SMS�߼�</td>
        <td><%= ojumun.FOneItem.Fbuyhp %></td>
    </tr>
<% end if %>

<% if isfailorder then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>[�ʼ�]�������</td>
        <td align="left">
            <%
            'if accountdiv<>"" and not(isnull(accountdiv)) and accountdiv<>"900" then
            '    IF not(accountdiv="100" or accountdiv="20" or accountdiv="400") THEN
            '        acctdivchg=" onFocus='this.initialSelect = this.selectedIndex;' onChange='this.selectedIndex = this.initialSelect;'"
            '    end if
            'end if
            %>
            <% DrawJumunMethod "acctdiv",accountdiv,acctdivchg %>
            &nbsp;&nbsp;
            <input type="checkbox" name="checkappexists" value="on" checked >PG����γ���üũ
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>[�ʼ�]PG�� TID</td>
        <td align="left"><input type="text" name="paygatetid" value="<%= ojumun.FOneItem.fpaygatetid %>" size="50" maxlength="50"></td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>���ι�ȣ</td>
        <td align="left"><input type="text" name="authcode" value="<%= ojumun.FOneItem.FAuthcode %>" size="32" maxlength="32"></td>
    </tr>
<% end if %>

</table>
</form>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
