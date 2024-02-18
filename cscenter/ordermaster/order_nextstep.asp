<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 이상구 생성
'			2020.01.17 한용민 수정
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
    response.write "정상적인 주문건이 아닙니다."
    dbget.close() : response.end
end if

cancelyn = ojumun.FOneItem.Fcancelyn
accountdiv = ojumun.FOneItem.Faccountdiv

IsAppExists = False
IsTempOrderExists = False
isfailorder = false

'// 승인내역 있는지
IsAppExists = ojumun.getAppLogExists()
IsTempOrderExists = ojumun.getTempOrderExists()

if C_CSPowerUser or C_ADMIN_AUTH then
    ' 주문실패 or 주문대기 일경우에만
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
        alert('취소된 내역 입니다.');
        return;
    }

    <%
    ' 오류 주문 정상 주문으로 변경
    if isfailorder then
    %>
        // 체크하면 안됨. 임시주문내역이 안들어가는 케이스 많음.
        //if (IsTempOrderExists==false){
        //    alert('임시 주문내역이 없습니다.');
        //    return;
        //}
        if (frm.checkappexists.checked){
            if (IsAppExists==false){
                if (ipkumdatediff>2){
                    alert('결제일이 2일이 초과 되었으나, PG사에서 넘어온 승인내역이 없는 주문 입니다.\n개발팀에 문의 하시거나, PG사승인내역체크를 해제해 주세요.');
                    return;
                }
            }
        }
        if (frm.acctdiv.value==''){
            alert('결제방법을 선택해 주세요.');
            return;
        }
        if (!(frm.acctdiv.value=='100' || frm.acctdiv.value=='20' || frm.acctdiv.value=='400' || frm.acctdiv.value=='7')){
            alert('결제방법은 신용카드,실시간이체,핸드폰결제,무통장 만 선택가능 합니다.');
            return;
        }
        if (frm.paygatetid.value==''){
            alert('PG사 TID를 입력해 주세요.');
            return;
        }
        if (frm.paygatetid.value.length<10){
            alert('정상적인 PG사 TID가 아닙니다.');
            frm.paygatetid.focus();
            return;
        }

        if(frm.authcode.value!=''){
            if (!IsDouble(frm.authcode.value)){
                alert('승인번호는 숫자만 가능합니다.');
                frm.authcode.focus();
                return;
            }
            if (frm.authcode.value.length>10){
                alert('정상적인 승인번호가 아닙니다.');
                frm.authcode.focus();
                return;
            }
        }

        if (confirm("정상주문으로 변경 진행 하시겠습니까?") == true) {
            frm.submit();
        }

    <% else %>
        //if ((frm.ipkumdiv.value!='2')&&(frm.ipkumdiv.value!='5')&&(frm.ipkumdiv.value!='6')){
        if ((frm.ipkumdiv.value != '2') && (IsallAppExists == false)) {
            alert('주문 접수내역만 다음 상태로 진행이 가능 합니다.');
            return;
        }

        if (confirm("다음 상태로 진행 하시겠습니까?") == true) {
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
                    <img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>다음 상태 진행</b>
                </td>
                <td align="right">
                    <input type="button" value="다음 상태 진행" class="csbutton" onclick="SubmitForm();">
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("topbar") %>">
    <td width="150" >현재상태</td>
    <td width="150" >다음상태</td>
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
        <td><input type="checkbox" name="emailok" checked >입금확인메일발송</td>
        <td ><%= ojumun.FOneItem.Fbuyemail %>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td><input type="checkbox" name="smsok" checked >입금확인SMS발송</td>
        <td><%= ojumun.FOneItem.Fbuyhp %></td>
    </tr>
<% end if %>

<% if isfailorder then %>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>[필수]결제방법</td>
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
            <input type="checkbox" name="checkappexists" value="on" checked >PG사승인내역체크
        </td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>[필수]PG사 TID</td>
        <td align="left"><input type="text" name="paygatetid" value="<%= ojumun.FOneItem.fpaygatetid %>" size="50" maxlength="50"></td>
    </tr>
    <tr height="25" align="center" bgcolor="#FFFFFF">
        <td>승인번호</td>
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
