<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_UTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib_UTF8.asp"-->
<!-- #include virtual="/cscenter/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim ojumun, orderserial, AlertMsg, IsOldOrder, ix
	orderserial = requestCheckVar(request("orderserial"),11)

set ojumun = new COrderMaster
	ojumun.FRectOrderSerial = orderserial
	ojumun.QuickSearchOrderMaster

	if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	    ojumun.FRectOldOrder = "on"
	    ojumun.QuickSearchOrderMaster

	    if (ojumun.FResultCount>0) then
	        IsOldOrder = true
	        AlertMsg = "6개월 이전 주문입니다."
	    end if

	end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

window.resizeTo(500,350);

function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".reqzipcode").value = post1 + "-" + post2;

    eval(frmname + ".reqzipaddr").value = addr;
    eval(frmname + ".reqaddress").value = dong;
}

document.title = "배송 정보";

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="order_info_edit_process_UTF8.asp">
<input type="hidden" name="mode" value="modifyreceiverinfo">
<input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.FOrderSerial %>">
<input type="hidden" name="acctdiv" value="<%=ojumun.FOneItem.FAccountDiv%>">
<input type="hidden" name="paygatetid" value="<%=ojumun.FOneItem.Fpaygatetid%>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>배송 정보</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="저장하기" class="csbutton" onclick="javascript:SubmitForm();" >
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">수령인명</td>
    <td><input type="text" class="text" name="reqname" id="[on,off,1,32][수령인명]" value="<%= ojumun.FOneItem.FReqName %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
    <td><input type="text" class="text" name="reqphone" id="[on,off,1,24][전화번호]" value="<%= ojumun.FOneItem.FReqPhone %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
    <td><input type="text" class="text" name="reqhp" id="[on,off,1,16][핸드폰]" value="<%= ojumun.FOneItem.FReqHp %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td rowspan="3" valign="top" bgcolor="<%= adminColor("topbar") %>">수령주소</td>
    <td>
        <input type="text" class="text" name="reqzipcode" value="<%= ojumun.FOneItem.FReqZipCode %>" size="7" readonly><!-- id="[on,off,7,7][우편번호]" -->
        <input type="button" class="button" value="검색" onClick="FnFindZipNew('frm','A')">
        <input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frm','A')">
        <% '<input type="button" class="button" value="검색(구)" onClick="PopSearchZipcode('frm')"> %>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td ><input type="text" class="text" name="reqzipaddr" id="[on,off,1,64][주소]" size="35" value="<%= ojumun.FOneItem.FReqZipAddr %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td>
        <input type="text" class="text" name="reqaddress" id="[on,off,1,200][주소]" size="35" value="<%= ojumun.FOneItem.FReqAddress %>">
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">기타사항</td>
    <td>
        <textarea class="textarea" rows="3" cols="35" name="comment" id="[off,off,off,off][기타사항]"><%= ojumun.FOneItem.FComment %></textarea>
	</td>
</tr>
</table>

<script type="text/javascript">
    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>

<%
session.codePage = 949

set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
