<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터
' History : 2009.04.17 이상구 생성
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim ojumun, orderserial, ix
dim AlertMsg, IsOldOrder
	orderserial= requestCheckVar(request("orderserial"),11)

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

function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

document.title = "구매자정보";

</script>

<!-- 구매자정보 -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="order_info_edit_process.asp">
<input type="hidden" name="mode" value="modifybuyerinfo">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>구매자 정보</b>
			    </td>    				    
			    <td align="right">
			        <input type="button" value="저장하기" class="csbutton" onClick="SubmitForm();" <%= chkIIF(IsOldOrder,"disabled","") %>>
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">구매자ID....</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="userid" id="[off,off,off,off][구매자ID]" value="<%= ojumun.FOneItem.FUserID %>" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">주문번호</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="orderserial" id="[off,off,off,off][주문번호]" value="<%= ojumun.FOneItem.FOrderSerial %>" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">구매자명</td>
    <td bgcolor="#FFFFFF">
        <input type="text" class="text" name="buyname" id="[on,off,1,32][구매자명]" value="<%= ojumun.FOneItem.FBuyName %>" size="8" >
        <font color="<%= getUserLevelColorByDate(ojumun.FOneItem.fUserLevel, left(ojumun.FOneItem.Fregdate,10)) %>">
        <%= getUserLevelStrByDate(ojumun.FOneItem.fUserLevel, left(ojumun.FOneItem.Fregdate,10)) %></a></font>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">전화번호</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyphone" id="[on,off,1,16][구매자전화번호]" value="<%= ojumun.FOneItem.FBuyPhone %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">핸드폰</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyhp" id="[on,off,1,16][구매자핸드폰]" value="<%= ojumun.FOneItem.FBuyHp %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">이메일</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="buyemail" id="[on,off,1,128][이메일]" value="<%= ojumun.FOneItem.FBuyEmail %>" ></td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">결재방법</td>
    <td bgcolor="#FFFFFF">
    	<%= ojumun.FOneItem.JumunMethodName %> / <font color="<%= ojumun.FOneItem.IpkumDivColor %>"><%= ojumun.FOneItem.IpkumDivName %></font>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">입금자명</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="accountname" id="[on,off,1,16][입금자명]" value="<%= ojumun.FOneItem.FAccountName %>" ></td>
</tr>
</form>
</table>
<!-- 구매자정보 -->

<script type="text/javascript">
    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>    

<%
set ojumun = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->