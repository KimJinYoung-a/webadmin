<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim gubun,idx

gubun   = requestCheckvar(request("gubun"),16)
idx     = requestCheckvar(request("id"),10)


dim ojungsanmaster, IsCommissionTax, itemvatyn
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = idx
ojungsanmaster.JungsanMasterList

if (ojungsanmaster.FResultCount<1) then
    dbget.Close(): response.end
end if

IsCommissionTax = ojungsanmaster.FItemList(0).IsCommissionTax
itemvatyn = ojungsanmaster.FItemList(0).Fitemvatyn
set ojungsanmaster=Nothing

if (IsCommissionTax and (gubun="upche" or gubun="witaksell")) then
    rw "수수료 정산인경우 배송비, 기타출고 정산만 기타내역으로 가능"
    ''dbget.Close() : response.end
end if
%>

<script language='javascript'>
function adddata(frm){
	if (frm.itemname.value.length<1){
		alert('내용을 입력하세요.');
		frm.itemname.focus();
		return;
	}

	if (frm.itemno.value.length<1){
		alert('갯수를 입력하세요.');
		frm.itemno.focus();
		return;
	}

	if (!IsDigit(frm.itemno.value)){
		alert('갯수는 숫자만 가능합니다.');
		frm.itemno.focus();
		return;
	}

	if (frm.sellcash.value.length<1){
		alert('판매가를 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return;
	}
    <% if (IsCommissionTax) then %>
	if (frm.reducedprice.value.length<1){
		alert('실매출을 입력하세요.');
		frm.reducedprice.focus();
		return;
	}

	if (!IsInteger(frm.reducedprice.value)){
		alert('실매출은 숫자만 가능합니다.');
		frm.reducedprice.focus();
		return;
	}

	if (frm.commission.value.length<1){
		alert('수수료를 입력하세요.');
		frm.commission.focus();
		return;
	}

	if (!IsInteger(frm.commission.value)){
		alert('수수료는 숫자만 가능합니다.');
		frm.commission.focus();
		return;
	}
    <% end if %>

	if (frm.suplycash.value.length<1){
		alert('매입가를 입력하세요.');
		frm.suplycash.focus();
		return;
	}

	if (!IsInteger(frm.suplycash.value)){
		alert('매입가는 숫자만 가능합니다.');
		frm.suplycash.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>기타내역추가</strong></font>
        </td>
        <td align="right">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmadd" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="mode" value="etcadd">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="idx" value="<%= idx %>">
    <input type="hidden" name="itemvatyn" value="<%= itemvatyn %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>내용</td>
		<td width="40">수량</td>
		<td width="80">판매가</td>
		<td width="80">실매출</td>
		<td width="80">수수료</td>
		<td width="80">공급가</td>
    </tr>
    <tr bgcolor="#FFFFFF">
		<td><input type="text" name="itemname" value="" size="60"></td>
		<td><input type="text" name="itemno" value="1" size="3" style="text-align:center"></td>
		<td><input type="text" name="sellcash" value="" size="8" style="text-align:right"></td>
		<td><input type="text" name="reducedprice" value="" size="8" style="text-align:right" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%>></td>
		<td><input type="text" name="commission" value="" size="8" style="text-align:right" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%>></td>
		<td><input type="text" name="suplycash" value="" size="8" style="text-align:right"></td>
    </tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="내역 추가" onclick="adddata(frmadd)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->