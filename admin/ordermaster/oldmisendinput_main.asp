<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%

dim deliveryno, oldmisend

dim orderserial,obalju
dim didx, mode

dim i

didx = request("didx")
mode = request("mode")
deliveryno = request("deliveryno")
orderserial = request("orderserial")

if (orderserial = "") then
        orderserial = "-"
end if

if (len(orderserial)=12) and (left(orderserial,2)="00") then
	orderserial = Right(orderserial,11)
end if


dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.FRectDeliveryNo = deliveryno
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

set obalju = New CBalju
obalju.FRectOrderSerial = orderserial

obalju.GetMiSendOrderDetail
%>
<script language='javascript'>
function DelMiSend(frm){
	var ret = confirm('삭제하시겠습니까?');

	if (ret){
		frm.mode.value="del";
		frm.submit();
	}
}

function SaveMiSend(frm){
	if (frm.ipgodate.value.length>0){
		if (frm.ipgodate.value.length!=10){
			alert('입고예정일을 정확히 입력하세요.');
			frm.ipgodate.focus();
			return;
		}
	}

	var ret = confirm('저장하시겠습니까?');

	if (ret){
		frm.submit();
	}
}
function calender_open() {
       document.all.cal.style.display="";
}
</script>
<style type="text/css">
<!--
td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "굴림", "돋움";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" >
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	주문번호 : <input type="text" name="orderserial" value="<%= request("orderserial") %>" size="12" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
				&nbsp;&nbsp;
				송장번호 : <input type="text" name="deliveryno" value="<%= request("deliveryno") %>" size="12" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
				&nbsp;&nbsp;
	        	<% if omasterwithcs.FOneItem.FCancelyn="Y" then %>
				<b><font color="#CC3333">취소된 주문건입니다.</font></b>
				<script language='javascript'>alert('취소된 거래 입니다.');</script>
				<% else %>
				정상 주문건입니다.
				<% end if %>
	        </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr bgcolor="DDDDFF" align="center">
<!--	<td width="20">-</td>	-->
		<td width="40">상품ID</td>
		<td width="50">이미지</td>
		<td>상품명</td>
		<td>상품옵션</td>
		<td width="50">주문수량</td>
		<td width="50">부족수량</td>
		<td width="30">취소<br>삭제</td>
		<td width="80">업체배송일</td>
		<td width="80">미배송사유</td>
		<td width="80">출고예정일</td>
	</tr>
<form name="frmmisend" method="post" action="domisendinput.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">
	<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
	
	<% if obalju.FBaljuDetailList(i).FItemLackNo<>0 then %>
	<tr bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr bgcolor="FFFFFF">
	<% end if %>
<!--	<td ><input type="checkbox" name="didxarr" value="<%=obalju.FBaljuDetailList(i).FDetailIDx %>" <% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then response.write "checked" end if %>></td>	-->
		<input type="hidden" name="tmporderserial" value="<%= orderserial %>">
		<input type="hidden" name="itemid" value="<%= obalju.FBaljuDetailList(i).FItemID %>">
		<input type="hidden" name="itemname" value="<%= DDotFormat(replace(obalju.FBaljuDetailList(i).FItemName,Chr(22),""),16) %>">
		<input type="hidden" name="itemoptionname" value="<%= replace(obalju.FBaljuDetailList(i).FItemOptionName,Chr(22),"") %>">
		<input type="hidden" name="makerid" value="<%= obalju.FBaljuDetailList(i).FMakerid %>">
		<input type="hidden" name="itemno" value="<%= obalju.FBaljuDetailList(i).FItemNo %>">

		<% if obalju.FBaljuDetailList(i).IsUpcheBeasong then %>
		<td ><font color="red"><%= obalju.FBaljuDetailList(i).FItemID %></font></td>
		<% else %>
		<td ><%= obalju.FBaljuDetailList(i).FItemID %></td>
		<% end if %>
		<td ><img src="<%= obalju.FBaljuDetailList(i).FImageSmall %>" width="50" height="50"></td>
		<td ><%= obalju.FBaljuDetailList(i).FItemName %></td>
		<td ><%= obalju.FBaljuDetailList(i).FItemOptionName %></td>
		<td align="center"><%= obalju.FBaljuDetailList(i).FItemNo %></td>
		<td align="center"><font color="red"><b><% if obalju.FBaljuDetailList(i).FItemLackNo=0 then response.write "-" else  response.write  obalju.FBaljuDetailList(i).FItemLackNo end if%></b></font></td>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).CancelYnColor %>"><%= obalju.FBaljuDetailList(i).CancelYnName %></font></td>
		<td align="center"><%= left(obalju.FBaljuDetailList(i).FUpcheBeasongdate,10) %></td>
		<% if Not IsNull(obalju.FBaljuDetailList(i).FmiSendCode) and (CStr(obalju.FBaljuDetailList(i).FDetailIDx)=Cstr(didx)) then %>
		<td align="center"><font color="red">입력중</font></td>
		<% else %>
		<td align="center"><font color="<%= obalju.FBaljuDetailList(i).getMiSendCodeColor %>"><%= obalju.FBaljuDetailList(i).getMiSendCodeName %></font></td>
		<% end if %>
		<td align="center"><%= obalju.FBaljuDetailList(i).FmiSendIpgodate %></td>
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type="button" value=" 미배송입력 " onclick="SubmitMiSend();">
			<input type="button" value=" 배송실완료처리 " onclick="SubmitFinish();">
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- 표 하단바 끝-->





<script>
function printAitem(orderserial, itemid, itemname,  optionname, makerid, brandname, itemno){
	//orderserial, itemid, itemname,  optionname, makerid, brandname, itemno
	ibarprt.printitem1(orderserial,makerid,itemname,optionname,itemid,brandname,itemno);
}

function SubmitMiSend()
{
	var frm = document.frmmisend;

	if (confirm("미배송입력합니다. 이미 미배송처리중인것은 초기화됩니다.\\n진행하시겠습니까?") != true) {
		return false;
	}

	//for (var i=0;i<frm.elements.length;i++){
	//	if ((frm.elements[i].name=="didxarr")&&(frm.elements[i].checked)){
	//		printAitem(frm.elements[i+1].value,frm.elements[i+2].value,frm.elements[i+3].value,frm.elements[i+4].value,frm.elements[i+5].value,"",frm.elements[i+6].value);
	//	}
	//}
	//return;

	frm.mode.value = "add";
	frm.submit();
}

function SubmitFinish()
{
	if (confirm("미배송완료처리합니다.\\n진행하시겠습니까?") != true) {
		return false;
	}
	document.frmmisend.mode.value = "finish";
	document.frmmisend.submit();
}
</script>
</table>
<script language='javascript'>
document.frm.orderserial.focus();
document.frm.orderserial.select();

</script>
<%
set omasterwithcs = Nothing
set obalju = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->