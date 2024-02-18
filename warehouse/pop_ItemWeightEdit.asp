<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 해외배송 상품 관리
' History : 2008.03.26 서동석 생성
'			2016.05.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/items/itemVolumncls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim itembarcode, prcAfter, sqlStr, mode, deliverOverseas, menupos, itemgubun, itemid, itemoption, i, regtype
dim optionsetYN

	itembarcode = requestCheckVar(request("itembarcode"),20)
	prcAfter = requestCheckVar(request("prcAfter"),32)
	mode = requestCheckVar(request("mode"),32)
	menupos		= requestCheckVar(getNumeric(request("menupos")),10)
	regtype = requestCheckVar(request("regtype"),32)
	optionsetYN=False
'' ByWeightProc/BySizeProc
if (mode = "") then
	mode = "ByWeightProc"
end if

if regtype="" then regtype="I"

'범용바코드 검색
if Len(itembarcode)>=12 then
	sqlStr = "select top 1 b.* " + VbCrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
	sqlStr = sqlStr + " where b.barcode='" + CStr(itembarcode) + "' " + VbCrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("itemid")
		itemoption = rsget("itemoption")
	else
		itemgubun 	= BF_GetItemGubun(itembarcode)
		itemid 		= BF_GetItemId(itembarcode)
		itemoption 	= BF_GetItemOption(itembarcode)
	end if
	rsget.Close
else
	itemgubun="10"
	itemid = itembarcode
	itemoption="0000"
end if

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim oitembar
set oitembar = new CItemInfo
	oitembar.FRectItemID = itemid
	if itemid<>"" then
		oitembar.GetOneItemInfo
	    if itemgubun = "10" and oitembar.FResultCount>0 then
		    regtype = oitembar.FOneItem.FitemManageType
	    end if
	end if

if itemgubun = "" then
	Response.Write "무게입력 불가 : 상품구분이 없음"
	Response.end
end if

if itemid<>"" and itemgubun <> "10" then
	oitembar.FRectItemGubun = itemgubun
	oitembar.FRectItemID =  itemid
	oitembar.FRectItemOption =  itemoption
	oitembar.GetOneItemInfoOffline
	regtype = "O"
	''Response.end
end if

dim k, oitemoption, oOptionMultipleType, oOptionMultiple

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
	oitemoption.GetItemOptionInfo
end If

set oOptionMultipleType = new CItemOptionMultiple
oOptionMultipleType.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
    oOptionMultipleType.GetOptionTypeInfo
end if

set oOptionMultiple = new CitemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
    oOptionMultiple.GetOptionMultipleInfo
end if
%>
<script language="javascript" type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function InputWeightInfo(frm){
//	//업배도 모두 오픈	'/2016.05.27 한용민 추가
//	if(frm.isUcDeli.value=='True') {
//		alert('텐바이텐 배송 상품만 무게를 입력할 수 있습니다.\n\n다른 상품을 선택해주세요.');
//		return;
//	}

	if (!frm.overSeaYn.value){
		alert('해외배송 여부를 선택해주세요.');
		frm.overSeaYn.focus();
		return;
	}

	if (confirm('정보를 저장하시겠습니까?')){
		frm.mode.value = "ByWeightProc";
		frm.submit();
	}
}

function InputSizeInfo(frm){
//	//업배도 모두 오픈	'/2016.05.27 한용민 추가
//	if(frm.isUcDeli.value=='True') {
//		alert('텐바이텐 배송 상품만 사이즈를 입력할 수 있습니다.\n\n다른 상품을 선택해주세요.');
//		return;
//	}

	if (!frm.itemWeight.value.length){
		alert('상품 무게를 정확히 입력하세요.');
		frm.itemWeight.focus();
		return;
	}
	if (frm.itemWeight.value*0 != 0) {
		alert('상품 무게는 숫자만 가능합니다.');
		frm.itemWeight.value="";
		frm.itemWeight.focus();
		return;
	}

	if (!frm.volX.value.length){
		alert('상품 사이즈를 정확히 입력하세요.');
		frm.volX.focus();
		return;
	}

	if (!frm.volY.value.length){
		alert('상품 사이즈를 정확히 입력하세요.');
		frm.volY.focus();
		return;
	}

	if (!frm.volZ.value.length){
		alert('상품 사이즈를 정확히 입력하세요.');
		frm.volZ.focus();
		return;
	}

	if (frm.volX.value*0 != 0) {
		alert('상품 사이즈는 숫자만 가능합니다.');
		frm.volX.value="";
		frm.volX.focus();
		return;
	}

	if (frm.volY.value*0 != 0) {
		alert('상품 사이즈는 숫자만 가능합니다.');
		frm.volY.value="";
		frm.volY.focus();
		return;
	}
	if (frm.volZ.value*0 != 0) {
		alert('상품 사이즈는 숫자만 가능합니다.');
		frm.volZ.value="";
		frm.volZ.focus();
		return;
	}

	if (confirm('상품 사이즈를 저장하시겠습니까?')){
	    frm.mode.value = "BySizeProc";
		frm.submit();
	}
}

function Research(frm){
	frm.submit();
}

function getOnliad() {
<%
	if (oitembar.FResultCount>0) then
		if Not(oitembar.FOneItem.IsUpcheBeasong) and (prcAfter="") then
			if (mode = "ByWeightProc") then
%>
    document.frmitemWeight.itemWeight.select();
    document.frmitemWeight.itemWeight.focus();
<%
			else
%>
    document.frmitemWeight.volX.select();
    document.frmitemWeight.volX.focus();
<%
			end if
		else
%>
    document.frmbar.itembarcode.select();
    document.frmbar.itembarcode.focus();
<%
		end if
	else
%>
    document.frmbar.itembarcode.select();
    document.frmbar.itembarcode.focus();
<% end if %>
}

function fnCheckRegType(regtype){
	if(regtype=="I"){
		$("#itemW").show();
		$("#itemS").show();
		$("#itemOPT").hide();
	}
	else{
		$("#itemW").hide();
		$("#itemS").hide();
		$("#itemOPT").show();
	}
}

function InputOptionSizeInfo(frm){
	if(frm.regtype[0].checked){
		if (!frm.itemWeight.value.length){
			alert('상품 무게를 정확히 입력하세요.');
			frm.itemWeight.focus();
			return;
		}
		if (frm.itemWeight.value*0 != 0) {
			alert('상품 무게는 숫자만 가능합니다.');
			frm.itemWeight.value="";
			frm.itemWeight.focus();
			return;
		}
		if (!frm.volX.value.length){
			alert('상품 사이즈를 정확히 입력하세요.');
			frm.volX.focus();
			return;
		}
		if (!frm.volY.value.length){
			alert('상품 사이즈를 정확히 입력하세요.');
			frm.volY.focus();
			return;
		}
		if (!frm.volZ.value.length){
			alert('상품 사이즈를 정확히 입력하세요.');
			frm.volZ.focus();
			return;
		}
		if (frm.volX.value*0 != 0) {
			alert('상품 사이즈는 숫자만 가능합니다.');
			frm.volX.value="";
			frm.volX.focus();
			return;
		}
		if (frm.volY.value*0 != 0) {
			alert('상품 사이즈는 숫자만 가능합니다.');
			frm.volY.value="";
			frm.volY.focus();
			return;
		}
		if (frm.volZ.value*0 != 0) {
			alert('상품 사이즈는 숫자만 가능합니다.');
			frm.volZ.value="";
			frm.volZ.focus();
			return;
		}
		if (confirm('상품 사이즈를 저장하시겠습니까?')){
			frm.mode.value = "ByOptionSameSizeProc";
			frm.submit();
		}
	}
	else{
		if (confirm('상품 사이즈를 저장하시겠습니까?')){
			frm.mode.value = "ByOptionSizeProc";
			frm.submit();
		}
	}
}

window.onload=getOnliad;
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<!-- 상단바 시작 -->
<form name="frmbar" method=get>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
					<font color="red">&nbsp;<strong>상품무게입력</strong></font>
				</td>
				<td align="right">
					<input type="text" class="text"  name="itembarcode" value="<%= itembarcode %>" size=14 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
					<input type="button" class="button" value="검색" onclick="Research(frmbar)" >
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->
</form>
<% if oitembar.FResultCount>0 then %>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
	<td colspan="2"><%= oitembar.FOneItem.Fmakername %>(<%= oitembar.FOneItem.Fmakerid %>)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
	<td colspan="2"><%= oitembar.FOneItem.FItemName %></td>
</tr>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">이미지</td>
	<td colspan="2"><img src="<%= oitembar.FOneItem.Flistimage %>" width="100" height="100" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
</tr>

<form name="frmitemWeight" method=post  action="/warehouse/itemWeight_process.asp" style="margin:0px;">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
	<td colspan="2"><%= FormatNumber(oitembar.FOneItem.FSellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">배송구분</td>
	<td colspan="2">
		<%=oitembar.FOneItem.GetDeliveryName%>
		<input type="hidden" name="isUcDeli" value="<%=oitembar.FOneItem.IsUpcheBeasong %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">해외배송 여부</td>
	<td colspan="2">
		<%
		' 해외직구 일경우
		if oitembar.FOneItem.Fdeliverfixday = "G" then
			deliverOverseas = "N"
		else
			deliverOverseas = oitembar.FOneItem.FdeliverOverseas
		end if

		if Not(oitembar.FOneItem.IsUpcheBeasong) and (oitembar.FOneItem.FitemWeight<=0) then
			drawSelectBoxUsingYN "overSeaYn", "Y"
			Response.Write "[현재 상태: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.FdeliverOverseas
			Response.Write "</b></font>]"
		else
			drawSelectBoxUsingYN "overSeaYn", deliverOverseas
			Response.Write "[현재 상태: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.FdeliverOverseas
			Response.Write "</b></font>]"
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">포장가능 여부</td>
	<td colspan="2">
	<%
		if Not(oitembar.FOneItem.IsUpcheBeasong) and (oitembar.FOneItem.FvolX<=0) then
			drawSelectBoxUsingYN "pojangok", "N"
			Response.Write " [현재 상태: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.Fpojangok
			Response.Write "</b></font>]"
		else
			drawSelectBoxUsingYN "pojangok", oitembar.FOneItem.Fpojangok
		end if
	%>
	<input type="button" class="button" value="저장" onclick="InputWeightInfo(frmitemWeight);">
	</td>
</tr>
<% If itemoption="0000" And oitemoption.FResultCount<1 Then %>
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#E6E6E6">상품 사이즈<br>등록유형</td>
	<td colspan="2">
		<input type="radio" name="regtype" value="I"<% If regtype="I" Then Response.write " checked" %> onClick="fnCheckRegType('I');">일괄등록 <input type="radio" name="regtype" value="O" <% If regtype="O" Then Response.write " checked" %> onClick="fnCheckRegType('O');">옵션개별등록 <input type="button" class="button" value="저장" onclick="InputOptionSizeInfo(frmitemWeight);">
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" id="itemW">
	<td bgcolor="<%= adminColor("tabletop") %>">상품무게</td>
	<td colspan="2">
		<%
		'/업배도 모두 오픈	'/2016.05.27 한용민 추가
		'if Not(oitembar.FOneItem.IsUpcheBeasong) then
		%>
			<input type="text" class="text" name="itemWeight" id="itemWeight" value="<%= oitembar.FOneItem.FitemWeight %>" size="6" AUTOCOMPLETE="off" style="text-align:right;">g
			&nbsp;※ 무게는 그램(g)으로 입력 (예:1.5Kg→1500g)
		<% 'else %>
			<!--<input type="text" class="text" name="itemWeight" value="<%'= oitembar.FOneItem.FitemWeight %>" size="6" readonly onKeyPress="if (event.keyCode == 13){ InputWeightInfo(frmitemWeight); return false;}" style="text-align:right;">g
			&nbsp;※ 텐바이텐배송 상품만 무게를 입력할 수 있습니다.-->
		<% 'end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="itemS">
	<td bgcolor="#E6E6E6">상품사이즈</td>
	<td colspan="2">
		<input type="text" class="text" name="volX" id="volX" value="<%= oitembar.FOneItem.FvolX %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		*
		<input type="text" class="text" name="volY" id="volY" value="<%= oitembar.FOneItem.FvolY %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		*
		<input type="text" class="text" name="volZ" id="volZ" value="<%= oitembar.FOneItem.FvolZ %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		cm
		&nbsp;※ 센티미터(cm)로 입력 <% If itemoption="0000" And oitemoption.FResultCount<1 Then %><input type="button" class="button" value="저장" onclick="InputSizeInfo(frmitemWeight);"><% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="itemOPT" style="display:none">
	<td bgcolor="#E6E6E6">옵션별</td>
	<td colspan="2">
		<% if oitemoption.FResultCount<1 then %>
		<% else %>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="#999999">
				<tr align="center" bgcolor="#E6E6E6">
					<td width="60">옵션코드</td>
					<td>옵션상세명</td>
					<td width="80">무게</td>
					<td width="150">상품사이즈</td>
				</tr>
				<% for k=0 to oitemoption.FResultCount -1 %>
				<tr align="center" bgcolor="#FFFFFF">
					<td><%= oitemoption.FItemList(k).Fitemoption %><input type="hidden" name="itemoption" value="<%= oitemoption.FItemList(k).Fitemoption %>"></td>
					<td><%= oitemoption.FItemList(k).Foptionname %></td>
					<td width="80"><input type="text" class="text" name="oitemWeight" id="MitemWeight" value="<%= oitemoption.FItemList(k).FitemWeight %>" size="6" style="text-align:right;">g</td>
					<td width="100"><input type="text" class="text" name="ovolX" id="MvolX" value="<%= oitemoption.FItemList(k).FvolX %>" size="2" style="text-align:right;">*<input type="text" class="text" name="ovolY" id="MvolY" value="<%= oitemoption.FItemList(k).FvolY %>" size="2" style="text-align:right;">*<input type="text" class="text" name="ovolZ" id="MvolZ" value="<%= oitemoption.FItemList(k).FvolZ %>" size="2" style="text-align:right;">cm</td>
				</tr>
				<% Next %>
				<% optionsetYN = True %>
			<table>
		<% end if %>
	</td>
</tr>
</form>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align="center">
		검색결과가 없습니다

		<!-- <br>
		현재 10코드(온라인등록상품)만 등록이 가능합니다.
		<br>90코드의 경우 오프상품관리를 이용하세요. -->
	</td>
</tr>
<% end if %>
</table>
<% if optionsetYN Then %>
<script>
$(function(){
	fnCheckRegType('<%= regtype %>');
	$("input[name='regtype']:radio[value='<%= regtype %>']").prop('checked',true);
});
</script>
<% End If %>

<form name="frmsavebar" method=post action="barcode_input_process.asp" style="margin:0px;">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="publicbarcode" value="">
</form>

<%
set oitembar = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
