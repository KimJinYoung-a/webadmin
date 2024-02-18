<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
if (session("ssBctID")="gogo27") then C_ADMIN_AUTH=true
dim idx, mastercode, imageon

idx = request("idx")
mastercode = request("mastercode")
imageon = request("imageon")

if (request("masterid") <> "") then
	idx = request("masterid")
end if

if imageon="" then
	imageon="on"
end if

dim oipchul, oipchuldetail, otmp

if (idx <> "") then
        set oipchul = new CIpChulStorage
        oipchul.FRectId = idx
        oipchul.GetIpChulMaster

		if (Left(oipchul.FOneItem.Fcode,2) <> "ST") then
			response.write "<script>alert('에러 : 입고코드가 아닙니다.');</script>"
			response.write "<br><br>에러 : 입고코드가 아닙니다." & oipchul.FOneItem.Fcode
			response.end
		end if

        set oipchuldetail = new CIpChulStorage
        oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
        if Left(oipchul.FOneItem.Fexecutedt,7) <> "" then
            oipchuldetail.FRectYYYYMM = Left(oipchul.FOneItem.Fexecutedt,7)
        end if
        oipchuldetail.GetIpChulDetail
else
        set otmp = new CIpChulStorage
        otmp.FRectStoragecode = mastercode
        idx = otmp.GetIdxFromMasterCode

        set oipchuldetail = new CIpChulStorage
        oipchuldetail.FRectStoragecode = mastercode
        oipchuldetail.GetIpChulDetail

        set oipchul = new CIpChulStorage
        oipchul.FRectId = idx
        oipchul.GetIpChulMaster
end if

dim i
dim sellcashtotal, suplycashtotal
dim itemsum

sellcashtotal  = 0
suplycashtotal = 0
itemsum =0

dim BasicMonth, IsExpireEdit
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))

if not IsNULL(oipchul.FOneItem.Fexecutedt) then
IsExpireEdit = Lcase(CStr(CDate(oipchul.FOneItem.Fexecutedt)<Cdate(BasicMonth)))
else
IsExpireEdit = false
end if

dim IsMaeipIpgo : IsMaeipIpgo = True
if (oipchul.FOneItem.Fdivcode <> "001" and oipchul.FOneItem.Fdivcode <> "801") then
	IsMaeipIpgo = False
end if

dim chkjungsan
if (Left(Now(), 7) = Left(oipchul.FOneItem.Fexecutedt, 7)) then
	'// 이번달 내역은 정산내역 체크 안함.
	chkjungsan = "N"
end if

%>
<script language='javascript'>

function publicbarreg(barcode){
	//var popwin = window.open('/common/popbarcode_input.asp?itembarcode=' + barcode,'popbarcode_input','width=500,height=400,resizable=yes,scrollbars=yes');
	var popwin = window.open('/admin/stock/popBarcodeManage.asp?itemcode=' + barcode,'popbarcode_input','width=550,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function upcheBarReg(barcode){
	var popwin = window.open('/admin/stock/popUpcheManageCode.asp?itemcode=' + barcode,'upcheBarReg_input','width=550,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function ModiRackIpgoYN(frm) {
	if (confirm('저장 하시겠습니까?')) {
		frm.mode.value = "editrackipgoyn";
		frm.action = "ipchuledit_process.asp";
		frm.submit();
	}
}

function ModiMaster(frm){

<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (frm.executedt.value>'<%= date() %>'){
		alert('입고일은 오늘날짜 보다 클수 없습니다.');
		return;
	}

	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (frm.divcode[0].checked){
		if (!checkAvail2(1,frm.executedt.value)){
			alert('매입이며 입고일이 이전달 날짜로 된 내역을 수정하실 경우 \r\n반드시 정산담당자에게 내용을 알려주시기 바랍니다.');
			//return;
		}
	}

	if (frm.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 수정 불가 합니다.');
		return;
	}
<% end if %>
	if (checkAvail3(frm.executedt.value) != true) {
		return;
	}

<% if (IsMaeipIpgo = True) then %>
	if (document.frmMaster.chkjungsan.value == "") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("먼저 정산내역을 체크하세요.");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역 체크없이 진행합니다.!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}

	if (document.frmMaster.chkjungsan.value == "Y") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("============== 수정불가 ==============\n\n정산내역이 있습니다!!!!!!");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역이 있습니다!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}
<% end if %>

	var divcodeChecked = false
	for (var i = 0; i < frm.divcode.length; i++) {
		if (frm.divcode[i].checked == true) {
			divcodeChecked = true;
		}
	}

	if (divcodeChecked == false) {
		alert("ERROR!!!! : 매입구분이 지정되어 있지 않습니다.");
		return;
	}

	if (confirm('저장 하시겠습니까?')){
		frm.mode.value = "editmaster";
		frm.action = "ipchuledit_process.asp";
		frm.submit();
	}
}

function DelMaster(frm){
<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (frm.divcode[0].checked){
		if (!checkAvail2(1,frm.executedt.value)){
			alert('매입이며 입고일이 이전달 날짜로 된 내역을 수정하실 경우 \r\n반드시 정산담당자에게 내용을 알려주시기 바랍니다.');
			//return;
		}
	}

	if (frm.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 삭제 불가 합니다. - 관리자 문의요망');
		return;
	}
<% end if %>

	if (checkAvail3(frm.executedt.value) != true) {
		return;
	}

<% if (IsMaeipIpgo = True) then %>
	if (document.frmMaster.chkjungsan.value == "") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("먼저 정산내역을 체크하세요.");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역 체크없이 진행합니다.!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}

	if (document.frmMaster.chkjungsan.value == "Y") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("============== 수정불가 ==============\n\n정산내역이 있습니다!!!!!!");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역이 있습니다!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}
<% end if %>

	if (confirm('삭제 하시겠습니까?')){
		frm.mode.value = "delmaster";
		frm.action = "ipchuledit_process.asp";
		frm.submit();
	}
}

function checkAvail2(monthdiff,orgdate){
	var nowdate = "<%= Left(now(),10) %>";
	var odate1 = new Date(orgdate.substring(0,4)*1,orgdate.substring(5,7)*1-1,orgdate.substring(8,10),0,0,0);
	var odate2 = new Date(nowdate.substring(0,4)*1,nowdate.substring(5,7)*1-1-(1-1*monthdiff),0,0,0,0);
	if (odate2>=odate1){
		return false;
	}else{
		return true;
	}
}

// 매월 5일까지 전월내역 수정가능
function checkAvail3(modiexecutedt) {
	var orgexecutedt = "<%= oipchul.FOneItem.Fexecutedt %>";
	var thisDate = "<%= Left(Now, 7) %>-01";
	var availDate = "<%= Left(Now, 7) %>-05";
	var nowdate = "<%= Left(now(),10) %>";
	var BasicMonth = "<%= BasicMonth %>";

	if ((orgexecutedt < BasicMonth) || (modiexecutedt < BasicMonth)) {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert('변경불가\n\n입고일이 두달 이전날짜입니다.');
		return false;
		<% else %>
		alert('관리자권한\n\n입고일이 두달 이전날짜입니다.');
		<% end if %>
	}

	if ((orgexecutedt < thisDate) || (modiexecutedt < thisDate)) {
		if (nowdate > availDate) {
			<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
			alert("변경불가\n\n매월 5일까지만 전월내역 변경가능합니다.");
			return false;
			<% else %>
			alert('관리자권한\n\n매월 5일까지만 전월내역 변경가능합니다.');
			<% end if %>
		}
	}

	return true;
}

//function checkAvail(diffdate,orgdate){
//	var nowdate = "<%= Left(now(),10) %>";
//	var odate1 = new Date(orgdate.substring(0,4)*1,orgdate.substring(5,7)*1-1,orgdate.substring(8,10),0,0,0);
//	var odate2 = new Date(nowdate.substring(0,4)*1,nowdate.substring(5,7)*1-1,nowdate.substring(8,10)-diffdate*1,0,0,0);
//	if (odate2>odate1){
//		return false;
//	}else{
//		return true;
//	}
//}

function DelDetail(masterfrm,iid){
<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (masterfrm.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 추가/수정/삭제 불가 합니다.');
		return;
	}
<% end if %>

	if (checkAvail3(masterfrm.executedt.value) != true) {
		return;
	}

<% if (IsMaeipIpgo = True) then %>
	if (document.frmMaster.chkjungsan.value == "") {
		alert("먼저 정산내역을 체크하세요.");
		return;
	}

	if (document.frmMaster.chkjungsan.value == "Y") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("============== 수정불가 ==============\n\n정산내역이 있습니다!!!!!!");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역이 있습니다!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}
<% end if %>

	var frm;
	var found = false;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "cksel") {
			if (frm.checked == true) {
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("선택된 상품을 삭제합니다.") == true) {
			frmDetail.mode.value = "deldetail";
			frmDetail.action = "ipchuledit_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("선택된 상품이 없습니다.");
	}
}

function ModiDetail(masterfrm,iid){
<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (masterfrm.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 추가/수정/삭제 불가 합니다.');
		return;
	}
<% end if %>

	if (checkAvail3(masterfrm.executedt.value) != true) {
		return;
	}

<% if (IsMaeipIpgo = True) then %>
	if (document.frmMaster.chkjungsan.value == "") {
		alert("먼저 정산내역을 체크하세요.");
		return;
	}

	if (document.frmMaster.chkjungsan.value == "Y") {
		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert("============== 수정불가 ==============\n\n정산내역이 있습니다!!!!!!");
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]정산내역이 있습니다!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}
<% end if %>

	var frm;
	var found = false;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "cksel") {
			if (frm.checked == true) {
				if (((frmDetail.elements[i+3].value*0) != 0) || ((frmDetail.elements[i+4].value*0) != 0) || ((frmDetail.elements[i+5].value*0) != 0)) {
					alert("입력값을 확인하세요.");
					return false;
				}
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("선택된 상품을 수정합니다.") == true) {
			frmDetail.mode.value = "editdetail";
			frmDetail.action = "ipchuledit_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("선택된 상품이 없습니다.");
	}

}


var popwin;
function AddItems(frm){
<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (frm.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 추가/수정/삭제 불가 합니다.');
		return;
	}
<% end if %>
	var suplyer, shopid;

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	popwin = window.open('popjumunitem.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&idx=' + frm.masterid.value,'ipgodetailadd','width=840,height=600,scrollbars=yes,resizable=no');
	popwin.focus();
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner,imwdiv){
	if (iidx!='<%= idx %>'){
		alert('주문서가 일치하지 않습니다. 다시시도해 주세요.');
		return;
	}
<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
	if (<%= IsExpireEdit %>){
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
	}

	if (frmMaster.executedt.value<'<%= BasicMonth %>'){
		alert('입고일이 전월 이전 날짜로는 추가/수정/삭제 불가 합니다.');
		return;
	}
<% end if %>

	var frm;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "itemid") {
			if ((iitemid.indexOf(frm.value + "|") == 0) || (iitemid.indexOf("|" + frm.value + "|") >= 0)) {
				if ((iitemoption.indexOf(frmDetail.elements[i+1].value + "|") == 0) || (iitemoption.indexOf("|" + frmDetail.elements[i+1].value + "|") >= 0)) {
					alert("중복된 상품이 있습니다.");
					popwin.focus();
					return false;
				}
			}
		}
	}

	// 상품추가 및 바로 저장
	frmDetail.itemgubunarr.value = igubun;
	frmDetail.itemarr.value = iitemid;
	frmDetail.itemoptionarr.value = iitemoption;
	frmDetail.sellcasharr.value = isellcash;
	frmDetail.suplycasharr.value = isuplycash;
	frmDetail.buycasharr.value = ibuycash;
	frmDetail.itemnoarr.value = iitemno;
	frmDetail.itemnamearr.value = iitemname;
	frmDetail.itemoptionnamearr.value = iitemoptionname;
	frmDetail.designerarr.value = iitemdesigner;
	frmDetail.mwdivarr.value = imwdiv;

	frmDetail.mode.value = "adddetail";
	frmDetail.action = "ipchuledit_process.asp";
	frmDetail.submit();
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

function jsCheckJungsan() {
	var mode = "chkjungsanexist";
	var designerid = "<%= oipchul.FOneItem.Fsocid %>";
	var mastercode = "<%= oipchul.FOneItem.Fcode %>";

	<% if Not IsMaeipIpgo then %>
	alert("매입입고가 아닙니다.");
	return;
	<% end if %>

	if (<%= IsExpireEdit %>){

		<% if Not(C_ADMIN_AUTH or C_MngPart) then %>
		alert('전월 이전 입고된 내역은 수정 불가능합니다.');
		return;
		<% else %>
		if (confirm("============== 경고 ==============\n\n[관리자권한]전월 이전 입고된 내역입니다!!!!!!\n\n계속 진행하시겠습니까?") != true) {
			return;
		}
		<% end if %>
	}

	var xmlhttp;
	if (window.XMLHttpRequest) {
		// code for IE7+, Firefox, Chrome, Opera, Safari
		xmlhttp = new XMLHttpRequest();
	} else {
		// code for IE6, IE5
		xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
	}

	xmlhttp.onreadystatechange = function() {
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			document.frmMaster.chkjungsan.value = trim(xmlhttp.responseText);
			if (document.frmMaster.chkjungsan.value == "Y") {
				alert("수정불가\n\n정산내역이 있습니다.");
			} else {
				alert("OK");
			}
		}
	}
	xmlhttp.open("GET","/admin/jungsan/actJungsanCheck.asp?mode=" + mode + "&designerid=" + designerid + "&mastercode=" + mastercode + "&t=" + Math.random(),true);
	xmlhttp.send();
}

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmDetail;

	if (frm.cksel.length) {
		for (var i = 0; i < frm.cksel.length; i++) {
			frm.cksel[i].checked = bool;
			AnCheckClick(frm.cksel[i]);
		}
	} else {
		frm.cksel.checked = bool;
		AnCheckClick(frm.cksel);
	}
}

function AGVIpgoProc(){
	var frm;
	var found = false;
	for (var i=0;i<frmDetail.elements.length;i++){
		frm = frmDetail.elements[i];
		if (frm.name == "cksel") {
			if (frm.checked == true) {
				if (((frmDetail.elements[i+3].value*0) != 0) || ((frmDetail.elements[i+4].value*0) != 0) || ((frmDetail.elements[i+5].value*0) != 0)) {
					alert("입력값을 확인하세요.");
					return false;
				}
				found = true;
			}
		}
	}

	if (found == true) {
		if (confirm("선택상품의 AGV 입고를 진행 하시겠습니까?") == true) {
			frmDetail.mode.value = "agvipgoitemdivisionorder";
			frmDetail.action = "ipchuledit_process.asp";
			frmDetail.submit();
		}
	} else {
		alert("선택된 상품이 없습니다.");
	}
}
function AGVIpgoDelProc(){
	if (confirm("AGV 입고를 삭제 하시겠습니까?") == true) {
		frmDetail.mode.value = "agvipgoitemdivisionorderdelete";
		frmDetail.action = "ipchuledit_process.asp";
		frmDetail.submit();
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
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>입고정보</strong></font>
	        &nbsp;
	        <b>[ <%= oipchul.FOneItem.Fcode %> ]</b>
        </td>
        <td align="right">

        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMaster" method="post" action="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="masterid" value="<%= oipchul.FOneItem.Fid %>">
	<input type="hidden" name="suplyer" value="<%= oipchul.FOneItem.Fsocid %>">
	<input type="hidden" name="shopid" value="10x10">
	<input type="hidden" name="chkjungsan" value="<%= chkjungsan %>">
	<tr bgcolor="#FFFFFF">
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >입출고코드</td>
		<td width="360"><%= oipchul.FOneItem.Fcode %></td>
		<td width="110" bgcolor="<%= adminColor("tabletop") %>" >브랜드</td>
		<td><%= oipchul.FOneItem.Fsocid %>&nbsp;(<%= oipchul.FOneItem.Fsocname %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >총소비자가</td>
		<td><%= FormatNumber(oipchul.FOneItem.Ftotalsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" >총매입가</td>
		<td><%= FormatNumber(oipchul.FOneItem.Ftotalsuplycash,0) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >명세서일자</td>
		<td ><input type="text" name="scheduledt" size="10" maxlength=10 readonly value="<%= Left(oipchul.FOneItem.Fscheduledt,10) %>"><a href="javascript:calendarOpen(frmMaster.scheduledt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a></td>
		<td bgcolor="<%= adminColor("tabletop") %>" >입고일자</td>
		<td><input type="text" name="executedt" size="10" maxlength=10 readonly value="<%= Left(oipchul.FOneItem.Fexecutedt,10) %>"><a href="javascript:calendarOpen(frmMaster.executedt);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (재고와 관련 있음)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >매입구분</td>
		<td colspan="3">
		<input type="radio" name="divcode" value="001" <% if oipchul.FOneItem.Fdivcode="001" then response.write "checked" %> >매입
		<input type="radio" name="divcode" value="002" <% if oipchul.FOneItem.Fdivcode="002" then response.write "checked" %> >위탁
		<input type="radio" name="divcode" value="801" <% if oipchul.FOneItem.Fdivcode="801" then response.write "checked" %> >OFF용매입
		<input type="radio" name="divcode" value="802" <% if oipchul.FOneItem.Fdivcode="802" then response.write "checked" %> >OFF용위탁
		<% if oipchul.FOneItem.Fexecutedt <> "" then %>
			<% if fnGetAGVCheckBalju(oipchul.FOneItem.Fcode) then %>
				<input type="button" class="button" value="AGV입고삭제" onClick="AGVIpgoDelProc();">
			<% else %>
				<input type="button" class="button" value="AGV입고" onClick="AGVIpgoProc();">
			<% end if %>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >랙입고</td>
		<td colspan="3">
		  <input type="radio" name="rackipgoyn" value="Y" <% if (oipchul.FOneItem.Frackipgoyn = "Y") then %>checked<% end if %>> 예
		  <input type="radio" name="rackipgoyn" value="N" <% if (oipchul.FOneItem.Frackipgoyn = "N") then %>checked<% end if %>> 아니요
		  <% if Not IsNull(oipchul.FOneItem.Fexecutedt) then %>
		  <input type=button value=" 수 정 " onClick="ModiRackIpgoYN(frmMaster)">
		  <% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >처리자</td>
		<td colspan="3"><%= oipchul.FOneItem.Fchargeid %>&nbsp;(<%= oipchul.FOneItem.Fchargename %>)</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" >기타사항</td>
		<td colspan="3">
			<textarea name="comment" cols=80 rows=6><%= (oipchul.FOneItem.Fcomment) %></textarea>
		</td>
	</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if (oipchul.FOneItem.Fdivcode="001") then %>
			<img src="/images/exclam.gif" align="absmiddle" width=21>매입 구분이 <font color=red>매입</font>인 경우 <b>정산 파일 작성 후 수정 한 경우</b> 반드시 정산 담당자에게 알려주시기 바랍니다.<br>
			<% else %>
			<img src="/images/exclam.gif" align="absmiddle" width=21>매입 구분이 <font color=red>위탁</font>인 경우 입고일 <b>최근 2달</b> 까지만 수정가능합니다.(재고 관련)<br>
			<% end if %>
			<input type=button value=" 수 정 " onClick="ModiMaster(frmMaster)">&nbsp;
			<input type=button value=" 삭 제 " onClick="DelMaster(frmMaster)">
			<% if (IsMaeipIpgo = True) then %>
			&nbsp;
			<input type=button value=" 정산내역 체크 " onClick="jsCheckJungsan()">
			<% end if %>
		</td>
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

<br>

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
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>상세내역</strong></font>
        	&nbsp;&nbsp;
        	<font color="<%= mwdivColor("M") %>">매입</font>&nbsp;
        	<font color="<%= mwdivColor("W") %>">위탁</font>&nbsp;
        	<font color="<%= mwdivColor("U") %>">업체배송</font>
        	&nbsp;&nbsp;
        	<!--
        	<input type=checkbox name=imageon <% if imageon="on" then response.write "checked" %>>이미지표시
        	<a href="javascript:document.frm.submit();"><img src="/images/button_reload.gif" width="60" height="20" border="0"></a>
            -->
        </td>
        <td align="right">
        	총건수:  <%= oipchuldetail.FResultCount %>
        	&nbsp;
        	<input type=button name="" value="상품추가" onClick="AddItems(frmMaster);">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmDetail" method="post" action="">
	<input type="hidden" name="masterid" value="<%= oipchul.FOneItem.Fid %>">
	<input type="hidden" name="code" value="<%= oipchul.FOneItem.Fcode %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="mwdivarr" value="">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<% if imageon="on" then %>
    	<td width="50">이미지</td>
    	<% end if %>
		<td width="100">상품코드</td>
		<td width="100">범용바코드</td>
		<td width="100">업체관리코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="50">수량</td>
		<td width="50">AGV수량</td>
		<td width="70"><font color="#AAAAAA">원소비자가</font></td>
		<td width="70"><font color="#AAAAAA">원공급가</font></td>
		<td width="70">판매가</td>
		<td width="70">매입가</td>
		<td width="50">마진</td>
        <td width="50">매입구분<br />(입고시)</td>
		<td width="50">매입구분(온라인)</td>
        <td width="50">매입구분<br />(월별)</td>
		<td width="50">CENTER<br>매입구분</td>
		<td width="50">비고</td>
	</tr>
	<% for i=0 to oipchuldetail.FResultCount -1 %>
	<%
	sellcashtotal = sellcashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsellcash
	suplycashtotal = suplycashtotal + oipchuldetail.FItemList(i).Fitemno * oipchuldetail.FItemList(i).Fsuplycash
	itemsum = itemsum + oipchuldetail.FItemList(i).Fitemno
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" value="<%= i %>" onClick="AnCheckClick(this);"></td>
		<% if imageon="on" then %>
    	<td width="50"><img src="<%= CHKIIF((oipchuldetail.FItemList(i).Fiitemgubun="10"), oipchuldetail.FItemList(i).Fsmallimage, oipchuldetail.FItemList(i).Foffimgsmall) %>"></td>
    	<% end if %>
		<td><font color="<%= mwdivColor(oipchuldetail.FItemList(i).FOnlineMwdiv) %>"><%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= oipchuldetail.FItemList(i).FItemId %>-<%= oipchuldetail.FItemList(i).FItemOption %></font></td>
		<td>
			<a href="javascript:publicbarreg('<%= oipchuldetail.FItemList(i).Fiitemgubun %><%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %><%= oipchuldetail.FItemList(i).FItemOption %>');">
			<% if oipchuldetail.FItemList(i).FPublicBarcode<>"" then %>
				<font color="#AAAAAA"><b><%= oipchuldetail.FItemList(i).FPublicBarcode %></b></font>
			<% else %>
				등록>>
			<% end if %>
			</a>
		</td>
		<td>
			<a href="javascript:upcheBarReg('<%= oipchuldetail.FItemList(i).Fiitemgubun %><%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %><%= oipchuldetail.FItemList(i).FItemOption %>');">
				<% if oipchuldetail.FItemList(i).FUpcheManageCode<>"" then %>
				<font color="#AAAAAA"><b><%= oipchuldetail.FItemList(i).FUpcheManageCode %></b></font>
				<% else %>
				등록>>
				<% end if %>
			</a>
		</td>
		<td align="left"><%= oipchuldetail.FItemList(i).FIItemName %></td>
		<td><%= oipchuldetail.FItemList(i).FIItemoptionName %></td>
		<input type="hidden" name="itemid" value="<%= oipchuldetail.FItemList(i).FItemId %>">
		<input type="hidden" name="itemoption" value="<%= oipchuldetail.FItemList(i).FItemOption %>">
		<td><input type=text name="itemno" value="<%= oipchuldetail.FItemList(i).Fitemno %>" size=4 maxlength=6 ></td>
		<td><input type="text" class="text" name="agvitemno" size="4" maxlength="6" value="<%= oipchuldetail.FItemList(i).Fitemno %>"></td>
		<td align="right"><font color="#AAAAAA"><%= oipchuldetail.FItemList(i).Forgprice %></font></td>
		<td align="right"><font color="#AAAAAA"><%= oipchuldetail.FItemList(i).Forgsuplycash %></font></td>
		<td align="right"><input type=text name="sellcash" value="<%= oipchuldetail.FItemList(i).Fsellcash %>" size=6 maxlength=9 style="text-align:right"></td>
		<td align="right"><input type=text name="suplycash" value="<%= oipchuldetail.FItemList(i).Fsuplycash %>" size=6 maxlength=9 style="text-align:right"></td>
		<input type="hidden" name="didx" value="<%= oipchuldetail.FItemList(i).Fid %>">
		<td align=center>
			<% if oipchuldetail.FItemList(i).Fsellcash<>0 then %>
			<%= 100-CLng(oipchuldetail.FItemList(i).Fsuplycash/oipchuldetail.FItemList(i).Fsellcash*100*100)/100 %>%
			<% end if %>
		</td>
		<td><font color="<%= mwdivColor(oipchuldetail.FItemList(i).Fmwgubun) %>"><%= oipchuldetail.FItemList(i).Fmwgubun %></font></td>
        <td><font color="<%= mwdivColor(oipchuldetail.FItemList(i).FOnlineMwdiv) %>"><%= oipchuldetail.FItemList(i).FOnlineMwdiv %></font></td>
        <td><font color="<%= mwdivColor(oipchuldetail.FItemList(i).Flastmwdiv) %>"><%= oipchuldetail.FItemList(i).Flastmwdiv %></font></td>
		<td><font color="<%= mwdivColor(oipchuldetail.FItemList(i).FCenterMwdiv) %>"><%= oipchuldetail.FItemList(i).FCenterMwdiv %></font></td>
		<td><% if oipchuldetail.FItemList(i).FDtComment<>"" then%><%= replace(oipchuldetail.FItemList(i).FDtComment," ","<br>") %><% end if %></td>
		<input type="hidden" name="itemgubun" value="<%= oipchuldetail.FItemList(i).Fiitemgubun %>">
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<% if imageon="on" then %>
    	<td></td>
    	<% end if %>
		<td colspan="6">총계</td>
		<td align="center"><%= itemsum %></td>
		<td colspan="3">&nbsp;</td>
		<td align="right"><b><%= FormatNumber(sellcashtotal,0) %></b></td>
		<td align="right"><b><%= FormatNumber(suplycashtotal,0) %></b></td>
		<td></td>
		<td></td>
		<td></td>
        <td></td>
        <td></td>
		<td></td>

	</tr>
	</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<input type=button value=" 선택상품수정 " onclick="ModiDetail(frmMaster,frmDetail)">
	    	<input type=button value=" 선택상품삭제 " onclick="DelDetail(frmMaster,frmDetail)">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<%
set oipchuldetail = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
