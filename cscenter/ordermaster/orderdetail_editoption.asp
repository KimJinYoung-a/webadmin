<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 상품변경
' History : 이상구 생성
'			2023.06.12 한용민 수정(표준코딩으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<%
dim i, idx, orderserial, result
	idx = requestCheckVar(request("idx"),10)

'상품옵션변경
'result = CSOrderModifyItemOption("10032647343", 178977, "0000", "7029")

'취소상품 정상화
'result = CSOrderRestoreCanceledItem("10032647343", 178977, "0000")

'상품취소
'result = CSOrderCancelItem("10032647343", 178977, "0000")

'response.write "aaaaaaaaaaaaaaaa" & CS_ORDER_FUNCTION_RESULT

dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if


if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if


If ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then

	Dim sqlStr, rsOption, k, optionText, itemStatus
	dim sqlsub

	sqlsub = "select top 1 optaddprice "
	sqlsub = sqlsub + "from [db_item].[dbo].tbl_item_option "
	sqlsub = sqlsub + "where 1 = 1 "
	sqlsub = sqlsub + "and itemid = " & CStr(ojumunDetail.FJumunDetail.Fitemid) & " "
	sqlsub = sqlsub + "and itemoption = '" & CStr(ojumunDetail.FJumunDetail.Fitemoption) & "' "

	'* 옵션변경은 <font color=red>옵션가</font>가 동일한 옵션상품만 가능합니다.<br>
	'* 주문당시 옵션가격에 상관없이 현재 상품정보 상의 옵션가격으로 비교합니다.<br>
	'* 상품할인정보(판매가,매입가 등)는 주문당시 정보가 유지됩니다.<br>
	' 주문후 사용안함 처리가 되어도 표시
	sqlStr = " select "
	sqlStr = sqlStr + " v.itemoption, v.optionname "
	sqlStr = sqlStr + " , v.optsellyn, v.optlimityn, v.optlimitno, v.optlimitsold "
	sqlStr = sqlstr + " , 0 as notused "
	sqlStr = sqlStr + " , case when v.optaddprice=IsNULL((" & sqlsub & "),0) " & " then 'T' else 'F' end "
	sqlStr = sqlStr + " , v.isusing "
	sqlStr = sqlStr + " , v.optaddprice "
	sqlStr = sqlStr + " , IsNull(P.regno, 0) as prevregno "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " on i.itemid=v.itemid "

	'이전 CS반품내역(접수+완료내역, 반품사유고려안함)
	sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
	sqlStr = sqlStr + "		    select d.itemid, d.itemoption, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
    sqlStr = sqlStr + "		    from" + VbCrlf
    sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
    sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
    sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
    sqlStr = sqlStr + "		    where a.orderserial='" + CStr(orderserial) + "'" + VbCrlf
    sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''반품 / 회수 / 상품변경 맞교환회수(텐바이텐배송) / 상품변경 맞교환반품(업체배송).
    sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
    'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'접수+완료 모두 계산
    sqlStr = sqlStr + "			group by d.itemid, d.itemoption" + VbCrlf
    sqlStr = sqlStr + " ) P " + VbCrlf
    sqlStr = sqlStr + "     ON i.itemid=P.itemid and v.itemoption=P.itemoption" + VbCrlf

	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & ojumunDetail.FJumunDetail.Fitemid & ""
	sqlStr = sqlStr + " order by i.itemid desc, v.itemoption"

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		rsOption = rsget.getrows
	End If
	rsget.close()

	'response.write sqlStr
End If


'==============================================================================

dim oordermaster, oorderdetail, selecteditemindex

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


set oorderdetail = new COrderMaster
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oorderdetail.FRectOldOrder = "on"
    oorderdetail.QuickSearchOrderDetail
end if


selecteditemindex = 0
for i = 0 to oorderdetail.FResultCount - 1
	if (CStr(oorderdetail.FItemList(i).Fidx) = CStr(ojumunDetail.FJumunDetail.Fdetailidx)) then
		selecteditemindex = i
	end if
next

dim currentitemoptionidx, currentitemoptionorgno
dim changedindex
dim prevregno

'==============================================================================
'// 옵션변경 맞교환의 경우 기존반품수량
prevregno = 0

For i = 0 To UBound(rsOption,2)
	if (rsOption(0,i) = ojumunDetail.FJumunDetail.Fitemoption) then
		if (rsOption(10,i) <> 0) then
			prevregno = rsOption(10,i)
		end if
	end if
Next

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript">
window.resizeTo(1400,800);
var oldConfirmDate = "";
var oldBeasongDate = "";
function CheckConfirmDate(comp){
    if (comp.value==""){
        oldConfirmDate = comp.form.upcheconfirmdate.value;
        oldBeasongDate = comp.form.beasongdate.value;
        comp.form.upcheconfirmdate.value = "";
    }else{
        if (oldConfirmDate!=""){
            comp.form.upcheconfirmdate.value = oldConfirmDate;
        }

        if (oldBeasongDate!=""){
            comp.form.beasongdate.value = oldBeasongDate;
        }
    }
}

function EditDetail(detailidx,mode,comp){
    var frm = document.frm;

	if (mode=="buycash"){
		if (!IsDigit(comp.value)){
			alert('매입가는 숫자만 가능합니다.');
			comp.focus();
			return;
		}
	}else if(mode=="isupchebeasong"){
	    if (frm.isupchebeasong.value=="Y"){
	        if (frm.omwdiv.value!="U"){
	            alert('매입구분과 배송구분이 일치하지 않습니다.');
	            return;
	        }
	    }else{
	        if (frm.omwdiv.value=="U"){
	            alert('매입구분과 배송구분이 일치하지 않습니다.');
	            return;
	        }
	    }

        if (frm.omwdiv.value=="U"){
            if ((frm.odlvType.value=="1")||(frm.odlvType.value=="4")){
                alert('매입구분과 배송구분이 일치하지 않습니다.');
	            return;
            }
        }else{
            if ((frm.odlvType.value!="1")&&(frm.odlvType.value!="4")){
                alert('매입구분과 배송구분이 일치하지 않습니다.');
	            return;
            }
        }


    }else if(mode=="songjangdiv"){


	}else if(mode=="currstate"){


    }else if(mode=="songjangdiv"){
        if (frm.songjangdiv.value.length<1){
            alert('택배사를 선택하세요.');
			frm.songjangdiv.focus();
			return;
        }

        if (!IsDigit(frm.songjangno.value)){
			alert('운송장번호는 숫자는 가능합니다.');
			frm.songjangdiv.focus();
			return;
		}
	}else if (mode=="requiredetail"){

	}else if (mode=="itemno"){

	}else if (mode=="itemOption"){
		var arr = comp.value.split("|");
	    if (frm.preItemOption.value==arr[0])
	    {
			alert("현재아이템과 동일한 옵션입니다. 변경하실 수 없습니다.");
			return;
	    }
	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
}

// ============================================================================
function EditItemOption(){
    var frm = document.frm;

	if (frm.contents_jupsu.value == "") {
		alert("변경할 옵션을 선택하세요.");
		return;
	}

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// 업체배송, 상품준비 이후
	if (confirm('업체배송이면서 상품준비 이후입니다.\n\n수정 하시겠습니까?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// 상품출고 이후
	alert('상품출고 이후입니다. 파티장에게 문의하세요.');
<% else  %>
	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
<% end if %>
}

function ForceEditItemOption(){
    var frm = document.frm;

	if (frm.contents_jupsu.value == "") {
		alert("변경할 옵션을 선택하세요.");
		return;
	}

	if (confirm('강제옵션변경 하시겠습니까?')){
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemRestoreCancel(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// 업체배송, 상품준비 이후
	if (confirm('업체배송이면서 상품준비 이후입니다.\n\n수정 하시겠습니까?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// 상품출고 이후
	alert('상품출고 이후입니다. 파티장에게 문의하세요.');
<% else  %>
	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품취소정상화";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 취소상태를 정상화 신청";

		frm.contents_jupsu.value = str;

		frm.mode.value="RestoreCancel";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemRestoreCancel(){
    var frm = document.frm;

	if (confirm('강제정상화 하시겠습니까?')){
		frm.title.value = "상품취소정상화";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 취소상태를 정상화 신청";


		frm.mode.value="RestoreCancel";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemCancel(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// 업체배송, 상품준비 이후
	if (confirm('업체배송이면서 상품준비 이후입니다.\n\n수정 하시겠습니까?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// 상품출고 이후
	alert('상품출고 이후입니다. 파티장에게 문의하세요.');
<% else  %>
	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품취소";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 취소 신청";

		frm.contents_jupsu.value = str;

		frm.mode.value="Cancel";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemCancel(){
    var frm = document.frm;

	if (confirm('강제정상화 하시겠습니까?')){
		frm.title.value = "상품취소";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 취소 신청";


		frm.mode.value="Cancel";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemNo(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// 업체배송, 상품준비 이후
	if (confirm('업체배송이면서 상품준비 이후입니다.\n\n수정 하시겠습니까?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// 상품출고 이후
	alert('상품출고 이후입니다. 파티장에게 문의하세요.');
<% else  %>
	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품수량변경";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 수량을 " + frm.preItemNo.value + " 에서 " + frm.itemno.value + " 로 변경 신청";

		frm.contents_jupsu.value = str;

		frm.mode.value="EditItemNo";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemNo(){
    var frm = document.frm;

	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품수량변경";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션중\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 수량을 " + frm.preItemNo.value + " 에서 " + frm.itemno.value + " 로 변경 신청";

		frm.mode.value="EditItemNo";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function ChangeJupsucontents(){
    var frm = document.frm;

	var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n의 옵션을\n";
	str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 에서\n";


	str = str + eval("frm.itemOption" + frm.itemOption.value).value + " 로 변경신청";

	frm.contents_jupsu.value = str;
}



// ============================================================================
// 옵션별 수량 자동조절(마이너스 입력불가)
// ============================================================================
function CheckItemOptionNoCount(changedindex){
    var frm = document.frm;
    var i;

	totalcount = 0;
	maxcount = parseInt(frm.currentitemoptionorgno.value);

	// for (i = 0; i < parseInt(frm.itemoptioncount.value); i++) {
	for (i = 0; i < parseInt(frm.itemoptionno.length); i++) {
		if ((frm.itemoptionno[i].value.length < 1) || (frm.itemoptionno[i].value*0 != 0)) {
			alert('수량에 숫자를 입력하세요.');
			return;
		}

		if (frm.itemoptionno[i].value*1 < 0) {
			alert('수량에 마이너스를 입력할 수 없습니다.');
			return;
		}

		if (i != changedindex) {
			maxcount = maxcount - parseInt(frm.itemoptionno[i].value);
		}

		if (i != parseInt(frm.currentitemoptionidx.value)) {
			totalcount = totalcount + parseInt(frm.itemoptionno[i].value);
		}
	}

	if ((parseInt(frm.currentitemoptionorgno.value) - totalcount) < 0) {
		alert('변경가능한 수량을 초과하였습니다.');
		frm.itemoptionno[changedindex].value = maxcount;
		return;
	}

	frm.itemoptionno[frm.currentitemoptionidx.value].value = parseInt(frm.currentitemoptionorgno.value) - totalcount;
}

// 옵션변경 주문변경
function SaveItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('변경할 수량이 0입니다.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// 업체배송, 상품준비 이후
	if (confirm('업체배송이면서 상품준비 이후입니다.\n\진행 하시겠습니까?') != true){
		return;
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// 상품출고 이후
	alert('상품출고 이후입니다. 파트장에게 문의하세요.');
	return;
<% end if %>
	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품옵션변경";

		// var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= ojumunDetail.FJumunDetail.Fitemname %>\n의 옵션중\n";
		// str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 수량을 " + frm.preItemNo.value + " 에서 " + frm.itemno.value + " 로 변경 신청";

		frm.contents_jupsu.value = "str";

		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

// 옵션변경 주문변경
function ForceSaveItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('변경할 수량이 0입니다.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

	if (confirm('수정 하시겠습니까?')){
		frm.title.value = "상품옵션변경";

		// var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= ojumunDetail.FJumunDetail.Fitemname %>\n의 옵션중\n";
		// str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " 의 수량을 " + frm.preItemNo.value + " 에서 " + frm.itemno.value + " 로 변경 신청";

		frm.contents_jupsu.value = "str";

		frm.forceedit.value="Y";
		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;

// 옵션변경 맞교환
function SaveChangeItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('변경할 수량이 0입니다.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

<% if (ojumunDetail.FJumunDetail.FcurrState < "7") then %>
	// 상품출고 이후
	alert('상품출고 이전 상품입니다. 교환(옵션변경)할 수 없습니다.');
	return;
<% end if %>
	if (confirm('교환 접수(옵션변경) 하시겠습니까?')){
		frm.title.value = "교환출고(옵션변경)";

		frm.contents_jupsu.value = "str";

		<% if (ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") then %>
			frm.requiremakerid.value="<%= ojumunDetail.FJumunDetail.Fmakerid %>";
		<% end if %>

		frm.mode.value="ChangeEditItemNoPart";
		frm.submit();
	}
}

// itemoptioncount

</script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript' SRC="/cscenter/js/newcsas.js"></script>

<form name="frm" method="post" action="/cscenter/ordermaster/orderdetail_process.asp" style="margin:0px;">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
<input type="hidden" name="mode" value="itemOption">
<input type="hidden" name="forceedit" value="N">
<input type="hidden" name="requiremakerid" value="">
<input type="hidden" name="itemId" value="<%= ojumunDetail.FJumunDetail.Fitemid %>">
<input type="hidden" name="preItemOption" value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
<input type="hidden" name="preItemOptionName" value="<%= Replace(ojumunDetail.FJumunDetail.FitemOptionName, ",", "") %>">
<input type="hidden" name="preItemNo" value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
<input type="hidden" name="title" value="상품옵션변경">
<input type="hidden" name="contents_jupsu" value="">
<input type="hidden" name="contents_finish" value="정상적으로 처리되었습니다.">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문아이템정보 수정</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="110" rowspan="4"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">기존옵션</td>
		<td>[<%= ojumunDetail.FJumunDetail.Fitemoption %>] <%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">취소상태</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
		<td>

		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<% if ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then %>
	<%
	currentitemoptionidx = 0
	changedindex = 0
	%>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			현재옵션
		</td>
		<td>
			<%= "[" & ojumunDetail.FJumunDetail.FitemOption & "] " & ojumunDetail.FJumunDetail.FitemOptionName %>
		</td>
		<input type=hidden name=itemoptioncode value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(ojumunDetail.FJumunDetail.FitemOptionName, ",", "") %>">
		<td>
			<input type="text" class="text_ro" name="itemoptionno" value="<%= (ojumunDetail.FJumunDetail.Fitemno - prevregno) %>" size="3" maxlength="9" readonly> 개(<%= ojumunDetail.FJumunDetail.Fitemno %>개)

			<% if (prevregno <> 0) then %>
				<font color=red>(기존반품 : <%= prevregno %> 개)</font>
			<% end if %>
		</td>
	</tr>
	<% For i = 0 To UBound(rsOption,2) %>
		<%
		If rsOption(2,i) = "N" Or ( (rsOption(3,i)="Y") and (rsOption(4,i) - rsOption(5,i) < 1) ) Then
			itemStatus = "판매중지"
		ElseIf rsOption(3,i)="Y" Then
			If ( rsOption(4,i) - rsOption(5,i) ) < 1 Then
				itemStatus = "한정:0"
			Else
				itemStatus = "한정:" & ( rsOption(4,i) - rsOption(5,i) )
			End If
		ElseIf rsOption(6,i) <> 0 Then
			itemStatus = "기주문:" & rsOption(6,i)
		Else
			itemStatus = ""
		End If

		If rsOption(8,i) = "N" Then
			If itemStatus <> "" Then
				itemStatus = itemStatus & ", " & "사용안함"
			else
				itemStatus = "사용안함"
			end if
		End If

		If itemStatus <> "" Then
			itemStatus = " (" & itemStatus & ")"
		End If

		optionText = "[" & rsOption(0,i) & "] " & rsOption(1,i) & itemStatus

		%>




        <% ''rw rsOption(0,i) & ".." & ojumunDetail.FJumunDetail.Fitemoption %>
		<% if (rsOption(0,i) = ojumunDetail.FJumunDetail.Fitemoption) then %>
			<!-- 옵션목록에서 선택하는 대신 주문디테일에서 현재 주문정보 가져온다. -->
		<% elseif (rsOption(7,i) = "F") then %>
			<% changedindex = changedindex + 1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			변경불가옵션(<%= i+1 %>)
		</td>
		<td>
			<%=optionText%><font color=red>(옵션가 다름)</font>
		</td>
		<input type=hidden name=itemoptioncode value="<%= rsOption(0,i) %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(rsOption(1,i), ",", "") %>">
		<td width="110">
			<input type="text" class="text_ro" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNoCount(<%= (changedindex) %>)" readonly> 개
		</td>
	</tr>
		<% else %>
			<% changedindex = changedindex + 1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			변경가능옵션(<%= i+1 %>)
		</td>
		<td>
			<%=optionText%>
		</td>
		<input type=hidden name=itemoptioncode value="<%= rsOption(0,i) %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(Replace(rsOption(1,i), ",", ""), "Perl ", "Perl") %>">
		<td width="110">
			<input type="text" class="text" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNoCount(<%= (changedindex) %>)"> 개
		</td>
	</tr>
		<% end If %>
	<% Next %>
	<input type=hidden name=itemoptioncount value="<%= (UBound(rsOption,2) + 1) %>">
	<input type=hidden name=currentitemoptionidx value="<%= currentitemoptionidx %>">
	<input type=hidden name=currentitemoptionorgno value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
<% end If %>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=35 bgcolor="<%= adminColor("tabletop") %>">
			사유구분
		</td>
		<td colspan=2>
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- 일부 사유 미리 표시 -->
                <%
                '참조쿼리
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">단순변심</a>]
                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">품절</a>]
                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">상품불량</a>]
                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">기타</a>]
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height=35>
		<td colspan="3" align="center">
<% if ojumunDetail.FJumunDetail.Fcancelyn <> "Y" then %>
			<input type="button" class="button" value="옵션변경" onclick="javascript:SaveItemOptionNo()">
			<% if (C_ADMIN_AUTH or C_CSPowerUser) then %>
			<!-- 파트장 이상 -->
		    <input type="button" class="button" value="강제변경" onclick="javascript:ForceSaveItemOptionNo()">
			<% end if %>
			<input type="button" class="button" value="옵션변경 맞교환" onclick="javascript:SaveChangeItemOptionNo()">
<% else %>
			취소된 상품은 수량변경 불가
<% end if %>
		</td>
	</tr>
</table>
</form>
<div>
* 옵션변경은 <font color=red>옵션가</font>가 동일한 옵션상품만 가능합니다.<br>
* 주문당시 옵션가격에 상관없이 현재 상품정보 상의 옵션가격으로 비교합니다.<br>
* 상품할인정보(판매가,매입가 등)는 주문당시 정보가 유지됩니다.<br>
</div>

<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->