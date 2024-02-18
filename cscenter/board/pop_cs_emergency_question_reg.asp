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
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/ClassEntityManager.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_emergencyQuestionCls.asp"-->
<%
dim idx, orderserial, makerid, title, contents
dim IS_EDIT_MODE : IS_EDIT_MODE = False
idx			= requestCheckVar(request("idx"),32)
''orderserial	= requestCheckVar(request("orderserial"),32)
''makerid		= requestCheckVar(request("makerid"),32)
''title		= requestCheckVar(request("title"),128)
''contents	= requestCheckVar(request("contents"),4000)

dim oCEmergencyQuestionMaster
Set oCEmergencyQuestionMaster = New CEmergencyQuestionMaster

if (idx <> "") then
	IS_EDIT_MODE = True
	Call oCEmergencyQuestionMaster.init(dbget, rsget)
	oCEmergencyQuestionMaster.LoadOne(idx)
else
	''if (orderserial <> "") then
	''	oCEmergencyQuestionMaster.FOneItem.Forderserial = orderserial
	''end if

	''if (makerid <> "") then
	''	oCEmergencyQuestionMaster.FOneItem.Fmakerid = makerid
	''end if

	''if (title <> "") then
	''	oCEmergencyQuestionMaster.FOneItem.Ftitle = title
	''end if

	''if (contents <> "") then
	''	oCEmergencyQuestionMaster.FOneItem.Fcontents = contents
	''end if
end if


dim oorderdetail
set oorderdetail = new COrderMaster
	oorderdetail.FRectOrderSerial = orderserial

	''if (orderserial <> "") and (IS_EDIT_MODE = False) then
	''	oorderdetail.QuickSearchOrderDetail
	''
	''	if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
	''		oorderdetail.FRectOldOrder = "on"
	''		oorderdetail.QuickSearchOrderDetail
	''	end if
	''end if

dim i, j, k
dim tmpDate

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function SubmitForm() {
	var frm = document.frm;

	<% if IS_EDIT_MODE = False then %>
	if (frm.makerid.value == '') {
		alert('브랜드를 입력하세요.');
		return;
	}

	if ($('input[name=categoryGubun]:checked', '#myForm').val() == undefined) {
		alert('구분을 입력하세요.');
		return;
	}
	<% end if %>

	if (frm.title.value == '') {
		alert('제목을 입력하세요.');
		return;
	}

	if (frm.contents.value == '') {
		alert('문의내용을 입력하세요.');
		return;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

function DelForm() {
	if (confirm("정말로 삭제하시겠습니까?") == true) {
		var frm = document.frm;
		frm.mode.value = "delEmergencyQuestion";
        frm.submit();
    }
}

function TnCSTemplateGubunChanged(gubun) {
	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=32&gubun=" + gubun;
}

function TnCSTemplateGubunProcess(v, errMSG) {
	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if(v == "") {
		//
	} else {
		document.frm.contents.value = v;
		// alert(v);
	}
}

function GetOrderInfo() {
	var frm = document.frm;
	frm.action = "";
	frm.submit();
}

$( document ).ready(function() {
    $('input[name=orderserial]').keyup(function() {
		//GetOrderInfo();
	});
    $('input[name=orderserial]').change(function() {
		//GetOrderInfo();
	});
	$('input[name=orderserial]').bind('input propertychange', function() {
		//GetOrderInfo();
	});
	$(document.body).css('padding-left', '5px');
});

</script>

<!-- 구매자정보 -->
<form name="frm" id="myForm" onsubmit="return false;" action="cs_emergency_question_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= CHKIIF(IS_EDIT_MODE, "modiEmergencyQuestion", "regEmergencyQuestion") %>">
<input type="hidden" name="idx" value="<%= oCEmergencyQuestionMaster.FOneItem.Fidx %>">
<table width="99%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>긴급문의작성</b>
			    </td>
			    <td align="right">

			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">주문번호</td>
    <td bgcolor="#FFFFFF">
		<% if (IS_EDIT_MODE) then %>
		<%= oCEmergencyQuestionMaster.FOneItem.Forderserial %>
		<% else %>
    	<input type="text" class="text" name="orderserial" value="<%= oCEmergencyQuestionMaster.FOneItem.Forderserial %>">
		<% end if %>
    </td>
</tr>
<% if False and (oorderdetail.FResultCount>0) and (IS_EDIT_MODE = False) then %>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">접수상품</td>
    <td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="<%= adminColor("topbar") %>" align="center">
				<td width="30">구분</td>
				<td width="60">진행상태</td>
				<td width="70">상품코드</td>
				<td>브랜드</td>
				<td align="left">상품명<br/>[옵션명]</td>
				<td width="30">수량</td>
				<td width="70">할인가</td>
			</tr>
    		<%
			for i = 0 to oorderdetail.FResultCount - 1
				if IsNull(oorderdetail.FItemList(i).FodlvType) then oorderdetail.FItemList(i).FodlvType = ""
			%>
			<% if oorderdetail.FItemList(i).FCancelyn ="Y" then %>
			<tr align="center" height="25" bgcolor="#EEEEEE" class="gray" align="center">
			<% else %>
			<tr align="center" height="25" align="center" bgcolor="#FFFFFF">
			<% end if %>
				<td><font color="<%= oorderdetail.FItemList(i).CancelStateColor %>"><%= oorderdetail.FItemList(i).CancelStateStr %></font></td>
				<td><font color="<%= oorderdetail.FItemList(i).GetStateColor %>"><%= oorderdetail.FItemList(i).GetStateName %></font></td>
				<td>
            	    <% if oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
            	    	<% if oorderdetail.FItemList(i).fodlvfixday="G" then %>
            	    		<font color="red"><%= oorderdetail.FItemList(i).Fitemid %><br>(해외직구)</font>
            	    	<% else %>
							<font color="red"><%= oorderdetail.FItemList(i).Fitemid %><br>(업체)</font>
						<% end if %>
					<% elseif oorderdetail.FItemList(i).Fisupchebeasong="N" and InStr(",2,9,7,", CStr(oorderdetail.FItemList(i).FodlvType)) > 0 then %>
						<font color="red"><%= oorderdetail.FItemList(i).Fitemid %><br>(해외)</font>
                    <% else %>
						<%= oorderdetail.FItemList(i).Fitemid %></a>
                    <% end if %>
				</td>
				<td><%= oorderdetail.FItemList(i).Fmakerid %></td>
				<td align="left"><%= Left(oorderdetail.FItemList(i).FItemName,20) %><br/><font color="blue">[<%= Left(oorderdetail.FItemList(i).FItemoptionName,20) %>]</font></td>
				<td><%= oorderdetail.FItemList(i).FItemNo %></td>
				<td>
                    <span title="<%= oorderdetail.FItemList(i).GetEtcDiscountText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(i).GetEtcDiscountColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(i).GetEtcDiscountPrice,0) %>
                    	</font>
                    </span>
				</td>
			</tr>
		<% next %>
		</table>
    </td>
</tr>
<% end if %>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">브랜드</td>
    <td bgcolor="#FFFFFF">
		<% if (IS_EDIT_MODE) then %>
		<%= oCEmergencyQuestionMaster.FOneItem.Fmakerid %>
		<% else %>
		<% drawSelectBoxDesignerwithName "makerid", oCEmergencyQuestionMaster.FOneItem.Fmakerid %>
		<% end if %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">구분</td>
    <td bgcolor="#FFFFFF">
		<% if (IS_EDIT_MODE) then %>
		<%= oCEmergencyQuestionMaster.FOneItem.FcategoryName %>
		<% else %>
		<% Call RadioBoxCsEmergencyQuestionCategoryGubun("categoryGubun", "", "N") %>
		<% end if %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">제목</td>
    <td bgcolor="#FFFFFF">
		<input type="text" class="text" name="title" value="<%= oCEmergencyQuestionMaster.FOneItem.Ftitle %>" size="40">
		<% SelectBoxCSTemplateGubunNew "32", "csreg_template", "" %>
		<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">문의내용</td>
    <td bgcolor="#FFFFFF">
		<textarea class="textarea" name="contents" rows="15" cols="80"><%= oCEmergencyQuestionMaster.FOneItem.Fcontents %></textarea>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">상태</td>
    <td bgcolor="#FFFFFF">
		<% if IS_EDIT_MODE then %>
		<%= CsEmergencyQuestionCurrStateToName(oCEmergencyQuestionMaster.FOneItem.FcurrState) %>
		<% end if %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>" width="120">작성일시</td>
    <td bgcolor="#FFFFFF">
		<%
		if IS_EDIT_MODE then
			tmpDate = CDate(oCEmergencyQuestionMaster.FOneItem.Fregdate)
			if DateDiff("d", tmpDate, Now()) then
				rw FormatDateTime(tmpDate, 4)
			else
				rw oCEmergencyQuestionMaster.FOneItem.Fregdate
			end if
		end if
		%>
    </td>
</tr>
<% if IS_EDIT_MODE = False or (IS_EDIT_MODE and oCEmergencyQuestionMaster.FOneItem.FcurrState="1") then %>
<tr height="40">
    <td bgcolor="#FFFFFF" colspan="2" align="center">
		<input type="button" value=" 저 장 하 기 " class="button" onClick="SubmitForm();">
		<% if IS_EDIT_MODE then %>
		&nbsp;
		&nbsp;
		<input type="button" value="삭제" class="button" onClick="DelForm();" <%= CHKIIF(oCEmergencyQuestionMaster.FOneItem.FcurrState="1", "", "disabled") %>>
		<% end if %>
    </td>
</tr>
<% end if %>
</table>
</form>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
