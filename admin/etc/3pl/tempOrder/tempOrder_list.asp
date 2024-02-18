<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.01.01 이상구 생성
'           2016.12.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/tempOrderCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn, companyid, matchState
Dim page, research
	companyid	= requestCheckVar(request("companyid"),32)
	page     	= requestCheckVar(request("page"),10)
	research    = requestCheckVar(request("research"),10)
	matchState  = requestCheckVar(request("matchState"),10)

If page = "" Then page = 1

if (research = "") then
	matchState = "I"
end if


dim oCTPLTempOrder
set oCTPLTempOrder = New CTPLTempOrder
	oCTPLTempOrder.FCurrPage				= page
	oCTPLTempOrder.FRectCompanyID			= companyid
	oCTPLTempOrder.FPageSize				= 20
	oCTPLTempOrder.FRectMatchState			= matchState

oCTPLTempOrder.GetTPLTempOrderList()


dim isCheckBoxDisable, pOrderSerial

%>
<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid, prdcode) {
	var popwin = window.open("pop_product_modify.asp?companyid=" + companyid + "&prdcode=" + prdcode,"jsPopModi","width=600 height=400 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit(frm) {
	frm.submit();
}

function apiOrderProcess(){
	var frm = document.frm;
	var companyid, partnercompanyid;

	var obj = frm.partnercompany;

	if (obj.value === '') {
		alert('제휴사를 선택하세요.');
		return;
	}

	var v = obj.options[obj.selectedIndex].text;

    if (confirm(""+v+"몰의 주문 연동 등록 하시겠습니까?")) {
		v = obj.value.split(',');
		companyid = v[0];
		partnercompanyid = v[1];

		frm = document.frmXSiteOrder;
		frm.mode.value = "getxsiteorderlist";
		frm.companyid.value = companyid;
		frm.partnercompanyid.value = partnercompanyid;
		frm.action = "3PLSiteOrder_Ins_Process.asp"
		frm.submit();
    }
}

function CheckProduct(o) {
	var frm;
	if (o.checked) {
		hL(o);
	} else {
		dL(o);
	}
}

function fnCheckValidAll(bool, comp) {
	var obj;
	for (var i = 0; ; i++) {
		obj = document.getElementById("cksel_" + i);
		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		obj.checked = bool;
		CheckProduct(obj);
	}
}

function SubmitInputOrder() {
    var checkedOrderSerial = "";
	var obj;
	var frm = document.frmXSiteOrder;

	for (var i = 0; ; i++) {
		obj = document.getElementById("cksel_" + i);

		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		if (obj.checked == true) {
			checkedOrderSerial = checkedOrderSerial + "," + obj.value;
		}
	}

    if (checkedOrderSerial == "") {
        alert('선택 주문이 없습니다.');
        return;
    }

    if (confirm('주문을 입력 하시겠습니까?')) {
        frm.mode.value = "add";
		frm.action = "orderInput_Process.asp";

		frm.arrOutMallOrderSerial.value = checkedOrderSerial;
        frm.submit();
    }
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" height="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 고객사 : <% Call SelectBoxCompanyID("companyid", companyid, CHKIIF(useyn="Y", "Y", "")) %>
	    &nbsp;&nbsp;
	    * 처리상태 :
		<select class="select" name="matchState">
			<option value='' <%= chkIIF(matchState="","selected","") %> >전체</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >엑셀등록</option>
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >주문입력완료</option>
     	</select>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
</table>

<div style="float: left; padding:5px;">
	<input type="button" class="button" value="1. 엑셀 등록" onClick="jsPopModi('', '')" disabled>
	<% Call SelectBoxApiInput(companyid, "partnercompany", "", "Y") %>
	<input type="button" class="button" value="1. API연동 등록" onClick="apiOrderProcess()">
</div>
<div style="float: right; padding:5px;">
	<input type="button" class="button" value="2. 선택내역 주문입력" onClick="SubmitInputOrder()">
</div>

</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		검색결과 : <b><%= FormatNumber(oCTPLTempOrder.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLTempOrder.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="20"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked);"></td>
	<td>고객사</td>
    <td>제휴몰</td>
    <td>제휴<br />주문번호</td>
    <td>제휴<br />주문상세</td>
    <td>주문자<br />수령인</td>
    <td>물류코드</td>
    <td>고객사<br />상품코드</td>
    <td>고객사<br />옵션</td>
    <td>제휴상품명<br />상품명</td>
    <td>제휴옵션명<br />옵션명</td>
    <td>수량</td>
	<td>물류<br />주문번호</td>
    <td>매칭상태</td>
    <td>비고</td>
</tr>
<% if (oCTPLTempOrder.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLTempOrder.FResultCount - 1) %>
<%
isCheckBoxDisable = False
if (oCTPLTempOrder.FItemList(i).Fprdcode = "") or IsNull(oCTPLTempOrder.FItemList(i).Fprdcode) then
	'// 물류코드 매칭이전
	isCheckBoxDisable = True
elseif (oCTPLTempOrder.FItemList(i).FOrderSerial <> "") then
	'// 주문입력완료
	isCheckBoxDisable = True
end if
%>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td>
			<input type="checkbox" id="cksel_<%= i %>" name="cksel" value="<%= oCTPLTempOrder.FItemList(i).FOutMallOrderSerial %>" onclick="CheckProduct(this);" <%= CHKIIF(isCheckBoxDisable = "Y", "disabled" ,"")%> >
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).Fcompanyid %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FSellSiteName %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOutMallOrderSerial %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOrgDetailKey %></td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).FOrderName %><br />
			<%= oCTPLTempOrder.FItemList(i).FReceiveName %>
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).Fprdcode %></td>
		<td><%= oCTPLTempOrder.FItemList(i).ForderItemID %></td>
		<td><%= oCTPLTempOrder.FItemList(i).ForderItemOption %></td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).ForderItemName %>
			<% if (oCTPLTempOrder.FItemList(i).ForderItemName<>oCTPLTempOrder.FItemList(i).Fprdname) then %>
			<br /><font color="#FF0000"><%= oCTPLTempOrder.FItemList(i).Fprdname %></font>
			<% end if %>
		</td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).ForderItemOptionName %>
			<% if (oCTPLTempOrder.FItemList(i).ForderItemOptionName<>oCTPLTempOrder.FItemList(i).Fprdoptionname) then %>
			<br /><font color="#FF0000"><%= oCTPLTempOrder.FItemList(i).Fprdoptionname %></font>
			<% end if %>
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).FItemOrderCount %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOrderSerial %></td>
		<td><%= oCTPLTempOrder.FItemList(i).getmatchStateString() %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="21" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLTempOrder.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLTempOrder.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLTempOrder.StartScrollPage to oCTPLTempOrder.FScrollCount + oCTPLTempOrder.StartScrollPage - 1 %>
	    		<% if i>oCTPLTempOrder.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLTempOrder.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="21">검색결과가 없습니다.</td>
    </tr>
<% end if %>

</table>

<form name="frmXSiteOrder" method="post" action="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="companyid" value="">
	<input type="hidden" name="partnercompanyid" value="">
	<input type="hidden" name="arrOutMallOrderSerial" value="">
</form>

<%
set oCTPLTempOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
