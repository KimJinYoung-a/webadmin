<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 원가결제리스트
' History : 2023.09.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%

dim i, research, page, ExcDel, productidx, sheetidx, makerid, purchasetype, codelist, reportIdx, itemid
	productidx = requestCheckVar(trim(getNumeric(request("productidx"))),8)
	sheetidx = requestCheckVar(trim(getNumeric(request("sheetidx"))),8)
	makerid = requestCheckVar(trim(request("makerid")),32)
	purchasetype = requestCheckVar(request("purchasetype"),2)
	codelist = requestCheckVar(request("codelist"),32)
	reportIdx = requestCheckVar(trim(getNumeric(request("reportIdx"))),8)
	itemid      = requestCheckvar(request("itemid"),1500)
page = requestCheckVar(request("page"),8)
ExcDel = requestCheckVar(request("ExcDel"),1)
research = requestCheckVar(request("research"),1)

if page = "" then page = "1"
if ExcDel = "" and research="" then ExcDel = "Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oCPurchasedProductPay
set oCPurchasedProductPay = new CPurchasedProduct
	oCPurchasedProductPay.FCurrPage = page
	oCPurchasedProductPay.Fpagesize = 50
    oCPurchasedProductPay.FRectExcDel = ExcDel
	oCPurchasedProductPay.FRectproductidx = productidx
    oCPurchasedProductPay.FRectSheetidx = sheetidx
	oCPurchasedProductPay.FRectpurchasetype = purchasetype
	oCPurchasedProductPay.FRectmakerid = makerid
	oCPurchasedProductPay.FRectcodelist = codelist
	oCPurchasedProductPay.FRectreportIdx = reportIdx
	oCPurchasedProductPay.FRectItemid       = itemid
	oCPurchasedProductPay.GetPurchasedProductItemAllPayList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type='text/javascript'>

function jsDetailView(idx) {
	var popwin = window.open('/admin/newstorage/PurchasedProductModify.asp?idx='+idx+'&menupos=<%= menupos %>','addreg','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function SubmitFrm(pg) {
	document.frm.page.value=pg;
	document.frm.target = "";
	document.frm.action = "";
	document.frm.submit();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/newstorage/PurchasedProductPayList_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<style>
textarea:-webkit-input-placeholder {color:#acacac;}
textarea:-moz-placeholder {color:#acacac;}
textarea:-ms-input-placeholder {color:#acacac;}
.placeholder { color: #acacac; }
</style>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 마스터IDX : <input type="text" class="text" name="productidx" value="<%= productidx %>" size="8" maxlength=10>
		&nbsp;
		* 원가상세IDX : <input type="text" class="text" name="sheetidx" value="<%= sheetidx %>" size="8" maxlength=10>
		&nbsp;
		* 브랜드ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		&nbsp;
		* 주문코드 : <input type="text" class="text" name="codelist" value="<%= codelist %>" size="20" maxlength=32>
		&nbsp;
		* 품의번호 : <input type="text" class="text" name="reportIdx" value="<%= reportIdx %>" size="8" maxlength=10>
		<Br><Br>
		* 상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<a href="#" onClick="SubmitFrm('1'); return false;"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" >
		* 구매유형 :
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
	</td>
</tr>
<tr>
    <td bgcolor="#FFFFFF" >
        <label><input type="checkbox" name="ExcDel" value="Y" <%=chkIIF(ExcDel="Y","checked","")%> /> 삭제건 제외</label>
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br />

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= oCPurchasedProductPay.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCPurchasedProductPay.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=60>마스터<br>IDX</td>
    <td width="60">품의번호</td>
    <td width="70">품의금액</td>
    <td width="80">결제요청서IDX</td>
    <td width="100">결제요청일</td>
    <td width="100">결제일</td>
    <td width="100">결제요청금액(원)</td>
    <td width="70">결제방법</td>
    <td>자금용도</td>
    <td>거래처</td>
    <td width="70">상태</td>
    <td width="50">비고</td>
</tr>
<% if oCPurchasedProductPay.FResultcount>0 then %>
<% for i=0 to oCPurchasedProductPay.FResultcount-1 %>
<tr bgcolor="<%= CHKIIF(IsNull(oCPurchasedProductPay.FItemList(i).Fdeldt), "#FFFFFF", "#EEEEEE") %>" align="center" height="25">
    <td><%= oCPurchasedProductPay.FItemList(i).Fidx %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).freportIdx %></td>
    <td align="right"><%= FormatNumber(oCPurchasedProductPay.FItemList(i).freportPrice, 0) %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestidx %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestdate %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpaydate %></td>
    <td align="right"><%= FormatNumber(oCPurchasedProductPay.FItemList(i).fpayRequestPrice, 0) %></td>
    <td align="center"><%= fnGetPayType(oCPurchasedProductPay.FItemList(i).fpaytype) %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).fpayRequestTitle %></td>
    <td align="center"><%= oCPurchasedProductPay.FItemList(i).fcust_nm %></td>
    <td align="center"><%= fnGetPayRequestState(oCPurchasedProductPay.FItemList(i).fpayrequeststate) %></td>
    <td align="center">
        <input type="button" class="button" value="상세" onclick="jsDetailView(<%= oCPurchasedProductPay.FItemList(i).Fidx %>);">
    </td>
</tr>
<% next %>
    <tr bgcolor="FFFFFF">
		<td colspan="12" align="center">
        	<% if oCPurchasedProductPay.HasPreScroll then %>
				<a href="javascript:SubmitFrm('<%= oCPurchasedProductPay.StartScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oCPurchasedProductPay.StartScrollPage to oCPurchasedProductPay.FScrollCount + oCPurchasedProductPay.StartScrollPage - 1 %>
				<% if i>oCPurchasedProductPay.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:SubmitFrm('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oCPurchasedProductPay.HasNextScroll then %>
				<a href="javascript:SubmitFrm('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oCPurchasedProductPay=nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
