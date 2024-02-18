<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/barcodeCls.asp"-->
<%
Dim itemid, page, itemgubun, useYN
Dim i, obarcode
page    				= request("page")
itemid  				= request("itemid")
useYN					= request("useYN")
'itemgubun				= request("itemgubun")

If page = "" Then page = 1

'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

SET obarcode = new CBarcode
	obarcode.FCurrPage			= page
	obarcode.FPageSize			= 20
	obarcode.FRectItemID		= itemid
	obarcode.FRectUseYN			= useYN
	obarcode.FRectItemGubun		= itemgubun
	obarcode.getBarcodelist
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function pop_BarcodeCont(idx){
	var pCM = window.open("/admin/itemmaster/pop_barcode.asp?idx="+idx,"pop_BarcodeCont","width=800,height=300,scrollbars=yes,resizable=yes");
	pCM.focus();
}

function popBarcodeMulti() {
	var pop = window.open("/admin/itemmaster/pop_barcode_multi.asp","popBarcodeMulti","width=500,height=500,scrollbars=yes,resizable=yes");
	pop.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
	<!--
		구분 :
		<select name="itemgubun" class="select">
			<option value="">전체</option>
		</select>
	-->
		&nbsp;
		등록여부 :
		<select name="useYN" class="select">
			<option value="">전체</option>
			<option value="Y" <%= CHkIIF(useYN="Y","selected","") %>>등록완료</option>
			<option value="N" <%= CHkIIF(useYN="N","selected","") %>>등록이전</option>
		</select>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<div style="height:5px;"></div>

<input type="button" class="button" value="일괄등록" onClick="popBarcodeMulti()">

<div style="height:5px;"></div>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" name="cmdparam" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= FormatNumber(obarcode.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(obarcode.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">idx</td>
	<td width="100"><b>범용바코드</b></td>
	<td width="30"><b>구분</b></td>
	<td width="80"><b>상품코드</b></td>
	<td width="40"><b>옵션<br />코드</b></td>
	<td>브랜드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="80">입력일</td>
	<td width="80">등록일</td>
	<td>등록상품명</td>
	<td width="100">등록자</td>
</tr>

<% For i=0 to obarcode.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF" onclick="pop_BarcodeCont('<%= obarcode.FItemList(i).FIdx %>')" style="cursor:pointer;">
	<td align="center"><%= obarcode.FItemList(i).FIdx %></td>
	<td align="center"><%= obarcode.FItemList(i).FBarcode %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemgubun %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemid %></td>
	<td align="center"><%= obarcode.FItemList(i).FItemoption %></td>
	<td align="center"><%= obarcode.FItemList(i).Fmakerid %></td>
	<td align="left"><%= obarcode.FItemList(i).Fitemname %></td>
	<td align="left"><%= obarcode.FItemList(i).Fitemoptionname %></td>
	<td align="center"><%= Left(obarcode.FItemList(i).FRegdate, 10) %></td>
	<td align="center"><%= Left(obarcode.FItemList(i).FReservedDate, 10) %></td>
	<td align="left"><%= nl2br(db2html(obarcode.FItemList(i).FReservedCont)) %></td>
	<td align="center"><%= obarcode.FItemList(i).Freguserid %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
        <% if obarcode.HasPreScroll then %>
		<a href="javascript:goPage('<%= obarcode.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + obarcode.StartScrollPage to obarcode.FScrollCount + obarcode.StartScrollPage - 1 %>
    		<% if i>obarcode.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if obarcode.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</form>
</table>
<% SET obarcode = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
