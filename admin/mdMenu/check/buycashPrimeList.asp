<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/itemcheckCls.asp" -->
<%
If (session("ssBctID")<>"kjy8517") AND (session("ssBctID")<>"eastone") AND (session("ssBctID")<>"hrkang97") Then
	response.write "관리자용"
	response.end
End If

Dim oCheck, research, i, page, dispCate, itemid, makerid, maxDepth
dispCate	= requestCheckvar(request("disp"),16)
research	= requestCheckvar(request("research"),2)
itemid  	= request("itemid")
makerid		= requestCheckvar(request("makerid"),32)
maxDepth	= 1

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

SET oCheck = new cCheck
	oCheck.FRectCateCode	= dispCate
	oCheck.FRectItemid		= itemid
	oCheck.FRectMakerid		= makerid
	oCheck.getBuycashPrimeList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function pop_optionEdit(v){
	var pwin = window.open('/admin/itemmaster/itemlist.asp?itemid='+v,'popOutMallEtcLink','width=1200,height=700,scrollbars=yes,resizable=yes');
	pwin.focus();
}
function pop_extsitejungsancheck(v){
	var pwin2 = window.open('/admin/etc/extsitejungsan_check.asp?itemid='+v,'pop_extsitejungsancheck','width=1500,height=700,scrollbars=yes,resizable=yes');
	pwin2.focus();
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<!-- #include virtual="/admin/mdMenu/check/checkTab.asp"-->
<input type="hidden" name="vTab" value="<%= vTab %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2"><h2>※상품 및 옵션추가금액의 공급가가 소수점으로 나타나 있는 것</h2></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		&nbsp;&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	</td>
	<td align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oCheck.FResultCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>판매가</td>
	<td>쿠폰제외가</td>
	<td>공급가</td>
	<td>상품최종수정일</td>
	<td>1Depth전시카테고리명</td>
	<td>관리</td>
</tr>
<% If oCheck.FResultCount > 0 Then %>
<% For i=0 to oCheck.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><a href="<%=wwwURL%>/<%= oCheck.FItemList(i).FItemID %>" target="_blank"><%= oCheck.FItemList(i).FItemID %></a><br /><%= oCheck.FItemList(i).FItemoption %></td>
	<td><%= oCheck.FItemList(i).FMakerid %></td>
	<td><%= oCheck.FItemList(i).FSellCash %></td>
	<td><%= oCheck.FItemList(i).FitemcostCouponNotApplied %></td>
	<td><%= oCheck.FItemList(i).FBuycash %></td>
	<td><%= oCheck.FItemList(i).FLastupdate %></td>
	<td><%= oCheck.FItemList(i).FCatename %></td>
	<td>
		<input type="button" class="button" value="Check" onclick="pop_optionEdit('<%= oCheck.FItemList(i).FItemID %>');">
		<input type="button" class="button" value="Check2" onclick="pop_extsitejungsancheck('<%= oCheck.FItemList(i).FItemID %>');">
	</td>
</tr>
<% Next %>
<% Else %>
<tr height="50">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
		데이터가 없습니다
    </td>
</tr>
<% End If %>
</table>
<% SET oCheck = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->