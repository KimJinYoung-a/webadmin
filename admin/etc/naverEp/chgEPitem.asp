<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim allEP, page, i, makerid, itemname, itemid, onlyValidMargin
page				= request("page")
makerid				= request("makerid")
itemname			= request("itemname")
itemid				= request("itemid")
onlyValidMargin		= request("onlyValidMargin")
If page = "" Then page = 1

Set allEP = new epShop
	allEP.FCurrPage				= page
	allEP.FRectMakerid			= makerid
	allEP.FRectItemname			= itemname
	allEP.FRectItemid			= itemid
	allEP.FRectOnlyValidMargin	= onlyValidMargin
	allEP.FPageSize	= 15
	allEP.ChgEpItemList
%>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function makeFile(){
	if(confirm("요약EP 모수 생성 기능입니다.\n\n데이터를 생성하시겠습니까?")){
		var popwin=window.open('<%=apiURL%>/outmall/naverEP/makeDailyNewVerEP_SugiFile.asp?epType=chg','makeFile','width=500,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}
}
function makeText(){
	if(confirm("먼저 모수 생성 후 EP를 생성해야 현 시간의 데이터가 반영 됩니다.\n\n팝업 생성 후 완료까지 1분정도 소요됩니다.\n\nEP데이터를 생성하시겠습니까?")){
		var popwin=window.open('<%=apiURL%>/outmall/naverEP/makeDailyNewVerEP_File.asp?epType=chg','makeText','width=500,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}
}
</script>
<!-- #include virtual="/admin/etc/potal/inc_naverHead.asp" -->
>> 요약EP리스트 &nbsp;
<input type="button" class="button" value="1.요약EP모수생성" onclick="makeFile();" style="color:blue;font-weight:bold;">
<input type="button" class="button" value="2.요약EP생성" onclick="makeText();" style="color:blue;font-weight:bold;">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		브 랜 드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		상품번호: <input type="text" name="itemid" value="<%= itemid %>" size="60" class="text"> &nbsp;
		<input type="checkbox" name="onlyValidMargin" <%= ChkIIF(onlyValidMargin="on","checked","") %> >마진 15%이상 상품만 보기
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>
<br>
※기본 검색조건<br>
1.당일 오전10시이전에는 상품최종수정일이 현재시간보다 12시간이하, 당일 오전10시이후부터는 상품최종수정일이 현재시간보다 5시간이하<br>
2.2Depth이상에 속한 상품<br>
3.성인상품 아닌것<br>
4.딜상품 제외<br>
5.감성채널 > BooK 관리카테고리 제외<br>
6.키덜트>성인토이>콘돔/젤 관리카테고리 제외<br>
7.뷰티 관리카테고리면서 상품명에 "뚜껑" 포함된 상품 제외<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(allEP.FTotalPage,0) %> 총건수: <%= FormatNumber(allEP.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>이미지</td>
    <td>상품코드</td>
    <td>상품명</td>
    <td>브랜드ID</td>
    <td>품절여부</td>
	<td>상품등록일</td>
	<td>상품최종수정일</td>
	<td>판매가</td>
	<td>마진</td>
</tr>
<% For i=0 to allEP.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20" align="center">
	<td><img src="<%= allEP.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= allEP.FItemList(i).FItemid %></td>
    <td><%= allEP.FItemList(i).FItemname %></td>
    <td><%= allEP.FItemList(i).FMakerid %></td>
    <td>
        <% if allEP.FItemList(i).IsSoldOut then %>
            <% if allEP.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td><%= allEP.FItemList(i).FRegdate %></td>
	<td><%= allEP.FItemList(i).FLastupdate %></td>
	<td>
        <%= FormatNumber(allEP.FItemList(i).FSellcash,0) %>
	</td>
	<td>
        <% if allEP.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-allEP.FItemList(i).Fbuycash/allEP.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if allEP.HasPreScroll then %>
		<a href="javascript:goPage('<%= allEP.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + allEP.StartScrollPage to allEP.FScrollCount + allEP.StartScrollPage - 1 %>
    		<% if i>allEP.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if allEP.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->