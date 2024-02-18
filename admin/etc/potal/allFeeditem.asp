<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/potal/potalCls.asp"-->
<%
Dim mallid, oItem, page, i, makerid, itemname, itemid, onlyValidMargin
mallid				= requestCheckvar(request("mallid"),32)
page				= requestCheckvar(request("page"),10)
makerid				= requestCheckvar(request("makerid"),32)
itemname			= request("itemname")
itemid				= request("itemid")
onlyValidMargin		= requestCheckvar(request("onlyValidMargin"),32)
research            = requestCheckvar(request("research"),32)

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

Set oItem = new CPotal
	oItem.FCurrPage				= page
	oItem.FRectMakerid			= makerid
	oItem.FRectItemname			= itemname
	oItem.FRectItemid			= itemid
	oItem.FRectOnlyValidMargin	= onlyValidMargin
	oItem.FRectMallGubun		= mallid
	oItem.FPageSize	= 15

	If (research <> "") Then
	    oItem.getAllItemList
	End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function makeFile(){
	if(confirm("팝업 생성 후 완료까지 10분정도 소요됩니다.\n\nEP데이터를 생성 하시겠습니까?")){
		var popwin=window.open('<%=apiURL%>/outmall/googleFeed/dailyFeedTxt.asp','mdConfirm','width=500,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}
}
</script>
<% If mallid = "ggshop" Then %>
<!-- #include virtual="/admin/etc/potal/inc_googleHead.asp" -->
<% End If %>

>> 전체리스트 &nbsp; <input type="button" class="button" value="Feed생성" onclick="makeFile();" style="color:blue;font-weight:bold;">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mallid" value="<%= mallid %>">
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE">
<tr>
	<td class="a">
		브 랜 드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		상품명: <input type="text" name="itemname" value="<%= itemname %>" size="20" maxlength="32" class="text">
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
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
1.상품이 판매중, 사용중<br>
<s>2.상품최종수정일이 현재시간보다 1년이하<br></s>
2.상품최종수정일이 현재시간보다 25개월이하이거나 최근판매가 1개이상<br>
3.판매제외 브랜드가 아닌것<br>
4.판매제외 상품이 아닌것<br>
5.2Depth이상에 속한 상품<br>
6.성인상품 아닌것<br>
7.감성채널 > BooK 관리카테고리 제외<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="16" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(oItem.FTotalPage,0) %> 총건수: <%= FormatNumber(oItem.FTotalCount,0) %></td>
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
<% For i=0 to oItem.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="20" align="center">
	<td><img src="<%= oItem.FItemList(i).Fsmallimage %>" width="50"></td>
    <td><%= oItem.FItemList(i).FItemid %></td>
    <td><%= oItem.FItemList(i).FItemname %></td>
    <td><%= oItem.FItemList(i).FMakerid %></td>
    <td>
        <% if oItem.FItemList(i).IsSoldOut then %>
            <% if oItem.FItemList(i).FSellyn="N" then %>
            <font color="red">품절</font>
            <% else %>
            <font color="red">일시<br>품절</font>
            <% end if %>
        <% end if %>
    </td>
	<td><%= oItem.FItemList(i).FRegdate %></td>
	<td><%= oItem.FItemList(i).FLastupdate %></td>
	<td>
        <%= FormatNumber(oItem.FItemList(i).FSellcash,0) %>
	</td>
	<td>
        <% if oItem.FItemList(i).Fsellcash<>0 then %>
        <%= CLng(10000-oItem.FItemList(i).Fbuycash/oItem.FItemList(i).Fsellcash*100*100)/100 %> %
        <% end if %>
	</td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="16" align="center" bgcolor="#FFFFFF">
        <% if oItem.HasPreScroll then %>
		<a href="javascript:goPage('<%= oItem.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + oItem.StartScrollPage to oItem.FScrollCount + oItem.StartScrollPage - 1 %>
    		<% if i>oItem.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if oItem.HasNextScroll then %>
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
