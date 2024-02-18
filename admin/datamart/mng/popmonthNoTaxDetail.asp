<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월별 면세 매출
' History : 2011.06.02 eastone 생성
'			2012.07.11 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/datamart/noTaxSummary.asp"-->
<%
Dim page, research ,yyyymm1,yyyymm2,yyyy1,mm1,yyyy2,mm2 ,makerid, pgn,ps ,placeALL ,olist ,TTLcnt, TTLsum ,i ,sellsite
dim grpsum
	page  = requestCheckvar(request("page"),10)
	research = requestCheckvar(request("research"),10)
	makerid = requestCheckvar(request("makerid"),32)
	pgn = requestCheckvar(request("pgn"),32)
	ps = requestCheckvar(request("ps"),32)
	placeALL = requestCheckvar(request("placeALL"),32)
	yyyymm1 = requestCheckvar(request("yyyymm1"),10)
	yyyymm2 = requestCheckvar(request("yyyymm2"),10)
	yyyy1 = requestCheckvar(request("yyyy1"),10)
	mm1 = requestCheckvar(request("mm1"),10)
	yyyy2 = requestCheckvar(request("yyyy2"),10)
	mm2 = requestCheckvar(request("mm2"),10)
	sellsite = requestCheckvar(request("sellsite"),32)
	grpsum = requestCheckvar(request("grpsum"),10)
	
IF (placeALL<>"") then
    pgn = Left(placeALL,5)
    ps = Mid(placeALL,6,255)
ELSE 
    placeALL = pgn + ps
END IF

if (yyyy1<>"") then
    yyyymm1 = yyyy1 + "-" + mm1
    yyyymm2 = yyyy2 + "-" + mm2
end if

if (yyyymm1="") then
    yyyymm1 = Left(CStr(now()),4) + "-" + Mid(CStr(now()),6,2)
end if

if (yyyymm2="") then
    yyyymm2 = yyyymm1
end if

IF page="" then page=1

set olist = new CNoTaxList
	olist.FPageSize=300
	olist.FCurrPage= page
	olist.FRectMakerid = makerid
	olist.FRectStYYYYMM = yyyymm1
	olist.FRectEdYYYYMM = yyyymm2
	olist.FRectplaceGubun = pgn
	olist.FRectplaceSub   = ps
	olist.FRectSellSite   = sellsite	
	olist.FRectGrpSum   = grpsum	
	olist.getMonthNoTaxDetailGroup

%>

<script language='javascript'>

function goPage(pg){
    iURL="?page=" + pg + "&yyyymm1=<%= yyyymm1 %>&yyyymm2=<%= yyyymm2 %>&makerid=<%= makerid %>&pgn=<%= pgn %>&ps=<%= ps %>&grpsum=<%=grpsum%>";
    location.href=iURL ;
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'simpleitemedit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function pop_itemedit_off_edit(ibarcode){
	var pop_itemedit_off_edit = window.open('/common/offshop/item/pop_itemedit_off_edit.asp?barcode=' + ibarcode,'pop_itemedit_off_edit','width=1024,height=768,resizable=yes,scrollbars=yes');
	pop_itemedit_off_edit.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		기간 :
		<% DrawYMYMBox LEFT(yyyymm1,4),RIGHT(yyyymm1,2),LEFT(yyyymm2,4),RIGHT(yyyymm2,2) %>
        <% if (FALSE) then %>
        <input type="checkbox" name="grpsum" <%=CHKIIF(grpsum="on","checked","")%>> 합계로보기
        <% end if %>
	    쇼핑몰 선택 :
	    
	    <% call drawSelectBoxXSiteOrderInputPartner("sellsite", sellsite) %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    구분     : <% call drawBoxNotaxPlaceGubun("placeALL",placeALL) %>
	    
		브랜드 ID : <% call drawSelectBoxDesigner("makerid",makerid) %>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= olist.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olist.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>년/월</td>
	<td>구분</td>
	<td>브랜드ID</td>
	<td>상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td>사이트<Br>구분</td>		
	<td>건수</td>
	<td>금액</td>
	<% if (FALSE) then %>
	<td></td>
	<% end if %>
</tr>
<% for i=0 to olist.FResultCount-1 %>
<%
TTLcnt = TTLcnt + olist.FItemList(i).Fitemno
TTLsum = TTLsum + olist.FItemList(i).FnotaxPrice
%>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= olist.FItemList(i).FYYYYMM %></td>
    <td><%= olist.FItemList(i).FPlaceSubName %></td>
    <td><%= olist.FItemList(i).Fmakerid %></td>
<% if (FALSE) then %>
    <% if olist.FItemList(i).Fitemgubun<>"10" then %>
    <td ><a target="blank" href="/common/offshop/item/pop_itemedit_off_edit.asp?barcode=<%= olist.FItemList(i).Fitemgubun + format00(8,olist.FItemList(i).Fitemid) + olist.FItemList(i).Fitemoption%>"><%= olist.FItemList(i).Fitemid %></a></td>
    <% else %>
    <td ><a target="blank" href="/common/pop_simpleitemedit.asp?itemid=<%= olist.FItemList(i).Fitemid %>"><%= olist.FItemList(i).Fitemid %></a></td>
    <% end if %>
<% else %>
    <td ><%= olist.FItemList(i).Fitemid %></td>
<% end if %>
    <td align="left"><%= olist.FItemList(i).FItemName %></td>
    <td align="left">
    	<%= olist.FItemList(i).FItemOptionName %>
    </td>
    <td ><%= olist.FItemList(i).fsitename %></td>
    <td><%= FormatNumber(olist.FItemList(i).Fitemno,0) %></td>
    <td align="right" ><%= FormatNumber(olist.FItemList(i).FnotaxPrice,0) %></td>
    <% if (FALSE) then %>
    <td>
        <% if olist.FItemList(i).FCurVatinclude="Y" then %>
        <%= olist.FItemList(i).FCurVatinclude %>
        <% end if %>
    </td>
    <% end if %>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF" >
    <td>합계</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td><%= FormatNumber(TTLcnt,0) %></td>
    <td align="right" ><%= FormatNumber(TTLsum,0) %></td>
    <% if (FALSE) then %>
    <td></td>
    <% end if %>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td colspan="10" align="center">
        <!-- 페이지 시작 -->
    	<%
    		if olist.HasPreScroll then
    			Response.Write "<a href='javascript:goPage(" & olist.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
    		else
    			Response.Write "[pre] &nbsp;"
    		end if
    
    		for i=0 + olist.StartScrollPage to olist.FScrollCount + olist.StartScrollPage - 1
    
    			if i>olist.FTotalpage then Exit for
    
    			if CStr(page)=CStr(i) then
    				Response.Write " <font color='red'>[" & i & "]</font> "
    			else
    				Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
    			end if
    
    		next
    
    		if olist.HasNextScroll then
    			Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
    		else
    			Response.Write "&nbsp; [next]"
    		end if
    	%>
    	<!-- 페이지 끝 -->
    </td>
</tr>
</table>
<%
set olist = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
