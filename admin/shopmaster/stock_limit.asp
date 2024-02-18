<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [검토]한정관리
' History : 		   이상구 생성
'			2016.03.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%
dim shopid, diff, isusing, diffdiv, mwdiv, orderby , currPage, OnlySellyn, BasicMonth, i, itemid
dim searchtype, rackcode2, fromrackcode2, torackcode2
dim excits
	shopid  		= request("shopid")
	diff    		= trim(request("diff"))
	isusing 		= requestcheckvar(request("isusing"),1)
	diffdiv 		= trim(request("diffdiv"))
	OnlySellyn 		= request("OnlySellyn")
	mwdiv   		= request("mwdiv")
	orderby 		= request("orderby")
	currPage 		= getNumeric(requestcheckvar(request("cp"),10))
	itemid 			= getNumeric(requestcheckvar(request("itemid"),10))
	searchtype  	= requestCheckvar(request("searchtype"),1)
	rackcode2   	= requestCheckvar(request("rackcode2"),2)
	fromrackcode2  	= requestCheckvar(request("fromrackcode2"),2)
	torackcode2  	= requestCheckvar(request("torackcode2"),2)
	excits  		= requestCheckvar(request("excits"),2)

IF currPage="" Then currPage = 1
if (diffdiv = "") then
	diffdiv = "percent"
end if

if (diff = "") then
	diff = "30"
end if

if (diff < 0) then
	diff = -1 * diff
end if

if searchtype="" then searchtype = "F"
if ((request("research") = "") and (isusing = "")) then
	isusing = "Y"
end if

if ((request("research") = "") and (OnlySellyn = "")) then
	OnlySellyn = "YS"
end if

if ((request("research") = "") and (mwdiv = "")) then
	mwdiv = "MW"
end if

if (request("research") = "") then
	excits = "Y"
end if


BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummarystock
set osummarystock = new CSummaryItemStock
	osummarystock.FCurrPage = currPage
	osummarystock.FPageSize    = 100
	osummarystock.FRectMakerid = shopid
	osummarystock.FRectParameter = diff
	osummarystock.FRectDiffDiv = diffdiv
	osummarystock.FRectOnlyIsUsing = isusing
	osummarystock.FRectOnlySellyn = OnlySellyn
	osummarystock.FRectMwDiv      = mwdiv
	osummarystock.FRectOrderBy      = orderby
	osummarystock.FRectitemid = itemid
	osummarystock.FRectSearchType = searchtype
	osummarystock.FRectRackCode = rackcode2
	osummarystock.FRectFromRackcode2 = fromrackcode2
	osummarystock.FRectToRackcode2 = torackcode2
	osummarystock.FRectExcIts = excits

	osummarystock.GetCurrentStockByOnlineBrandLimit

%>

<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function NextPage(v){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품코드는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	document.frm.cp.value=v;
	document.frm.submit();
}

//오차입력
function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=1024,height=768,scrollbar=yes,resizable=yes')
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="cp" value="<%= currPage %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "shopid",shopid %>
		&nbsp;&nbsp;
		<input type="radio" name="diffdiv" value="over" <% if (diffdiv = "over") then %>checked<% end if %>> 한정초과&nbsp;&nbsp;
		<input type="radio" name="diffdiv" value="number" <% if (diffdiv = "number") then %>checked<% end if %>> 수량&nbsp;&nbsp;
      	<input type="radio" name="diffdiv" value="percent" <% if (diffdiv = "percent") then %>checked<% end if %>> 퍼센트&nbsp;&nbsp;
    	범위 : <input type="text" class="text" name="diff" value="<%= diff %>" size="6">
    	&nbsp;&nbsp;
		<% if (diffdiv = "over") then %>
			* 한정비교 보다 현재한정수량이 많은 경우 입니다.
		<% elseif (diffdiv = "number") then %>
			* 한정비교재고와 현재한정수량과의 차이가 <strong><%= diff %>개</strong> 초과되는 상품입니다.
		<% else %>
			* 한정비교재고와 현재한정수량과의 차이가 <strong><%= diff %>%</strong> 초과되는 상품입니다.
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		판매 : <% drawSelectBoxSellYN "OnlySellyn", OnlySellyn %>
     	&nbsp;&nbsp;
     	사용 : <% drawSelectBoxUsingYN "isusing", isusing %>
     	&nbsp;&nbsp;
     	거래구분 : <% drawSelectBoxMWU "mwdiv", mwdiv %>
     	&nbsp;&nbsp;
		정렬 :
		<select class="select" name="orderby">
			<option  value="">실사재고</option> <!-- 초기값 -->
			<option  value="makerid" <%= ChkIIF(orderby="makerid","selected","") %> >브랜드ID</option> <!-- 알파벳순서 -->
			<option  value="itemrackcode" <%= ChkIIF(orderby="itemrackcode","selected","") %> >상품랙코드</option> <!-- 작은숫자부터 -->
			<option  value="itemid" <%= ChkIIF(orderby="itemid","selected","") %> >상품코드</option>
		</select>
		&nbsp;&nbsp;
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10">
		&nbsp;&nbsp;
		랙코드 :
		<input type="radio" name="searchtype" value="F" <% if (searchtype = "F") then %>checked<% end if %> >
		<input type="text" name=rackcode2 value="<%= rackcode2 %>" maxlength="2" size="2" class="text"> (앞 2자리)
		&nbsp;
		<input type="radio" name="searchtype" value="R" <% if (searchtype = "R") then %>checked<% end if %> >
    	<input type="text" name=fromrackcode2 value="<%= fromrackcode2 %>" maxlength="2" size="2" class="text">
		~
		<input type="text" name=torackcode2 value="<%= torackcode2 %>" maxlength="2" size="2" class="text"> (앞 2자리)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		<input type="checkbox" class="checkbox" name="excits" value="Y" <%= CHKIIF(excits="Y", "checked", "") %> > 아이띵소 제외
	</td>
</tr>
</form>
</table>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:1;">
<tr>
	<td align="left"></td>
	<td align="right"></td>
</tr>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= osummarystock.FresultCount %></b>
		&nbsp;
		페이지 : <b><%= currPage %>/ <%= osummarystock.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">브랜드</td>
	<td width="50">이미지</td>
    <td width="40">상품<br>랙코드</td>
	<td width="40">상품<br>코드</td>
	<td width="40">옵션<br>코드</td>
	<td>상품명<br>(옵션명)</td>
	<td width="35">계약<br>구분</td>
    <td width="35">전체<br>입고<br>반품</td>
    <td width="35">전체<br>판매<br>반품</td>
    <td width="35">전체<br>출고<br>반품</td>
    <td width="35">기타<br>출고<br>반품</td>
	<td width="50">총<br>실사<br>오차</td>
	<td width="50">실사<br>재고</td>
	<td width="50">총<br>불량</td>
	<td width="50">유효<br>재고</td>
	<!--<td width="30">총<br>불량</td>
    <td width="30">총<br>실사<br>오차</td>
    <td width="35">실사<br>재고</td>-->
	<td width="30">총<br>상품<br>준비</td>
    <td width="35">재고<br>파악<br>재고</td>
    <td width="30">ON<br>결제<br>완료</td>
    <td width="30">ON<br>주문<br>접수</td>
    <td width="35">한정<br>비교<br>재고</td>
    <td width="35">차이</td>
	<td width="35">한정<br>여부</td>
	<td width="35">판매<br>여부</td>
	<td width="50">오차<br>입력</td>
</tr>
<% if osummarystock.FresultCount > 0 then %>
	<% for i=0 to osummarystock.FresultCount - 1 %>
	<% if osummarystock.FItemList(i).Fisusing="Y" and (osummarystock.FItemList(i).GetLimitStr <> "") then %>
		<tr bgcolor="#FFFFFF" align="center">
	<% elseif IsNull(osummarystock.FItemList(i).GetLimitStr) then  %>
		<tr bgcolor="#FF0000" align="center">
	<% else %>
		<tr bgcolor="#EEEEEE" align="center">
	<% end if %>

		<td><%= osummarystock.FItemList(i).FMakerID %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
        <td><%= osummarystock.FItemList(i).Fitemrackcode %></td>
		<td><a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a></td>
		<td <%=CHKIIF(osummarystock.FItemList(i).Foptioncnt>0 and osummarystock.FItemList(i).FItemOption="0000"," bgcolor='#FF3333'","")%>><%= osummarystock.FItemList(i).FItemOption %></td>
		<td align="left">
	      	<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
	    	<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
	      	<br><font color="blue">[<%= osummarystock.FItemList(i).FItemOptionName %>]</font>
	    	<% end if %>
	    </td>
	    <td><%= fnColor(osummarystock.FItemList(i).Fmwdiv,"mw") %></td>
		<td><%= osummarystock.FItemList(i).Ftotipgono %></td>
		<td><%= -1*osummarystock.FItemList(i).Ftotsellno %></td>
		<td><%= osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono %></td>
	    <td><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></td>
		<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ferrrealcheckno, 0) %></b>&nbsp;</td>
	    <td align="right"><%= FormatNumber(osummarystock.FItemList(i).getErrAssignStock, 0) %>&nbsp;</td>
		<td align="right">
			<%= FormatNumber(osummarystock.FItemList(i).Ferrbaditemno, 0) %>&nbsp;
		</td>
	    <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %>&nbsp;</td>
		<!--<td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
	    <td><%= osummarystock.FItemList(i).Ferrrealcheckno %></td>
	    <td><b><%= osummarystock.FItemList(i).Frealstock %></b></td>-->
		<td><%= osummarystock.FItemList(i).Fipkumdiv5 + osummarystock.FItemList(i).Foffconfirmno %></td>
	    <td><b><%= osummarystock.FItemList(i).GetCheckStockNo %></b></td>
	    <td><%= osummarystock.FItemList(i).Fipkumdiv4 %></td>
	    <td><%= osummarystock.FItemList(i).Fipkumdiv2 %></td>
	    <td><b><%= osummarystock.FItemList(i).GetLimitStockNo %></b></td>
	   		<% if (diffdiv = "number") or (diffdiv = "over") then %>
	    <td><%= osummarystock.FItemList(i).GetLimitStockNo - osummarystock.FItemList(i).GetLimitStr %></td>
		<% elseif (diffdiv = "percent") then %>
	    <td><%= round((100 - (osummarystock.FItemList(i).GetLimitStr * 100 / osummarystock.FItemList(i).GetLimitStockNo)),1) %>%</td>
	    	<% end if %>
	    <td>
	      	한정<br>
	      	(<%= osummarystock.FItemList(i).GetLimitStr %>)

			<% 'if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
				<!--<br>(<% '= osummarystock.FItemList(i).Foptlimitno %>/<% '= osummarystock.FItemList(i).Foptlimitsold %>)-->
			<% 'else %>
				<!--<br>(<% '= osummarystock.FItemList(i).FLimitNo %>/<% '= osummarystock.FItemList(i).FLimitSold %>)-->
			<% 'end if %>
	    </td>
	    <td><%= fnColor(osummarystock.FItemList(i).Fsellyn,"yn") %></td>
		<td>
			<input type="button" class="button" value="오차" onclick="popRealErrInput('<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).Fitemoption %>');">
		</td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="25" align="center">
			<% if osummarystock.HasPreScroll then %>
	    		<a href="javascript:NextPage('<%= osummarystock.StartScrollPage-1 %>')">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + osummarystock.StartScrollPage to osummarystock.FScrollCount + osummarystock.StartScrollPage - 1 %>
	    		<% if i>osummarystock.FTotalpage then Exit for %>
	    		<% if CStr(currPage)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if osummarystock.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="dolimitsoldset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="dispyn" value="">
<input type="hidden" name="sellyn" value="">
</form>

<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
