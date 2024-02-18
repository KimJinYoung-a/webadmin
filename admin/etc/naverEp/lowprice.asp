<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim makerid, itemid, itemname, orderby, priceCompare, getday, regdate
Dim page, olow, i, suplycash, twentyhigh, sorting
Dim dispCate
page    				= requestCheckvar(request("page"),10)
itemid  				= request("itemid")
makerid					= requestCheckvar(request("makerid"),32)
itemname				= requestCheckvar(request("itemname"),100)
priceCompare			= requestCheckvar(request("priceCompare"),100)
regdate					= requestCheckvar(request("regdate"),10)
orderby					= requestCheckvar(request("orderby"),32)
dispCate 				= requestCheckvar(request("disp"),16)
suplycash				= requestCheckvar(request("suplycash"),4)
twentyhigh				= requestCheckvar(request("twentyhigh"),4)
sorting					= requestCheckvar(request("sorting"),16)

research = requestCheckvar(request("research"),10)
if (research="") and (regdate="") then
    regdate=LEFT(dateadd("d",-1,now()),10)
end if

If page = "" Then page = 1
If itemid <> "" then
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

If sorting <> "" Then orderby = ""
If orderby <> "" Then sorting = ""

SET olow = new epShop
	olow.FCurrPage				= page
	olow.FPageSize				= 20
	olow.FRectMakerid			= makerid
	olow.FRectItemid			= itemid
	olow.FRectItemname			= itemname
	olow.FRectPriceCompare		= priceCompare
	olow.FRectRegdate			= regdate
	olow.FRectCDL				= request("cdl")
	olow.FRectCDM				= request("cdm")
	olow.FRectCDS				= request("cds")
	olow.FRectDispCate			= dispCate
	olow.FRectOrderby			= orderby
	olow.FRectSorting			= sorting
	olow.FRectsuplycash			= suplycash
	olow.FRecttwentyhigh		= twentyhigh
    olow.getNaverLowpriceList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function pop_detail(midx, myrank){
	var popNvD = window.open('/admin/etc/naverEp/pop_lowprice_detail.asp?midx='+midx+'&myrank='+myrank,'notinItem','width=500,height=800,scrollbars=yes,resizable=yes');
	popNvD.focus();
}
function pop_cause(){
	var popCau = window.open('/admin/etc/naverEp/pop_cause.asp','cause','width=800,height=400,scrollbars=yes,resizable=yes');
	popCau.focus();
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;&nbsp;
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>&nbsp;&nbsp;
		상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"><br><br>
		관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->
		&nbsp;&nbsp;전시 카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		<br><br>
		매입가비교 : 
		<select name="suplycash" class="select">
			<option value="">-Choice-</option>
			<option value="high" <%= Chkiif(suplycash = "high", "selected", "") %> >매입가 > 최저가</option>
			<option value="low" <%= Chkiif(suplycash = "low", "selected", "") %> >매입가 < 최저가</option>
		</select>
		&nbsp;
		판매가비교 : 
		<select name="twentyhigh" class="select">
			<option value="">-Choice-</option>
			<option value="high" <%= Chkiif(twentyhigh = "high", "selected", "") %> >최저가보다 판매가의 20%이상</option>
			<option value="low" <%= Chkiif(twentyhigh = "low", "selected", "") %> >최저가보다 판매가의 20%미만</option>
		</select>
		<br><br>
		날짜 : <input type="text" name="regdate" id="regdate" size="10" value="<%=regdate%>" onClick="jsPopCal('regdate');" style="cursor:hand;">&nbsp;&nbsp;
		<!--
		네이버 판매가 : 
		<select name="priceCompare" class="select">
			<option value="">-Choice-</option>
			<option value="T" <%= Chkiif(priceCompare = "T", "selected", "") %> >네이버판매가 > 최저가</option>
			<option value="N" <%= Chkiif(priceCompare = "N", "selected", "") %> >네이버판매가 < 최저가</option>
			<option value="S" <%= Chkiif(priceCompare = "S", "selected", "") %> >네이버판매가 = 최저가</option>
		</select>&nbsp;&nbsp;
		-->
		정렬기준 : 
		<select name="orderby" class="select">
			<option value="">-Choice-</option>
			<option value="best" <%= Chkiif(orderby = "best", "selected", "") %> >베스트셀러순</option>
			<option value="wish" <%= Chkiif(orderby = "wish", "selected", "") %> >인기 위시순</option>
			<!-- <option value="myL" <%= Chkiif(orderby = "myL", "selected", "") %> >텐바이텐순위↓</option> -->
		</select>
		<input type="hidden" name="sorting" value="<%=sorting%>">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
</p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		1.네이버쇼핑에 연동되어 있으며, 가격비교페이지 내에 매칭된 상품만 검색 가능합니다.<br>
		2.리스트 중에 <font color="red">상세</font>는 현재 기준으로 일주일간만 확인 가능합니다.<br>
		3.텐바이텐 순위는 100위까지만 확인 가능하며, 그 이하일 경우 ‘최하위’로 표기됩니다.<br>
		4.판매갯수와 위시 찜갯수는 해당 상품이 텐바이텐에 최초 등록된 날로 부터 현재까지 누적갯수입니다.<br>
		5.당일 리스트는 오전 11시 10분경에 업데이트 됩니다.<br>
		6.<input type="button" class="button" value="가격비교 서비스 불가항목" onclick="javascrtip:pop_cause();">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(olow.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(olow.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">텐바이텐<br>상품번호</td>
	<td>브랜드<br>상품명</td>
	<td>상세</td>
	<!-- <td width="100">네이버<br>판매가</td> -->
	<td width="100" onClick="jstrSort('sellcash'); return false;" style="cursor:pointer;">
		텐바이텐<br>판매가
		<img src="/images/list_lineup<%=CHKIIF(sorting="sellcashD","_bot","_top")%><%=CHKIIF(instr(sorting,"sellcash")>0,"_on","")%>.png" id="imgsellcash">
	</td>
	<td width="100" onClick="jstrSort('samecashCnt'); return false;" style="cursor:pointer;">
		동일가<br>판매몰
		<img src="/images/list_lineup<%=CHKIIF(sorting="samecashCntD","_bot","_top")%><%=CHKIIF(instr(sorting,"samecashCnt")>0,"_on","")%>.png" id="imgsamecashCnt">
	</td>
	<td width="100" onClick="jstrSort('lowcash'); return false;" style="cursor:pointer;">
		최저가
		<img src="/images/list_lineup<%=CHKIIF(sorting="lowcashD","_bot","_top")%><%=CHKIIF(instr(sorting,"lowcash")>0,"_on","")%>.png" id="imglowcash">
	</td>
	<td width="100">Rank2판매가</td>
	<td width="100">Rank3판매가</td>
	<td width="100" onClick="jstrSort('myrank'); return false;" style="cursor:pointer;">
		텐바이텐<br>순위
		<img src="/images/list_lineup<%=CHKIIF(sorting="myrankD","_bot","_top")%><%=CHKIIF(instr(sorting,"myrank")>0,"_on","")%>.png" id="imgmyrank">
	</td>
	<td width="100" onClick="jstrSort('sellcount'); return false;" style="cursor:pointer;">
		판매갯수
		<img src="/images/list_lineup<%=CHKIIF(sorting="sellcountD","_bot","_top")%><%=CHKIIF(instr(sorting,"sellcount")>0,"_on","")%>.png" id="imgsellcount">
	</td>
	<td width="100" onClick="jstrSort('favcount'); return false;" style="cursor:pointer;">
		위시 찜갯수
		<img src="/images/list_lineup<%=CHKIIF(sorting="favcountD","_bot","_top")%><%=CHKIIF(instr(sorting,"favcount")>0,"_on","")%>.png" id="imgfavcount">
	</td>
	<td width="100" onClick="jstrSort('buycash'); return false;" style="cursor:pointer;">
		매입가
		<img src="/images/list_lineup<%=CHKIIF(sorting="buycashD","_bot","_top")%><%=CHKIIF(instr(sorting,"buycash")>0,"_on","")%>.png" id="imgbuycash">
	</td>
	<td width="100" onClick="jstrSort('margin'); return false;" style="cursor:pointer;">
		마진율
		<img src="/images/list_lineup<%=CHKIIF(sorting="marginD","_bot","_top")%><%=CHKIIF(instr(sorting,"margin")>0,"_on","")%>.png" id="imgmargin">
	</td>
	<!--
	<td width="100">최하위Rank</td>
	<td width="100">최상위Rank</td>
	<td width="100">최고가</td>
	-->
	<td width="100">업데이트<br>날짜</td>
</tr>
<% For i = 0 To olow.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center">
		<a href="<%=wwwURL%>/<%=olow.FItemList(i).FItemID%>" target="_blank"><%= olow.FItemList(i).FItemID %></a><br>
	</td>
	<td align="left"><%= olow.FItemList(i).FMakerid %><br><%= olow.FItemList(i).FItemName %></td>
	<td align="center">
		<% If datediff("d", olow.FItemList(i).FRegdate, now()) < 7 Then  %>
			<input type="button" class="button" value="확인" onclick="javascrtip:pop_detail('<%= olow.FItemList(i).FIdx %>', '<%=olow.FItemList(i).FMyrank%>')">
		<% Else %>
			열람불가
		<% End If %>
	</td>
	<!-- <td align="center"><%= FormatNumber(olow.FItemList(i).FNaverSellCash,0) %></td> -->
	<td align="center"><%= FormatNumber(olow.FItemList(i).FSellcash,0) %></td>
	<td align="center"><%= olow.FItemList(i).FSamecashCnt %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FLowcash,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FRank2Price,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FRank3Price,0) %></td>
	<td align="center">
	<%
		If olow.FItemList(i).FMyrank = "1000" Then
			response.write "<font color='RED'>최하위</font>"
		Else
			response.write olow.FItemList(i).FMyrank
		End If
	%>
	</td>
	
	<td align="center"><%= FormatNumber(olow.FItemList(i).FSellcount,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FFavcount,0) %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).Fbuycash,0) %></td>
	<td align="center">
	<%
		If olow.FItemList(i).Fsellcash<>0 Then
			response.write CLng(10000-olow.FItemList(i).Fbuycash/olow.FItemList(i).Fsellcash*100*100)/100 & "%"
		End If
	%>
	</td>
	
	<!--
	<td align="center"><%= olow.FItemList(i).FMallmaxrank %></td>
	<td align="center"><%= olow.FItemList(i).FMalllowrank %></td>
	<td align="center"><%= FormatNumber(olow.FItemList(i).FHighcash,0) %></td>
	-->
	<td align="center"><%= LEFT(olow.FItemList(i).FRegdate,10) %></td>
</tr>
<% Next %>
<tr height="20">
    <td colspan="19" align="center" bgcolor="#FFFFFF">
        <% if olow.HasPreScroll then %>
		<a href="javascript:goPage('<%= olow.StartScrollPage-1 %>');">[pre]</a>
    	<% else %>
    		[pre]
    	<% end if %>

    	<% for i=0 + olow.StartScrollPage to olow.FScrollCount + olow.StartScrollPage - 1 %>
    		<% if i>olow.FTotalpage then Exit for %>
    		<% if CStr(page)=CStr(i) then %>
    		<font color="red">[<%= i %>]</font>
    		<% else %>
    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
    		<% end if %>
    	<% next %>

    	<% if olow.HasNextScroll then %>
    		<a href="javascript:goPage('<%= i %>');">[next]</a>
    	<% else %>
    		[next]
    	<% end if %>
    </td>
</tr>
</table>
<% SET olow = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->