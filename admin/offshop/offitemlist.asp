<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  [OFF]오프_상품관리>>신상품관리
' History : 2008.04.17 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer, page, itemid, locationid, datesearch, sdate, edate, itemname, IsOnlineItem, imageList, offmain
Dim sort, i, offlist, offsmall, iTotCnt, itemgubun, isusing, vParam, inc3pl
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2, datefg, fromDate, toDate, cdl, cdm, cds
	designer = requestCheckVar(request("designer"),32)
	page = requestCheckVar(request("page"),10)
	itemid = requestCheckVar(request("itemid"),10)
	datesearch = requestCheckVar(request("datesearch"),10)
	sdate = requestCheckVar(request("sdate"),10)
	edate = requestCheckVar(request("edate"),10)
	itemname = requestCheckVar(request("itemname"),124)
	itemgubun = requestCheckVar(request("itemgubun"),2)
	isusing = requestCheckVar(request("isusing"),1)
	sort = requestCheckVar(request("sort"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	cdl     = requestCheckVar(request("cdl"),3)
	cdm     = requestCheckVar(request("cdm"),3)
	cds     = requestCheckVar(request("cds"),3)
    inc3pl = requestCheckVar(request("inc3pl"),32)
    
If page = "" Then page = 1
If sort = "" Then sort = "itemregdate"

if datefg = "" then datefg = "maechul"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장일경우 본인 매장만 사용가능
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	'if getlevel_sn("",session("ssBctId")) > 6 then
		locationid = C_STREETSHOPID
	'end if

else
	if (C_IS_Maker_Upche) then
		locationid = session("ssBctID")
	else
		locationid = request("locationid")
	end if
end if

vParam = "&locationid="&locationid&"&designer="&designer&"&itemid="&itemid&"&datesearch="&datesearch&"&sdate="&sdate&"&edate="&edate&"&itemgubun="&itemgubun&"&isusing="&isusing&"&itemname="&itemname&"&sort="&sort&"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&datefg="&datefg&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&inc3pl="&inc3pl

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 50
	ioffitem.FCurrPage = page
	ioffitem.FRectShopid = locationid
	ioffitem.FRectDesigner = designer
	ioffitem.FRectDateSearch = datesearch
	ioffitem.FRectSDate = sdate
	ioffitem.FRectEDate = edate
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemName = itemname
	ioffitem.FRectItemId = itemid
	ioffitem.FRectIsusing = isusing
	ioffitem.FRectSorting = sort
	ioffitem.frectdatefg = datefg	
	ioffitem.FRectStartDay = fromDate
	ioffitem.FRectEndDay = toDate
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectInc3pl = inc3pl
	
	If locationid <> "" Then
		ioffitem.GetOffLineNewItemList
	End If

iTotCnt = ioffitem.FTotalCount
%>

<script language='javascript'>

function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function goSort(a){
	document.frm.sort.value = document.getElementById("tmpsort").value;
	document.frm.submit();
}

function goExcelDown(){
	var ExcelDown = window.open('offitemlist_xls.asp?1=1<%=vParam%>','ExcelDown','width=600,height=400,scrollbars=yes,resizable=yes');
	ExcelDown.focus();
}

function pop_ipgomaechul(shopid, extbarcode, yyyy1, mm1, dd1, yyyy2, mm2, dd2){
	var pop_ipgomaechul = window.open('/admin/offshop/dayitemsellsum.asp?shopid='+shopid+'&extbarcode='+extbarcode+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2,'pop_ipgomaechul','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_ipgomaechul.focus();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value=1>
<input type="hidden" name="sort" value="<%=sort%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(locationid) <> "1" and locationid <> "" then %>
				* ShopID : <%=locationid%><input type="hidden" name="locationid" value="<%= locationid %>">
			<% else %>
				* ShopID : <% drawSelectBoxOffShopNotUsingAll "locationid",locationid %>
			<% end if %>
		<% else %>
			* ShopID : <% drawSelectBoxOffShopNotUsingAll "locationid",locationid %>
		<% end if %>
		&nbsp;&nbsp;
		* 기간 : 
		<select name="datesearch" class="select">
			<option value="" <%=CHKIIF(datesearch="","selected","")%>>-선택-</option>
			<option value="itemregdate" <%=CHKIIF(datesearch="itemregdate","selected","")%>>상품등록일</option>
			<option value="ipgodate" <%=CHKIIF(datesearch="ipgodate","selected","")%>>브랜드최초입고일</option>
			
			<% if locationid <> "" then %>
				<option value="stockipgodate" <%=CHKIIF(datesearch="stockipgodate","selected","")%>>상품최초입고일</option>
			<% end if %>
		</select>
		<input type="text" name="sdate" size="10" maxlength=10 value="<%=sdate%>">
		<a href="javascript:calendarOpen(frm.sdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;~&nbsp;
		<input type="text" name="edate" size="10" maxlength=10 value="<%=edate%>">
		<a href="javascript:calendarOpen(frm.edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
		&nbsp;&nbsp;
		* 사용여부 : 
		<input type="radio" name="isusing" value="Y" <%=CHKIIF(isusing="Y","checked","")%>>Y&nbsp;
		<input type="radio" name="isusing" value="N" <%=CHKIIF(isusing="N","checked","")%>>N
	</td>
	<td rowspan="4" class="a" align="center" valign="middle">
		<input type="button" onClick="frmsubmit('');" value="검색" class="button">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* 브랜드 : <% drawSelectBoxDesignerwithName "designer",designer %>
		&nbsp;&nbsp;
		* 상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size="8" maxlength="8">
		&nbsp;&nbsp;
		* 상품명 : <input type="text" name="itemname" value="<%= itemname %>" size="30">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* 상품구분 : 
		<input type="checkbox" name="itemgubun" value="10" <%=CHKIIF(InStr(itemgubun,"10")>0,"checked","")%>>온라인상품(10)&nbsp;
		<input type="checkbox" name="itemgubun" value="90" <%=CHKIIF(InStr(itemgubun,"90")>0,"checked","")%>>오프샵 전용상품(90)&nbsp;
		<input type="checkbox" name="itemgubun" value="70" <%=CHKIIF(InStr(itemgubun,"70")>0,"checked","")%>>소모품(70)&nbsp;
		<input type="checkbox" name="itemgubun" value="80" <%=CHKIIF(InStr(itemgubun,"80")>0,"checked","")%>>사은품(80)&nbsp;
		<input type="checkbox" name="itemgubun" value="60" <%=CHKIIF(InStr(itemgubun,"60")>0,"checked","")%>>할인권(60)&nbsp;
		<input type="checkbox" name="itemgubun" value="00" <%=CHKIIF(InStr(itemgubun,"00")>0,"checked","")%>>매장매입상품(00)&nbsp;
		<input type="checkbox" name="itemgubun" value="95" <%=CHKIIF(InStr(itemgubun,"95")>0,"checked","")%>>가맹점개별매입판매(95)
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
        <b>* 매출처구분</b>
        <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
        &nbsp;&nbsp;        
		* <!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</table>
<br>
<% If locationid = "" Then %>
	<center><font color="red"><b>※ ShopID(매장)를 선택하셔야 데이터가 나타납니다.</b></font></center><br>
<% End If %>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 검색된 결과중 상단부터 최대 5천건까지만 받아집니다.<br>
		<input type="button" onClick="goExcelDown();" value="엑셀다운" class="button">
	</td>
	<td align="right">	
		<% drawmaechuldatefg "datefg" ,datefg ,""%>
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>검색결과 : <b><%= FormatNumber(iTotCnt,0) %></b></td>
			<td align="right" valign="bottom">
				정렬 :
				<select name="tmpsort" id="tmpsort" class="select" style="margin-bottom:3px;" onChange="goSort(this.value);">
					<option value="itemregdate" <%=CHKIIF(sort="itemregdate","selected","")%>>상품등록일</option>
					<option value="ipgodate" <%=CHKIIF(sort="ipgodate","selected","")%>>입고일순</option>
				</select>			
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>ITEMID</td>
	<td>IMAGE</td>
	<td>BRANDID</td>
	<td>상품명[옵션명]</td>
	<td>사용<br>여부</td>
	<td>상품등록일</td>
	<td>최종업데이트일</td>
	<td>브랜드<Br>최초입고일</td>
	
	<% if locationid <> "" then %>
		<td>상품<Br>최초입고일</td>
	<% end if %>

	<td>
		매출액
	</td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>
			매입가
		</td>
	<% end if %>
	
	<td>판매<br>수량</td>
	
	<td width="60">비고</td>
</tr>
<%
If ioffitem.FResultCount > 0 Then

For i=0 To ioffitem.FResultCount -1
%>
<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" align="center">
	<td width=80><%=ioffitem.FItemList(i).Fitemgubun%><%=ioffitem.FItemList(i).Fshopitemid%><%=ioffitem.FItemList(i).Fitemoption%></td>	
	<td width=50>
		<%
			If ioffitem.FItemList(i).Fitemgubun = "10" Then
				Response.Write "<img src='" & ioffitem.FItemList(i).FimageSmall & "' width='50' height='50'>"
			Else
   				If ioffitem.FItemList(i).FOffimgSmall <> "" Then
   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgSmall & "' width='50' height='50'>"
   				ElseIf ioffitem.FItemList(i).FOffimgSmall = "" Then
	   				If ioffitem.FItemList(i).FOffimgMain <> "" Then
	   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgMain & "' width='50' height='50'>"
	   				ElseIf ioffitem.FItemList(i).FOffimgMain = "" Then
		   				If ioffitem.FItemList(i).FOffimgList <> "" Then
		   					Response.Write "<img src='" & ioffitem.FItemList(i).FOffimgList & "' width='50' height='50'>"
		   				End If
		   			End If
   				End If
	   		End If
		%></td>
	<td><%=ioffitem.FItemList(i).Fmakerid%></td>
	<td align="left">
		<%=ioffitem.FItemList(i).Fshopitemname%>
		
		<% if ioffitem.FItemList(i).Fshopitemoptionname <> "" then %>
		[<%= ioffitem.FItemList(i).Fshopitemoptionname %>]
		<% end if %>
	</td>
	<td width=30><%=ioffitem.FItemList(i).Fisusing%></td>
	<td width=140><%=ioffitem.FItemList(i).Fregdate%></td>
	<td width=140><%=ioffitem.FItemList(i).Fupdt%></td>
	<td width=80><%=ioffitem.FItemList(i).Ffirstipgodate%></td>

	<% if locationid <> "" then %>
		<td width=140>
			<%=ioffitem.FItemList(i).fstockregdate%>
			
			<% if ioffitem.FItemList(i).fstockregdate<>"" then %>
				<!--<Br><a href="javascript:pop_ipgomaechul('<%= locationid %>','<%=ioffitem.FItemList(i).Fitemgubun & Format00(6,ioffitem.FItemList(i).Fshopitemid) & ioffitem.FItemList(i).Fitemoption%>','<%= left(left(ioffitem.FItemList(i).fstockregdate,10),4) %>','<%= mid(left(ioffitem.FItemList(i).fstockregdate,10),6,2) %>','<%= right(left(ioffitem.FItemList(i).fstockregdate,10),2) %>','','','');" onfocus="this.blur()">
				날짜별상세매출보기</a>-->
			<% end if %>
		</td>
	<% end if %>

	<td align="right" bgcolor="#E6B9B8" width=80><%= FormatNumber(ioffitem.FItemList(i).fsellsum,0) %></td>
	
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right" width=80><%= FormatNumber(ioffitem.FItemList(i).fsuplyprice,0) %></td>
	<% end if %>
	
	<td align="right" width=60><%= FormatNumber(ioffitem.FItemList(i).fitemcnt,0) %></td>
	<td width=60>
		<input type="button" onclick="popShopCurrentStock('<%=ioffitem.FItemList(i).FShopID%>','<%=ioffitem.FItemList(i).Fitemgubun%>','<%=ioffitem.FItemList(i).Fshopitemid%>','<%=ioffitem.FItemList(i).Fitemoption%>');" value="재고" class="button">
	</td>
</tr>
<%
Next
%>

<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:frmsubmit('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
		<% if i>ioffitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:frmsubmit('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:frmsubmit('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<%
Else
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="center" colspan="20">검색된 상품이 없습니다.</td>
</tr>
<% End If %>
</table>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->