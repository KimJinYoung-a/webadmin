<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품별주문리스트
' History : 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, i, baljucode, itemid, makerid, mwdiv
dim purchasetype, blinkcode, datetype, statecd, tplgubun, productidx
dim sumRealItemnoSellcash, sumRealItemnoBuycash, sumBaljuItemno, sumRealItemno, sumCheckItemno
    page = RequestCheckVar(getNumeric(trim(request("page"))),10)
    research = RequestCheckVar(trim(request("research")),2)
    yyyy1 = RequestCheckVar(trim(request("yyyy1")),4)
    mm1   = RequestCheckVar(trim(request("mm1")),2)
    dd1   = RequestCheckVar(trim(request("dd1")),2)
    yyyy2 = RequestCheckVar(trim(request("yyyy2")),4)
    mm2   = RequestCheckVar(trim(request("mm2")),2)
    dd2   = RequestCheckVar(trim(request("dd2")),2)
    baljucode   = RequestCheckVar(trim(request("baljucode")),32)
    itemid      = requestCheckvar(trim(request("itemid")),1500)
    makerid   = RequestCheckVar(trim(request("makerid")),32)
    mwdiv       = requestCheckvar(trim(request("mwdiv")),10)
    purchasetype = RequestCheckVar(getNumeric(trim(request("purchasetype"))),10)
    blinkcode   = RequestCheckVar(trim(request("blinkcode")),32)
    datetype   = RequestCheckVar(trim(request("datetype")),32)
    statecd   = RequestCheckVar(trim(request("statecd")),10)
    tplgubun   = RequestCheckVar(trim(request("tplgubun")),32)
    productidx = RequestCheckVar(getNumeric(trim(request("productidx"))),10)

if datetype="" or isnull(datetype) then datetype="regdate"
if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-7,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-7,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-7,date())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

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

dim oItemOrder
set oItemOrder = new COrderSheet
oItemOrder.FCurrPage = page
oItemOrder.FPageSize = 100
oItemOrder.FRectStartDate = fromDate
oItemOrder.FRectEndDate   = toDate
oItemOrder.FRectbaljucode   = baljucode
oItemOrder.FRectblinkcode   = blinkcode
oItemOrder.FRectItemid       = itemid
oItemOrder.FRectmakerid       = makerid
oItemOrder.FRectmwdiv       = mwdiv
oItemOrder.FRectBrandPurchaseType = purchasetype
oItemOrder.FRectdatetype = datetype
oItemOrder.FRectstatecd = statecd
oItemOrder.FRecttplgubun = tplgubun
oItemOrder.FRectproductidx = productidx
oItemOrder.GetItemOrderList
%>
<script type='text/javascript'>

function NextPage(page){
	document.frm.target = "";
	document.frm.action = "";
    document.frm.page.value=page;
    document.frm.submit();
}

function jsItemStock(itemgubun, itemid,itemoption){
	var jsItemStock = window.open("/admin/stock/itemcurrentstock.asp?itemgubun="+itemgubun+"&itemid="+itemid+"&itemoption="+itemoption+"&menupos=709","jsItemStock","width=1400 height=800 scrollbars=yes resizable=yes");
	jsItemStock.focus();
}

function getOrderList(baljucode){
	var jsOrderList = window.open("/admin/newstorage/orderlist.asp?baljucode="+baljucode+"&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&dd1=<%=dd1%>&yyyy2=<%=yyyy2%>&mm2=<%=mm2%>&dd2=<%=dd2%>","jsOrderList","width=1400 height=800 scrollbars=yes resizable=yes");
	jsOrderList.focus();
}

function downloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/newstorage/itemOrderList_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
		* 주문코드 :
		<input type="text" class="text" name="baljucode" value="<%= baljucode %>" maxlength=32 size=10>
        &nbsp;
		* 브랜드ID :
		<% drawSelectBoxDesignerwithName "makerid",makerid  %>
        &nbsp;
		* 상품코드 :
        <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
        &nbsp;
		* 관련입고코드 :
		<input type="text" class="text" name="blinkcode" value="<%= blinkcode %>" maxlength=32 size=10>
        &nbsp;
		* 원가IDX :
		<input type="text" class="text" name="productidx" value="<%= productidx %>" maxlength=32 size=10>
    </td>
    <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="NextPage('1');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        * <select class="formSlt" name="datetype" id="datetype">
            <option value="regdate" <%= chkIIF(datetype="regdate","selected","") %> >주문일
            <option value="scheduledate" <%= chkIIF(datetype="scheduledate","selected","") %> >입고요청일
        </select>
        <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
        * 주문상태 :
		<select class="select" name="statecd" >
			<option value="">전체
			<option value="0" <% if statecd="0" then response.write "selected" %> >주문접수
			<option value="1" <% if statecd="1" then response.write "selected" %> >주문확인
			<option value="2" <% if statecd="2" then response.write "selected" %> >입금대기
			<option value="5" <% if statecd="5" then response.write "selected" %> >배송준비
			<option value="7" <% if statecd="7" then response.write "selected" %> >출고완료
			<option value="8" <% if statecd="8" then response.write "selected" %> >검품완료
			<option value="9" <% if statecd="9" then response.write "selected" %> >입고완료
			<option value="preorder" <% if statecd="preorder" then response.write "selected" %> >입고이전(기주문)
     	</select>
        &nbsp;
        * 매입구분 : <% drawSelectBoxMWU "mwdiv", mwdiv %>
        &nbsp;
        * 구매유형 :
        <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
        &nbsp;
        * 3PL구분 : 
        <% Call drawSelectBoxTPLGubun("tplgubun", tplgubun) %>
    </td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
    </td>
    <td align="right">
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
    </td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="27">
		검색결과 : <b><%= oItemOrder.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oItemOrder.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="80">주문코드</td>
    <td width="60">원가IDX</td>
    <td width="60">품의번호</td>
    <td width="70">주문일</td>
    <td width="70">입고요청일</td>
    <td width="70">주문상태</td>
    <td width="70">관련입고코드</td>
    <td width="80">브랜드ID</td>
    <td width="60">상품구분</td>
    <td width="60">상품코드</td>
    <td width="60">옵션코드</td>
    <td width="100">바코드</td>
    <td width="100">범용바코드</td>
    <td width="80">업체관리코드</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td width="60">소비자가</td>
    <td width="60">확정총소비자가</td>
    <td width="60">내역서매입가</td>
    <td width="60">확정총매입가</td>
    <td width="60">매입구분</td>
    <td width="60">주문수량</td>
    <td width="60">확정수량</td>
    <td width="60">검품수량</td>
    <td width="60">구매유형</td>
    <td>카테고리</td>
    <td width="60">최종입고월(물류)</td>
</tr>
<% if oItemOrder.FResultCount>0 then %>
<%
sumRealItemnoSellcash=0
sumRealItemnoBuycash=0
sumBaljuItemno=0
sumRealItemno=0
sumCheckItemno=0
for i=0 to oItemOrder.FResultCount-1
sumRealItemnoSellcash = sumRealItemnoSellcash + (oItemOrder.FItemList(i).fsellcash*oItemOrder.FItemList(i).frealitemno)
sumRealItemnoBuycash = sumRealItemnoBuycash + (oItemOrder.FItemList(i).fbuycash*oItemOrder.FItemList(i).frealitemno)
sumBaljuItemno = sumBaljuItemno + oItemOrder.FItemList(i).fbaljuitemno
sumRealItemno = sumRealItemno + oItemOrder.FItemList(i).frealitemno
sumCheckItemno = sumCheckItemno + oItemOrder.FItemList(i).fcheckitemno
%>
<tr bgcolor="#FFFFFF" align="center">
    <td>
        <a href="#" onclick="getOrderList('<%= oItemOrder.FItemList(i).fbaljucode %>'); return false;">        
        <%= oItemOrder.FItemList(i).fbaljucode %></a>
    </td>
    <td><%= oItemOrder.FItemList(i).fproductidxArr %></td>
    <td><%= oItemOrder.FItemList(i).freportidx %></td>
    <td><%= left(oItemOrder.FItemList(i).fregdate,10) %></td>
    <td><%= oItemOrder.FItemList(i).fscheduledate %></td>
    <td><%= oItemOrder.FItemList(i).fstatecdname %></td>
    <td><%= oItemOrder.FItemList(i).fblinkcode %></td>
    <td><%= oItemOrder.FItemList(i).fmakerid %></td>
    <td><%= oItemOrder.FItemList(i).fitemgubun %></td>
    <td>
        <a href="#" onclick="jsItemStock('<%= oItemOrder.FItemList(i).FItemgubun %>','<%= oItemOrder.FItemList(i).FItemID %>','<%= oItemOrder.FItemList(i).FItemOption %>'); return false;">
        <%= oItemOrder.FItemList(i).fitemid %></a>
    </td>
    <td><%= oItemOrder.FItemList(i).fitemoption %></td>
    <td><%= oItemOrder.FItemList(i).ftenbarcode %></td>
    <td><%= oItemOrder.FItemList(i).fbarcode %></td>
    <td><%= oItemOrder.FItemList(i).fupchemanagecode %></td>
    <td align="left"><%= oItemOrder.FItemList(i).fitemname %></td>
    <td align="left"><%= oItemOrder.FItemList(i).fitemoptionname %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fsellcash,0) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fsellcash*oItemOrder.FItemList(i).frealitemno,0) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fbuycash,0) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fbuycash*oItemOrder.FItemList(i).frealitemno,0) %></td>
    <td><%= mwdivName(oItemOrder.FItemList(i).fmwdiv) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fbaljuitemno,0) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).frealitemno,0) %></td>
    <td align="right"><%= FormatNumber(oItemOrder.FItemList(i).fcheckitemno,0) %></td>
    <td><%= oItemOrder.FItemList(i).fpurchaseTypename %></td>
    <td><%= oItemOrder.FItemList(i).fcateName %></td>
    <td><%= oItemOrder.FItemList(i).flastIpgoDate %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
    <td colspan=17>합계</td>
    <td align="right"><%= FormatNumber(sumRealItemnoSellcash,0) %></td>
    <td></td>
    <td align="right"><%= FormatNumber(sumRealItemnoBuycash,0) %></td>
    <td></td>
    <td align="right"><%= FormatNumber(sumBaljuItemno,0) %></td>
    <td align="right"><%= FormatNumber(sumRealItemno,0) %></td>
    <td align="right"><%= FormatNumber(sumCheckItemno,0) %></td>
    <td colspan=3></td>
</tr>
<tr bgcolor="FFFFFF">
	<td colspan="27" align="center">
		<% if oItemOrder.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oItemOrder.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oItemOrder.StartScrollPage to oItemOrder.FScrollCount + oItemOrder.StartScrollPage - 1 %>
			<% if i>oItemOrder.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oItemOrder.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="28" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oItemOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->