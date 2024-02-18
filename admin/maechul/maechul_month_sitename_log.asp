<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매출로그
' Hieditor : 2013.11.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
dim research
Dim i, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate ,toDate ,oCMaechulLog, page, targetGbn, makerid
Dim Adddategbn, AdddategbnYn, addyyyy1,addmm1,adddd1,addyyyy2,addmm2,adddd2, addfromDate ,addtoDate

dim searchfield, searchtext, dategbn, actDivCode, vatinclude, mwdiv_beasongdiv
dim excTPL

	research = requestCheckvar(request("research"),10)
	actDivCode = requestCheckvar(request("actDivCode"),10)
	page = request("page")
	dategbn     = requestCheckvar(request("dategbn"),10)
	Adddategbn  = requestCheckvar(request("Adddategbn"),10)
	AdddategbnYn= requestCheckvar(request("AdddategbnYn"),10)

	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)
	dd1     = requestcheckvar(request("dd1"),10)
	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	dd2     = requestcheckvar(request("dd2"),10)

	addyyyy1   = requestcheckvar(request("addyyyy1"),10)
	addmm1     = requestcheckvar(request("addmm1"),10)
	adddd1     = requestcheckvar(request("adddd1"),10)
	addyyyy2   = requestcheckvar(request("addyyyy2"),10)
	addmm2     = requestcheckvar(request("addmm2"),10)
	adddd2     = requestcheckvar(request("adddd2"),10)

	targetGbn     = requestcheckvar(request("targetGbn"),16)
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	vatinclude     = requestcheckvar(request("vatinclude"),1)
	mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),2)
	makerid   = requestcheckvar(request("makerid"),32)

	excTPL 	= request("excTPL")

if dategbn="" then dategbn="ActDate"
if adddategbn="" then adddategbn="ActDate"
if page = "" then page = 1
if (research = "") then
	excTPL = "Y"
end if



dim tmpDate
if (yyyy1="") then

	fromDate = Left(dateAdd("m",-1,now()),7)+"-01"
	toDate = Left(dateAdd("d",1,fromDate),10) ''Left(dateAdd("m",1,fromDate),10)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

if (addyyyy1="") then

	addfromDate = Left(dateAdd("m",-1,now()),7)+"-01"
	addtoDate = Left(dateAdd("d",1,addfromDate),10) ''Left(dateAdd("m",1,addfromDate),10)

	addyyyy1 = Cstr(Year(addfromDate))
	addmm1 = Cstr(Month(addfromDate))
	adddd1 = Cstr(day(addfromDate))

	tmpDate = DateAdd("d", -1, addtoDate)
	addyyyy2 = Cstr(Year(tmpDate))
	addmm2 = Cstr(Month(tmpDate))
	adddd2 = Cstr(day(tmpDate))
else
	addfromDate = DateSerial(addyyyy1, addmm1, adddd1)
	addtoDate = DateSerial(addyyyy2, addmm2, adddd2+1)
end if

set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 50
	oCMaechulLog.FCurrPage = page
	oCMaechulLog.FRectDategbn = dategbn
	oCMaechulLog.FRectStartDate = fromDate
	oCMaechulLog.FRectEndDate = toDate
	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
	oCMaechulLog.FRecttargetGbn = targetGbn
	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectmakerid = makerid
	if (AdddategbnYn="Y") then
	    oCMaechulLog.FRectAddDategbn    = Adddategbn
    	oCMaechulLog.FRectAddStartDate  = AddfromDate
    	oCMaechulLog.FRectAddEndDate    = AddtoDate
	end if

	oCMaechulLog.FRectExcTPL = excTPL

	oCMaechulLog.GetMaechul_month_sitename_Log

%>

<script language="javascript">

function searchSubmit(page){
	frm.page.value=page;
	frm.submit();
}

function pop_detail_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, sitename){
	<% if dategbn="ActDate" then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?actDate_yyyy1='+yyyy1+'&actDate_mm1='+mm1+'&actDate_dd1='+dd1+'&actDate_yyyy2='+yyyy2+'&actDate_mm2='+mm2+'&actDate_dd2='+dd2+'&chkActDate=Y&searchfield=sitename&searchtext='+sitename+'&targetGbn=<%= targetGbn %>&actDivCode=<%=actDivCode%>&vatinclude=<%=vatinclude%>&mwdiv_beasongdiv=<%=mwdiv_beasongdiv%>&makerid=<%=makerid%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% elseif (dategbn="chulgoDate") then %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?chulgoDate_yyyy1='+yyyy1+'&chulgoDate_mm1='+mm1+'&chulgoDate_dd1='+dd1+'&chulgoDate_yyyy2='+yyyy2+'&chulgoDate_mm2='+mm2+'&chulgoDate_dd2='+dd2+'&chkChulgoDate=Y&searchfield=sitename&searchtext='+sitename+'&targetGbn=<%= targetGbn %>&actDivCode=<%=actDivCode%>&vatinclude=<%=vatinclude%>&mwdiv_beasongdiv=<%=mwdiv_beasongdiv%>&makerid=<%=makerid%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% else %>
		var pop_detail_list = window.open('/admin/maechul/maechul_detail_log.asp?orgPay_yyyy1='+yyyy1+'&orgPay_mm1='+mm1+'&orgPay_dd1='+dd1+'&orgPay_yyyy2='+yyyy2+'&orgPay_mm2='+mm2+'&orgPay_dd2='+dd2+'&chkOrgPay=Y&searchfield=sitename&searchtext='+sitename+'&targetGbn=<%= targetGbn %>&actDivCode=<%=actDivCode%>&vatinclude=<%=vatinclude%>&mwdiv_beasongdiv=<%=mwdiv_beasongdiv%>&makerid=<%=makerid%>&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	<% end if %>

	pop_detail_list.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 날짜 :
				<select name="dategbn">
					<option value="ipkumdate" <%=CHKIIF(dategbn="ipkumdate","selected","")%> >원결제일자
					<option value="ActDate" <%=CHKIIF(dategbn="ActDate","selected","")%> >결제일자(처리일자)
					<option value="chulgoDate" <%=CHKIIF(dategbn="chulgoDate","selected","")%> >출고일자
				</select>
				<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
				&nbsp;
				<input type="checkbox" name="excTPL" value="Y" <% if (excTPL = "Y") then %>checked<% end if %> >
				3PL 매출 제외
<!--
				&nbsp;
                * <input type="checkbox" name="AdddategbnYn" value="Y" <% if (AdddategbnYn = "Y") then %>checked<% end if %> > 추가날짜검색

                <select name="Adddategbn">
					<option value="ipkumdate" <%=CHKIIF(Adddategbn="ipkumdate","selected","")%> >원결제일자
					<option value="ActDate" <%=CHKIIF(Adddategbn="ActDate","selected","")%> >결제일자(처리일자)
					<option value="chulgoDate" <%=CHKIIF(Adddategbn="chulgoDate","selected","")%> >출고일자
				</select>
				<% DrawDateBoxdynamic addyyyy1, "addyyyy1", addyyyy2, "addyyyy2", addmm1, "addmm1", addmm2, "addmm2", adddd1, "adddd1", adddd2, "adddd2" %>
-->
				<p>
				* 매출구분 : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
				&nbsp;&nbsp;
				* 과세구분 : <% drawSelectBoxVatYN "vatinclude", vatinclude %>
				&nbsp;&nbsp;
				* 매입구분 : <% drawmwdiv_beasongdiv "mwdiv_beasongdiv", mwdiv_beasongdiv , "" %>
				&nbsp;&nbsp;
				* 주문구분 :
				<select class="select" name="actDivCode">
					<option value=""></option>
					<option value="A" <% if (actDivCode = "A") then %>selected<% end if %> >원주문</option>
					<option value="C" <% if (actDivCode = "C") then %>selected<% end if %> >취소주문</option>
					<option value="H" <% if (actDivCode = "H") then %>selected<% end if %> >상품변경</option>
					<option value="E" <% if (actDivCode = "E") then %>selected<% end if %> >교환주문</option>
					<option value="M" <% if (actDivCode = "M") then %>selected<% end if %> >반품주문</option>
					<option value="CC" <% if (actDivCode = "CC") then %>selected<% end if %> >취소정상화주문</option>
					<option value="HH" <% if (actDivCode = "HH") then %>selected<% end if %> >상품변경취소주문</option>
					<option value="EE" <% if (actDivCode = "EE") then %>selected<% end if %> >교환취소주문</option>
					<option value="MM" <% if (actDivCode = "MM") then %>selected<% end if %> >반품취소주문</option>
				</select>
				<p>
				* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				&nbsp;&nbsp;
				* 검색조건 :
				<select class="select" name="searchfield">
					<option value=""></option>
					<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >주문번호</option>
					<option value="sitename" <% if (searchfield = "sitename") then %>selected<% end if %> >매출처</option>
				</select>
				<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit('');"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		※ 속도가 느려도 계속 누르지 마시고 기다려 주세요. 부하가 큰 페이지 입니다.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!--<h5>작업중</h5>-->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= oCMaechulLog.FTotalcount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2">출고처</td>

	<% if (dategbn="ActDate") then %>
		<td rowspan="2">결제일<br>(처리일)</td>
	<% elseif (dategbn="chulgoDate") then %>
		<td rowspan="2">출고일</td>
	<% else %>
		<td rowspan="2">원결제일</td>
	<% end if %>

	<% if (C_InspectorUser = False) then %>
	<td rowspan="2">소비자가<br>합계</td>
	<td rowspan="2">판매가<br>(할인가)</td>
	<td rowspan="2">상품쿠폰<br>적용가</td>
	<td colspan="3">보너스쿠폰</td>
	<td rowspan="2">기타할인<br>(올앳)</td>
	<% end if %>
	<td rowspan="2">매출총액</td>
	<td rowspan="2">업체<Br>정산액</td>
	<td rowspan="2"><b>회계매출</b></td>
	<td rowspan="2">구매<Br>마일리지</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="45">비율<br>쿠폰</td>
	<td width="45">정액<br>쿠폰</td>
	<td width="45">배송비<br>쿠폰</td>
	<% end if %>
</tr>

<% if oCMaechulLog.FresultCount >0 then %>
<% for i=0 to oCMaechulLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulLog.FItemList(i).fsitename %></td>
	<td>
		<a href="javascript:pop_detail_list('<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','01','<%= left(oCMaechulLog.fitemlist(i).fyyyymm,4) %>','<%= mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2) %>','<%= LastDayOfThisMonth( left(oCMaechulLog.fitemlist(i).fyyyymm,4),mid(oCMaechulLog.fitemlist(i).fyyyymm,6,2)) %>','<%= oCMaechulLog.FItemList(i).fsitename %>');" onfocus="this.blur()">
		<%= oCMaechulLog.FItemList(i).fyyyymm %></a>
	</td>
	<% if (C_InspectorUser = False) then %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).forgTotalPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalsum, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash, 0) %></td>
	<td align="right"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMileage, 0) %></td>
	<td></td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="50" align="center">
       	<% if oCMaechulLog.HasPreScroll then %>
			<span class="list_link"><a href="javascript:searchSubmit('<%= oCMaechulLog.StartScrollPage-1 %>')">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oCMaechulLog.StartScrollPage to oCMaechulLog.StartScrollPage + oCMaechulLog.FScrollCount - 1 %>
			<% if (i > oCMaechulLog.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oCMaechulLog.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:searchSubmit('<%=i%>')" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oCMaechulLog.HasNextScroll then %>
			<span class="list_link"><a href="javascript:searchSubmit('<%=i%>')">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="30">검색된 내용이 없습니다.</td>
</tr>
<% end if %>

<%
set oCMaechulLog = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
