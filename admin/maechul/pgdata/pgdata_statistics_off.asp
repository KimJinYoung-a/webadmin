<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : PG사 승인통계_OFF
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim excmatchfinish, onlyCardPriceNotSame
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy3,yyyy4,mm3,mm4,dd3,dd4
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim fromDate2 ,toDate2
dim shopid
dim appDivCode, cardReaderID, cardGubun, cardComp, cardAffiliateNo, ipkumdate
dim searchfield, searchtext
dim dategubun
dim chkSearchIpkumDate, chkSearchAppDate
dim chkAppPrice, chkIpkumPrice, chkDiffPrice
dim reasonGubun
dim pggubun, PGuserid

Dim i, j

	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	excmatchfinish = requestCheckvar(request("excmatchfinish"),10)
	onlyCardPriceNotSame = requestCheckvar(request("onlyCardPriceNotSame"),10)

	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")

	yyyy3   = request("yyyy3")
	mm3     = request("mm3")
	dd3     = request("dd3")
	yyyy4   = request("yyyy4")
	mm4     = request("mm4")
	dd4     = request("dd4")

	shopid 			= request("shopid")
	appDivCode 		= request("appDivCode")
	cardReaderID 	= request("cardReaderID")
	cardGubun 		= request("cardGubun")
	cardComp 		= request("cardComp")
	cardAffiliateNo = request("cardAffiliateNo")
	ipkumdate 		= request("ipkumdate")

	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")

	dategubun 			= request("dategubun")
	chkSearchIpkumDate 	= request("chkSearchIpkumDate")
	chkSearchAppDate 	= request("chkSearchAppDate")

	chkAppPrice		= request("chkAppPrice")
	chkIpkumPrice	= request("chkIpkumPrice")
	chkDiffPrice	= request("chkDiffPrice")
	reasonGubun 	= requestCheckvar(request("reasonGubun"),32)
	pggubun		 	= requestCheckvar(request("pggubun"),32)
	PGuserid 		= requestCheckvar(request("PGuserid"),32)

	select case PGuserid
		case "BC":
			PGuserid = "비씨카드사"
		case "LOTTE":
			PGuserid = "롯데카드사"
		case "SAMSUNG":
			PGuserid = "삼성카드사"
		case "SHINHAN":
			PGuserid = "신한카드"
		case "HANACARD":
			PGuserid = "하나카드"
		case "HYUNDAI":
			PGuserid = "현대카드사"
		case "ALI":
			PGuserid = "Alipay"
		case "KB":
			PGuserid = "KB국민카드"
		case "NH":
			PGuserid = "NH농협카드"
		case else:
			'//
	end select

if (chkSearchIpkumDate="") then chkSearchAppDate = "Y"
if (page="") then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))+1, 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))

	fromDate2 = fromDate
	toDate2 = toDate
	yyyy3 = yyyy1
	mm3 = mm1
	dd3 = dd1
	yyyy4 = yyyy2
	mm4 = mm2
	dd4 = dd2
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)

	fromDate2 = DateSerial(yyyy3, mm3, dd3)
	toDate2 = DateSerial(yyyy4, mm4, dd4+1)
end if

if (research="") then
	dategubun = "appdate"
	chkAppPrice = "Y"
	chkIpkumPrice = "Y"
	'// chkDiffPrice = "Y"
	reasonGubun = "001"
end if

if (chkAppPrice = "" and chkIpkumPrice = "" and chkDiffPrice = "") then
	chkAppPrice = "Y"
end if


Dim oCPGDataStatistics
set oCPGDataStatistics = new CPGData

	oCPGDataStatistics.FRectshopid = shopid

	oCPGDataStatistics.FRectDateGubun = dategubun
	oCPGDataStatistics.FRectReasonGubun = reasonGubun
	oCPGDataStatistics.FRectPGGubun = pggubun
	oCPGDataStatistics.FRectPGuserid = PGuserid

	if (chkSearchAppDate = "Y") then
		oCPGDataStatistics.FRectStartdate = fromDate
		oCPGDataStatistics.FRectEndDate = toDate
	end if

	if (chkSearchIpkumDate = "Y") then
		oCPGDataStatistics.FRectStartIpkumdate = fromDate2
		oCPGDataStatistics.FRectEndIpkumDate = toDate2
	end if

    oCPGDataStatistics.getPGDataStatisticList_OFF

dim arrCardComp, tmpCardComp

arrCardComp = oCPGDataStatistics.GetArrCardComp()

dim sumCardComp(), totSumCardPrice, sumCardCompIpkum(), totSumCardPriceIpkum, scmTotCardPrice
redim sumCardComp(UBound(arrCardComp))
redim sumCardCompIpkum(UBound(arrCardComp))


'// ============================================================================
dim numRows : numRows = 0
if (chkAppPrice = "Y") then
	numRows = numRows + 1
end if
if (chkIpkumPrice = "Y") then
	numRows = numRows + 1
end if
if (chkDiffPrice = "Y") then
	numRows = numRows + 1
end if

dim currRow, currCol
dim arrData(2, 100)
dim arrSumData(2, 100)
dim arrCardPrice, arrIpkumPrice

function resetArray(byRef arrData)
	dim i
	For i = 0 to UBound(arrData, 2)
		arrData(0, i) = 0
		arrData(1, i) = 0
		arrData(2, i) = 0
	Next
end function

dim totMeachulPrice, totEtcPrice

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popPGDataList(yyyy1, mm1, dd1, shopid, cardComp) {
	var popup = window.open("pgdata_off.asp?menupos=1562&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid + "&cardComp=" + cardComp,"popPGDataList","width=1024,height=768,scrollbars=yes,resizable=yes");
	popup.focus();
}

function popjumundetail(yyyy1, mm1, dd1, shopid) {
	var popjumundetail = window.open("popOffShopOrderList.asp?menupos=648&oldlist=&datefg=jumun&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid+"&buyergubun=","popjumundetail","width=1024,height=768,scrollbars=yes,resizable=yes");
	popjumundetail.focus();
}

function popPGDataListNotMatch(yyyy1, mm1, dd1, shopid) {
	<% if (dategubun = "ipkumdate") then %>
		alert("거래일자로 검색한 후 조회가능합니다.");
		return;
	<% end if %>

	var popup = window.open("pgdata_off.asp?menupos=1562&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid + "&excmatchfinish=Y","popPGDataListNotMatch","width=1024,height=768,scrollbars=yes,resizable=yes");
	popup.focus();
}

function changeBg(idName, onoff) {
	var objA = document.getElementById(idName + "a");
	var objB = document.getElementById(idName + "b");
	var objC = document.getElementById(idName + "c");

	if (onoff == "on") {
		if (objA != undefined) {
			objA.style.background="F1F1F1";
		}
		if (objB != undefined) {
			objB.style.background="F1F1F1";
		}
		if (objC != undefined) {
			objC.style.background="F1F1F1";
		}
	} else {
		if (objA != undefined) {
			objA.style.background="FFFFFF";
		}
		if (objB != undefined) {
			objB.style.background="FFFFFF";
		}
		if (objC != undefined) {
			objC.style.background="FFFFFF";
		}
	}
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		*날짜표시:
		<select class="select" name="dategubun">
			<option value="appdate" <% if (dategubun = "appdate") then %>selected<% end if %> >승인(취소)일</option>
			<option value="ipkumdate" <% if (dategubun = "ipkumdate") then %>selected<% end if %> >입금예정일</option>
		</select>
		&nbsp;
		<input type="checkbox" name="chkSearchAppDate"  value="Y" <% if (chkSearchAppDate = "Y") then %>checked<% end if %> > *승인(취소)일자:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" name="chkSearchIpkumDate"  value="Y" <% if (chkSearchIpkumDate = "Y") then %>checked<% end if %> > *입금예정일:
		<% DrawDateBoxdynamic yyyy3, "yyyy3", yyyy4, "yyyy4", mm3, "mm3", mm4, "mm4", dd3, "dd3", dd4, "dd4"  %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", "" %>
		&nbsp;
		* 표시금액 :
		<input type="checkbox" name="chkAppPrice"  value="Y" <% if (chkAppPrice = "Y") then %>checked<% end if %> >승인금액
		<input type="checkbox" name="chkIpkumPrice"  value="Y" <% if (chkIpkumPrice = "Y") then %>checked<% end if %> >입금예정액
		<input type="checkbox" name="chkDiffPrice"  value="Y" <% if (chkDiffPrice = "Y") then %>checked<% end if %> >수수료
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		&nbsp;
		* PG사 :
		<select class="select" name="pggubun">
		<option value=""></option>
		<option value="KICC" <% if (pggubun = "KICC") then %>selected<% end if %> >KICC</option>
		<option value="HAND" <% if (pggubun = "HAND") then %>selected<% end if %> >HAND</option>
		<option value="partner" <% if (pggubun = "partner") then %>selected<% end if %> >partner</option>
		<option value="inicis" <% if (pggubun = "inicis") then %>selected<% end if %> >inicis</option>
		<option value="uplus" <% if (pggubun = "uplus") then %>selected<% end if %> >uplus</option>
		</select>
		&nbsp;
		* PG사id :
		<% Call DrawSelectBoxPGUseridOff("PGuserid", PGuserid, "") %>
		&nbsp;
		* 상세사유 :
		<select class="select" name="reasonGubun">
		<option value=""></option>
		<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >선수금(매출)</option>
		<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >선수금(제휴사 매출)</option>
		<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >선수금(예치금)</option>
		<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >선수금(예치금환급)</option>
		<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >선수금(기프트)</option>
		<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >선수금(기프트환급)</option>
		<option value="">---------------</option>
		<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS서비스</option>
		<option value="">---------------</option>
		<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >무통장미확인</option>
		<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >취소매칭</option>
		<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >핑거스현금매출</option>
		<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >이자수익</option>
		<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >기타</option>
		<option value="">---------------</option>
		<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >입력이전</option>
		</select>
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

* 기타 : 외환카드사, 하나SK카드

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">매장</td>
	<td width="80">
		<% if (dategubun <> "ipkumdate") then %>
			거래일자
		<% else %>
			입금예정일
		<% end if %>
		<br>거래액<br>입금예정액
	</td>

	<%
	totSumCardPrice = 0
	totSumCardPriceIpkum = 0
	scmTotCardPrice = 0
	for j = 0 to UBound(arrCardComp)
		sumCardComp(j) = 0
	%>
	<td width="80"><%= arrCardComp(j) %></td>
	<% next %>

	<td width="90">합계</td>

    <td width="90">매출</td>
    <td width="90">매출제외</td>

	<td width="90">SCM금액<br />(매출)</td>
	<td width="90">오차</td>
	<td width="90">매칭이전<br>승인내역</td>
	<td>비고</td>
</tr>

<% for i=0 to oCPGDataStatistics.FresultCount -1 %>
<%
totSumCardPrice = totSumCardPrice + oCPGDataStatistics.FItemList(i).FtotSumCardPrice
totSumCardPriceIpkum = totSumCardPriceIpkum + (oCPGDataStatistics.FItemList(i).FtotSumCardIpkumPrice * 1)
scmTotCardPrice = scmTotCardPrice + oCPGDataStatistics.FItemList(i).FscmTotCardPrice

totMeachulPrice = totMeachulPrice + oCPGDataStatistics.FItemList(i).FmeachulPrice
totEtcPrice = totEtcPrice + oCPGDataStatistics.FItemList(i).FetcPrice

yyyy = Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 4)
mm = Right(Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 7), 2)
dd = Right(Left(oCPGDataStatistics.FItemList(i).Fyyyymmdd, 10), 2)

resetArray(arrData)
currRow = 0
currCol = 0

arrCardPrice = oCPGDataStatistics.FItemList(i).FarrSumCardPrice
arrIpkumPrice = oCPGDataStatistics.FItemList(i).FarrSumCardIpkumPrice

'// 승인액
if (chkAppPrice = "Y") then
	for currCol = 0 to UBound(arrCardPrice)
		arrSumData(currRow, currCol) = arrSumData(currRow, currCol) + arrCardPrice(currCol)*1
		arrData(currRow, currCol) = arrCardPrice(currCol)
	next
	arrData(currRow, currCol + 1) = oCPGDataStatistics.FItemList(i).FtotSumCardPrice
	arrSumData(currRow, currCol + 1) = arrSumData(currRow, currCol + 1) + oCPGDataStatistics.FItemList(i).FtotSumCardPrice
	currRow = currRow + 1
end if

'// 입금예정액
if (chkIpkumPrice = "Y") then
	for currCol = 0 to UBound(arrCardPrice)
		arrSumData(currRow, currCol) = arrSumData(currRow, currCol) + arrIpkumPrice(currCol)*1
		arrData(currRow, currCol) = arrIpkumPrice(currCol)
	next
	arrData(currRow, currCol + 1) = oCPGDataStatistics.FItemList(i).FtotSumCardIpkumPrice
	arrSumData(currRow, currCol + 1) = arrSumData(currRow, currCol + 1) + oCPGDataStatistics.FItemList(i).FtotSumCardIpkumPrice
	currRow = currRow + 1
end if

'// 차액(수수료)
if (chkDiffPrice = "Y") then
	for currCol = 0 to UBound(arrCardPrice)
		arrSumData(currRow, currCol) = arrSumData(currRow, currCol) + (arrCardPrice(currCol) - arrIpkumPrice(currCol))*1
		arrData(currRow, currCol) = (arrCardPrice(currCol) - arrIpkumPrice(currCol))
	next
	arrData(currRow, currCol + 1) = (oCPGDataStatistics.FItemList(i).FtotSumCardPrice - oCPGDataStatistics.FItemList(i).FtotSumCardIpkumPrice)
	arrSumData(currRow, currCol + 1) = arrSumData(currRow, currCol + 1) + (oCPGDataStatistics.FItemList(i).FtotSumCardPrice - oCPGDataStatistics.FItemList(i).FtotSumCardIpkumPrice)
	currRow = currRow + 1
end if

%>
<tr id="obj<%= i %>a" align="center" bgcolor="FFFFFF" onmouseover="changeBg('obj<%= i %>', 'on')"; onmouseout="changeBg('obj<%= i %>', 'off')">
	<td rowspan="<%= numRows %>"><%= oCPGDataStatistics.FItemList(i).Fshopid %></td>
	<td rowspan="<%= numRows %>">
		<a href="javascript:popPGDataList('<%= yyyy %>', '<%= mm %>', '<%= dd %>', '<%= oCPGDataStatistics.FItemList(i).Fshopid %>', '')">
			<%= oCPGDataStatistics.FItemList(i).Fyyyymmdd %>
		</a>
	</td>

	<%
	for j = 0 to UBound(arrCardPrice)
	%>
	<td align="right">
		<a href="javascript:popPGDataList('<%= yyyy %>', '<%= mm %>', '<%= dd %>', '<%= oCPGDataStatistics.FItemList(i).Fshopid %>', '<%= arrCardComp(j) %>')">
			<%= FormatNumber(arrData(0, j), 0) %>
		</a>
	</td>
	<% next %>

	<td align="right">
		<% if Not IsNull(arrData(0, j + 1)) then %>
			<%= FormatNumber(arrData(0, j + 1), 0) %>
		<% end if %>
	</td>
    <td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FmeachulPrice, 0) %></td>
    <td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FetcPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FscmTotCardPrice, 0) %></td>
	<td align="right">
		<% if Not IsNull(oCPGDataStatistics.FItemList(i).FtotSumCardPrice) and (dategubun <> "ipkumdate") then %>
			<a href="javascript:popjumundetail('<%= yyyy %>','<%= mm %>','<%= dd %>','<%= oCPGDataStatistics.FItemList(i).Fshopid %>')"><%= FormatNumber((oCPGDataStatistics.FItemList(i).FtotSumCardPrice - oCPGDataStatistics.FItemList(i).FscmTotCardPrice), 0) %></a>
		<% end if %>
	</td>
	<td align="right"><a href="javascript:popPGDataListNotMatch('<%= yyyy %>', '<%= mm %>', '<%= dd %>', '<%= oCPGDataStatistics.FItemList(i).Fshopid %>')"><%= FormatNumber(oCPGDataStatistics.FItemList(i).FcardPriceNotMatch, 0) %></a></td>
	<td>
	</td>
</tr>
<% if (numRows > 1) then %>
<tr id="obj<%= i %>b" align="center" bgcolor="FFFFFF" onmouseover="changeBg('obj<%= i %>', 'on')"; onmouseout="changeBg('obj<%= i %>', 'off')">
	<%
	for j = 0 to UBound(arrCardPrice)
	%>
	<td align="right">
		<%= FormatNumber(arrData(1, j), 0) %>
	</td>
	<% next %>
	<td align="right">
		<% if Not IsNull(arrData(1, j + 1)) then %>
			<%= FormatNumber(arrData(1, j + 1), 0) %></td>
		<% end if %>
	<td align="right"></td>
	<td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
<% if (numRows > 2) then %>
<tr id="obj<%= i %>c" align="center" bgcolor="FFFFFF" onmouseover="changeBg('obj<%= i %>', 'on')"; onmouseout="changeBg('obj<%= i %>', 'off')">
	<%
	for j = 0 to UBound(arrCardPrice)
	%>
	<td align="right">
		<%= FormatNumber(arrData(2, j), 0) %>
	</td>
	<% next %>
	<td align="right">
		<% if Not IsNull(arrData(2, j + 1)) then %>
			<%= FormatNumber(arrData(2, j + 1), 0) %></td>
		<% end if %>
	<td align="right"></td>
	<td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
<% next %>
<tr align="center" bgcolor="FFFFFF">
	<td rowspan="<%= numRows %>">합계</td>
	<td rowspan="<%= numRows %>"></td>

	<%
	for j = 0 to UBound(sumCardComp)
	%>
	<td align="right">
		<%= FormatNumber(arrSumData(0, j), 0) %>
	</td>
	<% next %>

	<td align="right"><%= FormatNumber(arrSumData(0, j+1), 0) %></td>
    <td align="right"><%= FormatNumber(totMeachulPrice, 0) %></td>
    <td align="right"><%= FormatNumber(totEtcPrice, 0) %></td>
	<td align="right"><%= FormatNumber(scmTotCardPrice, 0) %></td>
	<td align="right">
		<% if (dategubun <> "ipkumdate") then %>
			<%= FormatNumber((totSumCardPrice - scmTotCardPrice), 0) %>
		<% end if %>
	</td>
	<td></td>
	<td></td>
</tr>
<% if (numRows > 1) then %>
<tr align="center" bgcolor="FFFFFF">
	<%
	for j = 0 to UBound(sumCardCompIpkum)
	%>
	<td align="right">
		<%= FormatNumber(arrSumData(1, j), 0) %>
	</td>
	<% next %>

	<td align="right"><%= FormatNumber(arrSumData(1, j+1), 0) %></td>
	<td align="right"></td>
	<td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
<% if (numRows > 2) then %>
<tr align="center" bgcolor="FFFFFF">
	<%
	for j = 0 to UBound(sumCardCompIpkum)
	%>
	<td align="right">
		<%= FormatNumber(arrSumData(2, j), 0) %>
	</td>
	<% next %>

	<td align="right"><%= FormatNumber(arrSumData(2, j+1), 0) %></td>
	<td align="right"></td>
	<td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
	<td></td>
	<td></td>
</tr>
<% end if %>
</table>

<%
set oCPGDataStatistics = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
