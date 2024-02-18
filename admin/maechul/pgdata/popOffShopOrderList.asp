<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매출 주문건 상세 공용페이지 NO 페이징 버전
' History : 2009.04.07 서동석 생성
'			2010.03.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<%
dim shopid, oldlist , datefg , prejumunno , makerid , menupos ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim cardMinusTotal, cashMinusTotal, cardMinusCnt, cashMinusCnt, buyergubun
dim etcTotal, etcCnt, etcMinusTotal, etcMinusCnt ,i,totalsum ,cardtotal, cashtotal, cardcnt, cashcnt
dim cardpayonly, excmatchfinish, logidx, showdetail, cardsum
dim research
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	menupos = request("menupos")
	shopid = request("shopid")
	oldlist = request("oldlist")
	datefg = request("datefg")
	makerid = request("makerid")
	buyergubun = request("buyergubun")

	cardpayonly = request("cardpayonly")
	excmatchfinish = request("excmatchfinish")
	logidx = request("logidx")
	showdetail = request("showdetail")

	cardsum = request("cardsum")

	if (cardsum <> "") then
		cardsum = Trim(Replace(cardsum, ",", ""))

		if Not IsNumeric(cardsum) then
			response.write "<script>alert('금액은 숫자만 가능합니다.');</script>"
			cardsum = ""
		end if
	end if

	research = request("research")

if (research = "") then
	cardpayonly = "Y"
	excmatchfinish = "Y"
end if

if datefg = "" then datefg = "maechul"
''if datefg = "" then datefg = "jumun"

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

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

''두타쪽 매출조회 권한
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectOldData = oldlist
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.frectdatefg = datefg
    ooffsell.FRectDesigner = makerid
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectbuyergubun = buyergubun

	ooffsell.FRectCardPayOnly = cardpayonly
	ooffsell.FRectExcMatchFinish = excmatchfinish

	''ooffsell.FRectCardSum = cardsum
	ooffsell.FRectPaySum = cardsum
    ooffsell.FRectPgDataCheck ="on"  ''서동석추가

	ooffsell.GetDaylySellJumunList

totalsum =0
cardtotal =0
cashtotal =0
cardcnt   =0
cashcnt   =0
cardMinusTotal =0
cashMinusTotal =0
cardMinusCnt   =0
cashMinusCnt   =0
etcTotal        =0
etcCnt          =0
etcMinusTotal   =0
etcMinusCnt     =0

''response.write logidx & "aaa"

Dim oCPGData
set oCPGData = new CPGData

	oCPGData.FRectIdx = logidx

	if (logidx <> "") then
    	oCPGData.getPGDataOne_OFF
	end if

%>

<script language="javascript">

function frmsubmit(){

	frm.submit();
}

function jsMatchThis(orderno, cardsum) {
	var frm = document.frmAct;
	<% if (oCPGData.FResultCount > 0) then %>
	var pgcardsum = "<%= oCPGData.FOneItem.FcardPrice %>";
	<% else %>
	var pgcardsum = "-1";
	<% end if %>
alert(cardsum + '/' + pgcardsum);
	if (cardsum*1 != pgcardsum*1) {
		alert("매칭불가!!\n\n결제액과 승인액이 서로 다릅니다.");
		return;
	}

	frm.orderno.value = orderno;

	if ((frm.orderno.value == "") || (frm.logidx.value == "")) {
		alert("잘못된 접근입니다.");
		return;
	}

	if (confirm("매칭하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="logidx" value="<%= logidx %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="3">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="3">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<% if C_IS_Maker_Upche then %>
					* 브랜드 : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				&nbsp;&nbsp;
				* 국적구분: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* 결제액(카드 or 현금):
				<input type="text" class="text" name="cardsum" value="<%= cardsum %>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<input type="checkbox" name="excmatchfinish" value="Y" <% if (excmatchfinish = "Y") then %>checked<% end if %> > PG사 승인내역 매칭완료 제외
				&nbsp;
				<input type="checkbox" name="cardpayonly" value="Y" <% if (cardpayonly = "Y") then %>checked<% end if %> > 카드결제 내역만
				&nbsp;
				<input type="checkbox" name="showdetail" value="Y" <% if (showdetail = "Y") then %>checked<% end if %> > 상세내역 표시
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- 표 상단바 끝-->

<br>
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<% if (oCPGData.FResultCount > 0) then %>
			PG사 : <%= oCPGData.FOneItem.FPGgubun %>
			&nbsp;
			PG사KEY : <%= oCPGData.FOneItem.FPGkey %>
			&nbsp;
			거래일자 : <%= oCPGData.FOneItem.FappDate %>
			&nbsp;
			거래액 : <b><%= FormatNumber(oCPGData.FOneItem.FcardPrice, 0) %></b>
		<% end if %>

    </td>
    <td align="right">
    </td>
</tr>
</table>

<p><br>

<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		검색결과 : <b><%=ooffsell.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25" width="110">
		매장명
	</td>
	<td width="110">
		<% if datefg = "maechul" then %>
			<%= chkIIF(shopid="cafe002","매출일","주문번호") %>
		<% else %>
			<%= chkIIF(shopid="cafe002","주문일","주문번호") %>
		<% end if %>
	</td>
	<td></td>

	<td width="90">마일리지</td>
	<td width="90">기프트카드</td>
	<td width="90">상품권</td>
	<td width="110">신용카드</td>
	<td width="110">현금</td>

	<td width="70"></td>

	<td rowspan="2" width="70">판매가</td>
	<td rowspan="2" width="70">매출액</td>
	<% if shopid<>"cafe002" then %>
		<td width="150">주문일시</td>
	<% end if %>
	<td>KICC매칭</td>
    <td rowspan="2">승인번호</td>
	<td rowspan="2">비고<br>(관련주문번호)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25"></td>
	<td height="25"></td>
	<td>브랜드</td>

	<td colspan="5">상품명</td>

	<td>수량</td>

	<% if shopid<>"cafe002" then %>
		<td></td>
	<% end if %>
	<td></td>
</tr>
<%
if ooffsell.FresultCount > 0 then

for i=0 to ooffsell.FresultCount-1

if prejumunno<>ooffsell.FItemList(i).ForderNo then

	totalsum = totalsum + ooffsell.FItemList(i).Frealsum
	if (ooffsell.FItemList(i).Fcardsum>0) then
        cardtotal = cardtotal + ooffsell.FItemList(i).Fcardsum
        cardcnt   = cardcnt + 1
    elseif (ooffsell.FItemList(i).Fcardsum<0) then
        cardMinusTotal = cardMinusTotal + ooffsell.FItemList(i).Fcardsum
        cardMinusCnt   =cardMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).Fcashsum>0) then
        cashtotal = cashtotal + ooffsell.FItemList(i).Fcashsum
        cashcnt   = cashcnt + 1
    elseif (ooffsell.FItemList(i).Fcashsum<0) then
        cashMinusTotal = cashMinusTotal + ooffsell.FItemList(i).Fcashsum
        cashMinusCnt   =cashMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).FgiftcardPaysum>0) then
        etcTotal = etcTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcCnt   = etcCnt + 1
    elseif (ooffsell.FItemList(i).FgiftcardPaysum<0) then
        etcMinusTotal = etcMinusTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcMinusCnt   =etcMinusCnt + 1
    end if

	prejumunno = ooffsell.FItemList(i).ForderNo

	if IsNull(ooffsell.FItemList(i).Fpointuserno) then
		ooffsell.FItemList(i).Fpointuserno = 0
	end if

%>
<tr align="center" bgcolor="<% if (showdetail = "") then %>FFFFFF<% else %>EEEEEE<% end if %>">
	<td height="25">
		<%= ooffsell.FItemList(i).Fshopid %>
	</td>
	<td>
		<%= chkIIF(shopid="cafe002",ooffsell.FItemList(i).Fshopregdate,ooffsell.FItemList(i).ForderNo) %>
	</td>
	<td><font color="<%= ooffsell.FItemList(i).JumunMethodColor %>"><%= ooffsell.FItemList(i).JumunMethodName %></font></td>

	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Fspendmile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FgiftcardPaysum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FTenGiftCardPaySum,0) %></td>
	<td align="right"><b><%= FormatNumber(ooffsell.FItemList(i).Fcardsum,0) %></b></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Fcashsum,0) %></td>

	<td><input type="button" class="button" value="매칭" onClick="jsMatchThis('<%= ooffsell.FItemList(i).ForderNo %>', '<%= ooffsell.FItemList(i).Fcardsum %>')" <% if (ooffsell.FItemList(i).FmatchCount > 0) then %>disabled<% end if %> ></td>

	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Ftotalsum,0) %></td>
	<td align="right"><b><%= FormatNumber(ooffsell.FItemList(i).Frealsum,0) %></b></td>
	<% if shopid<>"cafe002" then %>
		<td><%= ooffsell.FItemList(i).Fshopregdate %></td>
	<% end if %>
	<td>
		<% if (ooffsell.FItemList(i).FmatchCount > 0) then %>Y<% end if %>
	</td>

    <td><%=ooffsell.FItemList(i).Fcardappno%></td>
	<td><%=ooffsell.FItemList(i).Freforderno%></td>
</tr>
<% end if %>
<% if (showdetail = "Y") then %>
<tr align="center" bgcolor="FFFFFF">
	<td height="25"></td>
	<td height="25"></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>

	<td colspan="5" align="left">&nbsp; <%= ooffsell.FItemList(i).FItemName %> <%= ooffsell.FItemList(i).FItemOptionName %></td>

	<% if ooffsell.FItemList(i).FItemNo<0 then %>
		<td align="center"><font color=red><%= ooffsell.FItemList(i).FItemNo %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></font></td>
	<% else %>
		<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
	<% end if %>

	<% if shopid<>"cafe002" then %>
	<td colspan="4"></td>
	<% else %>
	<td colspan="3"></td>
	<% end if %>
</tr>
<% end if %>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="3"><b>총계</b></td>
	<td colspan="12" align="right">
		<table width=440 border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
		    <td>현금 :</td>
		    <td align="right"><%= FormatNumber(cashtotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cashcnt,0) %> 건)</td>
		    <td width=10></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal,0) %> 원</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(cashtotal + cashMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>카드 :</td>
		    <td align="right"><%= FormatNumber(cardtotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cardcnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal,0) %> 원</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>상품권 :</td>
		    <td align="right"><%= FormatNumber(etcTotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(etccnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(etcMinusTotal,0) %> 원</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(etcMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(etcTotal + etcMinusTotal,0) %> 원</td>
		</tr>
		<tr>
		    <td>합계 :</td>
		    <td align="right"><%= FormatNumber(cashtotal + cardtotal + etcTotal,0) %> 원</td>
		    <td align="center">(<%= FormatNumber(cashcnt + cardcnt + etccnt,0) %> 건)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal + cardMinusTotal + etcMinusTotal,0) %> 원</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt + cardMinusCnt + etcMinusCnt,0) %> 건)</font></td>
		    <td align="right"><%= FormatNumber(totalsum,0) %> 원</td>
		</tr>
		</table>
	</td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<form name="frmAct" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
<input type="hidden" name="mode" value="matchoneorder">
<input type="hidden" name="logidx" value="<%= logidx %>">
<input type="hidden" name="orderno" value="">
</form>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
