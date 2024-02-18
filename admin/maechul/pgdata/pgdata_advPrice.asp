<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1
'', stplace, targetGbn, itemgubun
''dim ipgoMWdiv, itemMWdiv, itemid
''dim startYYYYMMDD, endYYYYMMDD
''dim addInfoType
''dim lastmwdiv, lastmakerid
dim tmpDate, nextMonth, prevMonth


page       	= requestCheckvar(request("page"),10)
research	= requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)


if (page="") then page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", 0, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
end if

'// ============================================================================
dim opgdataAdvPrice
set opgdataAdvPrice = new CPGData

opgdataAdvPrice.FPageSize = 100
opgdataAdvPrice.FCurrPage = 1
opgdataAdvPrice.FRectYYYYMM = yyyy1 + "-" + mm1

opgdataAdvPrice.getPGDataAdvPriceList

dim fromDate, toDate, showDiffPopup

fromDate = DateSerial(yyyy1, mm1, 1)
toDate = DateAdd("d", -1, DateAdd("m", 1, fromDate))

prevMonth = DateAdd("m", -1, fromDate)
nextMonth = DateAdd("m", 1, fromDate)

%>

<script language='javascript'>

function jsPopCheckPayLog(pggubun) {
    var pop;

    pop = window.open("/admin/maechul/payment_maechul_log_chk.asp?menupos=4161&pggubun=" + pggubun + "&yyyy1=<%= Year(fromDate) %>&mm1=<%= Month(fromDate) %>&dd1=<%= Day(fromDate) %>&yyyy2=<%= Year(toDate) %>&mm2=<%= Month(toDate) %>&dd2=<%= Day(toDate) %>")
}

function jsMakeAdvPrice(v) {
	var frm = document.getElementById('frmAct');
	var i;

	if (frm == undefined) {
		alert('============================== \n\n알 수 없는 오류입니다.\n\n ==============================');
		return;
	}

	frm.mode.value = "makeadvprc" + v;

	if (confirm("작성하시겠습니까?") == true) {
		frm.submit();
	}
}

function popAppPriceDetail(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var srcGbn;
	var lastDayOfMonth

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5);

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (pguserid) {
		case "balance":
		case "giftcard":
		case "mileage":
			// 통합포인트예치금관리
			switch (pguserid) {
				case "balance":
					srcGbn = "D";
					break;
				case "giftcard":
					srcGbn = "G";
					break;
				default:
					srcGbn = "M";
			}

			window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			break;
		case "bankipkum_10x10":
		case "bankrefund_10x10":
		case "bankipkum_fingers":
		case "bankrefund_fingers":
		case "gifticon":
		case "giftting":
		case "inicis":
		case "okcashbag":
		case "uplus":
		case "teenxteen3":
		case "teenxteen4":
		case "teenxteen6":
		case "teenxteen8":
		case "teenxteen9":
		case "tenbyten01":
		case "tenbyten02":
		case "R5523":
		case "KB":
		case "NH":
		case "LOTTE":
		case "BC":
		case "SAMSUNG":
		case "SHINHAN":
		case "KE":
		case "HANA":
		case "HYUNDAI":
			// 온라인/오프라인 승인내역
			if (targetGbn == "OF") {
				window.open("/admin/maechul/pgdata/pgdata_statistics_off.asp?menupos=1565&page=&research=on&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			} else {
				window.open("/admin/maechul/pgdata/pgdata_statistics_on.asp?menupos=1572&sumgubun=appMethod&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			}
			break;
		case "CASH":
		case "happymoney":
		case "streetshop018":
		case "partner":
			window.open("/common/offshop/maechul/statistic/statistic_checkmethod_datamart.asp?reload=on&menupos=1541&datefg=maechul&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&shopid=&offgubun=1&BanPum=&inc3pl=","popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			break;
		default:
			alert("표시할 정보없음");
			break;
	}
}

function popAppPriceDetailSUM(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var srcGbn;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5);

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	if (targetGbn == 'ON') {
		switch (pguserid) {
			case "balance":
			case "giftcard":
			case "mileage":
				// 통합포인트예치금관리
				switch (pguserid) {
					case "balance":
						srcGbn = "D";
						break;
					case "giftcard":
						srcGbn = "G";
						break;
					default:
						srcGbn = "M";
				}
				pop = window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
			default:
				pop = window.open("/admin/maechul/pgdata/pgdata_statistics_on.asp?menupos=1572&sumgubun=appMethod&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
		}
	} else if (targetGbn == 'OF') {
		switch (pguserid) {
			case "balance":
			case "giftcard":
			case "mileage":
				// 통합포인트예치금관리
				switch (pguserid) {
					case "balance":
						srcGbn = "D";
						break;
					case "giftcard":
						srcGbn = "G";
						break;
					default:
						srcGbn = "M";
				}
				pop = window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
			default:
				pop = window.open("/admin/maechul/pgdata/pgdata_statistics_off.asp?menupos=1565&page=&research=on&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
		}
	} else {
		alert("표시할 정보없음(" + targetGbn + ")");
		return;
	}
	pop.focus();
}

function popMeachulPriceDetail(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var lastDayOfMonth

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (targetGbn) {
		case "ON":
		case "AC":
			switch (pguserid) {
				case "balance":
				case "bankipkum_10x10":
				case "bankrefund_10x10":
				case "giftcard":
				case "gifticon":
				case "giftting":
				case "teenxteen3":
				case "teenxteen4":
				case "teenxteen6":
				case "teenxteen8":
				case "teenxteen9":
				case "mileage":
				case "okcashbag":
				case "tenbyten01":
				case "tenbyten02":
				case "R5523":
				case "bankrefund_fingers":
				case "":
					window.open("/admin/maechul/maechul_month_paymentPG_log.asp?menupos=1625&selD=2&selSY=" + yyyy + "&selSM=" + mm + "&selEY=" + yyyy + "&selEM=" + mm + "&selPGC=&selPGID=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
					break;
				default:
					alert("표시할 정보없음");
					break;
			}
			break;
		default:
			alert("표시할 정보없음");
			break;
	}
}

function popMeachulPriceReasonDetail(yyyymm, targetGbn, pggubun, pguserid, reasonGubun) {
	var yyyy, mm;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (targetGbn) {
		case "ON":
			pop = window.open("/admin/maechul/pgdata/pgdata_on.asp?menupos=1567&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&PGuserid=" + pguserid + "&reasonGubun=" + reasonGubun + "" + "&pggubun=" + pggubun,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
			break;
		case "AC":
		default:
			pop = window.open("/admin/maechul/pgdata/pgdata_off.asp?menupos=1562&dateType=A&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&reasonGubun=" + reasonGubun + "&pggubun=" + pggubun + "&pguserid=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
			break;
	}
	pop.focus();
}

function jsPopShowDiff(yyyymm, targetGbn, PGgubun, pguserid) {
	var yyyy, mm;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	pop = window.open("/admin/maechul/maechul_payment_log.asp?menupos=1606&dateGubun=payreqdate&matchState=Y&showOnlyPriceNotMatch=Y&targetGbn=" + targetGbn + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&PGuserid=" + pguserid + "&PGgubun=" + PGgubun,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
	pop.focus();
}

function jsGotoPrevMonth() {
    var frm = document.frm;
    var yyyy, mm;

    yyyy = <%= Year(prevMonth) %>;
    mm = <%= Month(prevMonth) %>;

    frm.yyyy1.value = yyyy;
    frm.mm1.value = (mm < 10 ? "0" : "") + mm;

    frm.submit();
}

function jsGotoNextMonth() {
    var frm = document.frm;
    var yyyy, mm;

    yyyy = <%= Year(nextMonth) %>;
    mm = <%= Month(nextMonth) %>;

    frm.yyyy1.value = yyyy;
    frm.mm1.value = (mm < 10 ? "0" : "") + mm;

    frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 월
            <input type="button" class="button" value="전월" onClick="jsGotoPrevMonth()">
            <input type="button" class="button" value="다음달" onClick="jsGotoNextMonth()">
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p />

* <font color="red">오프 현금매출이 누락</font>된 경우, 매장 브랜드정보에 기본정산방식이 설정되어 있는지 확인하세요.<br /><br />

* PG승인액(예치금, 기프트, 마일리지) : 통합예치금관리<br />
* PG승인액(무통장입금, 무통장환불, 신용카드, 카카오페이 등) : PG사 승인내역<br />
* 합계 : 선수금(매출) + 선수금(매출 이외), CS서비스 등<br />
* 선수금(매출) : 결제로그 실승인액<br />
* 선수금(매출 이외), CS서비스 등 : PG사 승인내역 사유 매출 이외 입력 건<br /><br />

* 오프라인 기프트카드 결제건 오류 있는 경우, 주문번호 매칭 및 사유입력 후 미수금재작성하면 정상입력됩니다.<br /><br />

* 검토(예치금, 기프트, 마일리지) : 결제로그 생성되었으나, 통합예치금 재작성 안된 케이스<br />
* 검토(신용카드 등) : 승인액 큰 경우 : 승인내역 있으나 결제로그 없는 케이스, 또는 결제로그-승인내역 매칭 안된 케이스<br />
* 검토(신용카드 등) : 승인액 작은 경우 : 결제로그 있으나 승인내역  없는 케이스<br />

<p />

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="재작성 01(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('01');">
			<input type="button" class="button" value="재작성 02(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('02');">
			<input type="button" class="button" value="재작성 03(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('03');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">년월</td>
		<td width="40">매출<br />구분</td>
		<td width="120">PG사</td>
		<td width="120">PG사ID</td>
		<td width="110">PG승인액</td>
		<td width="5"></td>
		<td width=110>검토</td>
		<td width="5"></td>
		<td width=110>합계</td>
		<td width=110>선수금<br />(매출)<br /><font color="red">(승인-매출)</font></td><!-- 001 -->
		<td width=110>선수금<br />(제휴사 매출)</td><!-- 002 -->
        <td width=110>선수금<br />(이니랜탈)</td><!-- 003 -->
		<td width=110>선수금<br />(예치금)</td><!-- 020 -->
		<td width=110>선수금<br />(예치금환급)</td><!-- 025 -->
		<td width=110>선수금<br />(기프트)</td><!-- 030 -->
		<td width=110>선수금<br />(기프트환급)</td><!-- 035 -->
        <td width=110>선수금<br />(B2B 매출)</td><!-- 004 -->
		<td width=110>CS서비스</td><!-- 040 -->
		<td width=110>이자수익</td><!-- 800 -->
		<td width=110>기타</td><!-- 900 -->
		<td width=110>핑거스<br />현금매출</td><!-- 901 -->
		<td width=110>무통장<br />미확인</td><!-- 950 -->
		<td width=110>취소매칭</td><!-- 999 -->
		<td width=110>사유<br />미입력</td><!-- XXX -->
		<td width="150">작성일</td>
		<td>비고</td>
	</tr>
	<% if opgdataAdvPrice.FResultCount >0 then %>
	<% for i=0 to opgdataAdvPrice.FResultcount-1 %>
    <%
    showDiffPopup = False
    if (opgdataAdvPrice.FItemList(i).FappPrice - opgdataAdvPrice.FItemList(i).GetAdvPriceSUM) <> 0 then
        showDiffPopup = True
    elseif opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice) <> "" then
        showDiffPopup = True
    end if

    %>
	<% if (opgdataAdvPrice.FItemList(i).FtargetGbn = "OF") then %>
	<tr bgcolor="#DDDDFF" height=25>
		<% else %>
		<tr bgcolor="#FFFFFF" height=25>
	<% end if %>
		<td align=center><%= opgdataAdvPrice.FItemList(i).Fyyyymm %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FtargetGbn %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FPGgubun %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FPGuserid %></td>
		<td align=right>
			<a href="javascript:popAppPriceDetailSUM('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FappPrice, 0) %>
			</a>
		</td>
		<td align=center></td>
		<td align=right>
			<% if showDiffPopup then %><a href="javascript:jsPopCheckPayLog('<%= opgdataAdvPrice.FItemList(i).FPGgubun %>')"><font color="red"><% end if %>
			<%= FormatNumber((opgdataAdvPrice.FItemList(i).FappPrice - opgdataAdvPrice.FItemList(i).GetAdvPriceSUM), 0) %>
            <% if showDiffPopup then %></font></a><% end if %>
		</td>
		<td align=center></td>
		<td align=right>
			<%= FormatNumber(opgdataAdvPrice.FItemList(i).GetAdvPriceSUM, 0) %>
		</td>
		<td align=right>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '001');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).GetMeachulPrice(opgdataAdvPrice.FItemList(i).FPGgubun, opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FreasonGubun001), 0) %>
			</a>
			<% Call opgdataAdvPrice.FItemList(i).ShowDiffIfExistWithPGgubun(opgdataAdvPrice.FItemList(i).FPGgubun, opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FreasonGubun001) %>
			<% if showDiffPopup then %>
			<a href="javascript:jsPopShowDiff('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>')">
				<%= opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice) %>
                <%= CHKIIF(opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice)="", "<br />(표시)", "") %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun002) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '002');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun002, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun003) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '003');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun003, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun020) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '020');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun020, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun025) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '025');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun025, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun030) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '030');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun030, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun035) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '035');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun035, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun004) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '004');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun004, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun040) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '040');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun040, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun800) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '800');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun800, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun900) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '900');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun900, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun901) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '901');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun901, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun950) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '950');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun950, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun999) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '999');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun999, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubunXXX) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', 'XXX');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubunXXX, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).Fregdate %></td>
		<td>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="26" align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>
</table>

<form id="frmAct" name="frmAct" method="post" action="https://scm.10x10.co.kr/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyymm" value="<%= yyyy1 %>-<%= mm1 %>">
</form>

<%
set opgdataAdvPrice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
