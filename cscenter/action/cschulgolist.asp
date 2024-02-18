<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

function CurrStateName(byval v)
	if v < "B006" then
		CurrStateName="접수"
	elseif v = "B006" then
		CurrStateName="업체처리완료"
	elseif v = "B007" then
		CurrStateName="처리완료"
	else
		CurrStateName = v
	end if
end function

function CurrStateColor(byval v)
	if v < "B006" then
		CurrStateColor="#000000"
	elseif v = "B006" then
		CurrStateColor="#000000"
	elseif v = "B007" then
		CurrStateColor="green"
	else
		CurrStateColor = "gray"
	end if
end function

function DivcdName(byval v)
	if v = "A004" or v = "A010" then
		DivcdName="반품"
	elseif v = "A000" then
		DivcdName="맞교환출고"
    elseif v = "A100" then
        DivcdName="교환출고"
	elseif v = "A002" then
		DivcdName="서비스"
	elseif v = "A011" then
	    DivcdName="맞교환회수"
	elseif v = "A012" or v = "A111" or v = "A112" then
		DivcdName="교환회수"
	elseif v = "CHG" then
	    DivcdName="교환CS"
	else
		DivcdName = v
	end if
end function


Const MaxRowSize = 1000
dim itemid, itemoption, itemgubun
dim currstate

dim datetype
dim startdate, enddate

itemid = request("itemid")
itemoption = request("itemoption")
currstate = request("currstate")

startdate = requestcheckvar(request("startdate"),10)
enddate = requestcheckvar(request("enddate"),10)

if startdate="" then
	startdate = Left(CStr(DateSerial(year(now), month(now), 1)),10)
end if
if enddate="" then
	enddate = date()
end if

datetype = request("datetype")

if (datetype="") then datetype="reg"
if (itemgubun="") then itemgubun="10"

datetype = "finish"
currstate = "finish"
''itemid = ""
''itemoption = ""


'상품코드 유효성 검사(2008.08.05;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim sqlStr, RowArr

'[코드정리]
'------------------------------------------------------------------------------
'A008			주문취소
'
'A004			반품접수(업체배송)
'A010			회수신청(텐바이텐배송)
'
'A001			누락재발송
'A002			서비스발송
'
'A200			기타회수
'
'A000			맞교환출고
'A100			상품변경 맞교환출고
'
'A009			기타사항
'A006			출고시유의사항
'A700			업체기타정산
'
'A003			환불
'A005			외부몰환불요청
'A007			카드,이체,휴대폰취소요청
'
'A011			맞교환회수(텐바이텐배송)
'A012			맞교환반품(업체배송)

'A111			상품변경 맞교환회수(텐바이텐배송)
'A112			상품변경 맞교환반품(업체배송)

''서비스발송, 맞교환출고 : 출고시 CS수량 마이너스
''상품변경 교환출고 : 출고시 CS수량 마이너스, 교환회수 완료시 수량 플러스
'' - 달을 넘겨 회수되는 경우도 고려해야 한다.
''기타회수 CS출고에 포함하지 않는다.
''6개월 이전 주문인 경우 검색이 되지 않는다.

sqlStr = " select top " & CStr(MaxRowSize) & " T.orderserial, T.sitename, T.chulgodate, T.itemid, T.itemoption, T.itemname, T.itemoptionname "

if (itemid <> "") then
	sqlStr = sqlStr + " , T.itemcnt "
	sqlStr = sqlStr + " , T.avgipgoprice "
	sqlStr = sqlStr + " , T.itemcost "
else
	sqlStr = sqlStr + " , isNull(sum(T.itemcnt),0) as itemcnt "
	sqlStr = sqlStr + " , isNull(sum(T.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " , isNull(sum(T.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 	, T.divcd "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	( "
sqlStr = sqlStr + " 		select "

if (itemid <> "") then
	sqlStr = sqlStr + " 			a.orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, a.finishdate as chulgodate "
	sqlStr = sqlStr + " 			, (case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end) as itemcnt "
	sqlStr = sqlStr + " 			, ((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*s.avgipgoprice) as avgipgoprice "
	sqlStr = sqlStr + " 			, ((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*d.itemcost) as itemcost "
else
	sqlStr = sqlStr + " 			'' as orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, '' as chulgodate "
	sqlStr = sqlStr + " 			, isNull(sum(case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end),0) as itemcnt "
	sqlStr = sqlStr + " 			, IsNull(sum((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*s.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " 			, IsNull(sum((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*d.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 			, d.itemid "
sqlStr = sqlStr + " 			, d.itemoption "
sqlStr = sqlStr + " 			, d.itemname "
sqlStr = sqlStr + " 			, d.itemoptionname "
sqlStr = sqlStr + " 			, a.divcd "
sqlStr = sqlStr + " 		from "
sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list a "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				a.id = d.masterid "
''sqlStr = sqlStr + " 			join [db_order].[dbo].tbl_order_master m "
''sqlStr = sqlStr + " 			on "
''sqlStr = sqlStr + " 				m.orderserial = a.orderserial "
sqlStr = sqlStr + "				join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
sqlStr = sqlStr + "				on "
sqlStr = sqlStr + "					1 = 1 "
sqlStr = sqlStr + "					and s.yyyymm = '" & Left(startdate, 7) & "' "
sqlStr = sqlStr + "					and s.itemgubun = '10' "
sqlStr = sqlStr + "					and s.itemid = d.itemid "
sqlStr = sqlStr + "					and s.itemoption = d.itemoption "
sqlStr = sqlStr + "					and s.lastmwdiv = 'M' "
sqlStr = sqlStr + " 		where "
sqlStr = sqlStr + " 			1 = 1 "
sqlStr = sqlStr + " 			and a.deleteyn <> 'Y' "
sqlStr = sqlStr + " 			and a.id >= 2500000 "
sqlStr = sqlStr + " 			and a.requireupche <> 'Y' "
sqlStr = sqlStr + " 			and a.divcd not in ('A008', 'A006', 'A001', 'A900', 'A010', 'A002', 'A111', 'A200', 'A999') "
sqlStr = sqlStr + " 			and a.currstate = 'B007' "
sqlStr = sqlStr + " 			and a.finishdate >= '" & startdate & "' "
sqlStr = sqlStr + " 			and a.finishdate < '" & enddate & "' "

if (itemid <> "") then
	sqlStr = sqlStr + " 	and d.itemid = " & itemid & " "
	if (itemoption <> "") then
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
	end if
else
	sqlStr = sqlStr + " 		group by "
sqlStr = sqlStr + " 			d.itemid, d.itemoption, d.itemname, d.itemoptionname,  a.divcd "
	sqlStr = sqlStr + " 		having sum(case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end) <> 0 "
end if

sqlStr = sqlStr + " 		union all "
sqlStr = sqlStr + " 		select "

if (itemid <> "") then
	sqlStr = sqlStr + " 			a.orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, a1.finishdate as chulgodate "
	sqlStr = sqlStr + " 			, d.confirmitemno as itemcnt "
	sqlStr = sqlStr + " 			, d.confirmitemno*s.avgipgoprice as avgipgoprice "
	sqlStr = sqlStr + " 			, d.confirmitemno*d.itemcost as itemcost "
else
	sqlStr = sqlStr + " 			'' as orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, '' as chulgodate "
	sqlStr = sqlStr + " 			, isNull(sum(d.confirmitemno),0) as itemcnt "
	sqlStr = sqlStr + " 			, IsNull(sum(d.confirmitemno*s.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " 			, IsNull(sum(d.confirmitemno*d.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 			, d.itemid "
sqlStr = sqlStr + " 			, d.itemoption "
sqlStr = sqlStr + " 			, d.itemname "
sqlStr = sqlStr + " 			, d.itemoptionname "
sqlStr = sqlStr + " 			, a.divcd "
sqlStr = sqlStr + " 		from "
sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list a "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_list a1 "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				1 = 1 "
sqlStr = sqlStr + " 				and a.id = a1.refasid "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				a.id = d.masterid "
''sqlStr = sqlStr + " 			join [db_order].[dbo].tbl_order_master m "
''sqlStr = sqlStr + " 			on "
''sqlStr = sqlStr + " 				m.orderserial = a.orderserial "
sqlStr = sqlStr + "				join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
sqlStr = sqlStr + "				on "
sqlStr = sqlStr + "					1 = 1 "
sqlStr = sqlStr + "					and s.yyyymm = '" & Left(startdate, 7) & "' "
sqlStr = sqlStr + "					and s.itemgubun = '10' "
sqlStr = sqlStr + "					and s.itemid = d.itemid "
sqlStr = sqlStr + "					and s.itemoption = d.itemoption "
sqlStr = sqlStr + "					and s.lastmwdiv = 'M' "
sqlStr = sqlStr + " 		where "
sqlStr = sqlStr + " 			1 = 1 "
sqlStr = sqlStr + " 			and a.requireupche <> 'Y' "
sqlStr = sqlStr + " 			and a.divcd = 'A100' "
sqlStr = sqlStr + " 			and a1.deleteyn <> 'Y' "
sqlStr = sqlStr + " 			and a1.id >= 2500000 "
sqlStr = sqlStr + " 			and a1.currstate = 'B007' "
sqlStr = sqlStr + " 			and a1.finishdate >= '" & startdate & "' "
sqlStr = sqlStr + " 			and a1.finishdate < '" & enddate & "' "

if (itemid <> "") then
	sqlStr = sqlStr + " 	and d.itemid = " & itemid & " "
	if (itemoption <> "") then
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
	end if
else
	sqlStr = sqlStr + " 		group by "
sqlStr = sqlStr + " 			d.itemid, d.itemoption, d.itemname, d.itemoptionname, a.divcd "
	sqlStr = sqlStr + " 		having sum(d.confirmitemno) <> 0 "
end if

sqlStr = sqlStr + " 	) T "

if (itemid = "") then
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	T.orderserial, T.sitename, T.chulgodate, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.divcd "
	sqlStr = sqlStr + " having sum(T.itemcnt) <> 0 "
end if

sqlStr = sqlStr + " order by "
sqlStr = sqlStr + " 	T.itemid, T.itemoption, T.chulgodate desc "

''response.write sqlStr
''response.end



IF application("Svr_Info")="Dev" THEN
    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        RowArr = rsget.getRows
    end if
    rsget.Close
ELSE
    db3_rsget.Open sqlStr,db3_dbget,1
    if not db3_rsget.Eof then
        RowArr = db3_rsget.getRows
    end if
    db3_rsget.Close
end if

dim RowCount, jumuncnt
RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

dim totno, i
totno = 0

%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script type="text/javascript">
function jsSubmitONE(itemid, itemoption) {
	var frm = document.frm;

	frm.itemid.value = itemid;
	frm.itemoption.value = itemoption;

	frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" rowspan="2" bgcolor="#EEEEEE">검색<br />조건</td>
        <td>
			검색기간 :
			<select class="select" name="datetype">
			    <option value="finish" <%= chkIIF(datetype="finish","selected","") %> >처리일</option>
			</select>
			<input id="sDt" name="startdate" value="<%=startdate%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="eDt" name="enddate" value="<%=enddate%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script>
				var CAL_Start = new Calendar({
					inputField : "sDt", trigger    : "sDt_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "eDt", trigger    : "eDt_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>

			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="12">
			&nbsp;
			옵션 <input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4">
        </td>
        <td align="center" width="50" rowspan="2" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
        <td>
			CS상태 : 처리완료 (최대 <%= MaxRowSize %>건 까지만 검색됩니다.)
        </td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p>
※ 재고기준월 : 검색시작월<br />
 - 서비스발송, 맞교환출고 : 출고시 CS수량 마이너스<br />
 - 상품변경 교환출고 : 출고시 CS수량 마이너스, 교환회수 완료시 수량 플러스<br />
 - 상품변경 교환회수 : 완료시 교환주문이 생성되므로 CS수량 변동없다.<br />
&nbsp;&nbsp; - 달을 넘겨 회수되는 경우도 고려<br />
 - 기타회수 : CS출고에 포함하지 않는다.<br />
 - 교환CS : 동일상품 교환출고 및 회수, 상품변경 교환출고 및 회수<br />
<!--
 - 6개월 이전 주문인 경우 검색이 되지 않는다.
-->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">주문번호</td>
		<td width="100">CS구분</td>
		<td width="30">구분</td>
		<td width="80">상품코드</td>
		<td width="50">옵션</td>
		<td width="100">바코드</td>
		<td width="300">상품명</td>
		<td width="200">옵션명</td>
		<td width="80">판매가</td>
		<td width="80">평균매입가</td>
		<td width="40">수량</td>
		<td width="150">출고일</td>
		<td>비고</td>
	</tr>
<%
if IsArray(RowArr) then
	for i=0 to RowCount
%>

	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= RowArr(0,i) %></td>
		<td><%= DivcdName(RowArr(10,i)) %></td>
		<td>10</td>
		<td><a href="javascript:jsSubmitONE('<%=RowArr(3,i)%>', '<%=RowArr(4,i)%>');"><%=RowArr(3,i)%></a></td>
		<td><a href="javascript:jsSubmitONE('<%=RowArr(3,i)%>', '<%=RowArr(4,i)%>');"><%=RowArr(4,i)%></a></td>
		<td align="left"><%= BF_MakeTenBarcode("10", RowArr(3,i), RowArr(4,i)) %></td>
		<td align="left"><%= DdotFormat(RowArr(5,i),25) %></td>
		<td align="left"><%= DdotFormat(RowArr(6,i),15) %></td>
		<td><%= RowArr(9,i) %></td>
		<td><%= RowArr(8,i) %></td>
		<td><%= RowArr(7,i) %></td>
		<td><%= RowArr(2,i) %></td>
		<td></td>
	</tr>
<%
			totno = totno + RowArr(7,i)
    next
end if

%>
    <tr height="25" bgcolor="#FFFFFF">
        <td align="right" colspan="13">총상품수 <%= totno %> 개 / 총주문건수 <%= jumuncnt %> 건</td>
    </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
