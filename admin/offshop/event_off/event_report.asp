<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인이벤트 통계
' History : 2010.03.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventreport_Cls.asp"-->
<%
Call fnSetEventCommonCode_off '공통코드 어플리케이션 변수에 세팅

dim evt_code,i,evt_kind,ReportType ,BasicDateSet, fromDate, toDate, page ,ttSellPrice ,shopid
dim evt_name , isgift, israck ,isprize ,issale ,ttsum_cnt ,datefg , totselljumuncnt
dim evt_cateL, evt_cateM, yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,datelen, inc3pl
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	datefg = requestCheckVar(request("datefg"),32)
	isgift	= requestCheckVar(Request("isgift"),1)
	israck	= requestCheckVar(Request("israck"),1)
	isprize	= requestCheckVar(Request("isprize"),1)
	issale	= requestCheckVar(Request("issale"),1)
	shopid = requestCheckVar(request("shopid"),32)
	ReportType = requestCheckVar(request("ReportType"),10)
	evt_name = requestCheckVar(request("evt_name"),60)
	evt_code = requestCheckVar(request("evt_code"),6)
	evt_kind = requestCheckVar(Request("evt_kind"),10)	'이벤트종류
	evt_cateL = requestCheckVar(request("selC"),10)
	evt_cateM = requestCheckVar(request("selCM"),10)
	menupos = requestCheckVar(request("menupos"),10)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "event"
IF ReportType="" THEN ReportType="e"
IF evt_kind = "" THEN
	evt_kind="1"
END IF

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-90)
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

	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

dim oReport
set oReport = new Cevtreport_list
	oReport.frectevt_startdate = fromDate
	oReport.frectevt_enddate = toDate
	oReport.FRectShopID = shopid
	oReport.FRectevt_code = evt_code
	oReport.frectevt_kind = evt_kind
	oReport.FRectReportType= ReportType
	oReport.FRectevt_name = evt_name
	oReport.frectissale 	= issale
	oReport.frectisgift 	= isgift
	oReport.frectisrack 	= israck
	oReport.frectisprize 	= isprize
	oReport.frectevt_cateL	= evt_cateL
	oReport.frectevt_cateM	= evt_cateM
	oReport.frectdatefg = datefg
	oReport.FRectInc3pl = inc3pl

	'/통계 가져옴
	oReport.getevent_sum

Dim arreventkind
	'공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	arreventkind= fnSetCommonCodeArr_off("evt_kind",False)

ttSellPrice = 0
ttsum_cnt = 0
totselljumuncnt = 0
%>

<script language="javascript">

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function submitfrm(){
		frmEvt.submit();
	}

	//상세보기
	function pop_detail(datefg,SType,evt_code,shopid,yyyy1,mm1,dd1,yyyy2,mm2,dd2){
		 var pop_detail = window.open('/admin/offshop/event_off/event_report_detail.asp?datefg='+datefg+'&SType='+SType+'&evt_code='+evt_code+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&shopid='+shopid+'&menupos=<%=menupos%>','pop_detail','width=1024,height=768,scrollbars=yes,resizable=yes');
		 pop_detail.focus();
	}

	function enddatech(v){
		if (v=='event'){
			frmEvt.yyyy2.style.background='EEEEEE';
			frmEvt.mm2.style.background='EEEEEE';
			frmEvt.dd2.style.background='EEEEEE';
		}else{
			frmEvt.yyyy2.style.background='FFFFFF';
			frmEvt.mm2.style.background='FFFFFF';
			frmEvt.dd2.style.background='FFFFFF';
		}
	}

</script>

<!-- 표 상단바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frmEvt" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 :
				<% draweventmaechul_datefg "datefg" ,datefg ," onchange='submitfrm()'"%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>

					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
				<% end if %>
				<p>
				<!--<input type="radio" name="ReportType" value="e" <% IF ReportType="e" Then response.write "checked" %>>이벤트 기간별 매출
				<input type="radio" name="ReportType" value="s" <% IF ReportType="s" Then response.write "checked" %>>선택 기간별 매출-->
				* 이벤트종류 : <%sbGetOptEventCodeValue_off "evt_kind", evt_kind, True," onchange='submitfrm()'"%>
				<!--이벤트번호 : <input type="text" size="10" name="evt_code" value="<%=evt_code%>">//-->
				&nbsp;&nbsp;
				* 이벤트명 : <input type="text" size="30" name="evt_name" value="<%=evt_name%>">
				&nbsp;&nbsp;
		    	* 이벤트타입 :
		    	<input type="checkbox" name="issale" value="Y" onclick='submitfrm()' <% if issale = "Y" then response.write " checked"%>>할인
		    	<input type="checkbox" name="isgift" value="Y" onclick='submitfrm()' <% if isgift = "Y" then response.write " checked"%>>사은품
		    	<input type="checkbox" name="israck" value="Y" onclick='submitfrm()' <% if israck = "Y" then response.write " checked"%>>매대
		    	<input type="checkbox" name="isprize" value="Y" onclick='submitfrm()' <% if isprize = "Y" then response.write " checked"%>>당첨
				<p>
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
	            &nbsp;&nbsp;
				<!-- #include virtual="/common/module/categoryselectbox_event.asp"-->

				<script>
					enddatech('<%=datefg%>');
				</script>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmEvt.submit();">
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
        	<font color="red">[필독]</font> ※ 이벤트에 상품이 등록되지 않은경우, 통계가 산출되지 않습니다.
			<Br>상세보기에 상품별 데이터만 실시간 매출 내역이며, 나머지 모든 통계 데이터는 하루에 한번 새벽에 업데이트 됩니다.
	    </td>
	    <td align="right">
        </td>
	</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=oReport.FResultCount%></b>개 ※ 500건 까지 검색가능
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>이벤트<br>번호</td>
	<td>이미지</td>
	<td>이벤트<br>타입</td>
	<!--
	<td>브랜드</td>
	<td>이벤트<br>종류</td>
	//-->
	<td>카테고리</td>
	<td>이벤트명</td>
	<td>시작일<br>종료일</td>
	<td>
		<% if datefg = "event" then %>
			이벤트기간<br>일수
		<% else %>
			시작일대비<br>지난일수
		<% end if %>
	</td>
	<td>매장</td>
	<td>매출액</td>
	<td>
		<% if datefg = "event" then %>
			이벤트기간<br>일평균매출액
		<% else %>
			시작일대비<br>일평균매출액
		<% end if %>
	</td>
	<!--<td>주문<br>건수</td>-->
	<td>판매<br>수량</td>
	<td>담당MD</td>
	<td>비고</td>
</tr>
<%
if oReport.FresultCount>0 then

For i = 0 To oReport.FResultCount - 1

ttSellPrice = ttSellPrice + oReport.FItemList(i).fsellsum
ttsum_cnt = ttsum_cnt + oReport.FItemList(i).fsum_cnt
totselljumuncnt = totselljumuncnt + oReport.FItemList(i).ftotselljumuncnt

'/이벤트일수
datelen = ""

if datefg = "event" then
	datelen = datediff("d",oReport.FItemList(i).fevt_startdate,oReport.FItemList(i).fevt_enddate)
else
	datelen = datediff("d",oReport.FItemList(i).fevt_startdate,date())
end if
%>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td width=50><%= oReport.FItemList(i).fevt_code %></td>
	<td width=60>
		<%= CHKIIF(oReport.FItemList(i).fimgbasic<>"","<img src=""" & oReport.FItemList(i).fimgbasic & """ width=""50"" height=""50"" style=""cursor:pointer;"" onClick=""jsImgView('" & oReport.FItemList(i).fimgbasic & "');"">","") %>
	</td>
  	<td>
  		<%
  			if oReport.FItemList(i).fissale = "Y" then
  				response.write " <img src='http://fiximage.10x10.co.kr/web2008/category/icon_sale.gif'> "
  			end if
  			if oReport.FItemList(i).fisgift = "Y" then
  				response.write " <img src='http://fiximage.10x10.co.kr/web2008/category/icon_gift.gif'> "
  			end if
  			if oReport.FItemList(i).fisrack = "Y" then
  				response.write " 매대("&oReport.FItemList(i).fisracknum&") "
  			end if

  			if oReport.FItemList(i).fisprize = "Y" then
  				response.write " 당첨 "
  			end if
  		%>
  	</td>
  	<!--
	<td><%= oReport.FItemList(i).fmakerid %></td>
	<td><%=fnGetCommCodeArrDesc_off(arreventkind,oReport.FItemList(i).fevt_kind)%></td>
	//-->
	<td>
		<%= oReport.FItemList(i).fcate_nm1 %>
		<% if oReport.FItemList(i).fcate_nm2 <> "" then %>
			(<%= oReport.FItemList(i).fcate_nm2 %>)
		<% end if %>
	</td>
	<td align="left"><%= oReport.FItemList(i).fevt_name %></td>
	<td width=100 align="center">
		<%= oReport.FItemList(i).fevt_startdate %>
		<br><%= oReport.FItemList(i).fevt_enddate %>
	</td>
	<td width=70 align="right"><%= datelen %></td>
  	<td width=120>
  		<%
  		if oReport.FItemList(i).fshopid = "all" then
  			response.write "전체매장"
  		else
  			response.write oReport.FItemList(i).fshopname
  		end if
  		%>
  	</td>
	<td width=80 align="right" bgcolor="#E6B9B8"><%= FormatNumber(oReport.FItemList(i).fsellsum,0) %></td>
	<td width=80 align="right">
		<% if datelen <> "" and datelen <> 0 then %>
			<%= FormatNumber(oReport.FItemList(i).fsellsum / datelen,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<!--<td width=50 align="right"><%'= oReport.FItemList(i).ftotselljumuncnt %></td>-->
	<td width=50 align="right"><%= oReport.FItemList(i).fsum_cnt %></td>
	<td width=100><%= oReport.FItemList(i).fpartmdname %></td>
	<td width=150>
		<% if datefg = "jumun" then %>
			<input type="button" onclick="pop_detail('<%=datefg%>','D','<%= oReport.FItemList(i).fevt_code %>','<%= oReport.FItemList(i).fshopid %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');" value="날짜별" class="button">
			<input type="button" onclick="pop_detail('<%=datefg%>','T','<%= oReport.FItemList(i).fevt_code %>','<%= oReport.FItemList(i).fshopid %>','<%= yyyy1 %>','<%= mm1 %>','<%= dd1 %>','<%= yyyy2 %>','<%= mm2 %>','<%= dd2 %>');" value="상품별" class="button">
		<% else %>
			<input type="button" onclick="pop_detail('<%=datefg%>','D','<%= oReport.FItemList(i).fevt_code %>','<%= oReport.FItemList(i).fshopid %>','<%= left(oReport.FItemList(i).fevt_startdate,4) %>','<%= mid(oReport.FItemList(i).fevt_startdate,6,2) %>','<%= right(oReport.FItemList(i).fevt_startdate,2) %>','<%= left(oReport.FItemList(i).fevt_enddate,4) %>','<%= mid(oReport.FItemList(i).fevt_enddate,6,2) %>','<%= right(oReport.FItemList(i).fevt_enddate,2) %>');" value="날짜별" class="button">
			<input type="button" onclick="pop_detail('<%=datefg%>','T','<%= oReport.FItemList(i).fevt_code %>','<%= oReport.FItemList(i).fshopid %>','<%= left(oReport.FItemList(i).fevt_startdate,4) %>','<%= mid(oReport.FItemList(i).fevt_startdate,6,2) %>','<%= right(oReport.FItemList(i).fevt_startdate,2) %>','<%= left(oReport.FItemList(i).fevt_enddate,4) %>','<%= mid(oReport.FItemList(i).fevt_enddate,6,2) %>','<%= right(oReport.FItemList(i).fevt_enddate,2) %>');" value="상품별" class="button">
		<% end if %>
	</td>
</tr>
<%	Next %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=8 align="left">
		총평균매출액 : <%=FormatNumber(ttSellPrice/oReport.FResultCount,0) %>원
	</td>
	<td align="right"><%= FormatNumber(ttSellPrice,0) %></td>
	<td></td>
	<!--<td align="right"><%'= FormatNumber(totselljumuncnt,0) %></td>-->
	<td align="right"><%= FormatNumber(ttsum_cnt,0) %></td>
	<td colspan=2></td>
</tr>
<% ELSE %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="15">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>

<%
set oReport = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
