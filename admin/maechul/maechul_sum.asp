<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 매출통계
' History : 2007.12.06 한용민 생성
'			2011.05.18 서동석 수정(소비가, 할인금액, 상품쿠폰사용액등 추가/ 마진 = 매입액/실결제액으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechul_class.asp" -->

<%
dim dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc, exceptChangeOrd, research, grpTp
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim i ,defaultdate,defaultdate1 , olddata
dim channelDiv, inc3pl
	ipkumdatesucc = request("ipkumdatesucc")
	olddata = request("olddata")
	sitename = request("sitenamebox")
	accountdiv = request("accountdiv")
	bancancle = request("bancancle")
	if bancancle = "" then bancancle = "1"
	datecancle = request("datecancle")
	dateview1 = request("dateview1")
	if dateview1 = "" then dateview1 = "yes"
	defaultdate1 = dateadd("d",-60,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 60이전까지 검색
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then
		mm2 = month(now)
	else
		if Len(mm2) = 2 then
			mm2 = request("mm2")
		else
			mm2 = "0"&request("mm2")
		end if
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)

	channelDiv = request("channelDiv")

	research   = request("research")
	exceptChangeOrd = request("exceptChangeOrd")
    grpTp = request("grpTp")
    inc3pl = request("inc3pl")
	if (research="") then exceptChangeOrd="on"
	if (grpTp="") then grpTp="d"

dim Omaechul_list
set Omaechul_list = new Cmaechul_list
	Omaechul_list.FRectStartdate = yyyy1 & "-" & mm1 & "-" & dd1
	Omaechul_list.FRectEndDate = yyyy2 & "-" & mm2 & "-" & dd2
	Omaechul_list.frectdatecancle = datecancle
	Omaechul_list.frectbancancle = bancancle
	Omaechul_list.frectaccountdiv = accountdiv
	Omaechul_list.frectsitename = sitename
	Omaechul_list.frectipkumdatesucc = ipkumdatesucc
	Omaechul_list.fRectChannelDiv=channelDiv
	Omaechul_list.fRectexceptChangeOrd=exceptChangeOrd
	Omaechul_list.FRectGroupType=grpTp
	Omaechul_list.FRectInc3pl = inc3pl  ''2013/12/02 추가
	Omaechul_list.fmaechul_list()

if olddata = "no" then
	dim Omaechul_list_old
'	set Omaechul_list_old = new Cmaechul_list
'		Omaechul_list_old.FRectStartdate = (yyyy1-1) & "-" & mm1 & "-" & dd1
'		Omaechul_list_old.FRectEndDate = (yyyy2-1) & "-" & mm2 & "-" & dd2
'		Omaechul_list_old.frectdatecancle = datecancle
'		Omaechul_list_old.frectbancancle = bancancle
'		Omaechul_list_old.frectaccountdiv = accountdiv
'		Omaechul_list_old.frectsitename = sitename
'		Omaechul_list_old.fRectChannelDiv=channelDiv
'		Omaechul_list_old.fmaechul_list()
end if

''사이트구분
Sub Drawsitename(selectboxname, sitename)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str

	response.write "<select name='" & selectboxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if sitename ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">전체</option>"								'선택이란 단어가 나오도록.

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select id from [db_partner].[dbo].tbl_partner"
	userquery = userquery + " where 1=1"
	userquery = userquery + " and id <> ''"
	userquery = userquery + " and id is not null"
	userquery = userquery + " and userdiv= '999'"
	userquery = userquery + " group by id"

	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(sitename) = Lcase(rsget("id")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if

			response.write "<option value='" & rsget("id") & "' " & tem_str & ">" & rsget("id") & "</option>"
			tem_str = ""				'rsget에 id를 선택하고 검색할 값으로 선택
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub

Dim vParameter
	vParameter = "yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&datecancle="&datecancle&"&bancancle="&bancancle&"&accountdiv="&accountdiv&"&sitename="&sitename&"&dateview1="&dateview1&"&ipkumdatesucc="&ipkumdatesucc&""
%>

<script language="javascript" src="/admin/maechul/daumchart/FusionCharts.js"></script>		<!-- 그래프를 위한 자바스크립트파일-->
<script language="javascript">

function submit()
{
	frm.submit();
}

<!--월별 매출 상세보기 시작-->
function monthsum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var monthsum = window.open('/admin/maechul/maechul_month_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'monthsum','width=1024,height=768,scrollbars=yes,resizable=yes');
	monthsum.focus();
}
<!--월별 매출 상세보기 끝-->

<!--월별 매출 상세보기 시작-->
function weeksum(yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var weeksum = window.open('/admin/maechul/maechul_week_sum.asp?yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'weeksum','width=1024,height=768,scrollbars=yes,resizable=yes');
	weeksum.focus();
}
<!--월별 매출 상세보기 끝-->

<!--엑셀출력 시작-->
function excelprint(olddata,yyyy1,yyyy2,dateview1,datecancle,bancancle,accountdiv,sitename,ipkumdatesucc,menupos){
	var excelprint = window.open('/admin/maechul/maechul_sum_excel.asp?olddata='+olddata+'&yyyy1='+yyyy1+'&yyyy2='+yyyy2+'&dateview1='+dateview1+'&datecancle='+datecancle+'&bancancle='+bancancle+'&accountdiv='+accountdiv+'&sitename='+sitename+'&ipkumdatesucc='+ipkumdatesucc+'&menupos='+menupos ,'excelprint','width=1024,height=768,scrollbars=yes,resizable=yes');
	excelprint.focus();
}
<!--엑셀 출력  끝-->

function goOpenGraph()
{
	var graph = window.open('pop_graph.asp','graph','width=1024, height=768, scrollbars=yes, resizable=yes');
	graph.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
            * 기간 :
			<select name="dateview1" class="select">
				<option value="yes" <%=CHKIIF(dateview1="yes","selected","")%>>주문일</option>
				<option value="no" <%=CHKIIF(dateview1="no","selected","")%>>결제일</option>
			</select>
			<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			<input type="radio" name="grpTp" value="d" <%=CHKIIF(grpTp="d","checked","") %> >일별
			<input type="radio" name="grpTp" value="m" <%=CHKIIF(grpTp="m","checked","") %> >월별

		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

        	<!--<input type=checkbox name="datecancle" value="on" <% if datecancle="on" then  response.write "checked" %>>취소건만-->
        	* 사이트구분 : <% Drawsitename "sitenamebox",sitename %>
			* 주문구분 :
			<select name="bancancle" class="select">
				<option value="1" <%=CHKIIF(bancancle="1","selected","")%>>반품포함</option>
				<option value="3" <%=CHKIIF(bancancle="3","selected","")%>>반품제외</option>
				<option value="2" <%=CHKIIF(bancancle="2","selected","")%>>반품건만</option>
			</select>
        	* 결제구분 <select name="accountdiv">
        		<option value="" <% if accountdiv = "" then response.write "selected" %>>전체</option>
        		<option value="7" <% if accountdiv = "7" then response.write "selected" %>>무통장</option>
				<option value="14" <% if accountdiv = "14" then response.write "selected" %>>편의점결제</option>
        		<option value="20" <% if accountdiv = "20" then response.write "selected" %>>실시간</option>
        		<option value="50" <% if accountdiv = "50" then response.write "selected" %>>외부몰</option>
        		<option value="80" <% if accountdiv = "80" then response.write "selected" %>>올엣</option>
        		<option value="100" <% if accountdiv = "100" then response.write "selected" %>>신용카드</option>
        	</select>
        	* 채널구분
        	<select name="channelDiv">
	        	<option value="" <%=CHKIIF(channelDiv="","selected","") %> >전체</option>
	        	<option value="w" <%=CHKIIF(channelDiv="w","selected","") %> >웹</option>
	        	<option value="j" <%=CHKIIF(channelDiv="j","selected","") %> >제휴</option>
	        	<option value="m" <%=CHKIIF(channelDiv="m","selected","") %> >모바일웹</option>
        	</select>
        	<input type=checkbox name="exceptChangeOrd" value="on" <% if exceptChangeOrd="on" then  response.write "checked" %>>교환주문제외
        	<input type=checkbox name="ipkumdatesucc" value="on" <% if ipkumdatesucc="on" then  response.write "checked" %>>미결제건포함
            &nbsp;
            <b>* 매출처구분</b>
        	<% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
		</td>
	</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left" style="padding:10 0 10 0">
		&nbsp;
		<!-- 수정중
			<input type="button" class="button" value="엑셀출력" onclick="javascript:excelprint('<%=olddata%>','<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
	    -->
		</td>
		<td align="right">
		<% if (NOT C_InspectorUser) then %>
			<input type="button" class="button" value="그래프 통계" onclick="javascript:goOpenGraph();">
		<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 월별 주별 매출통계 상세내역 보기 시작-->
<!-- radio버튼으로 통합
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor=FFFFFF>
		<td align="left" style="padding:3 0 3 0">
			&nbsp;&nbsp;<a href="javascript:monthsum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			월별 매출통계 상세내역 보기 [클릭]</a>
		</td>
		<td align="right">
			<a href="javascript:weeksum('<%=yyyy1%>','<%=yyyy2%>','<%=dateview1%>','<%=datecancle%>','<%=bancancle%>','<%=accountdiv%>','<%=sitename%>','<%=ipkumdatesucc%>','<%=menupos%>');">
			주별 매출통계 상세내역 보기 [클릭]</a>&nbsp;&nbsp;
		</td>
	</tr>
</table>
-->
<!-- 월별 주별 매출통계 상세내역 보기 끝-->

<!-- 리스트 시작 -->
<%
dim totalsum_totalsum, totalcount_totalsum, subtotalprice_totalsum, totalbuysum_totalsum, spendScoupon_totalsum, spendMileage_totalsum
dim discountEtc_totalsum, sumpaymentEtc_totalsum, tendeliverBuysum_totalsum, tendeliverCount_totalsum, sunsuik_totalsum, magin_totalsum
Dim TTLtotalorgitemcostsum,TTLtotalOrgDlvPay,TTLtotalitemcostcouponNotApplied,TTLtotalCouponNotAppliedDlvPay
Dim TTLtotalitemcostsum,TTLtotalDlvPay,TTLtotalreducedDlvPay,TTLupchepartDeliverBuySum
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if dateview1 = "yes" then %>
		<td align="center" width="70" rowspan="2">주문일</td>
	<% elseif dateview1 = "no" then %>
		<td align="center" width="70" rowspan="2">입금일</td>
	<% end if %>
	<% if datecancle <> "" then %>
		<td align="center" width="70" rowspan="2">취소일</td>
	<% end if %>
    <td align="center" width="50" rowspan="2">총주문<br>건수</td>
<% if (NOT C_InspectorUser) THEN %>
    <td align="center" colspan="2">소비자가<br>A</td>
    <td align="center" colspan="2">할인금액<br>B</td>
    <td align="center" colspan="2">판매가(할인가)<br>C=A-B</td>
    <td align="center" colspan="2">상품쿠폰사용액<br>D</td>
    <td align="center" colspan="2">구매총액<br>E=C-D</td>
    <td align="center" colspan="2">보너스쿠폰사용액<br>F</td>
	<td align="center" width="70" rowspan="2">기타할인<br>H</td>
<% end if %>
	<td align="center" width="70" rowspan="2">매출액<br>E-F-H</td>
	<td align="center" width="70" rowspan="2">마일리지<br>G</td>
	<td align="center" width="70" rowspan="2">예치금<br>사용<br>M</td>
	<td align="center" width="70" rowspan="2"><strong>결제총액</strong><br>I=E-G</td>
	<td align="center" width="70" rowspan="2">매입가<br>(상품쿠폰)<br>J</td>
	<td align="center" colspan="2">배송비용<br>K</td>
	<td align="center" rowspan="2">마진<br>L=(J+K)/I</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (NOT C_InspectorUser) THEN %>
    <td>상품</td>
    <td>배송비</td>
    <td>상품</td>
    <td>배송비</td>
    <td>상품</td>
    <td>배송비</td>
    <td>상품</td>
    <td>배송비</td>
    <td>상품</td>
    <td>배송비</td>
    <td>상품</td>
    <td>배송비</td>
<% end if %>
    <td>텐배</td>
    <td>업배</td>
</tr>

<% for i = 0 to Omaechul_list.ftotalcount -1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td align="center">
		<% if (grpTp="d") and right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "토요일" then %>
			<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
		<% elseif (grpTp="d") and right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "일요일" then %>
			<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>
		<% else %>
			<%= Omaechul_list.flist(i).forderdate %>
		<% end if %>
	</td>
    <td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
<% if (NOT C_InspectorUser) THEN %>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalOrgDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum-Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalOrgDlvPay-Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied-Omaechul_list.flist(i).ftotalitemcostsum) %></td>
    <td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay-Omaechul_list.flist(i).ftotalDlvPay) %></td>

    <% if IsNULL(Omaechul_list.flist(i).ftotalitemcostsum) then %>
    	<td align="right" colspan="2" bgcolor="#9DCFFF"><%= CurrFormat(Omaechul_list.flist(i).ftotalsum) %></td>
    <% else %>
	    <td align="right" bgcolor="#9DCFFF"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostsum) %></td>
	    <td align="right" bgcolor="#9DCFFF"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalDlvPay) %></td>
    <% end if %>

    <% if IsNULL(Omaechul_list.flist(i).ftotalreducedDlvPay) then %>
    	<td align="right" colspan="2"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
    <% else %>
    	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fspendScoupon-(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay)) %></td>
    	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay) %></td>
    <% end if %>
	<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
<% end if %>
    <td align="right" bgcolor="#E6B9B8">
    	<%= CurrFormat((Omaechul_list.flist(i).ftotalitemcostsum+Omaechul_list.flist(i).ftotalDlvPay)-((Omaechul_list.flist(i).fspendScoupon-(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay))+(Omaechul_list.flist(i).ftotalDlvPay-Omaechul_list.flist(i).ftotalreducedDlvPay))-(Omaechul_list.flist(i).fdiscountEtc)) %>
    </td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsumpaymentetc) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>
    <td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliverBuysum) %></td>
	<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fupchepartDeliverBuySum) %></td>
    <td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
</tr>
    <%
    totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount
    TTLtotalorgitemcostsum = TTLtotalorgitemcostsum + Omaechul_list.flist(i).ftotalorgitemcostsum
    TTLtotalOrgDlvPay      = TTLtotalOrgDlvPay      + Omaechul_list.flist(i).ftotalOrgDlvPay
    TTLtotalitemcostcouponNotApplied = TTLtotalitemcostcouponNotApplied + Omaechul_list.flist(i).ftotalitemcostcouponNotApplied
    TTLtotalCouponNotAppliedDlvPay = TTLtotalCouponNotAppliedDlvPay + Omaechul_list.flist(i).ftotalCouponNotAppliedDlvPay
    TTLtotalitemcostsum = TTLtotalitemcostsum + Omaechul_list.flist(i).ftotalitemcostsum
    TTLtotalDlvPay = TTLtotalDlvPay + Omaechul_list.flist(i).ftotalDlvPay

    TTLtotalreducedDlvPay = TTLtotalreducedDlvPay + Omaechul_list.flist(i).ftotalreducedDlvPay

    totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum

	subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice
	totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum
	spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon
	spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage
	discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc
	sumpaymentEtc_totalsum = sumpaymentEtc_totalsum + Omaechul_list.flist(i).fsumpaymentetc
	tendeliverBuysum_totalsum = tendeliverBuysum_totalsum + Omaechul_list.flist(i).ftendeliverBuysum
	tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount
	TTLupchepartDeliverBuySum = TTLupchepartDeliverBuySum + Omaechul_list.flist(i).fupchepartDeliverBuySum
	sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik
	%>
<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td align="center" rowspan="2">
			총계
		</td>
		<td align="center"  rowspan="2"><%= totalcount_totalsum %></td>
<% if (NOT C_InspectorUser) THEN %>
		<td align="right"><%= NullOrCurrFormat(TTLtotalorgitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalOrgDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalCouponNotAppliedDlvPay) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay) %></td>

		<% IF IsNULL(TTLtotalitemcostsum) then %>
			<td align="right" colspan="2" rowspan="2"><%= CurrFormat(totalsum_totalsum) %></td>
		<% else %>
			<td align="right"><%= NullOrCurrFormat(TTLtotalitemcostsum) %></td>
			<td align="right"><%= NullOrCurrFormat(TTLtotalDlvPay) %></td>
		<% end if %>

		<% IF IsNULL(TTLtotalreducedDlvPay) then %>
			<td align="right" colspan="2" rowspan="2"><%= CurrFormat(spendScoupon_totalsum) %></td>
		<% else %>
			<td align="right"><%= NullOrCurrFormat(spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay)) %></td>
			<td align="right"><%= NullOrCurrFormat(TTLtotalDlvPay-TTLtotalreducedDlvPay) %></td>
		<% end if %>

		<td align="right" rowspan="2"><%= CurrFormat(discountEtc_totalsum) %></td>
<% end if %>
		<td align="right" rowspan="2"><%= CurrFormat((TTLtotalitemcostsum+TTLtotalDlvPay)-((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))+(TTLtotalDlvPay-TTLtotalreducedDlvPay))-(discountEtc_totalsum)) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(spendMileage_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(sumpaymentEtc_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(subtotalprice_totalsum) %></td>
		<td align="right" rowspan="2"><%= CurrFormat(totalbuysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(tendeliverBuysum_totalsum) %></td>
		<td align="center"><%= NullOrCurrFormat(TTLupchepartDeliverBuySum) %></td>
		<!-- <td align="right"><%= CurrFormat(sunsuik_totalsum) %></td>-->
		<td align="center" rowspan="2">
		<% if (subtotalprice_totalsum<>0) then %>
		    <% if Not IsNULL(sunsuik_totalsum) then %>
			<% magin_totalsum = CLNG((sunsuik_totalsum / subtotalprice_totalsum)*100*100)/100 %>
			<%= round(magin_totalsum,2) %>%
			<% end if %>
		<% end if %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<% if (NOT C_InspectorUser) THEN %>
	    <td colspan="2"><%= NullOrCurrFormat(TTLtotalorgitemcostsum+TTLtotalOrgDlvPay) %></td>
	    <td colspan="2"><%= NullOrCurrFormat((TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied)+(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay)) %></td>
	    <td colspan="2"><%= NullOrCurrFormat(TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay) %></td>
	    <td colspan="2"><%= NullOrCurrFormat((TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum)+(TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay)) %></td>

	    <% IF IsNULL(TTLtotalitemcostsum) then %>
	    <% else %>
	    	<td colspan="2"><%= NullOrCurrFormat(TTLtotalitemcostsum+TTLtotalDlvPay) %></td>
	    <% end if %>

	    <% IF IsNULL(TTLtotalreducedDlvPay) then %>
	    <% else %>
	   	 <td colspan="2"><%= NullOrCurrFormat((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))+(TTLtotalDlvPay-TTLtotalreducedDlvPay)) %></td>
	    <% end if %>
     <% end if %>
	    <td colspan="2"><%= NullOrCurrFormat(tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum) %></td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" rowspan="2">
		점유율
		</td>
		<td align="center" rowspan="2"></td>
	<% if (NOT C_InspectorUser) THEN %>
		<td align="right" colspan="2" rowspan="2">소비가대비=&gt</td>
		<td align="center">
			<% if TTLtotalorgitemcostsum<>0 then %>
			    <%= CLNG((TTLtotalorgitemcostsum-TTLtotalitemcostcouponNotApplied)/TTLtotalorgitemcostsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if TTLtotalOrgDlvPay<>0 then %>
			    <%= CLNG((TTLtotalOrgDlvPay-(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay))/TTLtotalOrgDlvPay*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="right" colspan="2" rowspan="2">판매가대비=&gt</td>
		<td align="center">
			<% if TTLtotalitemcostcouponNotApplied<>0 then %>
			    <%= CLNG((TTLtotalitemcostcouponNotApplied-TTLtotalitemcostsum)/TTLtotalitemcostcouponNotApplied*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if TTLtotalCouponNotAppliedDlvPay<>0 then %>
			    <%= CLNG((TTLtotalCouponNotAppliedDlvPay-TTLtotalDlvPay)/TTLtotalCouponNotAppliedDlvPay*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="right" colspan="2" rowspan="2">구매총액대비=&gt</td>

		<% IF IsNULL(TTLtotalreducedDlvPay) then %>
		    <td align="center" colspan="2" rowspan="2">
		    <% if (totalsum_totalsum<>0) then %>
		        <%= CLNG(spendScoupon_totalsum/totalsum_totalsum*100*100)/100 %> %
		    <% end if %>
		    </td>
		<% else %>
    		<td align="center">
    		<% if TTLtotalitemcostsum<>0 then %>
    		    <%= CLNG((spendScoupon_totalsum-(TTLtotalDlvPay-TTLtotalreducedDlvPay))/TTLtotalitemcostsum*100*100)/100 %> %
    		<% end if %>
    		</td>
    		<td align="center">
    		<% if TTLtotalDlvPay<>0 then %>
    		    <%= CLNG((TTLtotalDlvPay-TTLtotalreducedDlvPay)/TTLtotalDlvPay*100*100)/100 %> %
    		<% end if %>
    		</td>
		<% end if %>

		<td align="center" rowspan="2">
			<% if totalsum_totalsum<>0 then %>
			    <%= CLNG(discountEtc_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
	<% end if %>
		<td align="center" rowspan="2">
		</td>
		<td align="center" rowspan="2">
			<% if totalsum_totalsum<>0 then %>
			    <%= CLNG((spendMileage_totalsum)/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
			<% if sumpaymentEtc_totalsum<>0 then %>
			    <%= CLNG(sumpaymentEtc_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
		    <% if totalsum_totalsum<>0 then %>
			    <%= CLNG(subtotalprice_totalsum/totalsum_totalsum*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2">
			매입총액대비=&gt<br>
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(totalbuysum_totalsum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(tendeliverBuysum_totalsum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center">
			<% if (totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)<>0 then %>
			    <%= CLNG(TTLupchepartDeliverBuySum/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
			<% end if %>
		</td>
		<td align="center" rowspan="2" > </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
	<% if (NOT C_InspectorUser) THEN %>
	    <td colspan="2">
	    <% if (TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)<>0 then %>
	        <%= CLNG(((TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)-(TTLtotalitemcostcouponNotApplied+(TTLtotalOrgDlvPay-TTLtotalCouponNotAppliedDlvPay)))/(TTLtotalorgitemcostsum+TTLtotalOrgDlvPay)*100*100)/100 %> %
	    <% end if %>
	    </td>
	    <td colspan="2">
	    <% if (TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)<>0 then %>
	        <%= CLNG(((TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)-(TTLtotalitemcostsum+TTLtotalDlvPay))/(TTLtotalitemcostcouponNotApplied+TTLtotalCouponNotAppliedDlvPay)*100*100)/100 %> %
	    <% end if %>
	    </td>

	    <% IF IsNULL(TTLtotalreducedDlvPay) then %>
	    <% else %>
		    <td colspan="2">
		        <% if (TTLtotalitemcostsum+TTLtotalDlvPay)<>0 then %>
		            <%= CLNG(((spendScoupon_totalsum+TTLtotalDlvPay)-((TTLtotalDlvPay-TTLtotalreducedDlvPay)+TTLtotalreducedDlvPay))/(TTLtotalitemcostsum+TTLtotalDlvPay)*100*100)/100 %> %
		        <% end if%>
		    </td>
	    <% end if %>
    <% end if %>
		<td colspan="2">
	        <% if ((totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum))<>0 then %>
	            <%= CLNG((tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)/(totalbuysum_totalsum+tendeliverBuysum_totalsum+TTLupchepartDeliverBuySum)*100*100)/100 %> %
	        <% end if%>
	    </td>

	</tr>
</table>


<!-- Not Using .. OLD Ver -->
<% IF (FALSE) THEN %>
<p>------------------------<p>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if Omaechul_list.ftotalcount > 0 then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<% if dateview1 = "yes" then %>
			<td align="center" width="80">주문일</td>
		<% elseif dateview1 = "no" then %>
			<td align="center" width="80">입금일</td>
		<% end if %>
		<% if datecancle <> "" then %>
			<td align="center" width="80">취소일</td>
		<% end if %>
		<td align="center" width="70">총주문<br>건수</td>
		<td align="center" width="90">소비자가<br>매출총액<!-- <br>(배송비포함) --></td>
		<td align="center" width="80">할인금액</td>
		<td align="center" width="90">판매가<br>매출총액<br>(할인가)</td>
		<td align="center" width="80">상품쿠폰<br>사용액</td>
		<td align="center"><strong>상품<br>매출총액</strong></td>
		<td align="center" width="70">보너스<br>쿠폰<br>사용액</td>
		<td align="center" width="70">마일리지<br>사용액</td>
		<td align="center" width="70">기타할인</td>
		<td align="center" width="70">배송비<br>매출</td>
		<td align="center"><strong>매출총액<br>(실결제액)<br>배송비포함</strong></td>
		<td align="center">매입가<br>총액<br>(상품)</td>
		<td align="center"  width="60">텐바이텐<br>배송비<br>(매입)</td>
		<td align="center"  width="60">업체개별<br>배송비<br>(매입)</td>
		<td align="center"><strong>매출수익</strong></td>
		<td align="center">마진</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td></td>
        <td></td>
        <td>A</td>
        <td>B</td>
        <td>C=A-B</td>
        <td>D</td>
        <td>E=C-D</td>
        <td>F</td>
        <td>G</td>
        <td>H</td>
        <td>I</td>
        <td>J=E-(F+G+H)+I</td>
        <td>K</td>
        <td>L</td>
        <td>M</td>
        <td>N=J-(K+L+M)</td>
        <td>N/J</td>
    </tr>
	<% for i = 0 to Omaechul_list.ftotalcount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center">
		<% if right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "토요일" then %>
			<font color="blue"><%= Omaechul_list.flist(i).forderdate %></font>
		<% elseif right(FormatDateTime(Omaechul_list.flist(i).forderdate,1),3) = "일요일" then %>
			<font color="red"><%= Omaechul_list.flist(i).forderdate %></font>
		<% else %>
			<%= Omaechul_list.flist(i).forderdate %>
		<% end if %>
		</td>
		<% if datecancle <> "" then %>
			<td align="center"><%= Omaechul_list.flist(i).fcanceldate %></td>
		<% end if %>
		<td align="center"><%= Omaechul_list.flist(i).ftotalcount %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalorgitemcostsum-Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostcouponNotApplied-Omaechul_list.flist(i).ftotalitemcostsum) %></td>

		<td align="right">
		<% if IsNULL(Omaechul_list.flist(i).ftotalitemcostsum) then %>
		<%= CurrFormat(Omaechul_list.flist(i).ftotalsum-(Omaechul_list.flist(i).ftendeliversum+Omaechul_list.flist(i).fupchepartDeliverSum)) %>
		<% else %>
		<%= NullOrCurrFormat(Omaechul_list.flist(i).ftotalitemcostsum) %>
		<% end if %>
		</td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendScoupon) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fspendMileage) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fdiscountEtc) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).ftendeliversum+Omaechul_list.flist(i).fupchepartDeliverSum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsubtotalprice) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftotalbuysum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).ftendeliverBuysum) %></td>
		<td align="right"><%= NullOrCurrFormat(Omaechul_list.flist(i).fupchepartDeliverBuySum) %></td>
		<td align="right"><%= CurrFormat(Omaechul_list.flist(i).fsunsuik) %></td>
		<td align="center"><%= Omaechul_list.flist(i).fmagin %>%</td>
    </tr>
	<% totalsum_totalsum = totalsum_totalsum + Omaechul_list.flist(i).ftotalsum %>
	<% totalcount_totalsum = totalcount_totalsum + Omaechul_list.flist(i).ftotalcount %>
	<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list.flist(i).fsubtotalprice %>
	<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list.flist(i).ftotalbuysum %>
	<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list.flist(i).fspendScoupon %>
	<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list.flist(i).fspendMileage %>
	<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list.flist(i).fdiscountEtc %>
	<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list.flist(i).ftendeliversum %>
	<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list.flist(i).ftendeliverCount %>
	<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list.flist(i).fsunsuik %>
	<% next %>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
		총 합계
		</td>
		<td align="center"><%= totalcount_totalsum %></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"></td>
		<td align="right"><%= CurrFormat(totalsum_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendScoupon_totalsum) %></td>
		<td align="right"><%= CurrFormat(spendMileage_totalsum) %></td>
		<td align="right"><%= CurrFormat(discountEtc_totalsum) %></td>
		<td align="center"></td>
		<td align="right"><%= CurrFormat(subtotalprice_totalsum) %></td>
		<td align="right"><%= CurrFormat(totalbuysum_totalsum) %></td>
		<td align="right"><%= CurrFormat(tendeliversum_totalsum) %></td>
		<td align="center"></td>
		<td align="right"><%= CurrFormat(sunsuik_totalsum) %></td>
		<td align="center">
			<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
			<%= round(magin_totalsum,2) %>%
		</td>
		<%
		totalsum_totalsum = 0
		totalcount_totalsum = 0
		subtotalprice_totalsum = 0
		totalbuysum_totalsum = 0
		spendScoupon_totalsum = 0
		spendMileage_totalsum = 0
		discountEtc_totalsum = 0
		tendeliversum_totalsum = 0
		tendeliverCount_totalsum = 0
		sunsuik_totalsum = 0
		magin_totalsum = 0
		%>
	</tr>
	<!--
	<tr bgcolor="#DDDDFF">
		<td colspan="15">
			&nbsp;&nbsp;&nbsp;전년도 비교 내역 표시
			<input type=checkbox name="olddata" value="no" onclick=
			"submit();"<% if olddata="no" then  response.write "checked" %>>
		</td>
	</tr>
	//-->
	<% if (FALSE) and (olddata = "no") then %>
		<% if Omaechul_list_old.ftotalcount > 0 then %>
			<% for i = 0 to Omaechul_list_old.ftotalcount -1 %>
			<tr bgcolor="#FFFFFF">
				<td align="right">
				<% if right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "토요일" then %>
					<font color="blue"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% elseif right(FormatDateTime(Omaechul_list_old.flist(i).forderdate,1),3) = "일요일" then %>
					<font color="red"><%= Omaechul_list_old.flist(i).forderdate %></font>
				<% else %>
					<%= Omaechul_list_old.flist(i).forderdate %>
				<% end if %>
				</td>
				<% if datecancle <> "" then %>
					<td align="center"><%= Omaechul_list_old.flist(i).fcanceldate %></td>
				<% end if %>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalsum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftotalcount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsubtotalprice) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftotalbuysum) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendScoupon) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fspendMileage) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fdiscountEtc) %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).ftendeliversum) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).ftendeliverCount %></td>
				<td align="right"><%= CurrFormat(Omaechul_list_old.flist(i).fsunsuik) %></td>
				<td align="center"><%= Omaechul_list_old.flist(i).fmagin %>%</td>
			</tr>
			<% totalsum_totalsum = totalsum_totalsum + Omaechul_list_old.flist(i).ftotalsum %>
			<% totalcount_totalsum = totalcount_totalsum + Omaechul_list_old.flist(i).ftotalcount %>
			<% subtotalprice_totalsum = subtotalprice_totalsum + Omaechul_list_old.flist(i).fsubtotalprice %>
			<% totalbuysum_totalsum = totalbuysum_totalsum + Omaechul_list_old.flist(i).ftotalbuysum %>
			<% spendScoupon_totalsum = spendScoupon_totalsum + Omaechul_list_old.flist(i).fspendScoupon %>
			<% spendMileage_totalsum = spendMileage_totalsum + Omaechul_list_old.flist(i).fspendMileage %>
			<% discountEtc_totalsum = discountEtc_totalsum + Omaechul_list_old.flist(i).fdiscountEtc %>
			<% tendeliversum_totalsum = tendeliversum_totalsum + Omaechul_list_old.flist(i).ftendeliversum %>
			<% tendeliverCount_totalsum = tendeliverCount_totalsum + Omaechul_list_old.flist(i).ftendeliverCount %>
			<% sunsuik_totalsum = sunsuik_totalsum + Omaechul_list_old.flist(i).fsunsuik %>
			<% next %>
			<tr bgcolor="#F4F4F4">
				<td align="center" <% if datecancle = "on" then response.write "colspan=2" %>>
				총 합계
				</td>
				<td align="right">
					<%= CurrFormat(totalsum_totalsum) %>
				</td>
				<td align="center">
					<%= totalcount_totalsum %>
				</td>
				<td align="right">
					<%= CurrFormat(subtotalprice_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(totalbuysum_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(spendScoupon_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(spendMileage_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(discountEtc_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(tendeliversum_totalsum) %>
				</td>
				<td align="center">
					<%= CurrFormat(tendeliverCount_totalsum) %>
				</td>
				<td align="right">
					<%= CurrFormat(sunsuik_totalsum) %>
				</td>
				<td align="center">
					<% magin_totalsum = (sunsuik_totalsum / totalsum_totalsum)*100 %>
					<%= round(magin_totalsum,2) %>%
				</td>
			</tr>
		<% else %>
			<tr align="center" bgcolor="#DDDDFF">
		    	<td align=center bgcolor="#FFFFFF" colspan="15">전년도 검색 결과가 없습니다.</td>
		    </tr>
		<% end if %>
	<% end if %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>


</table>
<% end if %>

<%
	set Omaechul_list = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
