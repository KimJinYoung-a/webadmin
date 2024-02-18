<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]통계>>[1:1상담]당일답변율
' Hieditor : 2021.07.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/board/customer_board_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, tmpDateStr, startDateStr, endDateStr, chkGrpByReplyUser
Dim replyUser, i, userlevel, sitename, totalcount, d0pro, d1pro, d2pro, d3pro, d4pro, d0_1pro, d2_3_4pro
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)
	userlevel = requestcheckvar(getNumeric(request("userlevel")),10)
	sitename = requestcheckvar(request("sitename"),32)

chkGrpByReplyUser	= req("chkGrpByReplyUser","")
replyUser	= req("replyUser","")

if (yyyy1="") then
	chkGrpByReplyUser = "Y"
	startdateStr = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	startdateStr = DateSerial(yyyy1, mm1, dd1)
end if

yyyy1 = left(startdateStr,4)
mm1 = Mid(startdateStr,6,2)
dd1 = Mid(startdateStr,9,2)
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

endDateStr = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CReportMaster
	oreport.FRectuserlevel = userlevel
	oreport.FRectStart = startdateStr
	oreport.FRectEnd = endDateStr
	oreport.FRectReplyUser = replyUser
	oreport.FRectGroupByReplyUser = chkGrpByReplyUser
	oreport.FRectsitename = sitename
	oreport.FPageSize = 1000
	oreport.FCurrPage = 1
	oreport.getsameday_report

%>
<script type="text/javascript">

function searchSubmit(){
	//날짜 비교
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

	if (diffDay > 31){
		alert('검색기간은 1달 단위로 검색 가능 합니다.');
		return;
	}

	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간(답변일기준) <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		* 답변자ID : <input type="text" class="text" name="replyUser" value="<%=replyUser%>" size="12" maxlength="32">
		&nbsp;
		<input type="checkbox" name="chkGrpByReplyUser" value="Y" <%if (chkGrpByReplyUser = "Y") then %>checked<% end if %> > 답변자ID 표시
		&nbsp;
		* 회원등급 : <% DrawselectboxUserLevel "userlevel", userlevel, "" %>
		&nbsp;
		* 판매처 : 
		<select name="sitename">
			<option value="" <% if sitename="" then response.write " selected" %>>전체</option>
			<option value="10x10" <% if sitename="10x10" then response.write " selected" %>>10x10</option>
			<option value="10x10not" <% if sitename="10x10not" then response.write " selected" %>>제휴몰</option>
		</select>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="searchSubmit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		※ 1시간 지연 데이터 입니다. 부하가 있는 매뉴 입니다. 검색하신후 재차 누르지 마시고 기다려 주세요.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oreport.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (chkGrpByReplyUser = "Y") then %>
		<td rowspan=2>답변자ID</td>
	<% end if %>

	<td rowspan=2>답변월일</td>
	<td rowspan=2>기준</td>
	<td colspan=2>기준적합</td>
	<td colspan=3>기준미달</td>
	<td rowspan=2>합계</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>D+0</td>
	<td>D+1</td>
	<td>D+2</td>
	<td>D+3</td>
	<td>D+4 이상</td>
</tr>

<% if oreport.FresultCount > 0 then %>
	<%
	for i=0 to oreport.FresultCount -1
	totalcount=oreport.FItemList(i).fd0+oreport.FItemList(i).fd1+oreport.FItemList(i).fd2+oreport.FItemList(i).fd3+oreport.FItemList(i).fd4
	d0pro=(oreport.FItemList(i).fd0 / totalcount)*100
	d1pro=(oreport.FItemList(i).fd1 / totalcount)*100
	d2pro=(oreport.FItemList(i).fd2 / totalcount)*100
	d3pro=(oreport.FItemList(i).fd3 / totalcount)*100
	d4pro=(oreport.FItemList(i).fd4 / totalcount)*100
	d0_1pro=((oreport.FItemList(i).fd0+oreport.FItemList(i).fd1)/totalcount)*100
	d2_3_4pro=((oreport.FItemList(i).fd2+oreport.FItemList(i).fd3+oreport.FItemList(i).fd4)/totalcount)*100
	%>
	<tr align="center" bgcolor="FFFFFF">
		<% if (chkGrpByReplyUser = "Y") then %>
			<td rowspan=3><%= oreport.FItemList(i).freplyuser %></td>
		<% end if %>

		<td rowspan=3><%= oreport.FItemList(i).freplydate %></td>
		<td>답변건수</td>
		<td><%= oreport.FItemList(i).fd0 %></td>
		<td><%= oreport.FItemList(i).fd1 %></td>
		<td><%= oreport.FItemList(i).fd2 %></td>
		<td><%= oreport.FItemList(i).fd3 %></td>
		<td><%= oreport.FItemList(i).fd4 %></td>
		<td><%= CurrFormat(totalcount) %></td>
	</tr>
	<tr align="center" bgcolor="FFFFFF">
		<td>비율</td>
		<td><%= round(d0pro,2) %>%</td>
		<td><%= round(d1pro,2) %>%</td>
		<td><%= round(d2pro,2) %>%</td>
		<td><%= round(d3pro,2) %>%</td>
		<td><%= round(d4pro,2) %>%</td>
		<td rowspan=2>100%</td>
	</tr>
	<tr align="center" bgcolor="FFFFFF">
		<td>합계</td>
		<td colspan=2><%= round(d0_1pro,2) %>%</td>
		<td colspan=3><%= round(d2_3_4pro,2) %>%</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
