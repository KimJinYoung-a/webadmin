<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 traffic analysis  
' History : 2007.09.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->

<% 
dim yyyy , mm ,dd,buy_date,buy_date1 , defaultnow, vYYYY, vMM
dim yyyy1,mm1,dd1 ,fpageview_sum,ftotalcount_sum,fnewcount_sum,frecount_sum,frealcount_sum
	yyyy = left(now(),4)		'오늘날짜에서 년
	mm = mid(now(),6,2)			'오늘날짜에서 달
	dd = mid(now(),9,2) 		'오늘날짜에서 일
	defaultnow = dateadd("d",-30,yyyy &"-"& mm &"-"& dd)		'오늘날짜에서 -30일
	menupos = request("menupos") 
	buy_date = request("buy_date")
	if buy_date = "" then		'값 지정 안할경우 시작일 기본값
		buy_date = left(defaultnow,4) &"-"&  mid(defaultnow,6,2) &"-"& mid(defaultnow,9,2)    	 
	end if
	
	buy_date1 = request("buy_date1")	
	if buy_date1 = "" then			'값지정 안할경우 마지막일 기본값
		buy_date1 = yyyy &"-"& mm &"-"& dd    	 
	end if	

	vYYYY	= left(buy_date1,4)
	vMM		= mid(buy_date1,6,2)

dim otrafficlist , i
	set otrafficlist = new Ctrafficlist
	otrafficlist.frectbuy_date = left(buy_date,4)&mid(buy_date,6,2)&mid(buy_date,9,2)
	otrafficlist.frectbuy_date1 =left(buy_date1,4)&mid(buy_date1,6,2)&mid(buy_date1,9,2)	
	otrafficlist.Ftrafficlist()

fpageview_sum = 0
ftotalcount_sum = 0
fnewcount_sum = 0
frecount_sum = 0
frealcount_sum = 0
%>

<script language="javascript" src="/admin/traffic/daumchart/FusionCharts.js"></script>		<!-- 그래프를 위한 자바스크립트파일-->
<script language="javascript">

function popup()
{
	var popup = window.open('/admin/traffic/traffic_analysis.asp','popup','width=1024,height=768,scrollbars=yes,resizable=yes');
	popup.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input type="text" name="buy_date" size=10 value="<%= buy_date %>">			
		<a href="javascript:calendarOpen3(frm.buy_date,'시작일',frm.buy_date.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a> -
		<input type="text" name="buy_date1" size=10  value="<%= buy_date1 %>">
		<a href="javascript:calendarOpen3(frm.buy_date1,'마지막일',frm.buy_date1.value)">
		<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">		
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		- 해당 조회 기간중 처음 방문한 최초 1회만 카운트되며 이후의 방문은 인식하지 않습니다. <br>
		&nbsp;&nbsp;예를 들어 오전에 1번 방문하고 오후에 1번 방문하여도 중복으로 제거되므로 일 실방문자 수는 1이 됩니다.<br>
		- 방문자 구분은 사용자의 쿠키정보를 통해 구분합니다.		
	</td>
	<td align="right">
		<input type="button" value="traffic analysis 추출" onclick="popup()" class="button">				
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if otrafficlist.ftotalcount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= otrafficlist.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td >날짜</td>
	<td >페이지뷰</td>
	<td >전체방문자수</td>
	<td >신규방문자수</td>
	<td >재방문자수</td>
	<td >실제방문자수</td>
</tr>
<% for i=0 to otrafficlist.ftotalcount -1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td ><%= otrafficlist.flist(i).fyyyymmdd %></td>
	<td >
		<%= otrafficlist.flist(i).fpageview %>
		<% fpageview_sum = fpageview_sum + clng(otrafficlist.flist(i).fpageview) %>
	</td>
	<td >
		<%= otrafficlist.flist(i).ftotalcount %>
		<% ftotalcount_sum = ftotalcount_sum + clng(otrafficlist.flist(i).ftotalcount) %>
	</td>
	<td >
		<%= otrafficlist.flist(i).fnewcount %>
		<% fnewcount_sum = fnewcount_sum + clng(otrafficlist.flist(i).fnewcount) %>
	</td>
	<td >
		<%= otrafficlist.flist(i).frecount %>
		<% frecount_sum = frecount_sum + clng(otrafficlist.flist(i).frecount) %>
	</td>
	<td >
		<%= otrafficlist.flist(i).frealcount %>
		<% frealcount_sum = frealcount_sum + clng(otrafficlist.flist(i).frealcount) %>
	</td>
</tr>   
<% next %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td >합계</td>
	<td ><%= fpageview_sum %></td>
	<td ><%= ftotalcount_sum %></td>
	<td><%= fnewcount_sum %></td>
	<td><%= frecount_sum %></td>
	<td><%= frealcount_sum %></td>
</tr>
<tr align="center" bgcolor="FFFFFF" >
	<td colspan=10>
		
		<!-- 메일진 그래프 시작-->	
		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr bgcolor=FFFFFF>
			<td> &nbsp;<%= left(buy_date1,4) %>년 <%= mid(buy_date1,6,2) %>월 내역
			</td>			
			<td><div align="right"><input type="button" value="그래프프린트" onclick="javascript:window.print();" class="button"></div>
			</td>
		</tr>
		<tr bgcolor=FFFFFF>		
			<td align="center" colspan="2">		
				<br>
				<div id="chartdiv3" align="center"></div>
				<script type="text/javascript">	
				var chart = new FusionCharts("/admin/traffic/daumchart/MSCombiDY2D.swf", "chartdiv3", "640", "480", "0", "0");
				chart.setDataURL("/admin/traffic/daumchart/MSCombiDY2D.asp?param=yyyy=<%=vYYYY%>^^mm=<%=vMM%>");
				chart.render("chartdiv3");
				</script>
			</td>
		</tr>
		</table>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set otrafficlist = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->