<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  고객데이터 통계(숨은 페이지 이며, scm에 매뉴로 노출 되어 있지 안음.)
' History : 2017.03.10 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<%
'if not(session("ssBctId")="tozzinet" or session("ssBctId")="djjung") then
'	response.write "권한이 없습니다. 관리자에게 문의 하세요."
'	session.codePage = 949
'	dbget.close() : response.end
'end if

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, fromDate,toDate
dim beforemonthly_firstday, beforemonthly_lastday
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)

'/이전달의 첫날
beforemonthly_firstday = DateSerial(Cstr(Year(dateadd("m", -1, date()))), Cstr(Month(dateadd("m", -1, date()))), "01")

'/이전달의 마지막날
beforemonthly_lastday = DateSerial(Cstr(Year(dateadd("m", -1, date()))), Cstr(Month(dateadd("m", -1, date()))), LastDayOfThisMonth(Year(dateadd("m", -1, date())), Month(dateadd("m", -1, date()))))

if (yyyy1="") then
	fromDate = beforemonthly_firstday
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(beforemonthly_lastday))
if (mm2="") then mm2 = Cstr(Month(beforemonthly_lastday))
if (dd2="") then dd2 = Cstr(day(beforemonthly_lastday))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)


%>

<script type='text/javascript'>

	function downloadfile(vmode){
		frm.action='/admin/datamart/mkt/customerdata_mkt_process.asp';
		frm.target='view';
		frm.mode.value=vmode;
		frm.submit();
		frm.action='';
		frm.target='';
		frm.mode.value='';
	}

	function categorylistview(vmode){
		frm.action='/admin/datamart/mkt/customerdata_mkt_process.asp';
		frm.target='categorylist';
		frm.mode.value=vmode;
		frm.submit();
		frm.action='';
		frm.target='';
		frm.mode.value='';
	}

	function gosubmit(){
		frm.action='';
		frm.target='';
		frm.mode.value='';
		frm.submit();
	}

</script>

<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<font color="red">※부하가 큰 페이지 입니다. 막누르지 마시고 기다리세요.</font>
    </td>
    <td align="right"></td>
</tr>
</table>
<br>
<!--<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<input type="button" onclick="categorylistview('categorylist');" value="관리카테고리리스트" class="button">
		<- 누르시고 하단에 리스트가 나오면, 하얀색영역에 마우스 왼쪽클릭 아무곳이나 한번 누르시고, 전체선택(CTRL+A) 하신후에 , 복사(CTRL+C) 하셔서 엑셀에 붙여넣기(CTRL+V) 하세요.
		<iframe id="categorylist" name="categorylist" src="" width="100%" height=100 frameborder="0" scrolling="no"></iframe>
    </td>
    <td align="right"></td>
</tr>
</table>-->

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="fromDate" value="<%= fromDate %>">
<input type="hidden" name="toDate" value="<%= toDate %>">
<!--<input type="button" onclick="downloadfile('bonuscoupon');" value="보너스쿠폰데이터CSV다운" class="button"> 한달단위로만 검색 하시고, 날짜 선택후 검색버튼을 꼭 누르고 다운로드 받아 주세요.-->
<!--<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="gosubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">		
	</td>
</tr>
</table>-->

</form>

<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/commonbodytail.asp"-->