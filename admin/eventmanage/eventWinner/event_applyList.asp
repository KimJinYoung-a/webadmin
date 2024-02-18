<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_applyList.asp
' Description :  이벤트 응모자 리스트
' History : 2007.09.19 김정인
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body topmargin="0" >

<%


dim evtCode
evtCode =request("eC")

dim UserLevelOpt , winCountOpt, AreaOpt , OrderCashOpt, SelectingOpt
UserLevelOpt = request("uLOpt")
winCountOpt  = request("winOpt")
AreaOpt  = request("arOpt")
OrderCashOpt  = request("ordOpt")
SelectingOpt  = request("selOpt")

dim arrList,intLoop
dim appList

set appList = new ClsEventApply
appList.FECode = evtCode
appList.FUserLevelOpt 	= UserLevelOpt
appList.FwinCountOpt 	= winCountOpt
appList.FAreaOpt 		= AreaOpt
appList.FOrderCashOpt 	= OrderCashOpt
appList.FSelectingOpt 	= SelectingOpt

arrList = appList.fnGetApplyList
set appList = nothing

'arrList(evtcom_idx ,evt_code ,userid ,evtcom_txt ,evtcom_regdate)

%>
<script language='javascript'>

function showTXT(divVal){

	var mx = document.body.scrollLeft + event.clientX+10;
	var my = document.body.scrollTop + event.clientY +10;

	var vDIV = document.getElementById(divVal);

	var iTooltd = document.getElementById("tooltd");
		iTooltd.innerHTML = vDIV.innerHTML;

	var iTool = document.getElementById("tool");

		iTool.style.left=mx;
		iTool.style.top=my;
		iTool.style.display="";


	//setTimeout(showTXT(divVal),10000);
}

function hideTXT(vDIV){

	var iTool = document.getElementById("tool");
	iTool.style.display="none";

}
</script>

<!-- 테이블 상단 검색바 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
        	<select name="uLOpt">
				<option value="">회원등급별</option>
				<option value="">VIP</option>
				<option value="">블루</option>
				<option value="">그린</option>
				<option value="">옐로우</option>
				<option value="">오렌지</option>
				<option value="">매니아</option>
			</select>
			<select name="winOpt">
				<option value="">당첨횟수별</option>
				<option value="">당첨적은순</option>
				<option value="">당첨많은순</option>
			</select>
			<select name="arOpt">
				<option value="">거주지역별</option>
				<option value="">서울</option>
				<option value="">경기</option>
				<option value="">충청도</option>
				<option value="">강원도</option>
				<option value="">경상도</option>
				<option value="">전라도</option>
				<option value="">제주도</option>
			</select>
			<select name="ordOpt">
				<option value="">구매금액별</option>
				<option value="">구매 많은순</option>
				<option value="">구매 적은순</option>
			</select>
			<select name="selOpt">
				<option value="">응모자전체</option>
				<option value="">당첨선택고개</option>
				<option value="">당첨확정고객</option>
				<option value="">당첨안된고객</option>
			</select>
			<input type="button" class="button" value="검색" onclick="">
			<input type="text" class="button" value="" >명
        </td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 테이블 상단 검색바 끝 -->
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center"><input type="checkbox" name="chk"></td>
		<td width="70" align="center">등록일</td>
		<td width="110" align="center">아이디</td>
		<td align="center">코멘트내용</td>
		<td width="90" align="center">당첨/참여횟수</td>
		<td width="80" align="center">최근당첨일</td>
		<td width="80" align="center">구매금액<br>(3개월)</td>
		<td width="70" align="center">거주지역</td>
		<td width="70" align="center">가입일</td>
	</tr>
	<% if isArray(arrList) then %>
	<% for intLoop=0 to Ubound(arrList,2) %>
	<tr>
		<td bgcolor="#FFFFFF" align="center"><input type="checkbox" name="chk"></td>
		<td bgcolor="#FFFFFF" align="center"><%= formatdatetime(arrList(4,intLoop),2) %></td>
		<td bgcolor="#FFFFFF" align="center"><%= arrList(2,intLoop) %></td>
		<td bgcolor="#FFFFFF" style="cursor:pointer" onmousemove="showTXT('txt<%= intLoop %>');" onmouseover="showTXT('txt<%= intLoop %>');" onmouseout="hideTXT('txt<%= intLoop %>');"><%= left(db2html(arrList(3,intLoop)),35) %>..<div id="txt<%= intLoop %>" style="postion:absolute;display:none;"><%= nl2br(db2html(arrList(3,intLoop))) %></div></td>
		<td bgcolor="#FFFFFF" align="center">1/2</td>
		<td bgcolor="#FFFFFF" align="center">20070510</td>
		<td bgcolor="#FFFFFF" align="center">100,000</td>
		<td bgcolor="#FFFFFF" align="center">서울</td>
		<td bgcolor="#FFFFFF" align="center">20070505</td>
	</tr>

	<% next %>
	<% end if %>
	<tr>
		<td colspan="9">paging</td>
	</tr>
</table>
<div id="tool" style="position:absolute;display:none;">
<table width="450" height="50"border="0" cellpadding="3" cellspacing="0" class="a" style="border:1px solid #CCCCCC;" bgcolor="#FFFF96">
	<tr>
		<td valign="top" id="tooltd"></td>
	</tr>
</table>
</div>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->