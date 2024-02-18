<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_EntryList.asp
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

function optSelStr(byval optStr,byval OptVal)
	if (CStr(optStr) = CStr(OptVal)) then
		optSelStr="selected"
	end if
end function



dim evtCode,UserLevelOpt , SortingOpt, AreaOpt , OrderCashOpt, SelectingOpt

evtCode =request("eC")
UserLevelOpt = request("uLOpt")
AreaOpt  = request("arOpt")
SelectingOpt  = request("selOpt")
SortingOpt  = request("sortOpt")


dim Page,PageSize,ScrollCount,i
Page = request("Page")
if Page="" then Page=1
PageSize =30
ScrollCount = 20

dim Param

Param = "&eC=" & evtCode & "&uLOpt=" & UserLevelOpt & "&arOpt=" & AreaOpt & "&selOpt=" & SelectingOpt & "&sortOpt=" & SortingOpt

dim arrList,intLoop
dim appList

dim TotalCount ,TotalPage,SelUsrCnt


set appList = new ClsEventEntry
appList.FECode = evtCode
appList.FUserLevelOpt 	= UserLevelOpt
appList.FSortingOpt 	= SortingOpt
appList.FAreaOpt 		= AreaOpt
appList.FSelectingOpt 	= SelectingOpt
appList.FPageSize = PageSize
appList.FCurrPage = Page
'// 1차 선택 고객 수

SelUsrCnt = appList.fnGetSelectedIdCount
'//전체갯수,전체 페이지
arrList = appList.fnGetEntryListCount

TotalCount = arrList(0,0)
TotalPage  = arrList(1,0)
'//리스트

arrList = appList.fnGetEntryList
set appList = nothing
'     0             1         2        3              4
' c.evtcom_idx,c.evt_code,c.userid,c.evtcom_txt,c.evtcom_regdate
'      5          6         7      8        9          10          11          12         13         14
',g.userName,g.userlevel,g.age,g.sexflag,g.wincnt,g.entrycnt,g.lastWindate,g.orderSum,g.address,g.joinDate
%>

<script language='javascript'>

function AnSelectAllChk(bool){
	var frm = document.getElementsByName('cksel');
	for (var i=0;i<frm.length;i++){
		if (frm[i].disabled!=true){
			frm[i].checked = bool;
			AnCheckClick(frm[i]);

		}
	}

	SelCounting();
}

function SelCounting(frm){

	var sel = document.getElementById('selectCnt');

	var frm = document.getElementsByName('cksel');

	var cnt =0 ;
	for(i=0;i<frm.length;i++){
		if(frm[i].checked){
			cnt = cnt + 1;
		}
	}
	sel.value = cnt;


}

function showTXT(divVal){

	var mx = document.body.scrollLeft + event.clientX+10;
	var my = document.body.scrollTop + event.clientY -40;

	var vDIV = document.getElementById(divVal);

	var iTooltd = document.getElementById("tooltd");
		iTooltd.innerHTML = vDIV.innerHTML;

	var iTool = document.getElementById("tool");

		iTool.style.left=mx;
		iTool.style.top=my;
		iTool.style.display="";


	//setTimeout(showTXT(divVal),10000);
}

function hideTXT(){

	var iTool = document.getElementById("tool");
	iTool.style.display="none";

}

function showUserinfo(sc,mc,wc,fc,ec,age,sex,nm){

	var mx = document.body.scrollLeft + event.clientX+10;
	var my = document.body.scrollTop + event.clientY -40;


	var iTool 	= document.getElementById("tool");
	var iToolTD = document.getElementById("tooltd");

	var iToolID = document.getElementById("toolid");

	iToolTD.innerHTML = iToolID.innerHTML ;

	iToolTD.innerHTML = iToolTD.innerHTML.replace("##SCOUNT##",sc);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##MCOUNT##",mc);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##WCOUNT##",wc);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##FCOUNT##",fc);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##ECOUNT##",ec);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##AGE##",age);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##SEX##",sex);
	iToolTD.innerHTML = iToolTD.innerHTML.replace("##NAME##",nm);

		iTool.style.left=mx;
		iTool.style.top=my;
		iTool.style.display="";

}

function checkedValue(){
	var tgvalue="";
	var chkbx = document.getElementsByName('cksel');

	for (var i=0;i<chkbx.length;i++) {
		if (chkbx[i].checked){
			tgvalue=tgvalue  + chkbx[i].value + ",";
		}
	}

	if (tgvalue.length < 1){
		alert('하나 이상 응모자를 선택해 주세요');
		return '';
	}else{
		return tgvalue;
	}
}

function selEntry(strSel){

	var arridx = checkedValue();

	if (arridx.length < 1){
		return;

	} else {

		var conf;

		if (strSel=='S'){
			conf = confirm('선택 고객을 당첨자로 선정합니다.');
		} else if(strSel=='N'){
			conf = confirm('선택 고객을 제외합니다.');
		} else if (strSel=='C'){
			conf = confirm('선택 고객을 이벤트 당첨자로 확정 합니다.');
		}

		if (conf){
			selFrm.arridx.value = arridx;
			selFrm.selStr.value = strSel;
			//window.open("event_entry_process.asp", "pop", "width=10,height=10,scrollbars=no,status=no,resizable=yes");
			//selFrm.target="pop";
			selFrm.submit();
		}
	}
}
function fnGoSelectedList(){
	document.location.href="event_confirmList.asp?eC=<%= evtCode %>";
}
function fnSearch(){
	document.ListFrm.submit();
}
</script>

<!-- 테이블 상단 검색바 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="ListFrm" method="get" action="">
	<input type="hidden" name="eC" value="<%= evtCode %>">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<select name="selOpt">
				<option value="" <%=optSelStr(SelectingOpt,"") %>>응모자전체</option>
				<option value="S" <%=optSelStr(SelectingOpt,"S") %>>선택 고객</option>
				<option value="N" <%=optSelStr(SelectingOpt,"N") %>>선택 고객 제외</option>
			</select>

        	<select name="uLOpt">
				<option value="" <%=optSelStr(UserLevelOpt,"") %>>회원등급전체</option>
				<option value="3" <%=optSelStr(UserLevelOpt,"3") %>>VIP</option>
				<option value="2" <%=optSelStr(UserLevelOpt,"2") %>>블루</option>
				<option value="1" <%=optSelStr(UserLevelOpt,"1") %>>그린</option>
				<option value="0" <%=optSelStr(UserLevelOpt,"0") %>>옐로우</option>
				<option value="5" <%=optSelStr(UserLevelOpt,"5") %>>오렌지</option>
				<option value="9" <%=optSelStr(UserLevelOpt,"9") %>>매니아</option>
			</select>

			<select name="arOpt">
				<option value=""  <%=optSelStr(AreaOpt,"") %>>거주지 전체</option>
				<option value="서울" <%=optSelStr(AreaOpt,"서울") %>>서울</option>
				<option value="경기" <%=optSelStr(AreaOpt,"경기") %>>경기</option>
				<option value="충청" <%=optSelStr(AreaOpt,"충청") %>>충청도</option>
				<option value="강원" <%=optSelStr(AreaOpt,"강원") %>>강원도</option>
				<option value="경상" <%=optSelStr(AreaOpt,"경상") %>>경상도</option>
				<option value="전라" <%=optSelStr(AreaOpt,"전라") %>>전라도</option>
				<option value="제주" <%=optSelStr(AreaOpt,"제주") %>>제주도</option>
			</select>

			<select name="sortOpt">
				<option value=""  <%=optSelStr(SortingOpt,"") %>>정렬선택</option>
				<option value="cL"  <%=optSelStr(SortingOpt,"cL") %>>당첨적은순</option>
				<option value="cH"  <%=optSelStr(SortingOpt,"cH") %>>당첨많은순</option>
				<option value="oL"  <%=optSelStr(SortingOpt,"oL") %>>구매 적은순</option>
				<option value="oH"  <%=optSelStr(SortingOpt,"oH") %>>구매 많은순</option>
			</select>


			<input type="button" class="button" value="검색" onclick="fnSearch();">
			현재 선택 고객수<input type="text" name="selectCnt" size="2" value="0" >명
			전체 선택 고객수 :<%= SelUsrCnt %>
        </td>
        <td align="right"><input type="button" class="button" value="당첨자 보기" onclick="fnGoSelectedList();"></td>

		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 테이블 상단 검색바 끝 -->
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center"><input type="checkbox" name="ckselm" onClick="AnSelectAllChk(this.checked);"></td>
		<td width="70" align="center">등록일</td>
		<td width="110" align="center">아이디</td>
		<td align="center">코멘트내용</td>
		<td width="90" align="center">당첨/참여횟수</td>
		<td width="80" align="center">최근당첨일</td>
		<td width="80" align="center">구매금액<br>(5개월)</td>
		<td width="70" align="center">거주지역</td>
		<td width="70" align="center">가입일</td>
	</tr>
	<% if isArray(arrList) then %>
	<% for intLoop=0 to Ubound(arrList,2) %>

	<% if arrList(15,intLoop)<>"" then %>
		<tr class="H" bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" value="<%= arrList(0,intLoop) %>" onClick="AnCheckClick(this);SelCounting(this);" checked></td>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" value="<%= arrList(0,intLoop) %>" onClick="AnCheckClick(this);SelCounting(this);"></td>
	<% end if %>

		<td align="center"><%= DateValue (arrList(4,intLoop)) %></td>
		<td align="center" onmouseover="showUserinfo(0,0,0,0,0,'<%= arrList(5,intLoop) %>','<%=arrList(8,intLoop)%>','<%=arrList(7,intLoop)%>');" onmouseout="hideTXT();">
			<%= GetUserLevelColorStr(arrList(6,intLoop),arrList(2,intLoop)) %>
		</td>
		<td onmousemove="showTXT('txt<%= intLoop %>');" onmouseover="showTXT('txt<%= intLoop %>');" onmouseout="hideTXT();"><%= DDotFormat(db2html(arrList(3,intLoop)),35) %><div id="txt<%= intLoop %>" style="postion:absolute;display:none;"><%= nl2br(db2html(arrList(3,intLoop))) %></div></td>
		<td align="center"><%= arrList(9,intLoop) %>/<%= arrList(10,intLoop) %></td>
		<td align="center"><%if not isnull(arrList(11,intLoop)) then response.write DateValue(arrList(11,intLoop)) %></td>
		<td align="center"><%= FormatNumber(arrList(12,intLoop),0) %></td>
		<td align="center"><%= left(arrList(13,intLoop),10) %></td>
		<td align="center"><%= DateValue(arrList(14,intLoop)) %></td>
	</tr>
	<% next %>
	<% end if %>
</table>

<!-- 하단 시작 -->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
				<tr>
					<td width="150" align="left">

						<input type="button" class="button" value="선택" onclick="selEntry('S');">&nbsp;&nbsp;&nbsp;
						<% if SelectingOpt="S" then %>
						<input type="button" class="button" value="삭제" onclick="selEntry('N');">&nbsp;&nbsp;&nbsp;
						<% end if %>
					</td>
					<td align="center">
						<% if ((page-1)\SCrollCount)+1 > 1 then %>
							<a href="?page=<%= ((page-1)\SCrollCount)-1 %><%= Param %>">[pre]</a>
						<% else %>
							[pre]
						<% end if %>

						<% for i=0 + ((page-1)\SCrollCount)+1 to ScrollCount + (page-1)\SCrollCount %>
							<% if i>Totalpage then Exit for %>
							<% if CStr(page)=CStr(i) then %>
							<font color="red">[<%= i %>]</font>
							<% else %>
							<a href="?page=<%= i %><%= Param %>">[<%= i %>]</a>
							<% end if %>
						<% next %>

						<% if TotalPage  > (page-1)\SCrollCount+1 + ScrollCount -1then %>
							<a href="?page=<%= i %><%= Param %>">[next]</a>
						<% else %>
							[next]
						<% end if %>
					</td>
					<td width="120" align="center">

					</td>
				</tr>
			</table>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<form name="selFrm" method="get" target="subFrame" action="event_entry_process.asp">
<input type="hidden" name="eC" value="<%= evtCode %>">
<input type="hidden" name="arridx" value="">
<input type="hidden" name="selStr" value="">
</form>

<iframe name="subFrame" src="" frameborder="0" width="300" height="100"></iframe>

<div id="tool" style="position:absolute;display:none;">
<table border="0" cellpadding="5" cellspacing="0" class="a" style="border:1px solid #CCCCCC;" bgcolor="#FFFF96">
	<tr>
		<td valign="top" align="left" id="tooltd"></td>
	</tr>
</table>
</div>

<div id="toolid" style="position:absolute;display:none;">
<ul>
<li>참여이벤트 당첨내역
	<ul type="circle">
		<li>쇼핑찬스 코멘트 이벤트 : ##SCOUNT## </li>
		<li>감성매니아 : ##MCOUNT##</li>
		<li>위클리코디네이터 : ##WCOUNT##</li>
		<li>핑거스 : ##FCOUNT##</li>
		<li>기타 참여이벤트 : ##ECOUNT##</li>
	</ul>
</li>
<li>나이 : ##AGE##</li>
<li>성별 : ##SEX##</li>
<li>이름 : ##NAME##</li>
</ul>
</div>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->