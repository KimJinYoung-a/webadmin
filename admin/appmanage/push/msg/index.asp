<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' PageName : index.asp
' Description : 푸시 메시지
' Hieditor : 2014.05.08 이종화 생성
'			 2016.06.15 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
Dim page, i, opush, reload, pkey, ckey, isusing , state , reservationdate, pushtitle, pushurl, targetKey
dim repeatpushyn, idx, repeatidx, resetyn
	menupos = requestcheckvar(request("menupos"),10)
	page = requestcheckvar(request("page"),10)
	reload = requestcheckvar(request("reload"),2)
	isusing = requestcheckvar(request("isusing"),1)
	state = requestcheckvar(request("state"),1)
	reservationdate = requestcheckvar(request("reservationdate"),10)
	pushtitle = requestcheckvar(request("pushtitle"),300)
	pushurl = requestcheckvar(request("pushurl"),600)
	targetKey = requestcheckvar(request("targetKey"),10)
	repeatpushyn = requestcheckvar(request("repeatpushyn"),1)
	idx = requestcheckvar(request("idx"),10)
	repeatidx = requestcheckvar(request("repeatidx"),10)
	resetyn = requestcheckvar(request("resetyn"),1)

if page = "" then page = 1
if repeatpushyn="" then repeatpushyn="N"
if reload="" and isusing="" then isusing="Y" ''사용중 기본

' 재검색일경우 푸시구분을 변경했을경우 반복푸시번호와 발송타켓을 리셋시킴
if resetyn="Y" then
	repeatidx=""
	targetKey=""
end if

set opush = new cpush_msg_list
	opush.FPageSize = 50
	opush.FCurrPage = page
	opush.Frectdate = reservationdate
	opush.Fstate = state
	opush.Fisusing = isusing
	opush.Frectpushtitle = pushtitle
	opush.Frectpushurl = pushurl
	opush.FrecttargetKey = targetKey
	opush.Frectrepeatpushyn = repeatpushyn
	opush.frectidx = idx
	opush.frectrepeatidx = repeatidx
	opush.fpushmsglist()
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	<% ' 재검색일경우 푸시구분을 변경했을경우 푸시번호와 발송타켓을 리셋시킴 %>
	if ( frm.repeatpushyn.value!='<%=repeatpushyn%>' ){
		frm.resetyn.value='Y';
	}

	frm.page.value=page;
	frm.submit();
}

//예약등록
function AddNewContents(idx){
	var poppushin;
	poppushin = window.open('/admin/appmanage/push/msg/poppushmsg_edit.asp?idx='+ idx,'poppushin','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushin.focus();
}

//푸쉬메시지테스트발송
function example_msg(idx, repeatpushyn){
	var poppushexam;
	poppushexam = window.open('/admin/appmanage/push/msg/poppushmsg_example.asp?idx='+ idx + '&repeatpushyn=' + repeatpushyn,'poppushexam','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushexam.focus();
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function popPushReport(iidx){
	var poppushreport;
	poppushreport = window.open('/admin/appmanage/push/msg/poppushmsg_report.asp?idx='+ iidx,'poppushreport','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushreport.focus();
}

// 타켓쿼리관리
function targetqueryreg(){
	var poppushtarget;
	poppushtarget = window.open('/admin/appmanage/push/msg/pushtarget.asp','poppushtarget','width=1600,height=800,scrollbars=yes,resizable=yes');
	poppushtarget.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<input type="hidden" name="reload" value="ON">
<input type="hidden" name="resetyn" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 푸시구분 : <% Drawpushgubun "repeatpushyn", repeatpushyn, " onchange='frmsubmit("""");'", "" %>
		<% if repeatpushyn="N" then %>
			&nbsp;&nbsp;
			* 발송타켓 : <% drawSelectBoxTarget "targetKey", targetKey, " onchange='frmsubmit("""");'", repeatpushyn, "" %>
		<% else %>
			&nbsp;&nbsp;
			* 발송종류 : <% drawSelectBoxTarget "targetKey", targetKey, " onchange='frmsubmit("""");'", repeatpushyn, "Y" %>
		<% end if %>
		&nbsp;&nbsp;
		* 푸시제목 : <input type="text" name="pushtitle" value="<%= pushtitle %>" size=25 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		&nbsp;&nbsp;
		* 푸시번호 : <input type="text" name="idx" value="<%= idx %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		<br><br>
		* 발송일 : 
		<input type="text" id="termSdt" name="reservationdate" size="7" maxlength=10 value="<%= reservationdate %>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;" />
		<img src="/images/admin_calendar.png" alt="달력으로 검색" id="ChkStart_trigger" onclick="return false;" />
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "termSdt", trigger    : "ChkStart_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					//CAL_End.args.min = date;
					//CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d" <%=chkIIF(reservationdate<>"",", max: " & replace(reservationdate,"-",""),"")%>
			});
		</script>
		&nbsp;&nbsp;
		* 링크 : <input type="text" name="pushurl" value="<%= pushurl %>" size=25 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		&nbsp;&nbsp;
		* 사용여부 : 
		<% drawSelectBoxisusingYN "isusing",isusing, " onchange='frmsubmit("""");'" %>
		&nbsp;&nbsp;
		* 상태
		<% Drawpushstatename "state" , state , " onchange='frmsubmit("""");'", "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('');">
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
	</td>
	<td align="right">
		<input type="button" class="button" value="신규등록" onclick="AddNewContents('0');">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="타켓쿼리관리" onclick="targetqueryreg();">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= opush.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= opush.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=55>푸시구분</td>
	<td width=70>번호</td>
	<td width=145>발송일</td>	
	<td width=210>제목</td>
	<td>링크</td>
	<td width=70>상태</td>	
	<td width=40>사용<br>여부</td>
	<% '<td width=60>개인화<br>푸시<br>여부</td> %>
	<td width=70>타겟상태</td>
	<td width=200>타겟</td>
	<!--<td width=50>타겟메인</td>-->
	<td width=70>타겟수량</td>
	<td width=40>이미지</td>
	<!--<td width=90>최초등록</td>-->
	<!--<td width=90>마지막수정</td>-->
	<td width=160>비고</td>
</tr>
<% if opush.FresultCount>0 then %>
    <% for i=0 to opush.FresultCount-1 %>
    <% 
	ckey = CStr(CHKIIF(isNULL(opush.FItemList(i).fbaseIdx),"",opush.FItemList(i).fbaseIdx))
	if (ckey="") then ckey=CStr(opush.FItemList(i).Fidx)
    %>
    <% if (pkey<>"") and (pkey<>ckey) then %>
		<tr align="center" bgcolor="cccccc" ><td colspan="12"><%=pkey%>,<%=ckey%></td></tr>
    <% end if %>
    
    <% if (opush.FItemList(i).fisusing="N") then %>
		<tr align="center" bgcolor="cccccc" >    
    <% else %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background="#FFFFFF";>
    <% end if %>

    	<td>
    		<%= Selectpushgubunname(opush.FItemList(i).frepeatpushyn) %>
    	</td>
    	<td>
    		<%= opush.FItemList(i).Fidx %>
    	</td>
    	<td>
			<%= dateconvert(opush.FItemList(i).Freservedate) %>
    	</td>
    	<td align="left">
    		<%= chrbyte(opush.FItemList(i).fpushtitle,20,"Y") %>
    	</td>
    	<td align="left">
    		<%= opush.FItemList(i).fpushurl %>
    	</td>
    	<td>
    		<%= pushmsgstate(opush.FItemList(i).fstate)%>
    	</td>
    	<td>
    		<%= opush.FItemList(i).fisusing %>
    	</td>
    	<!--<td><%'= opush.FItemList(i).fprivateYN %></td>-->
    	<td><% if (opush.FItemList(i).fistargetMsg=1) then %><%=opush.FItemList(i).getTargetStateName%><% end if %></td>
    	<td>
    	    <% if (opush.FItemList(i).fistargetMsg=1) then %>
    	        <%= opush.FItemList(i).ftargetName %>
    	    <% else %>
				전체
			<% end if %>
    	</td>
    	<!--<td><%=opush.FItemList(i).fbaseIdx%></td>-->
    	<td><%=FormatNumber(opush.FItemList(i).fmayTargetCnt,0)%></td>
		<td>
			<% if opush.FItemList(i).fpushimg<>"" and not(isnull(opush.FItemList(i).fpushimg)) then %>
				<img src="<%=opush.FItemList(i).fpushimg%>" width=40 height=40>
			<% end if %>
		</td>
		<!--<td>-->
			<%'= left(opush.FItemList(i).fregdate,10) %>
			<!--<br><%'= mid(opush.FItemList(i).fregdate,12,11) %>-->
			<!--<br><%'=opush.FItemList(i).fregadminid%>-->
		<!--</td>-->
		<!--<td>-->
			<%'= left(opush.FItemList(i).flastupdate,10) %>
			<!--<br><%'= mid(opush.FItemList(i).flastupdate,12,11) %>-->
			<!--<br><%'=opush.FItemList(i).flastadminid%>-->
		<!--</td>-->
    	<td>
			<input type="button" value="수정" onclick="AddNewContents('<%= opush.FItemList(i).Fidx %>');" class="button" />
    		<input type="button" value="테스트(<%= opush.FItemList(i).ftestpush %>건)" onclick="example_msg(<%= opush.FItemList(i).Fidx %>,'N');" class="button" />
			<input type="button" value="통계" onClick="popPushReport('<%= opush.FItemList(i).Fidx %>');" class="button" />
    	</td>
    </tr>
    <% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if opush.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= opush.StartScrollPage-1 %>')">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + opush.StartScrollPage to opush.StartScrollPage + opush.FScrollCount - 1 %>
				<% if (i > opush.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(opush.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if opush.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
session.codePage = 949
set opush = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->