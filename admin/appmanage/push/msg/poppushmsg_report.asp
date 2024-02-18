<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 푸시 메시지 통계
' Hieditor : 2014.05.08 이종화 생성
'			 2019.06.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->

<%
Dim cAppPushInfo, cAppPushReport, vIdx, vArr, i, vPushTargetKey, vPushReviewCount
dim ttlCnt, waitCnt, sentCnt, clickCnt, succCnt, failCnt, diffnomuts, diffseconds, reservationdate
dim repeatpushyn, onepushdispyn, page, pushtitle, pushurl, targetKey, appkey, resetyn, reload
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	vIdx	= requestcheckvar(getNumeric(request("idx")),10)
	repeatpushyn = requestcheckvar(request("repeatpushyn"),1)
	reservationdate = requestcheckvar(request("reservationdate"),10)
	pushtitle = requestcheckvar(request("pushtitle"),300)
	pushurl = requestcheckvar(request("pushurl"),600)
	targetKey = requestcheckvar(request("targetKey"),10)
	appkey = requestcheckvar(getNumeric(request("appkey")),10)
	reload = requestcheckvar(request("reload"),2)
	resetyn = requestcheckvar(request("resetyn"),1)

if page = "" then page = 1
onepushdispyn="N"
if repeatpushyn="" then repeatpushyn="N"

' 재검색일경우 푸시구분을 변경했을경우 푸시번호와 발송타켓을 리셋시킴
if resetyn="Y" then
	vidx=""
	targetKey=""
end if

if repeatpushyn="Y" then

else
	if vIdx<>"" and not(isnull(vIdx)) then
		onepushdispyn="Y"
	end if
end if

ttlCnt=0 : waitCnt=0 : sentCnt=0 : clickCnt=0 : succCnt=0 : failCnt=0 : diffnomuts=0 : diffseconds=0

Set cAppPushReport = New cpush_msg_list
	cAppPushReport.FPageSize = 100
	cAppPushReport.FCurrPage = page
	cAppPushReport.Frectidx = vIdx
	cAppPushReport.Frectdate = reservationdate
	cAppPushReport.Frectpushtitle = pushtitle
	cAppPushReport.Frectpushurl = pushurl
	cAppPushReport.FrecttargetKey = targetKey
	cAppPushReport.Frectrepeatpushyn = repeatpushyn
	cAppPushReport.Frectappkey = appkey

' 반복푸시 아닌거
if repeatpushyn="N" then
	' 한개푸시만
	if onepushdispyn="Y" then
		Set cAppPushInfo = New cpush_msg_list
			cAppPushInfo.Frectidx = vIdx
			cAppPushInfo.FPageSize = 1
			cAppPushInfo.FCurrPage = 1
			cAppPushInfo.fpushmsglist()		'정태훈 추가 
			vPushTargetKey = cAppPushInfo.FItemList(0).ftargetKey
			vPushReviewCount = cAppPushInfo.FItemList(0).freviewCount
		Set cAppPushInfo = Nothing

		cAppPushReport.fpushmessage_report
	else
		cAppPushReport.fpushsummary_report
	end if

' 반복푸시 일경우
else
	cAppPushReport.fpushsummary_Repeat_report
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function frmsubmit(page){
	<% if reload="ON" then %>
		<% ' 재검색일경우 푸시구분을 변경했을경우 푸시번호와 발송타켓을 리셋시킴 %>
		if ( frm.repeatpushyn.value!='<%=repeatpushyn%>' ){
			frm.resetyn.value='Y';
		}
	<% end if %>

	frm.page.value=page;
	frm.submit();
}

function AddNewContents(idx){
	var poppushin;
	poppushin = window.open('/admin/appmanage/push/msg/poppushmsg_edit.asp?idx='+ idx,'poppushin','width=1280,height=960,scrollbars=yes,resizable=yes');
	poppushin.focus();
}

function fnpush_clickdetail(multipskey, targetkey, repeatpushyn, appkey){
	var popclickdetail;
	popclickdetail = window.open('/admin/appmanage/push/msg/poppushmsg_report_clickdetail.asp?multipskey='+ multipskey + '&targetkey=' + targetkey + '&repeatpushyn=' + repeatpushyn + '&appkey=' + appkey + '&reload=ON&menupos=<%=menupos%>','popclickdetail','width=1600,height=960,scrollbars=yes,resizable=yes');
	popclickdetail.focus();
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
			<% if onepushdispyn="N" then %>
				&nbsp;&nbsp;
				* 발송종류 : <% drawSelectBoxTarget "targetKey", targetKey, " onchange='frmsubmit("""");'", repeatpushyn, "Y" %>
			<% end if %>

			&nbsp;&nbsp;
			* 푸시번호 : <input type="text" name="idx" value="<%= vidx %>" size=8 maxlength=10 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
		<% else %>
			&nbsp;&nbsp;
			* 발송종류 : <% drawSelectBoxTarget "targetKey", targetKey, " onchange='frmsubmit("""");'", repeatpushyn, "" %>
			<!--&nbsp;&nbsp;-->
			<!--* 발송종류 : --><%' drawSelectBoxrepeatpush "idx", vidx, " onchange='frmsubmit("""");'" %>
		<% end if %>

		<% if onepushdispyn="N" then %>
			&nbsp;&nbsp;
			* 푸시제목 : <input type="text" name="pushtitle" value="<%= pushtitle %>" size=25 onKeyPress="if(window.event.keyCode==13) frmsubmit('');" >
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
			* OS : <% Drawpushappkeyname "appkey", appkey, " onchange='frmsubmit("""");'" %>
		<% else %>
			<input type="hidden" name="targetKey" value="<%= targetKey %>" >
		<% end if %>
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
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= cAppPushReport.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= cAppPushReport.FTotalPage %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td width=55>푸시구분</td>
	<% if repeatpushyn="N" then %><td width=70>푸시번호</td><% end if %>
	<td width=145>발송일</td>
	<td>푸시제목</td>
	<td width=100>OS</td>
	<td>발송종류</td>
    <td width=60>총 발송<br>A</td>
    <td width=60>대기<br>B</td>
    <td width=60>발송완료<br>C</td>
    <td width=60>실패<br>D</td>
    <td width=60>기타오류<br>E</td>
    <td width=60>성공<br>F</td>
    <td width=60>클릭수<br>G</td>
    <td width=50>
		클릭률
		<br>G / F
	</td>    

	<%
	if vPushTargetKey="99995" then		'정태훈 추가 
	%>
		<td width=60>리뷰 작성 수</td>
	<% end if %>

    <td width=145>첫발송시간</td>
    <td width=145>마지막<br>발송시간</td>
    <td width=50>걸린시간<Br>(분,초)</td>
    <td width=50>발송속도<Br>(통/분)</td>
    <td width=50>예상시간<Br>(분)</td>
</tr>

<% if cAppPushReport.FresultCount>0 then %>
	<% for i=0 to cAppPushReport.FresultCount-1 %>
	<%
	ttlCnt = ttlCnt + cAppPushReport.FItemList(i).fttlCnt
	waitCnt = waitCnt + cAppPushReport.FItemList(i).fwaitCnt
	sentCnt = sentCnt + cAppPushReport.FItemList(i).fsentCnt
	clickCnt = clickCnt + cAppPushReport.FItemList(i).fclickCnt
	succCnt = succCnt + cAppPushReport.FItemList(i).fsuccCnt
	failCnt = failCnt + cAppPushReport.FItemList(i).ffailCnt
	diffnomuts = diffnomuts + cAppPushReport.FItemList(i).fdiffnomuts
	diffseconds = diffseconds + cAppPushReport.FItemList(i).fdiffseconds
	%>
		<tr bgcolor="#FFFFFF" align="center">
            <td><%= Selectpushgubunname(repeatpushyn) %></td>

			<% if repeatpushyn="N" then %>
				<td>
					<% if repeatpushyn="N" or repeatpushyn="" or isnull(repeatpushyn) then %>
						<%= cAppPushReport.FItemList(i).fmultiPskey %>
					<% else %>
						<%= cAppPushReport.FItemList(i).frepeatidx %>
					<% end if %>
				</td>
			<% end if %>

			<td>
				<%= dateconvert(cAppPushReport.FItemList(i).Freservedate) %>
			</td>
			<td align="left">
				<%= chrbyte(cAppPushReport.FItemList(i).fpushtitle,30,"Y") %>
			</td>
			<td><%= Selectappname(cAppPushReport.FItemList(i).fappkey) %></td>
			<td>
				<% if repeatpushyn="N" then %>

						<% if (cAppPushReport.FItemList(i).fistargetMsg=1) then %>
							<%= cAppPushReport.FItemList(i).ftargetName %>
						<% else %>
							회원전체
						<% end if %>		

				<% else %>
					<%= cAppPushReport.FItemList(i).ftargetName %>
				<% end if %>
			</td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).fttlCnt,0)%></td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).fwaitCnt,0)%></td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).fsentCnt,0)%></td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).ffailCnt,0)%></td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).fttlCnt-cAppPushReport.FItemList(i).fwaitCnt-cAppPushReport.FItemList(i).fsuccCnt-cAppPushReport.FItemList(i).ffailCnt,0)%></td>
			<td><%=FormatNumber(cAppPushReport.FItemList(i).fsuccCnt,0)%></td>
			<td>
				<!--<a href="#" onclick="fnpush_clickdetail('<%'= cAppPushReport.FItemList(i).fmultiPskey %>','<%'= cAppPushReport.FItemList(i).ftargetkey %>','<%'= repeatpushyn %>','<%'= cAppPushReport.FItemList(i).fappkey %>'); return false;">-->
				<%=FormatNumber(cAppPushReport.FItemList(i).fclickCnt,0)%>
				<!--</a>-->
			</td>

			<% if cAppPushReport.FItemList(i).fsuccCnt<>0 then %>
				<td><%=FormatPercent(cAppPushReport.FItemList(i).fclickCnt/cAppPushReport.FItemList(i).fsuccCnt,1)%><!--100%--></td>		
			<% else %>
				<td></td>
			<% end if %>
			<%
			if vPushTargetKey="99995" then		'정태훈 추가 
			%>
				<td></td>
			<% end if %>
			<td>
				<%= dateconvert(cAppPushReport.FItemList(i).ffirstSentDate) %>
			</td>
			<td>
				<%= dateconvert(cAppPushReport.FItemList(i).flastSentDate) %>
			</td>
			<td><%= cAppPushReport.FItemList(i).fdiffnomuts %> (<%= cAppPushReport.FItemList(i).fdiffseconds %>)</td>
			<td>
				<% if (cAppPushReport.FItemList(i).fdiffseconds<>0) then %>
					<%= FormatNumber(cAppPushReport.FItemList(i).fsentCnt /cAppPushReport.FItemList(i).fdiffseconds*60,0) %>
				<% end if %>
			</td>
			<td>
				<% if (cAppPushReport.FItemList(i).fdiffnomuts<>0) then %>
					<%= FormatNumber(cAppPushReport.FItemList(i).fttlCnt/(cAppPushReport.FItemList(i).fsentCnt/cAppPushReport.FItemList(i).fdiffnomuts),0) %>
				<% end if %>
			</td>
		</tr>
	<% Next %>

	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>합계</td>
		<td colspan=<% if repeatpushyn="N" then %>5<% else %>4<% end if %>>
		<td><%= FormatNumber(ttlCnt,0) %></td>
		<td><%= FormatNumber(waitCnt,0) %></td>
		<td><%= FormatNumber(sentCnt,0) %></td>
		<td><%= FormatNumber(failCnt,0) %></td>
		<td><%= FormatNumber(ttlCnt-waitCnt-succCnt-failCnt,0) %></td>
		<td><%= FormatNumber(succCnt,0) %></td>
		<td><%= FormatNumber(clickCnt,0) %></td>
		<td>
			<% if clickCnt<>0 and sentCnt<>0 then %>
				<%= FormatPercent(clickCnt/sentCnt,1) %>
			<% else %>
				0
			<% end if %>
		</td>
			<%
			if vPushTargetKey="99995" then		'정태훈 추가 
			%>
				<td><%=FormatNumber(vPushReviewCount,0)%></td>
			<% end if %>
		<td></td>
		<td></td>
		<td><%= diffnomuts %> (<%= diffseconds %>)</td>
		<td>
			<% if (diffseconds<>0) then %>
				<%= FormatNumber(sentCnt /diffseconds*60,0) %>
			<% end if %>
		</td>
		<td>
			<% if (diffnomuts<>0) then %>
				<%= FormatNumber(ttlCnt/(sentCnt/diffnomuts),0) %>
			<% end if %>
		</td>
	</tr>

	<% if onepushdispyn="N" then %>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="25" align="center">
				<% if cAppPushReport.HasPreScroll then %>
					<span class="list_link"><a href="javascript:frmsubmit('<%= cAppPushReport.StartScrollPage-1 %>')">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + cAppPushReport.StartScrollPage to cAppPushReport.StartScrollPage + cAppPushReport.FScrollCount - 1 %>
					<% if (i > cAppPushReport.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(cAppPushReport.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if cAppPushReport.HasNextScroll then %>
					<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	<% end if %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% End If %>
</table>

<%
Set cAppPushReport = Nothing
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->