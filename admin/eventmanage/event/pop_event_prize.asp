<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event
' History : 2008.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<%
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype,estatedesc, ekinddesc, prizeyn
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind

	egKindCode = request("evt_code")


dim oip, i

	'--이벤트 개요
	set oip = new ClsEvent
		oip.FECode = egKindCode	'이벤트 코드

		oip.fnGetEventCont	 '이벤트 내용 가져오기
		ekind 		=	oip.FEKind
		ekinddesc	=	oip.FEKindDesc
		eman 		=	oip.FEManager
		escope 		=	oip.FEScope
		ename 		=	db2html(oip.FEName)
		esday 		=	oip.FESDay
		eeday 		=	oip.FEEDay
		epday 		=	oip.FEPDay
		elevel 		=	oip.FELevel
		estate 		=	oip.FEState
		estatedesc 	= oip.FEStateDesc
		eregdate 	=	oip.FERegdate
		prizeyn		=	oip.FPrizeYN
	set oip = nothing
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frm" action="/common/event_prize_process.asp" method="get">
<input type="hidden" name="egKindCode" value="egKindCode">
<input type="hidden" name="eCode" value="<%=egKindCode%>">
<tr>
	<td><!-- 당첨자 등록-->
	<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨관리 : 한번 등록된 당첨자는 취소할 수 없습니다. 입력시 주의해 주세요</span><br>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
			<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=egKindCode%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
		</tr>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ekinddesc%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
		</tr>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">기간</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
		</tr>
		<% If prizeyn = "N" Then %>
		<tr height="25">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨자등록</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5" colspan="3">
				<input type="button" class="button" value="당첨자없음" onclick="frm.submit();">
				* 당첨자가 없는 경우에만 입력하세요.
			</td>
		</tr>
		<% End If %>
		</table>
	</td>
</tr>
</form>

<!-- /당첨자 등록-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
