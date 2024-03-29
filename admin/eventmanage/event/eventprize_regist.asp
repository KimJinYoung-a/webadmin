<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/eventprize_regist.asp
' Description :  이벤트 당첨자 등록
' History : 2007.02.13 정윤정 생성
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
'--------------------------------------------------------
' 변수선언
'--------------------------------------------------------
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind
eCode = Request("eC")
	
	IF eCode = "" THEN	'이벤트 코드값이 없을 경우 back
%>
		<script language="javascript">
		<!--
			alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
			history.back();
		//-->
		</script>
<%	dbget.close()	:	response.End
	END IF	
'--------------------------------------------------------
' 이벤트 데이터 가져오기
'--------------------------------------------------------
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	
	cEvtCont.fnGetEventCont	 '이벤트 내용 가져오기
	ekind		= cEvtCont.FEKind 
	sEKindDesc 	= cEvtCont.FEKindDesc 
	eman 		= cEvtCont.FEManager 
	escope 		= cEvtCont.FEScope 
	ename 		= db2html(cEvtCont.FEName)
	esday 		= cEvtCont.FESDay
	eeday 		= cEvtCont.FEEDay
	epday 		= cEvtCont.FEPDay
	elevel 		= cEvtCont.FELevel
	estate 		= cEvtCont.FEState
	sStateDesc  = cEvtCont.FEStateDesc
	eregdate 	= cEvtCont.FERegdate 
	set cEvtCont = nothing
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td><!-- 당첨자 등록-->				
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨관리 : 한번 등록된 당첨자는 취소할 수 없습니다. 입력시 주의해 주세요</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
					<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sEKindDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트기간</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%> ~ <%=eeday%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sStateDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
				</tr>			
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /당첨자 등록-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->