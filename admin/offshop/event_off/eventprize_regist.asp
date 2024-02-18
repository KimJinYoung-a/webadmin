<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  이벤트 당첨자 등록
' History : 2010.03.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_Cls.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->

<%
dim evt_code , chkdisp , evt_using , evt_kind , evt_name , evt_startdate ,evt_enddate
dim evt_state , evt_prizedate , opendate ,closedate , brand , isgift ,partMDid ,evt_forward
dim evt_comment , regdate , shopid , sEKindDesc ,sStateDesc
	evt_code= requestCheckVar(Request("evt_code"),10)	'이벤트코드	

	IF evt_code = "" THEN	'이벤트 코드값이 없을 경우 back
%>
		<script language="javascript">
		<!--
			alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
			history.back();
		//-->
		</script>
<%	dbget.close()	:	response.End
	END IF	
	
dim cEvtCont
set cEvtCont = new cevent_list
	cEvtCont.frectevt_code = evt_code	'이벤트 코드
	
	'//수정일경우에만 쿼리
	if evt_code <> "" then
		
	'이벤트 내용 가져오기	
	cEvtCont.fnGetEventCont_off
	evt_kind = cEvtCont.FOneItem.fevt_kind
	evt_name = cEvtCont.FOneItem.fevt_name
	evt_startdate = cEvtCont.FOneItem.Fevt_startdate
	evt_enddate = cEvtCont.FOneItem.Fevt_enddate
	evt_prizedate =	cEvtCont.FOneItem.Fevt_prizedate
	evt_state =	cEvtCont.FOneItem.Fevt_state
	IF datediff("d",now,evt_enddate) <0 THEN evt_state = 9 '기간 초과시 종료표기
	regdate	= cEvtCont.FOneItem.fevt_regdate
	evt_using = cEvtCont.FOneItem.Fevt_using
	shopid = cEvtCont.FOneItem.fshopid
	sEKindDesc = cEvtCont.FOneItem.fevt_kinddesc
	sStateDesc = cEvtCont.FOneItem.fevt_statedesc
	end if	
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td><!-- 당첨자 등록-->				
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨관리 : 한번 등록된 당첨자는 취소할 수 없습니다. 입력시 주의해 주세요</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
					<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_code%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_name%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sEKindDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트기간</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_startdate%> ~ <%=evt_enddate%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">상태</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sStateDesc%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=evt_prizedate%></td>
				</tr>			
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- #include virtual="/admin/offshop/event_off/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /당첨자 등록-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->