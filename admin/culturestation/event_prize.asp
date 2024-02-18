<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station Event  
' History : 2008.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->

<%
Dim eCode,egKindCode
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,stype
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
Dim strparm
Dim sStateDesc, sEKindDesc
Dim arrEvtStatus, arrEvtKind
	eCode = 4
	egKindCode = request("evt_code")

	
dim oip, i 
	set oip = new cevent_list
	oip.frectevt_code = request("evt_code")
	oip.fevent_oneitem()
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
<form name="frm" action="/common/event_prize_process.asp" method="get">
<input type="hidden" name="egKindCode" value="<%= oip.foneitem.fevt_code %>">
<input type="hidden" name="eCode" value=4>
<tr>
	<td><!-- 당첨자 등록-->				
	<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨관리 : 한번 등록된 당첨자는 취소할 수 없습니다. 입력시 주의해 주세요</span><br>		
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트코드</td>
				<td width="200" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= oip.foneitem.fevt_code %></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트명</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%= oip.foneitem.fevt_name %></td>
			</tr>	
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">종류</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">culturestation Event</td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">이벤트기간</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=oip.foneitem.fstartdate%> ~ <%=oip.foneitem.fenddate%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨 발표일</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=oip.foneitem.feventdate%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨자등록</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=oip.foneitem.fprizeyn%>
					<input type="button" class="button" value="수동Y로전환하기" onclick="frm.submit();">
				</td>
			</tr>			
		</table>
	</td>
</tr>
</form>
<tr>
		<td>
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
	</td>
</tr>	
<!-- /당첨자 등록-->
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

