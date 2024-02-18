<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
	Dim cEvent, vEvtCode, vEvtSubj, vEvtStatus, vEvtRealStatus, vEvtSDate, vEvtEDate, vEvtMDName, vEvtIsSale, vEvtIsGift, vEvtSaleCnt, vEvtGiftCnt, vQuery
	vEvtCode = requestCheckVar(Request("evt_code"),6)
	
	IF vEvtCode <> "" Then
		If IsNumeric(vEvtCode) = True Then
			Set cEvent = new cOnlySys
			cEvent.FEvtcode = vEvtCode
			cEvent.fnEventCont
			
			vEvtSubj = cEvent.FEvtSubject
			vEvtStatus = cEvent.FEvtStatus
			vEvtRealStatus = cEvent.FEvtRealStatus
			vEvtSDate = cEvent.FEvtSDate
			vEvtEDate = cEvent.FEvtEDate
			vEvtMDName = cEvent.FEvtMDname
			vEvtIsSale = cEvent.FEvtIsSale
			vEvtIsGift = cEvent.FEvtIsGift
			vEvtSaleCnt = cEvent.FEvtSaleCnt
			vEvtGiftCnt = cEvent.FEvtGiftCnt
			Set cEvent = Nothing
		End If
	End IF

	vQuery = "select * from" & vbCrLf
	vQuery = vQuery & "db_event.dbo.tbl_event" & vbCrLf
	vQuery = vQuery & "where evt_code = '" & vEvtCode & "'" & vbCrLf & vbCrLf
	vQuery = vQuery & "--update db_event.dbo.tbl_event set" & vbCrLf
	vQuery = vQuery & "evt_state = '', evt_enddate = ''" & vbCrLf
	vQuery = vQuery & "where evt_code = '" & vEvtCode & "'" & vbCrLf
%>

<script language="javascript">
function jsEventSearch()
{
	if(frm1.evt_code.value == "")
	{
		alert("이벤트 코드 입력하세요.");
		frm1.evt_code.focus();
		return;
	}
	if(isNaN(frm1.evt_code.value))
	{
		alert("이벤트 코드를 숫자로만 입력하세요.");
		frm1.evt_code.value = "";
		frm1.evt_code.focus();
		return;
	}
	frm1.submit();
}
function jsEventUpdate()
{
	if(frm1.evt_code.value == "")
	{
		alert("이벤트 코드 입력하세요.");
		frm1.evt_code.focus();
		return;
	}
	if(isNaN(frm1.evt_code.value))
	{
		alert("이벤트 코드를 숫자로만 입력하세요.");
		frm1.evt_code.value = "";
		frm1.evt_code.focus();
		return;
	}
	if(frm1.evt_status.value != "" && isNaN(frm1.evt_status.value))
	{
		alert("이벤트 진행상태값을 숫자로만 입력하세요.");
		frm1.evt_status.value = "";
		frm1.evt_status.focus();
		return;
	}
	if(frm1.evt_edate.value == "")
	{
		alert("이벤트 종료일을 선택하세요.");
		return;
	}
	
	if(confirm("이대로 진행하시겠습니까?") == true) {
		frm1.method = "post";
		frm1.action = "event_date_update_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				이벤트 코드 : <input type="text" name="evt_code" value="<%=vEvtCode%>" maxlength="5">
				<input type="button" class="button" value="검 색" onClick="jsEventSearch()">
			</td>
		</tr>
		<% If vEvtCode <> "" Then %>
		<tr>
			<td style="padding:15 10 0 0;">
				<b>진행상태값</b> : <input type="text" name="evt_status" value="" maxlength="1" size="3">(수정할때만 입력. <u>필요없을시 공란.</u> 종료된것을 다시 진행하려면 7)&nbsp;
				<b>종료일</b> : <input type="text" name="evt_edate" size="10" maxlength=10 readonly value="">
				<a href="javascript:calendarOpen(frm1.evt_edate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
			</td>
			<td style="padding:15 0 0 0;">
				<div id="btn" style="display:block;"><input type="button" value="바로변경하기" onClick="jsEventUpdate()"></div>
			</td>
		</tr>
		<% End If %>
		</table>
		</form>
	</td>
</tr>
</table>

<% If vEvtSubj <> "" Then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td nowrap>이벤트코드</td>
  	<td nowrap>진행상태</td>
  	<td nowrap>이벤트명</td>
  	<td width="60">시작일</td>
  	<td width="60">종료일</td>
  	<td nowrap>담당MD</td>
  	<td nowrap></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=vEvtCode%></td>
	<td><%=vEvtRealStatus%> (<%=fnGetCommCodeArrDesc(fnSetCommonCodeArr("eventstate",False),vEvtStatus)%>)</td>
	<td><%=vEvtSubj%></td>
	<td><%=vEvtSDate%></td>
	<td><%=vEvtEDate%></td>
	<td><%=vEvtMDName%></td>
	<td>
		<%if vEvtIsSale then%>할인(<%=vEvtSaleCnt%>)<br><font color="red">할인이 걸려있습니다. 따로 처리하세요.</font><script language="javascript">document.getElementById("btn").style.display="none";</script><%end if%>&nbsp;
		<%if vEvtIsGift then%>사은품(<%=vEvtGiftCnt%>)<br><font color="red">할인이 걸려있습니다. 따로 처리하세요.</font><script language="javascript">document.getElementById("btn").style.display="none";</script><%end if%>
	</td>
</tr>
</table>
<br><br>* 쿼리구문<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->