<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  세미나실 관리
' History : 2009.04.07 서동석 생성
'			2010.12.27 한용민 수정
'           2012.01.10 허진원 수정; 세미나실 정리/추가
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/shedulecls.asp"-->
<%
Dim rno,mode,tdate,ttime,roomid , osemi, query1, arrScheList
rno = request("rno")
mode = request("mode")
roomid = request("roomid")
tdate = request("tdate")
ttime = request("ttime")

Dim SPhoto, Sstartdate, Senddate, Scomment

If mode = "modify" Then
	Set osemi = new CSeminarRoomCalendar
		osemi.FReqNo = rno
		osemi.fnGetSchedule
		arrScheList = osemi.fnGetSchedule
		
		SPhoto = arrScheList(1,0)
		Sstartdate = arrScheList(2,0)
		Senddate = arrScheList(3,0)
		Scomment = arrScheList(4,0)
	Set osemi = Nothing
End If
%>
<script language='javascript'>
function SubmitForm(){
	if (document.SubmitFrm.req_photo_user.value == 0){
		alert('담당포토를 선택하세요.');
		document.SubmitFrm.req_photo_user.focus();
		return;
	}

	if (document.SubmitFrm.start_time.value.length < 1){
		alert('시간을 선택해주세요');
		document.SubmitFrm.start_time.focus();
		return;
	}

	if (document.SubmitFrm.start_time.value >= document.SubmitFrm.end_time.value)
	{
		alert('촬영일을 확인해 주세요.');
		return;
	}
	if (document.SubmitFrm.req_comment.value == "")
	{
		alert('내용을 입력하세요');
		document.SubmitFrm.req_comment.focus();
		return;
	}

	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}
</script>
<table width="400" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center">
<form name="SubmitFrm" method="post" action="request_cal_proc.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="rno" value="<% = rno %>">
<tr>
	<td width="100">담당포토</td>
	<td>
		<select class="select" name="req_photo_user">
			<option value=''>-- 포토그래퍼 선택 --</option>
<%
	query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user with (nolock)"
	query1 = query1 + " where user_useyn = 'Y'"		' user_type='1'

	rsget.CursorLocation = adUseClient
	rsget.Open query1, dbget, adOpenForwardOnly, adLockReadOnly

	If  not rsget.EOF  Then
	rsget.Movefirst
		Do until rsget.EOF
		   response.write("<option value='"&rsget("user_id")& "' "& chkIIF(SPhoto=""&rsget("user_id")&"","selected","") &">" & rsget("user_name") & "" & "</option>")
		   rsget.MoveNext
		Loop
	End if
	rsget.close
%>
		</select>
	</td>
</tr>
<tr>
	<td width="100">시간</td>
	<td><input type="text" name="req_date" value="<% = tdate %>" class="input_b" size="10">일
		<select name="start_time">
			<option value="10:00">10:00</option>
			<option value="10:30">10:30</option>
			<option value="11:00">11:00</option>
			<option value="11:30">11:30</option>
			<option value="12:00">12:00</option>
			<option value="12:30">12:30</option>
			<option value="13:00">13:00</option>
			<option value="13:30">13:30</option>
			<option value="14:00">14:00</option>
			<option value="14:30">14:30</option>
			<option value="15:00">15:00</option>
			<option value="15:30">15:30</option>
			<option value="16:00">16:00</option>
			<option value="16:30">16:30</option>
			<option value="17:00">17:00</option>
			<option value="17:30">17:30</option>
		</select>~
		<select name="end_time">
			<option value="10:30">10:30</option>
			<option value="11:00">11:00</option>
			<option value="11:30">11:30</option>
			<option value="12:00">12:00</option>
			<option value="12:30">12:30</option>
			<option value="13:00">13:00</option>
			<option value="13:30">13:30</option>
			<option value="14:00">14:00</option>
			<option value="14:30">14:30</option>
			<option value="15:00">15:00</option>
			<option value="15:30">15:30</option>
			<option value="16:00">16:00</option>
			<option value="16:30">16:30</option>
			<option value="17:00">17:00</option>
			<option value="17:30">17:30</option>
			<option value="18:00">18:00</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">입력내용</td>
	<td bgcolor="#FFFFFF"><input type="text" name="req_comment" size="30" maxlength="128" class="text" value="<%=Scomment%>"></td>
</tr>
<% If mode = "modify" Then %>
<tr>
	<td width="100">삭제</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" value="Y">Y
		<input type="radio" name="useyn" value="N" checked>N
	</td>
</tr>
<%End If%>
<tr>
	<td colspan="2" align="center"><input type="button" value="저장" onClick="SubmitForm()"  class="button"></td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->