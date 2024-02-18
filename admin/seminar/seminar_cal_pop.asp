<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  세미나실 캘린더
' History : 2012.10.24 김진영 생성
'################################# 달력시작 #################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/seminar/seminarCls.asp"-->
<%
Dim idx,mode,tdate,ttime,roomid , osemi, query1, arrScheList
idx = request("idx")
mode = request("mode")
roomid = request("roomidx")
tdate = request("tdate")
ttime = request("ttime")

Dim Sroomidx,Sgroupname, Susepurpose, Susercell, Susesu, Setc, Slecnum, Sisusing, Sstart_date, Send_date
Dim SH, SM, EH, EM, StartHM, EndHM

If mode = "modify" Then
	Set osemi = new CSeminarRoomCalendar
		osemi.FIdx = idx
		osemi.fnGetSchedule
		arrScheList = osemi.fnGetSchedule
		Sroomidx	= arrScheList(1,0)
		Sstart_date	= arrScheList(2,0)
		Send_date	= arrScheList(3,0)
		Sgroupname	= arrScheList(4,0)
		Susepurpose = arrScheList(5,0)
		Susercell	= arrScheList(6,0)
		Susesu		= arrScheList(7,0)
		Setc		= arrScheList(8,0)
		Slecnum		= arrScheList(9,0)
		Sisusing	= arrScheList(10,0)

		SH = hour(Sstart_date)
		If SH < 10 Then
			SH = "0"&SH
		End If
		
		SM = minute(Sstart_date)
		If SM < 10 Then
			SM = "0"&SM
		End If
		StartHM = SH&":"&SM
		
		EH = hour(Send_date)
		If EH < 10 Then
			EH = "0"&EH
		End If
		
		EM = minute(Send_date)
		If EM < 10 Then
			EM = "0"&EM
		End If
		EndHM = EH&":"&EM
	Set osemi = Nothing
Else
	'신규 저장시 전송된 값 사용
	Sroomidx = roomid
	Sstart_date = tdate
	StartHM = UseTimeName(chkIIF(ttime>=35,"34",ttime))
	EndHM = UseTimeName(chkIIF(ttime>=35,"35",cStr(ttime+1)))
End If

Function UseTimeName(ttime)
	Select Case ttime
		Case "6"	UseTimeName = "09:00"
		Case "7"	UseTimeName = "09:30"
		Case "8"	UseTimeName = "10:00"
		Case "9"	UseTimeName = "10:30"
		Case "10"	UseTimeName = "11:00"
		Case "11"	UseTimeName = "11:30"
		Case "12"	UseTimeName = "12:00"
		Case "13"	UseTimeName = "12:30"
		Case "14"	UseTimeName = "13:00"
		Case "15"	UseTimeName = "13:30"
		Case "16"	UseTimeName = "14:00"
		Case "17"	UseTimeName = "14:30"
		Case "18"	UseTimeName = "15:00"
		Case "19"	UseTimeName = "15:30"
		Case "20"	UseTimeName = "16:00"
		Case "21"	UseTimeName = "16:30"
		Case "22"	UseTimeName = "17:00"
		Case "23"	UseTimeName = "17:30"
		Case "24"	UseTimeName = "18:00"
		Case "25"	UseTimeName = "18:30"
		Case "26"	UseTimeName = "19:00"
		Case "27"	UseTimeName = "19:30"
		Case "28"	UseTimeName = "20:00"
		Case "29"	UseTimeName = "20:30"
		Case "30"	UseTimeName = "21:00"
		Case "31"	UseTimeName = "21:30"
		Case "32"	UseTimeName = "22:00"
		Case "33"	UseTimeName = "22:30"
		Case "34"	UseTimeName = "23:00"
		Case "35"	UseTimeName = "23:30"
	End Select
End Function
%>
<script language='javascript'>
function SubmitForm(){
	if (document.SubmitFrm.roomidx.value == 0){
		alert('세미나실을 선택하세요.');
		document.SubmitFrm.roomidx.focus();
		return;
	}

	if (document.SubmitFrm.start_time.value.length < 1){
		alert('시간을 선택해주세요');
		document.SubmitFrm.start_time.focus();
		return;
	}

	if (document.SubmitFrm.start_time.value >= document.SubmitFrm.end_time.value)
	{
		alert('예약일을 확인해 주세요.');
		return;
	}
	if (document.SubmitFrm.groupname.value == "")
	{
		alert('그룹명을 입력하세요');
		document.SubmitFrm.groupname.focus();
		return;
	}
	if (document.SubmitFrm.useSu.value == "")
	{
		alert('사용인원을 입력하세요');
		document.SubmitFrm.useSu.focus();
		return;
	}
	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}
function onlyNumber(){
	if((event.keyCode<48)||(event.keyCode>57))
		event.returnValue=false;
}

</script>
<table width="400" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center">
<form name="SubmitFrm" method="post" action="seminar_cal_proc.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<tr>
	<td width="100">세미나실 선택</td>
	<td>
		<select class="select" name="roomidx">
			<option value=''>-- 세미나실 선택 --</option>
<%
	query1 = " select idx, roomname, MaxSu, orderNo, isusing from db_partner.dbo.tbl_seminarRoom "
	query1 = query1 + " where isusing='Y' Order by orderNo ASC "
	rsget.Open query1,dbget,1

	If  not rsget.EOF  Then
	rsget.Movefirst
		Do until rsget.EOF
%>
			<option value="<%=rsget("idx")%>" <% If cStr(Sroomidx) = cStr(rsget("idx")) Then response.write "selected" End If %>>
				<%=rsget("roomname")%>

				<% if rsget("roomname") <> "홀" then %>
					(<%=rsget("MaxSu")%>)
				<% end if %>
			</option>
<%
		   rsget.MoveNext
		Loop
	End if
	rsget.close
%>
		</select>
	</td>
</tr>
<tr>
	<td width="100">예약일</td>
	<td><input type="text" name="reserdate" value="<% = tdate %>" class="input_b" size="10">일
		<select name="start_time">
			<option value="09:00" <%=chkiif(StartHM = "09:00","selected","")%>>09:00</option>
			<option value="09:30" <%=chkiif(StartHM = "09:30","selected","")%>>09:30</option>
			<option value="10:00" <%=chkiif(StartHM = "10:00","selected","")%>>10:00</option>
			<option value="10:30" <%=chkiif(StartHM = "10:30","selected","")%>>10:30</option>
			<option value="11:00" <%=chkiif(StartHM = "11:00","selected","")%>>11:00</option>
			<option value="11:30" <%=chkiif(StartHM = "11:30","selected","")%>>11:30</option>
			<option value="12:00" <%=chkiif(StartHM = "12:00","selected","")%>>12:00</option>
			<option value="12:30" <%=chkiif(StartHM = "12:30","selected","")%>>12:30</option>
			<option value="13:00" <%=chkiif(StartHM = "13:00","selected","")%>>13:00</option>
			<option value="13:30" <%=chkiif(StartHM = "13:30","selected","")%>>13:30</option>
			<option value="14:00" <%=chkiif(StartHM = "14:00","selected","")%>>14:00</option>
			<option value="14:30" <%=chkiif(StartHM = "14:30","selected","")%>>14:30</option>
			<option value="15:00" <%=chkiif(StartHM = "15:00","selected","")%>>15:00</option>
			<option value="15:30" <%=chkiif(StartHM = "15:30","selected","")%>>15:30</option>
			<option value="16:00" <%=chkiif(StartHM = "16:00","selected","")%>>16:00</option>
			<option value="16:30" <%=chkiif(StartHM = "16:30","selected","")%>>16:30</option>
			<option value="17:00" <%=chkiif(StartHM = "17:00","selected","")%>>17:00</option>
			<option value="17:30" <%=chkiif(StartHM = "17:30","selected","")%>>17:30</option>
			<option value="18:00" <%=chkiif(StartHM = "18:00","selected","")%>>18:00</option>
			<option value="18:30" <%=chkiif(StartHM = "18:30","selected","")%>>18:30</option>
			<option value="19:00" <%=chkiif(StartHM = "19:00","selected","")%>>19:00</option>
			<option value="19:30" <%=chkiif(StartHM = "19:30","selected","")%>>19:30</option>
			<option value="20:00" <%=chkiif(StartHM = "20:00","selected","")%>>20:00</option>
			<option value="20:30" <%=chkiif(StartHM = "20:30","selected","")%>>20:30</option>
			<option value="21:00" <%=chkiif(StartHM = "21:00","selected","")%>>21:00</option>
			<option value="21:30" <%=chkiif(StartHM = "21:30","selected","")%>>21:30</option>
			<option value="22:00" <%=chkiif(StartHM = "22:00","selected","")%>>22:00</option>
			<option value="22:30" <%=chkiif(StartHM = "22:30","selected","")%>>22:30</option>
			<option value="23:00" <%=chkiif(StartHM = "23:00","selected","")%>>23:00</option>
		</select>~
		<select name="end_time">
			<option value="09:30" <%=chkiif(EndHM = "09:30","selected","")%>>09:30</option>
			<option value="10:00" <%=chkiif(EndHM = "10:00","selected","")%>>10:00</option>
			<option value="10:30" <%=chkiif(EndHM = "10:30","selected","")%>>10:30</option>
			<option value="11:00" <%=chkiif(EndHM = "11:00","selected","")%>>11:00</option>
			<option value="11:30" <%=chkiif(EndHM = "11:30","selected","")%>>11:30</option>
			<option value="12:00" <%=chkiif(EndHM = "12:00","selected","")%>>12:00</option>
			<option value="12:30" <%=chkiif(EndHM = "12:30","selected","")%>>12:30</option>
			<option value="13:00" <%=chkiif(EndHM = "13:00","selected","")%>>13:00</option>
			<option value="13:30" <%=chkiif(EndHM = "13:30","selected","")%>>13:30</option>
			<option value="14:00" <%=chkiif(EndHM = "14:00","selected","")%>>14:00</option>
			<option value="14:30" <%=chkiif(EndHM = "14:30","selected","")%>>14:30</option>
			<option value="15:00" <%=chkiif(EndHM = "15:00","selected","")%>>15:00</option>
			<option value="15:30" <%=chkiif(EndHM = "15:30","selected","")%>>15:30</option>
			<option value="16:00" <%=chkiif(EndHM = "16:00","selected","")%>>16:00</option>
			<option value="16:30" <%=chkiif(EndHM = "16:30","selected","")%>>16:30</option>
			<option value="17:00" <%=chkiif(EndHM = "17:00","selected","")%>>17:00</option>
			<option value="17:30" <%=chkiif(EndHM = "17:30","selected","")%>>17:30</option>
			<option value="18:00" <%=chkiif(EndHM = "18:00","selected","")%>>18:00</option>
			<option value="18:30" <%=chkiif(EndHM = "18:30","selected","")%>>18:30</option>
			<option value="19:00" <%=chkiif(EndHM = "19:00","selected","")%>>19:00</option>
			<option value="19:30" <%=chkiif(EndHM = "19:30","selected","")%>>19:30</option>
			<option value="20:00" <%=chkiif(EndHM = "20:00","selected","")%>>20:00</option>
			<option value="20:30" <%=chkiif(EndHM = "20:30","selected","")%>>20:30</option>
			<option value="21:00" <%=chkiif(EndHM = "21:00","selected","")%>>21:00</option>
			<option value="21:30" <%=chkiif(EndHM = "21:30","selected","")%>>21:30</option>
			<option value="22:00" <%=chkiif(EndHM = "22:00","selected","")%>>22:00</option>
			<option value="22:30" <%=chkiif(EndHM = "22:30","selected","")%>>22:30</option>
			<option value="23:00" <%=chkiif(EndHM = "23:00","selected","")%>>23:00</option>
			<option value="23:30" <%=chkiif(EndHM = "23:30","selected","")%>>23:30</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">그룹명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="groupname" size="30" maxlength="128" class="text" value="<%=Sgroupname%>"></td>
</tr>
<tr>
	<td width="100">사용목적</td>
	<td bgcolor="#FFFFFF">
		<select name="usepurpose" class="select">
			<option value="1" <%=chkiif(Susepurpose = "1","selected","")%>>강좌</option>
			<option value="2" <%=chkiif(Susepurpose = "2","selected","")%>>회의</option>
			<option value="3" <%=chkiif(Susepurpose = "3","selected","")%>>미팅</option>
			<option value="4" <%=chkiif(Susepurpose = "4","selected","")%>>면접</option>
			<option value="5" <%=chkiif(Susepurpose = "5","selected","")%>>기타</option>
		</select>
	</td>
</tr>
<tr>
	<td width="100">연락처</td>
	<td bgcolor="#FFFFFF"><input type="text" name="usercell" size="30" maxlength="128" class="text" value="<%=Susercell%>"></td>
</tr>
<tr>
	<td width="100">사용인원</td>
	<td bgcolor="#FFFFFF"><input type="text" style="IME-MODE:disabled;" name="useSu" size="10" maxlength="5" class="text" value="<%=Susesu%>" onkeypress="onlyNumber();">명</td>
</tr>
<tr>
	<td width="100">기타사항</td>
	<td bgcolor="#FFFFFF"><textarea name="etc" cols="40" rows="5"><%=Setc%></textarea></td>
</tr>
<tr>
	<td width="100">강좌번호</td>
	<td bgcolor="#FFFFFF"><input type="text" style="IME-MODE:disabled;" name="lecnum" size="10" maxlength="10" class="text" value="<%=Slecnum%>" onkeypress="onlyNumber();">※없으면 공란</td>
</tr>
<% If mode = "modify" Then %>
<tr>
	<td width="100">사용여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%=chkiif(Sisusing = "Y","checked","")%>>Y
		<input type="radio" name="isusing" value="N" <%=chkiif(Sisusing = "N","checked","")%>>N
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