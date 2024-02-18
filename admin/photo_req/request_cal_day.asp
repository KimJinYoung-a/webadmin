<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 요청 스케줄
' History : 2012.03.15 김진영 생성
'			2018.02.09 한용민 수정
'####################################################
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/shedulecls.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
Dim ttime, query1, user_cnt, roomid, room_id, i, stylist_name, code_name, code_detail_name

Function UseTimeName(ttime)
	Select Case ttime
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
	End Select
End Function

Dim getday, pre_getday, next_getday
getday = request("getday")
pre_getday = CDate(getday) - 1
next_getday = CDate(getday) + 1

Dim PhotoCnt, cPhotoreq
set cPhotoreq = new Photoreq
PhotoCnt = cPhotoreq.fnGetPhotoUser
'response.write pre_getday&"<BR>"
'response.write next_getdate&"<BR>"
Set osemi = new CSeminarRoomCalendar
%>
<style>
.verdana-m {
	font-family: "돋움체";
	font-size: 9pt;
}
.box-rb td {border-right:1px solid #888; border-bottom:1px solid #888;}
.box-tt tr {height:62px;}
.box-tt td {text-align:right; vertical-align:top;}
</style>
<script language="JavaScript">
<!--
function OpenRoomWrite(tdate,ttime,roomid,mm,rno){
	var popwin = window.open('request_cal_pop.asp?rno='+rno+'&tdate=' + tdate + '&ttime=' + ttime + '&mode='+mm+'&roomid='+roomid,'seminar','width=450, height=350,resizable=1');
	popwin.focus();
}

function request_modi(idx){
	location.href= 'request_modi.asp?req_no='+idx+'&udate=A';
}
function detail_status(idx,sno){
	var popwin = window.open('request_cal_pop2.asp?rno='+idx+'&sno='+sno,'seminar2','width=450, height=250,resizable=1');
	popwin.focus();
}
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
function goDay(){
	location.href='request_cal_day.asp?getday='+document.getElementById("Movedate").value;
}
//-->
</script>
<table border="0" cellpadding="0" cellspacing="0" class="verdana-mid">
<tr align="center" height="40">
	<td colspan="8" valign="middle" style="border-bottom:2px solid #888;">
		<A HREF="request_cal_day.asp?getday=<%=pre_getday%>"><img src="http://testwebadmin.10x10.co.kr/images/icon_arrow_left.gif" border="0"></A>&nbsp;&nbsp;<b><% = FormatDateTime(getday,1) %></b>&nbsp;&nbsp;<A HREF="request_cal_day.asp?getday=<%=next_getday%>"><img src="http://testwebadmin.10x10.co.kr/images/icon_arrow_right.gif" border="0"></A>
		&nbsp;&nbsp;
		<input type="text" name="Movedate" id="Movedate" size="10" value="<%=getday%>" onClick="jsPopCal('Movedate');" style="cursor:pointer;" />
		<input type="button" class="button" value="이동" onclick="goDay();" />
	</td>
</tr>
<tr height="25">
	<td align="center" width="100" style="border-right:1px solid #888; border-left:1px solid #888; border-bottom:2px solid #888;">담당포토</td>
<%
	   query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user"
	   query1 = query1 + " where user_type='1' and user_useyn = 'Y'"
	   rsget.Open query1,dbget,1

		If  not rsget.EOF  Then
			user_cnt = 0
			roomid = ""

			Do until rsget.EOF
				user_cnt = user_cnt + 1
				roomid =  roomid&","&rsget("user_id")
%>
	<td align="center" width="190" bgcolor="#F0FFF8" style="border-right:1px solid #888; border-bottom:2px solid #888;"><%=rsget("user_name")%></td>
<%
				rsget.MoveNext
	       	Loop
		End if
		rsget.close
%>
</tr>
<tr>
	<td align="center" width="100" style="vertical-align:top; border-right:1px solid #888; border-left:1px solid #888; border-bottom:1px solid #888;">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="verdana-small box-tt">
		<tr><td>10:00</td></tr>
		<tr><td>11:00</td></tr>
		<tr><td>12:00</td></tr>
		<tr><td>13:00</td></tr>
		<tr><td>14:00</td></tr>
		<tr><td>15:00</td></tr>
		<tr><td>16:00</td></tr>
		<tr><td style="height:61px;">17:00</td></tr>
		</table>
	</td>
    <td colspan="9">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
<%
			Dim osemi
			Dim ix,iy,iz,k, ii, start_time, end_time
			Dim strTRBgCr, strTimeCont, strTDClick, strRowspan, blnRspan
			Dim basictime, FStartDate, basictime2, FEndDate, fusetime
			room_id = Split(roomid, ",")

			For ii = 1 To UBOUND(room_id)
				osemi.FRectYYYYMM = getday
				osemi.FRectRoom = room_id(ii)
				osemi.DailyList
%>
			<td bgcolor="<%=strTRBgCr%>">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" class="verdana-m box-rb" style="cursor:pointer">
<%
				For iy = 8 to 23 			'14시간 루프
					If osemi.FResultCount > 0 then
						strTimeCont = ""
						blnRspan = false
						For iz = 0 To osemi.FResultCount - 1

							start_time 			= Cstr(osemi.FItemList(iz).Fbasictime)
							end_time 			= Cstr(osemi.FItemList(iz).Fbasictime2)
							stylist_name		= UserCodeType("user", "", osemi.FItemList(iz).Freq_stylist)
							code_name			= UserCodeType("code", "doc_status", osemi.FItemList(iz).FReqUse)
							code_detail_name	= UserCodeType("code", "doc_status_detail", osemi.FItemList(iz).FReqUseDetail)

							'if Cstr(iy) >= start_time And Cstr(iy) < end_time Then
							If Cstr(iy) = start_time Then
								'해당 시간에 일정이 있으면
								strRowspan	= osemi.FItemList(iz).Fusetime		'시간 간격
								If strRowspan<1 Then strRowspan=1				'종료시간보다 시작시간이 나중일 수 없음

								If osemi.FItemList(iz).FUse_yn = "Y" Then
									strTimeCont = strTimeCont &"<center><b>"& DDotFormat(osemi.FItemList(iz).FReqDepartment,6)+DDotFormat(osemi.FItemList(iz).FWritename,20) & "</b></center>"
									If code_detail_name = "" then
										strTimeCont = strTimeCont & code_name& "<BR />"
									Else
										strTimeCont = strTimeCont & code_name&"("&code_detail_name & ")<BR />"
									End If
									strTimeCont = strTimeCont & DDotFormat(osemi.FItemList(iz).FPrdName,16) & "<BR />"
									strTimeCont = strTimeCont & DDotFormat(osemi.FItemList(iz).FCode_nm,6) & "<BR />"
									strTimeCont = strTimeCont & "[stylist]"&stylist_name & "<BR />"

									Select Case osemi.FItemList(iz).FStatus
										Case "4"	strTRBgCr	= "#00CC00"
										Case "1"	strTRBgCr	= "#6699FF"
										Case "2"	strTRBgCr	= "#FFCC33"
										Case "3"	strTRBgCr	= "GRAY"
									End Select
									'strTDClick	= "request_modi(" & osemi.FItemList(iz).FReqNo & ",'" & getday & "','" & iy & "','" & roomid & "');"
									strTDClick	= "detail_status(" & osemi.FItemList(iz).FReqNo & ", " & osemi.FItemList(iz).FSchedule_no & ")"

								ElseIf osemi.FItemList(iz).FUse_yn = "S" Then
									strTimeCont = strTimeCont & osemi.FItemList(iz).FReq_comment & "<BR />"
									strTRBgCr = "#CCFFFF"
									'If CInt(PhotoCnt) > 0 or (session("ssBctId") = "kjy8517" OR session("ssBctId") = "tozzinet" or session("ssBctId") = "hrkang97") Then
										strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & roomid & "','modify'," & osemi.FItemList(iz).FReqNo & ");"
									'Else
									'	strTDClick = "alert('포토그래퍼만 수정가능합니다.');"
									'End If
								End If

									blnRspan = false
							end if

							If Cint(iy) > Cint(start_time) and Cint(iy) < Cint(end_time) Then
								blnRspan = true
							End if
						Next

						'일정걸리게 없으면 기본값
						If strTimeCont="" Then
							strTRBgCr = "#FFFFFF"
							strTimeCont = UseTimeName(iy)
							'If CInt(PhotoCnt) > 0 or (session("ssBctId") = "kjy8517" OR session("ssBctId") = "tozzinet" or session("ssBctId") = "hrkang97") Then
								strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & roomid & "','write','');"
							'Else
							'	strTDClick = "alert('포토그래퍼만 수정가능합니다.');"
							'End If
							strRowspan = "1"
						End If
					Else
						strTRBgCr = "#FFFFFF"
						strTimeCont = UseTimeName(iy)
						'If CInt(PhotoCnt) > 0 or (session("ssBctId") = "kjy8517" OR session("ssBctId") = "tozzinet" or session("ssBctId") = "hrkang97") Then
							strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & roomid & "','write','');"
						'Else
						'	strTDClick = "alert('포토그래퍼만 수정가능합니다.');"
						'End If
						strRowspan = "1"
						blnRspan = false
					End If

					If Not(blnRspan) Then
%>
						<tr>
							<td onclick="<%=strTDClick%>" style="height:<%=(strRowspan*30)+(chkIIF(strRowspan>1,strRowspan-1,0))%>px; background-color:<%=strTRBgCr%>; vertical-align:top; text-align:left; padding: 0 3px;">
								<div style="position:absolute; width:190px; height:<%=(strRowspan*30)-2%>px; overflow:auto;"><% = strTimeCont %></div>
							</td>
						</tr>
<%
					End If
				Next
%>
			</table>
			</td>
		<% Next %>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<table width="300" border="0" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="verdana-mid">
<tr><td align="center"><b>진행상황 상태 색상표</td></tr>
<tr><td bgcolor="WHITE"><b>진행상황 상태 선택 안 함(WHITE)</b></td></tr>
<tr><td bgcolor="#00CC00"><b>추가기입 요청(GREEN)</b></td></tr>
<tr><td bgcolor="#6699FF"><b>촬영스케줄 지정(BLUE)</b></td></tr>
<tr><td bgcolor="#FFCC33"><b>촬영중(GOLD)</b></td></tr>
<tr><td bgcolor="GRAY"><b>촬영완료(GRAY)</b></td></tr>
<tr><td bgcolor="#CCFFFF"><b>포토그래퍼 스케쥴 지정(SKY_BLUE)</b></td></tr>
</table>
<% Set osemi = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->