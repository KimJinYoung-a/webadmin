<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 세미나실 관리
' History : 2012.10.24 김진영 생성
'			2013.10.23 한용민 수정(세미나실 위치 변경)
'			2015.09.24 허진원 수정(선택한 방,날짜,시간 쓰기창에 전달)
'####################################################
%>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/seminar/seminarCls.asp"-->
<%
Dim ttime, query1, RoomCnt, roomidx, room_idx, i, stylist_name, code_name, code_detail_name
Dim purpose

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

Dim getday, pre_getday, next_getday
getday		= request("getday")
pre_getday	= CDate(getday) - 1
next_getday	= CDate(getday) + 1

Dim PhotoCnt, cPhotoreq
Dim osemi
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
function OpenRoomWrite(tdate,ttime,roomidx,mm,rno){
	var popwin = window.open('seminar_cal_pop.asp?idx='+rno+'&tdate=' + tdate + '&ttime=' + ttime + '&mode='+mm+'&roomidx='+roomidx,'seminar','width=450, height=350,resizable=1');
	popwin.focus();
}
//-->
</script>
<div><a href="/admin/seminar/seminar_calendar.asp?menupos=1482&ThisDate=<%=getday%>"><< 달력으로 이동</a></div>
<table border="0" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="verdana-mid">
<tr align="center" height="40">
	<td colspan="12" valign="middle" style="border-bottom:2px solid #888;"><A HREF="seminar_room_daily.asp?getday=<%=pre_getday%>"><img src="http://webadmin.10x10.co.kr/images/icon_arrow_left.gif" border="0"></A>&nbsp;&nbsp;<b><% = FormatDateTime(getday,1) %></b>&nbsp;&nbsp;<A HREF="seminar_room_daily.asp?getday=<%=next_getday%>"><img src="http://webadmin.10x10.co.kr/images/icon_arrow_right.gif" border="0"></A></td>
</tr>
<!--<tr height="25">
	<td align="center" width="100">&nbsp;</td>
	<td align="center" colspan="4">에버리치 홀딩스</td>
	<td align="center" colspan="4">자유빌딩</td>
</tr>-->
<tr height="25">
	<td width="100" align="center" style="border-right:1px solid #888; border-left:1px solid #888; border-bottom:2px solid #888;">방이름</td>
<%
	   query1 = " select idx, roomname, MaxSu, orderNo, isusing from db_partner.dbo.tbl_seminarRoom "
	   query1 = query1 + " where isusing='Y' Order by orderNo ASC "
	   rsget.Open query1,dbget,1

		If  not rsget.EOF  Then
			RoomCnt = 0
			roomidx = ""

			Do until rsget.EOF
				RoomCnt = RoomCnt + 1
				roomidx =  roomidx&","&rsget("idx")
%>
	<td width="194" align="center" <%=chkIIF(rsget("orderNo") <= 100,"bgcolor='YELLOW'","bgcolor='LIGHTGREEN'")%> style="border-right:1px solid #888; border-bottom:2px solid #888;"><%=rsget("roomname")%>
		<% if rsget("roomname")<>"홀" then %>
			(<%=rsget("MaxSu")%>)
		<% end if %>
	</td>
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
		<tr><td>09:00</td></tr>
		<tr><td>10:00</td></tr>
		<tr><td>11:00</td></tr>
		<tr><td>12:00</td></tr>
		<tr><td>13:00</td></tr>
		<tr><td>14:00</td></tr>
		<tr><td>15:00</td></tr>
		<tr><td>16:00</td></tr>
		<tr><td>17:00</td></tr>
		<tr><td>18:00</td></tr>
		<tr><td>19:00</td></tr>
		<tr><td>20:00</td></tr>
		<tr><td>21:00</td></tr>
		<tr><td>22:00</td></tr>
		<tr><td style="height:61px;">23:00</td></tr>
		</table>
	</td>
    <td colspan="11">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
<%
			Dim ix,iy,iz,k, ii, start_time, end_time
			Dim strTRBgCr, strTimeCont, strTDClick, strRowspan, blnRspan
			Dim basictime, FStartDate, basictime2, FEndDate, fusetime
			room_idx = Split(roomidx, ",")

			For ii = 1 To UBOUND(room_idx)
				osemi.FRectYYYYMM = getday
				osemi.FRectRoom = room_idx(ii)
				osemi.DailyList
%>
			<td bgcolor="<%=strTRBgCr%>" style="vertical-align:top; width:194px;">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" class="verdana-m box-rb" style="cursor:pointer;">
<%
				For iy = 6 to 35
					If osemi.FResultCount > 0 then
						strTimeCont = ""
						blnRspan = false
						strTRBgCr = "#E4F7BA"
						For iz = 0 To osemi.FResultCount - 1
							start_time 			= Cstr(osemi.FItemList(iz).Fbasictime)
							end_time 			= Cstr(osemi.FItemList(iz).Fbasictime2)

							If Cstr(iy) = start_time Then
								'해당 시간에 일정이 있으면
								strRowspan	= osemi.FItemList(iz).Fusetime		'시간 간격
								If strRowspan<1 Then strRowspan=1				'종료시간보다 시작시간이 나중일 수 없음

								If osemi.FItemList(iz).Fisusing = "Y" Then

									Select Case osemi.FItemList(iz).Fusepurpose
										Case "1"	purpose	= "[강좌]"
										Case "2"	purpose	= "[회의]"
										Case "3"	purpose	= "[미팅]"
										Case "4"	purpose	= "[면접]"
										Case "5"	purpose	= "[기타]"
									End Select

									strTimeCont = strTimeCont &"<b>"& purpose & DDotFormat(osemi.FItemList(iz).Fgroupname,16)&"</b><BR />"
									strTimeCont = strTimeCont &"사용인원 : "& DDotFormat(osemi.FItemList(iz).Fusesu,6) & "명<BR />"
									strTimeCont = strTimeCont & ReplaceScript(nl2br(osemi.FItemList(iz).Fetc)) & "<BR />"
									strTimeCont = strTimeCont &"등록자 : " & osemi.FItemList(iz).Fusername & "<BR />"
									strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & room_idx(ii) & "','modify'," & osemi.FItemList(iz).Fidx & ");"
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
							strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & room_idx(ii) & "','write','');"
							strRowspan = 1
						End If
					Else
						strTRBgCr = "#FFFFFF"
						strTimeCont = UseTimeName(iy)
						strTDClick = "OpenRoomWrite('" & getday &"','" & iy & "','" & room_idx(ii) & "','write','');"
						strRowspan = 1
						blnRspan = false
					End If

					If Not(blnRspan) Then
%>
						<tr>
							<td onclick="<%=strTDClick%>" style="height:<%=(strRowspan*30)+(chkIIF(strRowspan>1,strRowspan-1,0))%>px; background-color:<%=strTRBgCr%>; vertical-align:top; text-align:left; padding: 0 3px;">
								<div style="position:absolute; width:190px; height:<%=(strRowspan*30)-2%>px; overflow:auto;">
									<% = strTimeCont %>
								</div>
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

<% Set osemi = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->