<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seminar_roomcls.asp"-->
<%
dim ttime

Function UseTimeName(ttime)
	if ttime = 12 then
		UseTimeName = "12:00"
	elseif ttime = 13 then
		UseTimeName = "12:30"
	elseif ttime = 14 then
		UseTimeName = "1:00"
	elseif ttime = 15 then
		UseTimeName = "1:30"
	elseif ttime = 16 then
		UseTimeName = "2:00"
	elseif ttime = 17 then
		UseTimeName = "2:30"
	elseif ttime = 18 then
		UseTimeName = "3:00"
	elseif ttime = 19 then
		UseTimeName = "3:30"
	elseif ttime = 20 then
		UseTimeName = "4:00"
	elseif ttime = 21 then
		UseTimeName = "4:30"
	elseif ttime = 22 then
		UseTimeName = "5:00"
	elseif ttime = 23 then
		UseTimeName = "5:30"
	elseif ttime = 24 then
		UseTimeName = "6:00"
	elseif ttime = 25 then
		UseTimeName = "6:30"
	elseif ttime = 26 then
		UseTimeName = "7:00"
	elseif ttime = 27 then
		UseTimeName = "7:30"
	elseif ttime = 28 then
		UseTimeName = "8:00"
	elseif ttime = 29 then
		UseTimeName = "8:30"
	elseif ttime = 30 then
		UseTimeName = "9:00"
	elseif ttime = 31 then
		UseTimeName = "9:30"
	elseif ttime = 32 then
		UseTimeName = "10:00"
	elseif ttime = 33 then
		UseTimeName = "10:30"
	elseif ttime = 34 then
		UseTimeName = "11:00"
	elseif ttime = 35 then
		UseTimeName = "11:30"
	end if
End Function

Dim getday

getday = request("getday")

set osemi = new CSeminarRoomCalendar

%>
<style>
.verdana-m {

	font-family: "돋움체";
	font-size: 9pt;
}
</style>
<script language="JavaScript">
<!--
	function OpenRoomWrite(tdate,ttime,roomid){
		window.open('seminar_room_edit.asp?tdate=' + tdate + '&ttime=' + ttime + '&mode=write&roomid='+roomid,'seminar','width=450, height=350,resizable=1');
	}
	function OpenRoomEdit(idx,tdate,ttime,roomid){
		window.open('seminar_room_edit.asp?idx=' + idx + '&tdate=' + tdate + '&ttime=' + ttime + '&mode=modify&roomid='+roomid,'seminar','width=450, height=350,resizable=1');
	}
//-->
</script>
<table>
<tr>
	<td><font color="red"><b><% = FormatDateTime(getday,1) %></b></font></td>
</tr>
</table>
<table width="580" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="verdana-mid">
  <tr>
    <td align="center">방이름</td>
    <td align="center" width="90">Heart</td>
    <td align="center" width="90">Fingers</td>
    <td align="center" width="90">Chocolate</td>
    <td align="center" width="90">Moon</td>
  </tr>
  <tr>
    <td align="center">가능인원</td>
    <td align="center" class="verdana-small">8-10</td>
    <td align="center" class="verdana-small">8-10</td>
    <td align="center" class="verdana-small">8-10</td>
    <td align="center" class="verdana-small">8-10</td>
  </tr>
  <tr>
    <td align="center">
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="verdana-small">
		<tr>
			<td align="right" height="50" valign="top">12:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">13:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">14:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">15:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">16:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">17:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">18:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">19:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">20:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">21:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">22:00</td>
		</tr>
		<tr>
			<td align="right" height="50" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" class="verdana-small" width="100%" height="100%">
			<tr>
				<td align="right" valign="top">23:00</td>
			</tr>
			<tr>
				<td align="right" valign="bottom">24:00</td>
			</tr>
			</table>
			</td>
		</tr>
		</table>
	</td>
    <td colspan="9">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
			<tr>
<%
dim osemi
dim ix,iy,iz,k
dim roomid
dim basictime

for ix = 1 to 9 '9개의 방만큼 루프
If ix=3 Or ix=4 Or ix=6 Or ix=8 then
roomid = "0" + Cstr(ix)

	osemi.FRectYYYYMM = getday
	osemi.FRectRoom = roomid
	osemi.DailyList

%>
				<td>
					<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080"  width="100%" height="100%" class="verdana-m" style="cursor:hand">
				
				<% for iy = 12 to 35 '12시간 루프 %>
					<% if osemi.FResultCount < 1 Then %>
							<tr>
								<td height="25" width="90" onclick="OpenRoomWrite('<% = getday  %>','<% = iy %>','<% = roomid %>');" valign="top" align="left"><font color="#808080"><% = UseTimeName(iy) %></font></td>
							</tr>
					<% Else %>
						<% For iz = 0 To osemi.FResultCount - 1 %>
						<% basictime = osemi.FItemList(iz).Fbasictime %>
							<% if Cstr(iy) = Cstr(osemi.FItemList(iz).Fusetime) Then %>
								<tr>
									<td height="25" bgcolor="#CCFF00" width="90" align="center" onclick="OpenRoomEdit('<% = osemi.FItemList(iz).Fidx %>','<% = getday  %>','<% = iy %>','<% = roomid %>');"><% = DDotFormat(osemi.FItemList(iz).Fusername,6) %></td>
								</tr>
								<tr>
									<td height="25" bgcolor="#CCFF00" align="center"><% = osemi.FItemList(iz).Fusepeople %>명</td>
								</tr>
								<tr>
									<td height="25" bgcolor="#CCFF00" align="center"><% = osemi.FItemList(iz).Fuserphone %></td>
								</tr>
								<% For k = 1 To (2*(basictime-1))-1 %>
								<tr>
									<td height="25" bgcolor="#CCFF00" align="center">&nbsp;</td>
								</tr>
								<% Next %>
							<% iz = iz + (2 * (basictime - 1)) '걸리면 한칸 보내기%>
							<% iy = iy + (2 * (basictime)) - 1 '2칸 보내기%>
							<% Else %>
								<% If Cstr(iy) <> Cstr(osemi.FItemList(iz).Fusetime) And iz = osemi.FResultCount - 1 Then %>
								<tr>
									<td height="25" width="90" onclick="OpenRoomWrite('<% = getday  %>','<% = iy %>','<% = roomid %>');"  valign="top" align="left"><font color="#808080"><% = UseTimeName(iy) %></font></td>
								</tr>
								<% End If %>
							<% End If %>
						<% Next %>
					<% End If %>
				<% Next %>
					</table>
				</td>
<% End If %>
<% Next %>
			</tr>
			</table>
	</td>
  </tr>
</table>
<% Set osemi = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->