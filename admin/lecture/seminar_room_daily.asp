<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/academy/seminar_roomcls.asp"-->
<%
dim ttime

Function UseTimeName(ttime)
	if ttime = 8 then
		UseTimeName = "10:00"
	elseif ttime = 9 then
		UseTimeName = "10:30"
	elseif ttime = 10 then
		UseTimeName = "11:00"
	elseif ttime = 11 then
		UseTimeName = "11:30"
	elseif ttime = 12 then
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

	font-family: "����ü";
	font-size: 9pt;
}
</style>
<script language="JavaScript">
<!--
	function OpenRoomWrite(tdate,ttime,roomid){
		var popwin = window.open('seminar_room_edit.asp?tdate=' + tdate + '&ttime=' + ttime + '&mode=write&roomid='+roomid,'seminar','width=450, height=350,resizable=1');
		popwin.focus();
	}

	function OpenRoomEdit(idx,tdate,ttime,roomid){
		var popwin = window.open('seminar_room_edit.asp?idx=' + idx + '&tdate=' + tdate + '&ttime=' + ttime + '&mode=modify&roomid='+roomid,'seminar','width=450, height=350,resizable=1');
		popwin.focus();
	}
//-->
</script>
<table>
<tr>
	<td><font color="red"><b><% = FormatDateTime(getday,1) %></b></font></td>
</tr>
</table>
<table width="800" border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="black" class="verdana-mid">
  <tr>
    <td align="center">���̸�</td>
<!--
    <td align="center" width="90">Television</td>
    <td align="center" width="90">Bingo</td>
    <td align="center" width="90">Chocolate</td>
-->
    <td align="center" width="90" bgcolor="#F0FFF8">Idea</td>
    <td align="center" width="90" bgcolor="#F0FFF8">Paper</td>
    <td align="center" width="90" bgcolor="#FFF8F0">Heart</td>
    <td align="center" width="90" bgcolor="#FFF8F0">Fingers</td>
    <td align="center" width="90" bgcolor="#FFF8F0">Moon</td>
    <td align="center" width="90" bgcolor="#FFF8F0">Star</td>
    <td align="center" width="90" bgcolor="#F0FFF8">Step</td>
    <td align="center" width="90" bgcolor="#F0FFF8">Music</td>
  </tr>
  <tr>
    <td align="center">�����ο�</td>
    <td align="center" class="verdana-small" bgcolor="#F0FFF8">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#F0FFF8">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#FFF8F0">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#FFF8F0">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#FFF8F0">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#FFF8F0">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#F0FFF8">8-10</td>
    <td align="center" class="verdana-small" bgcolor="#F0FFF8">8-10</td>
  </tr>
  <tr>
    <td align="center">
		<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="verdana-small">
		<tr><td align="right" height="50" valign="top">10:00</td></tr>
		<tr><td align="right" height="50" valign="top">11:00</td></tr>
		<tr><td align="right" height="50" valign="top">12:00</td></tr>
		<tr><td align="right" height="50" valign="top">13:00</td></tr>
		<tr><td align="right" height="50" valign="top">14:00</td></tr>
		<tr><td align="right" height="50" valign="top">15:00</td></tr>
		<tr><td align="right" height="50" valign="top">16:00</td></tr>
		<tr><td align="right" height="50" valign="top">17:00</td></tr>
		<tr><td align="right" height="50" valign="top">18:00</td></tr>
		<tr><td align="right" height="50" valign="top">19:00</td></tr>
		<tr><td align="right" height="50" valign="top">20:00</td></tr>
		<tr><td align="right" height="50" valign="top">21:00</td></tr>
		<tr><td align="right" height="50" valign="top">22:00</td></tr>
		<tr>
			<td align="right" height="50" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" class="verdana-small" width="100%" height="100%">
			<tr><td align="right" valign="top">23:00</td></tr>
			<tr><td align="right" valign="bottom">24:00</td></tr>
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
dim roomid, strTRBgCr
dim basictime

for ix = 1 to 11 '8���� �游ŭ ����
If ix=1 Or ix=2 Or ix=3 Or ix=4 Or ix=8 or ix=9 or ix=10 or ix=11 then
roomid = Num2Str(ix,2,"0","R")

	Select Case roomid
		Case "01","02", "10", "11"
			strTRBgCr = "#F0F4F0"
		Case "03","04", "08", "09"
			strTRBgCr = "#F4F0F0"
		Case Else
			strTRBgCr = ""
	End Select

	osemi.FRectYYYYMM = getday
	osemi.FRectRoom = roomid
	osemi.DailyList

%>
				<td bgcolor="<%=strTRBgCr%>">
					<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080"  width="100%" height="100%" class="verdana-m" style="cursor:hand">

				<% for iy = 8 to 35 '14�ð� ���� %>
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
									<td height="25" bgcolor="#CCFF00" align="center"><% = osemi.FItemList(iz).Fusepeople %>��</td>
								</tr>
								<tr>
									<td height="25" bgcolor="#CCFF00" align="center"><% = osemi.FItemList(iz).Fuserphone %></td>
								</tr>
								<% For k = 1 To (2*(basictime-1))-1 %>
								<tr>
									<td height="25" bgcolor="#CCFF00" align="center">&nbsp;</td>
								</tr>
								<% Next %>
							<% iz = iz + (2 * (basictime - 1)) '�ɸ��� ��ĭ ������%>
							<% iy = iy + (2 * (basictime)) - 1 '2ĭ ������%>
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