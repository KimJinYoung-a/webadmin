<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventAppReport.asp"-->
<%
Dim eventNo, sDate, eDate, offEvent, arrList, i, ageStr, eventName
Dim TotalUserCnt, TotalManCnt, TotalWomenCnt, TotalVVIPCnt, TotalVIPGOLDCnt, TotalVIPSILVERCnt, TotalBLUECnt, TotalGREENCnt, TotalYELLOWCnt, TotalORANGECnt
Dim TermUserCnt, TermManCnt, TermWomenCnt, TermVVIPCnt, TermVIPGOLDCnt, TermVIPSILVERCnt, TermBLUECnt, TermGREENCnt, TermYELLOWCnt, TermORANGECnt

eventNo	= requestCheckVar(request("eventNo"),10)
sDate	= requestCheckVar(request("sDate"),10)
eDate	= requestCheckVar(request("eDate"),10)
If (NOT isDate(sDate)) OR (NOT isDate(eDate)) Then
	response.write "<script>alert('날짜 형식이 잘 못 되었습니다.');window.close()</script>"
	response.end
end if

If NOT isnumeric(eventNo) Then
	response.write "<script>alert('이벤트가 잘 못 되었습니다.');window.close()</script>"
	response.end
End If

SET offEvent = new COffEvent
	offEvent.FRectSdate			= sDate
	offEvent.FRectEdate			= eDate
	offEvent.FRectEventNo		= eventNo
	arrList = offEvent.fnOffEventUserReport
	'누적 전체 TR Data
	TotalUserCnt		= offEvent.FTotalUserCnt
	TotalManCnt			= offEvent.FTotalManCnt
	TotalWomenCnt		= offEvent.FTotalWomenCnt
	TotalVVIPCnt		= offEvent.FTotalVVIPCnt
	TotalVIPGOLDCnt		= offEvent.FTotalVIPGOLDCnt
	TotalVIPSILVERCnt	= offEvent.FTotalVIPSILVERCnt
	TotalBLUECnt		= offEvent.FTotalBLUECnt
	TotalGREENCnt		= offEvent.FTotalGREENCnt
	TotalYELLOWCnt		= offEvent.FTotalYELLOWCnt
	TotalORANGECnt		= offEvent.FTotalORANGECnt

	'기간 전체 TR Data
	TermUserCnt			= offEvent.FTermUserCnt
	TermManCnt			= offEvent.FTermManCnt
	TermWomenCnt		= offEvent.FTermWomenCnt
	TermVVIPCnt			= offEvent.FTermVVIPCnt
	TermVIPGOLDCnt		= offEvent.FTermVIPGOLDCnt
	TermVIPSILVERCnt	= offEvent.FTermVIPSILVERCnt
	TermBLUECnt			= offEvent.FTermBLUECnt
	TermGREENCnt		= offEvent.FTermGREENCnt
	TermYELLOWCnt		= offEvent.FTermYELLOWCnt
	TermORANGECnt		= offEvent.FTermORANGECnt

	Select Case eventNo
		Case "1"		eventName = "앱 설치 이벤트"
		Case "2"		eventName = "아트토이컬쳐"
		Case "3"		eventName = "사은품 증정 이벤트"
		Case Else 		eventName = "정의되지 않은 이벤트"
	End Select
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="15">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>회원 통계 [<%= eventName %>] <%= sDate %> ~ <%= eDate %></strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="50">
	<td width="9%"></td>
	<td width="9%">전체 회원(명)</td>
	<td width="9%">남자(명)</td>
	<td width="9%">여자(명)</td>
	<td width="9%">VVIP</td>
	<td width="9%">VIP Gold</td>
	<td width="9%">VIP Silver</td>
	<td width="9%">BLUE</td>
	<td width="9%">GREEN</td>
	<td width="9%">YELLOW</td>
	<td width="9%">ORANGE</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td bgcolor="<%= adminColor("tabletop") %>"><strong>누적 전체</strong></td>
	<td><%= TotalUserCnt %></td>
	<td><%= TotalManCnt %></td>
	<td><%= TotalWomenCnt %></td>
	<td><%= TotalVVIPCnt %></td>
	<td><%= TotalVIPGOLDCnt %></td>
	<td><%= TotalVIPSILVERCnt %></td>
	<td><%= TotalBLUECnt %></td>
	<td><%= TotalGREENCnt %></td>
	<td><%= TotalYELLOWCnt %></td>
	<td><%= TotalORANGECnt %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td bgcolor="<%= adminColor("tabletop") %>"><strong>기간 전체</strong></td>
	<td><%= TermUserCnt %></td>
	<td><%= TermManCnt %></td>
	<td><%= TermWomenCnt %></td>
	<td><%= TermVVIPCnt %></td>
	<td><%= TermVIPGOLDCnt %></td>
	<td><%= TermVIPSILVERCnt %></td>
	<td><%= TermBLUECnt %></td>
	<td><%= TermGREENCnt %></td>
	<td><%= TermYELLOWCnt %></td>
	<td><%= TermORANGECnt %></td>
</tr>
<%
If IsArray(arrList) Then
	For i=0 To Ubound(arrList, 2) 
		Select Case arrList(0, i)
			Case "v20"	ageStr = "0 ~ 19세"
			Case "v24"	ageStr = "20 ~ 24세"
			Case "v29"	ageStr = "25 ~ 29세"
			Case "v34"	ageStr = "30 ~ 34세"
			Case "v39"	ageStr = "35 ~ 39세"
			Case "v49"	ageStr = "40 ~ 49세"
			Case "v50"	ageStr = "50세 ~"
			Case Else	ageStr = "탈퇴회원"
		End Select
%>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td bgcolor="<%= adminColor("tabletop") %>"><strong><%= ageStr %></strong></td>
	<td><%= arrList(1, i) %></td>
	<td><%= arrList(2, i) %></td>
	<td><%= arrList(3, i) %></td>
	<td><%= arrList(4, i) %></td>
	<td><%= arrList(5, i) %></td>
	<td><%= arrList(6, i) %></td>
	<td><%= arrList(7, i) %></td>
	<td><%= arrList(8, i) %></td>
	<td><%= arrList(9, i) %></td>
	<td><%= arrList(10, i) %></td>
</tr>
<%
	Next 
End If
%>
<tr align="center" height="30" bgcolor="#FFFFFF">
	<td colspan="15">
		<input type="button" class="button" onclick="self.close();" value="닫기">
	</td>
</tr>
</table>
<% SET offEvent = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->