<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/outmallSellCls.asp"-->
<%
Dim vMallid, oOutMall, page, i, vMakerid
vMallid		= requestCheckvar(Request("mallid"),16)
vMakerid	= requestCheckvar(request("makerid"),50)
If page = "" Then page = 1

SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= page
	oOutMall.FPageSize			= 100
	oOutMall.FRectMallid		= vMallid
	oOutMall.FRectMakerid		= vMakerid
	oOutMall.getOutmallLogList
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr><td width="90%"></td></tr>
		<tr>
			<td>이전 판매상태 내역</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="mallgubun" value="<%=vMallid%>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="90">구분</td>
	<td width="70">판매설정</td>
	<td>내용</td>
	<td width="70">등록자</td>
	<td width="170">등록일</td>
</tr>
<% If oOutMall.FResultCount > 0 Then %>
<% For i = 0 To oOutMall.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
	<%
		Select Case oOutMall.FItemList(i).FMallid
			Case "all"		response.write "제휴몰 전체"
			Case "daumep"	response.write "다음"
			Case "naverep"	response.write "네이버"
			Case "shodocep"	response.write "쇼닥"
			Case Else		response.write oOutMall.FItemList(i).FMallid
		End Select
	%>
	</td>
	<td>
	<%
		Select Case oOutMall.FItemList(i).FUseYN
			Case "Y"		response.write "판매"
			Case "N"		response.write "판매안함"
			Case Else 		response.write "취소"
		End Select
	%>
	</td>
	<td>
	<%
		If Instr(oOutMall.FItemList(i).FHopestr, "[관리자]") > 0 Then
			response.write "<strong>"&oOutMall.FItemList(i).FHopestr&"</strong>"
		Else
			response.write oOutMall.FItemList(i).FHopestr
		End If
	%>
	 </td>
	<td><%= oOutMall.FItemList(i).FReguserid %></td>
	<td width="170"><%= oOutMall.FItemList(i).FRegdate %></td>
</tr>
<% Next %>
<% Else %>
<tr align="center" bgcolor="#FFFFFF" height="50">
	<td colspan="5">설정 내역이 없습니다.</td>
</tr>
<% End If %>
</table>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
