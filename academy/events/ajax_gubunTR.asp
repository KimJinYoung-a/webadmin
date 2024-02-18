<%@ Language=VBScript %>
<%
Option Explicit
Response.Expires = -1440
response.Charset="euc-kr" 
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/event/eventCls.asp"-->
<%
Dim gubun, idx, actid, company_name, regCode
Dim oEvent
gubun	= RequestCheckvar(request("gubun"),1)
idx		= RequestCheckvar(request("idx"),10)

If idx <> "" Then
	Set oEvent = new CEvent
		oEvent.FRectIdx = idx
		oEvent.getEventOneItem

		regCode			= oEvent.FOneItem.FContentsCode
		actid			= oEvent.FOneItem.FActid
		company_name	= oEvent.FOneItem.FCompany_name
	Set oEvent = nothing
End If
%> 
<% If gubun = "D" Then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">작가</td>
	<td>
		<input type="text" name="actid" class="text" id="tecid" value="<%=actid%>">
		<input type="text" name="company_name" class="text" id="diy_name" value="<%=company_name%>" readonly>
		<input type="button" value="작가 전체보기" class="button" onclick="pop_lecture('D');">
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">작품코드</td>
	<td>
		<input type="text" name="diycode" id="diycode" class="text">
		<input type="button" value="판매중인 작품보기" class="button" id="btnDiyView" <% If idx = "" Then response.write "disabled" End If %> onclick="pop_art();">
		&nbsp;&nbsp;<% If idx <> "" Then response.write "기존코드 : ("&regCode&") 수정시에는 작품코드가 공란입니다. 재입력해주세요" End If %>
	</td>
<% ElseIf gubun = "L" Then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">강사</td>
	<td>
		<input type="text" name="lectureid" class="text" id="lecid" value="<%=actid%>" readonly>
		<input type="text" name="company_name" class="text" id="company_name" value="<%=company_name%>" readonly>
		<input type="button" value="강사 전체보기" class="button" onclick="pop_lecture('L');">
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">강좌코드</td>
	<td>
		<input type="text" name="lecidx" class="text" id="lecidx" readonly>
		<input type="button" value="진행중인 강좌보기" class="button" id="btnView" <% If idx = "" Then response.write "disabled" End If %> onclick="pop_lec();">
		&nbsp;&nbsp;<% If idx <> "" Then response.write "기존코드 : ("&regCode&") 수정시에는 강좌코드가 공란입니다. 재입력해주세요" End If %>
	</td>
</tr>
<% End If %>