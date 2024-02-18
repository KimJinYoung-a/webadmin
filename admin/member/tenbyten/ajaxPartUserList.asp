<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim part_sn
part_sn = requestCheckvar(Request("part_sn"),10)
IF part_sn = "" THEN part_sn =0
	'// 직원정보 리스트
	dim oMember, arrList,intLoop
	Set oMember = new CTenByTenMember 
	oMember.Fpart_sn 		= part_sn 
	arrList = oMember.fnGetPartUserList 
	set oMember = nothing 
%>
<select  name="selUL" id="selUL" multiple size="20" style="width:200px">
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBOund(arrList,2)
%>
	<option value="<%=arrList(2,intLoop)&"-"&arrList(4,intLoop)%>"><%=arrList(1,intLoop)%>&nbsp;<%=arrList(7,intLoop)%> </option>
<%	Next
END IF%>
</select>
<!-- #include virtual="/lib/db/dbclose.asp" -->