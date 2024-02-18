<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->

<%
	Dim arrList, intLoop, vWorkType, vWorkTarget, vReqEquipment
	vWorkType 		= requestCheckVar(Request("work_type"),2)
	vWorkTarget		= requestCheckVar(Request("work_target"),20)
	vReqEquipment	= requestCheckVar(Request("req_equipment"),2)

	If vWorkType = "3" Then
		vWorkTarget = vWorkTarget & "_break"
	End If
	If vWorkTarget = "etc_break" Then
		vWorkTarget = "etc"
	End IF

	Dim breakdownCodeList
	Set breakdownCodeList = New CBreakCommonCode
	breakdownCodeList.FCodeType = vWorkTarget
	arrList = breakdownCodeList.fnGetBreakCodeList
	Set breakdownCodeList = nothing 
%>

<script language="javascript">
document.domain = "10x10.co.kr";
function selCode(code,codename)
{
	parent.document.frm.req_equipment.value = code;
	parent.document.frm.req_equipment_name.value = codename;
}
</script>

<%If isArray(arrList) THEN%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-top:10px;">* <%=fnWorkTargetCode3(vWorkType,vWorkTarget)%></td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%For intLoop = 0 To UBound(arrList,2)%>
<tr>
	<td bgcolor="#FFFFFF"><input type="radio" name="code_value" value="<%=arrList(1,intLoop)%>" onClick="selCode('<%=arrList(1,intLoop)%>','<%=Replace(arrList(4,intLoop),vbCrLf," ")%>');" <%=CHKIIF(CStr(vReqEquipment)=CStr(arrList(1,intLoop)),"checked","")%>></td>
	<td bgcolor="#FFFFFF"><%=arrList(4,intLoop)%></td>
</tr>
<%Next%>
</table>
<br>
<%End if%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->