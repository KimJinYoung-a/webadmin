<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
dim cAgit
Dim areaDiv, retData, agitSmsCont, agitSmsUpdate, agitSmsRegUser, mode
	areaDiv =requestCheckvar(request("areaDiv"),1)
if areaDiv="" then areaDiv="3"
	
 set cAgit = new CAgitUse
 	 	cAgit.FRectAreadiv = areaDiv
		retData = cAgit.fnGetAgitSmsCont
		if isArray(retData) then
			agitSmsCont = retData(1,0)
			agitSmsRegUser = retData(3,0)
			agitSmsUpdate = retData(4,0)
			mode = "SU"
		else
			mode = "SI"
		end if
 set cAgit = nothing
%>
<script type="text/javascript">
//���õ��
function jsSubmit(){
	if(confirm("������ �����Ͻðڽ��ϱ�?")){
		document.frm.submit();
	}
}
</script>
<h3>����Ʈ �ȳ� ���� ����</h3>
<div class="a" style="text-align:right;">�� �̿� �Ϸ��� ����ڿ��� ���� ���� 3�� �ڵ� �߼۵˴ϴ�.</div>
<form name="frm" method="POST" action="procAgit.asp">
<input type="hidden" name="hidM" value="<%=mode%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
	<td width="70" bgcolor="<%= adminColor("tabletop") %>">
		����Ʈ
	</td>
	<td align="left">
		<select name="areaDiv" class="select">
			<!--<option value="1">���ֵ�</option>-->
			<!--<option value="2">����</option>-->
			<option value="3" <%=chkIIF(areaDiv="3","checked","")%>>����</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">
		����
	</td>
	<td align="left">
		<textarea name="agitSmsCont" style="width:98%; height:350px;"><%=agitSmsCont%></textarea>
	</td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="�� ��" onClick="jsSubmit();"> 
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->