<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��Ƽ3�� �̺�Ʈ ����
' History : 2018.11.05 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/Multi3Cls.asp" -->
<%
dim evt_code, content_idx 
evt_code = request("evtcode")
content_idx = request("contentIdx")
%>

<script type="text/javascript">
function addUnit(){
	var frm = document.unitFrm;
	if(!chkValidation(frm))return false;
	var link = "multi3_process.asp"
	frm.action = link;
	frm.submit();
}
function chkValidation(frm){
	if(frm.unit_class.value==""){
		alert("�з��� �Է����ּ���.");
		return false;
	}
	return true;
}
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<h3>��Ƽ3�� ���������� �߰�</h3>
�̺�Ʈ�ڵ� : <%=evt_code%>
<div>			
	<form name="unitFrm">
	<input type="hidden" name="mode" value="unitadd">
	<input type="hidden" name="evt_code" value="<%=evt_code %>">
	<input type="hidden" name="content_idx" value="<%=content_idx %>">	
	<table width="100%" border="0" align="left" style="margin-top:10px" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">										
		<tr>
			<td width="80" align="center" bgcolor="<%= adminColor("tabletop") %>">
			�з�<b style="color:red">*</b>					
			</td>
			<td bgcolor="#FFFFFF">
			#<input type="text" name="unit_class" size="40" value="" maxlength="32">					
			</td>
		</tr>				
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">���ּ���</td>
			<td bgcolor="#FFFFFF">
			<input type="number" style="width:50px" name="unit_order" size="40" value="" maxlength="32">					
			</td>
		</tr>								
		<tr> 
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ī��</td>
			<td bgcolor="#FFFFFF"><textarea name="unit_main_copy" style="width:90%; height:40px;" value=""></textarea>					
			</td>
		</tr>	
		<tr> 
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF"><textarea name="unit_main_content" style="width:90%; height:40px;" value=""></textarea>					
			</td>
		</tr>		
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">�±�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="tag" value="" maxlength="100"></td>
		</tr>																					
	</table>
	</form>
</div>
<div align="center">
<input type="button" onclick="addUnit();" value="����">
<input type="button" onclick="window.close();" value="���">
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
