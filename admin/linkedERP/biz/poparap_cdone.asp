<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׸�
' History : 2011.04.21 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<%
dim menupos, i, searcharap_cdname, clsAccount, iOpExpPartIdx, arrAccount
menupos = requestCheckvar(getNumeric(Request("menupos")),10)
searcharap_cdname = requestCheckvar(Request("searcharap_cdname"),50) 
iOpExpPartIdx = requestCheckvar(Request("selP"),10)

IF iOpExpPartIdx = "" THEN iOpExpPartIdx = 0

 '�����׸� ����Ʈ
set clsAccount = new COpExpAccount
	clsAccount.FOpExpPartIdx = iOpExpPartIdx
	clsAccount.frectarap_nm = searcharap_cdname
	arrAccount = clsAccount.fnGetArapRegList
set clsAccount = nothing
%>  
 
<script type="text/javascript">

//�˻�
function jsSearch(){  
document.frm.submit();

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">  
	<tr>
	<td><strong>�μ�  ����</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<form name="frm" method="post" action="" style="margin:0px;">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<tr align="center" bgcolor="#FFFFFF" > 
			<td align="left">&nbsp; 
			 �����׸�: <input type="text" name="searcharap_cdname" size="30" value="<%=searcharap_cdname%>">
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>  
<tr>
	<td>
		<!-- ��� �� ���� --> 
		<form name="frmReg" method="post" style="margin:0px;">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   
				<td>�����׸�</td>  
				<td>ó��</td> 
			</tr>
			<%
			If isArray(arrAccount) THEN
			For i = 0 To UBound(arrAccount,2)
			%>
			<tr height=30 align="center" bgcolor="#FFFFFF"> 
				<td align="left">
					 <input type="hidden" name="arap_cd" value="<%=arrAccount(0,i)%>">
					 <input type="hidden" name="arap_name" value="<%=chkIIF(arrAccount(2,i),"[���]","[����]") & arrAccount(1,i)%>">
					 <%=chkIIF(arrAccount(2,i),"[���]","[����]") & arrAccount(1,i)%>
				</td>	 
				<td>
					<input type="button" class="button" value="����" onClick="opener.jsSetarap_cd('<%=arrAccount(0,i)%>','<%=chkIIF(arrAccount(2,i),"[���]","[����]") & arrAccount(1,i)%>');self.close();">
				</td>
			</tr>
			<%
			Next
			ELSE
			%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="2">��ϵ� ������ �����ϴ�.</td>	
			</tr>
		<%END IF%>
		</table>	
		</form>
	</td> 
</tr>  
</table>
<!-- ������ �� -->
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	



	