<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׸� ����Ʈ - ������
' History : 2011.11.15 ������  ����
'	jsSetARAP ��ũ��Ʈ �Լ� opener���� �����ؼ� ����ó��
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/arapCls.asp"-->
<%
Dim clsARAP
Dim arrList, intLoop
Dim sARAP_GB,sCASH_FLOW,sARAP_NM

sARAP_GB = requestCheckvar(Request("rdoGB"),3)  
sCASH_FLOW = requestCheckvar(Request("selFlow"),3)  
sARAP_NM = requestCheckvar(Request("sNM"),50)   

Set clsARAP = new COpExpAccount
	 clsARAP.FARAP_GB		=sARAP_GB 	
	 clsARAP.FCASH_FLOW =sCASH_FLOW 
	 clsARAP.FARAP_NM   =sARAP_NM 	
	arrList = clsARAP.fnGetArapOutList 
Set clsARAP = nothing
%>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>�����׸�  ����</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popGetOpExpARAP.asp">  
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
				<td align="left">
					 ����:
						<input type="radio" name="rdoGB" value=""<%IF sARAP_GB="" THEN%>checked<%END IF%>>��ü
						<input type="radio" name="rdoGB" value="1" <%IF sARAP_GB="1" THEN%>checked<%END IF%>>����
						<input type="radio" name="rdoGB" value="2" <%IF sARAP_GB="2" THEN%>checked<%END IF%>>����
						&nbsp; &nbsp; &nbsp;
						�з�:
						<select name="selFlow">
							<option value="">��ü</option>
							<option value="001"  <%IF sCASH_FLOW="001" THEN%>selected<%END IF%>>����</option>
							<option value="002"  <%IF sCASH_FLOW="002" THEN%>selected<%END IF%>>����</option>
							<option value="003"  <%IF sCASH_FLOW="003" THEN%>selected<%END IF%>>�繫</option>
						</select>
					</td>
					
						
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>�����׸��: <input type="text" name="sNM" value="<%=sARAP_NM%>" size="20">
				</td>
			</tr>				
		</form>
		</table>
	</td>
</tr> 
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td>�ڵ�</td> 
		 	<td>����</td>  
			<td>�з�</td>  
			<td>�����׸�</td>  
			<td>�����������</td>  
			<td>����/����ŷ�����</td> 
			<td>����</td>  
		</tr> 
		<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=arrList(0,intLoop)%></td>
		 	<td><%=fnGetARAP_GB(arrList(1,intLoop))%></td> 
		 	<td><%=fnGetARAP_Cash(arrList(3,intLoop))%></td> 
		 	<td><%=arrList(2,intLoop)%></td> 
		 	<td align="left">[<%=arrList(4,intLoop)%>] <%=arrList(5,intLoop)%></td>  
		 	<td><%=arrList(7,intLoop)%></td> 
		 	<td><input type="button" class="button" value="����" onClick="opener.jsSetARAP('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>');self.close();"> </td>
		</tr>  
	<%	Next %>
	<%ELSE%>
	<tr bgcolor="#FFFFFF"  align="center">
			<td colspan="7" align="Center">��ϵ� ������ �����ϴ�.</td>
		</tr>
	<% 	END IF%>
		</table>	
	</td> 
</tr>  
</table>
<!-- ������ �� -->
</body>
</html>
 