<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : pay manager list
' History : 2011.03.26 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"--> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<%
Dim clsPayManager, arrList, intLoop  
Set clsPayManager = new CPayManager 
	arrList = clsPayManager.fnGetPayManagerList 	
Set clsPayManager = nothing	 
%>  
<!-- #include virtual="/lib/db/dbclose.asp" --> 	
<script language="javascript">
<!--
// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	  
function jsChangeGroup(){
	document.frm.submit();
}
	  
//���ε��
function jsNewReg(){
	var winC = window.open("popPayManager.asp","popC","width=600, height=600, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//����
function jsModReg(payManagerIdx){
	var winC = window.open("popPayManager.asp?ipm="+payManagerIdx,"popC","width=600, height=600, resizable=yes, scrollbars=yes");
	winC.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  

<tr>
	<td><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>���</td>
				<td>���̵�</td>
				<td>�̸�</td>
				<td>��å</td> 
				<td>�⺻�����</td>  	
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=fnGetPayManagerTypeDesc(arrList(2,intLoop))%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(1,intLoop)%></td>			
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(3,intLoop)%></a></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(6,intLoop)%></td>	 
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%IF arrList(7,intLoop) THEN%><font color="red">Y</font><%ELSE%>N<%END IF%></td>	
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="4">��ϵ� ������ �����ϴ�.</td>	
			</tr>
			<%END IF%>
		</table>	
	</td> 
</tr> 
 
</table>
<!-- ������ �� -->
</body>
</html>
 



	