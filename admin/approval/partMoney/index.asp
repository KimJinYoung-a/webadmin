<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڱݰ��� �μ�
' History : 2011.04.21 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"--> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim clsBS
Dim arrList, intLoop 
Dim sUSEYN,sBS_NM
sBS_NM = requestCheckvar(Request("sBS_NM"),100)  
sUSEYN = requestCheckvar(Request("sUSEYN"),3) 
 
Set clsBS = new CBizSection 
	clsBS.FBS_NM 	= sBS_NM
	clsBS.FUSE_YN = sUSEYN  
	arrList = clsBS.fnGetBizSectionList  
Set clsBS = nothing	  
%>  
 
<script language="javascript">
<!-- 
	 
	//����
	function jsModReg(eapppartidx){
		var winC = window.open("popPart.asp?iepidx="+eapppartidx,"popC","width=600, height=600, resizable=yes, scrollbars=yes");
		winC.focus();
	}
  
   //�˻�
   function jsSearch(){  
    document.frm.submit();
   }
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<form name="frm" method="post" action="index.asp">
		<input type="hidden" name="menupos" value="<%=menupos%>"> 
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="100" bgcolor="#EEEEEE">�˻� ����</td>
			<td align="left">&nbsp; 
			 �μ���: <input type="text" name="sBS_NM" size="20" value="<%=sBS_NM%>">
		 	&nbsp;
		 	<input type="checkbox" name="sUSEYN" value="A" <%IF cStr(sUSEYN) ="A" THEN%>checked<%END IF%>>��Ȱ������
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
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
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   
				<td width="50">Ȱ������</td>  
				<td>�μ���</td> 
			  <td>���ͺμ�����</td> 
			  <td>���ÿ���</td> 
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="<%IF arrList(3,intLoop) ="N" THEN%>#EFEFEF<%ELSE%>#FFFFFF<%END IF%>">	 
				<td><%=arrList(3,intLoop)%></td>
				<td align="left"><%IF arrList(2,intLoop) <> "" THEN%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					�� <%END IF%><%=arrList(0,intLoop)%>&nbsp; <%=arrList(1,intLoop)%></td>  
					<td></td>
					<td></td>
			</tr>
		<%		Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="2">��ϵ� ������ �����ϴ�.</td>	
			</tr>
		<%END IF%>
		</table>	
	</td> 
</tr>  
</table>
<!-- ������ �� -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" --> 	



	