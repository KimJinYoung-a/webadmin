<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׸� ���� ����  ��� 
' History : 2011.11.15 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp" --> 
<!-- #include virtual="/lib/classes/linkedERP/arapCls.asp" --> 
<!-- #include virtual="/lib/classes/approval/edmsCls.asp" --> 
<%
Dim clsALE
Dim sMode
Dim iedmsidx,iaraplinkedmsIdx
Dim sARAP_CD,sARAP_NM,sedmsname, blnUsing 
Dim sACC_USE_CD,sACC_NM

sARAP_CD= requestCheckvar(Request("dAc"),10) 
 
 
if iedmsidx = "" THEN iedmsidx = 0	
	sMode ="I"
 

Set clsALE = new CArapLinkEdms 
	clsALE.FARAP_CD = sARAP_CD
  clsALE.fnGetArapLinkEdmsData 	 
  sARAP_CD       	= clsALE.FARAP_CD     	
  sARAP_NM       	= clsALE.FARAP_NM   
  iaraplinkedmsIdx= clsALE.FaraplinkedmsIdx    	
	iedmsidx       	= clsALE.Fedmsidx       	
  sedmsname       = clsALE.Fedmsname  
  sACC_USE_CD 		= clsALE.FACC_USE_CD			 
	sACC_NM 				= clsALE.FACC_NM			
Set clsALE =  nothing
IF iaraplinkedmsIdx <> "" THEN 	sMode ="U" 
 
%>  
  
<script language="javascript">
<!--  
 	
 	//������ �ҷ�����
 	function jsGetEdms(){
 		var winEdms = window.open("/admin/approval/edms/popGetEdms.asp","popEdms","width=600,height=600,resizable=yes, scrollbars=yes");
 			winEdms.focus();
 	}
 	
 	//���� ������ ��������
 	function jsSetEdms(ieidx, sENM){
 		document.frmReg.ieidx.value =ieidx;
 		document.frmReg.sENM.value = sENM; 
 	}
 	
	//����� �ʵ� üũ
	function jsSubmit(){ 
	 
	  if(document.frmReg.ieidx.value==""){
	 alert("�������� �������ּ���"); 
	 return false;
	 } 
	  
	 return true;
	}
	 
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�����׸� ���� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frmReg" method="post" action="procArapEdms.asp" OnSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="idx" value="<%=iaraplinkedmsIdx%>">   
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����׸�</td>
			<td bgcolor="#FFFFFF"><input type="hidden" name="dAC" value="<%=sARAP_CD%>"><input type="text" name="sANM" size="30" value="<%=sARAP_NM%>" readonly style="border:0" ></td>
		</tr> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����������</td>
			<td bgcolor="#FFFFFF">[<%=sACC_USE_CD%>] <%=sACC_NM%> </td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">������</td>
			<td bgcolor="#FFFFFF"><input type="hidden" name="ieidx" value="<%=iedmsidx%>">
				<input type="text" name="sENM" size="30" value="<%=sedmsname%>"  onClick="jsGetEdms();" style="cursor:hand;" >&nbsp;<input type="button" class="button" value="����" onClick="jsGetEdms();" style="cursor:hand;"></td>
		</tr>
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��뿩��</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoU" value="1"  checked >��� <input type="radio" name="rdoU" value="0" >������</td>
		</tr>
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="���" class="button"></td>
</tr>
</form>
</table> 
</body>
</html> 

<!-- #include virtual="/lib/db/dbclose.asp" --> 	
	