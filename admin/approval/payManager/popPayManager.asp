<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : manager regist
' History : 2011.03.26 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp" -->
<%
Dim clsPayManager
Dim ipaymanageridx, ipaymanagertype, suserid, susername,sjob_name,ijob_sn,blnusing,ipart_sn,blnDef
Dim sMode,menupos
  
ipaymanageridx= requestCheckvar(Request("ipm"),10) 
menupos		= requestCheckvar(Request("menupos"),10) 

sMode = "I"

Set clsPayManager= new CPayManager
IF ipaymanageridx <> "" THEN
	sMode ="U"
	clsPayManager.Fpaymanageridx = ipaymanageridx
	clsPayManager.fnGetPayManagerData	
	  
	suserid  	= clsPayManager.Fuserid 		 
	ipaymanagertype  	= clsPayManager.FpayManagerType  
	susername  	= clsPayManager.Fusername  	     
	sjob_name  		= clsPayManager.Fjob_name 		
	ijob_sn   		= clsPayManager.Fjob_sn 			
	ipart_sn   	= clsPayManager.Fpart_sn		   
  blnusing    = clsPayManager.FisUsing 		
  blnDef      = clsPayManager.FisDef	
 END IF    
%>  
<%Set clsPayManager= nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	
<script language="javascript">
<!-- 
//���̵� ���
	function jsRegID(iMode){  
		var winRI = window.open('/admin/approval/eapp/popSetID.asp?iM='+iMode+'&part_sn=8' ,'popAL','width=600, height=400, resizable=yes, scrollbars=yes');
		winRI.focus();
	} 
	
	//����� �ʵ� üũ
	function jsSubmit(){
	 if(document.frm.sALN.value==""){
	 alert("������� �Է����ּ���");
	 document.frm.sALN.focus();
	 return false;
	 }
	  
	 return true;
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>������ûó�� ����� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frm" method="post" action="procPayManager.asp" OnSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="ipm" value="<%=ipaymanageridx%>"> 
		<input type="hidden" name="menupos" value="<%=menupos%>">		
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ڵ� IDX</td>
			<td bgcolor="#FFFFFF" width="180"><%=ipaymanageridx%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">������</td>
			<td bgcolor="#FFFFFF" width="180"> 
			<select name="selPMT">
			<option value="1" <%IF ipaymanagertype="1" THEN%>selected<%END IF%>>��������</option>
			<option value="2"  <%IF ipaymanagertype="2" THEN%>selected<%END IF%>>�繫ȸ����</option>
			</select> 
			</td>
		</tr> 		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�����</td>
			<td bgcolor="#FFFFFF" width="180">
				<input type="hidden" name="hidAI" value="<%=trim(suserid)%>">
				<input type="hidden" name="hidAJ" value="<%=sjob_name%>">
				<input type="text" name="sALN" size="30" maxlength="32" value="<%=susername&" "&sjob_name%>" readonly style="border:0;" > &nbsp;<input type="button" name="btnID" value="����� ���" onClick="jsRegID(3);" class="button">
			</td>
		</tr> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�⺻�����</td>
			<td bgcolor="#FFFFFF" width="180">
				<input type="checkbox" name="chkD" value="1" <%If blnDef THEN%>checked<%END IF%>> ����
			</td>
		</tr> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�������</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="rdoU" value="1" checked>��� <input type="radio" name="rdoU" value="0">������</td>
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

	