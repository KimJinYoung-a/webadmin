<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ڰ��� �����ڵ� ��� 
' History : 2011.03.09 ������  ����
'			2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp" -->
<%
Dim clscomm
Dim icomm_cd, iparentkey, scomm_name, scomm_desc, idispnum,ierpCode, blnactiveYn
Dim sMode,menupos
  
icomm_cd= requestCheckvar(Request("icc"),10) 
menupos		= requestCheckvar(Request("menupos"),10) 

sMode = "I"

Set clscomm= new CcommCode
IF icomm_cd <> "" THEN
	sMode ="U"
	clscomm.Fcomm_cd = icomm_cd
	clscomm.fnGetCommCDData	
	  
	iparentkey  	= clscomm.Fparentkey  
	scomm_name  	= clscomm.Fcomm_name  
	scomm_desc  	= clscomm.Fcomm_desc  
	ierpCode  	= clscomm.FerpCode  	
	idispnum   	= clscomm.Fdispnum   	
	blnactiveYN   	= clscomm.FactiveYN     
END IF
 
%>  
<script type='text/javascript'>
<!-- 
	//����� �ʵ� üũ
	function jsSubmit(){
	 if(document.frmReg.sCN.value==""){
	 alert("�ڵ���� �Է����ּ���");
	 document.frmReg.sCN.focus();
	 return false;
	 }
	  
	 return true;
	}
//-->
</script>
<form name="frmReg" method="post" action="procCommCode.asp" OnSubmit="return jsSubmit();" style="margin:0px;">
<input type="hidden" name="hidM" value="<%=sMode%>">
<input type="hidden" name="icc" value="<%=icomm_cd%>"> 
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�����ڵ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ڵ� IDX</td>
			<td bgcolor="#FFFFFF" width="180"><%=icomm_cd%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�׷��</td>
			<td bgcolor="#FFFFFF" width="180"> 
			<select name="selPK">
			<option value="0">--�ֻ���--</option>
			<%	clsComm.FRectParentKey = iparentkey
				clsComm.sbOptCommCDGroup%>
			</select> 
			</td>
		</tr>
		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ڵ��</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="sCN" size="30" maxlength="32" value="<%= ReplaceBracket(scomm_name) %>"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�߰��ڵ�</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="iEC" size="5" maxlength="10" value="<%=ierpCode%>" style="text-align:right;"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�ڵ弳��</td>
			<td bgcolor="#FFFFFF" width="180"><textarea name="sCD" cols="60" rows="3"><%= ReplaceBracket(scomm_desc) %></textarea></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">ǥ�ü���</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="iDN" size="3"  value="<%=idispnum%>" style="text-align:right;"></td>
		</tr>
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�������</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="blnAYN" value="1" checked>��� <input type="radio" name="blnAYN" value="0">������</td>
		</tr>	
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="���" class="button"></td>
</tr>
</table>
</form>
</body>
</html> 
<%Set clscomm= nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	
	