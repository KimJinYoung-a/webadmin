<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ī�װ�  ���
' History : 2010.12.23 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp" -->
<%
Dim clsedms 
Dim icategoryidx, icatedepth, scatename,scatecode,ipcateidx,dregdate
Dim sMode,menupos
  
icategoryidx= requestCheckvar(Request("icidx"),10)
ipcateidx	= requestCheckvar(Request("selCL"),10)
menupos		= requestCheckvar(Request("menupos"),10)
 


sMode = "I"

Set clsedms = new Cedms
IF icategoryidx <> "" THEN
	sMode ="U"
	clsedms.Fcategoryidx = icategoryidx
	clsedms.fnGetedmsCategoryData	
	 
	icatedepth  = clsedms.Fcatedepth
  	scatename 	= clsedms.Fcatename	
  	scatecode 	= clsedms.Fcatecode 
  	IF ipcateidx = "" THEN 
  	ipcateidx 	= clsedms.Fpcateidx 
	END IF
  	dregdate 	= clsedms.Fregdate	
 ELSE
 	IF ipcateidx = "" THEN ipcateidx = 0
	IF ipcateidx = 0 THEN
		icatedepth	= 1
	ELSE
		icatedepth	= 2
	END IF

	IF icatedepth = 2 THEN
		clsedms.Fpcateidx =ipcateidx
		scatecode = Format00(3,clsedms.fnGetCatecode)
	END IF 
END IF
 
%>  
<script language="javascript">
<!--
	//ī�װ� ����� ����Ʈ�� �缳��
	function jsChPCategory(){
		document.frmReg.action = "popcategorydata.asp"; 
		document.frmReg.submit();
	}
	
	//����� �ʵ� üũ
	function jsSubmit(){
	 if(document.frmReg.sCN.value==""){
	 alert("ī�װ����� ������ּ���");
	 document.frmReg.sCN.focus();
	 return false;
	 }
	 
	  if(document.frmReg.sCC.value==""){
	 alert("ī�װ��ڵ带 ������ּ���");
	 document.frmReg.sCC.focus();
	 return false;
	 }
	 return true;
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�������� ī�װ� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frmReg" method="post" action="proccategory.asp" OnSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="icidx" value="<%=icategoryidx%>">
		<input type="hidden" name="icd" value="<%=icatedepth%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">����ī�װ�</td>
			<td bgcolor="#FFFFFF" width="180"> 
			<select name="selCL" onChange="jsChPCategory();">
			<option value="0">--�ֻ���--</option>
			<%clsedms.sbGetOptedmsCategory 1,0,ipcateidx %>			
			</select> 
			</td>
		</tr>
		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">ī�װ���</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="sCN" size="30" maxlength="60" value="<%=scatename%>"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">ī�װ��ڵ�</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="sCC" size="3" maxlength="3" value="<%=scatecode%>"></td>
		</tr>
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�������</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="blnU" value="1" checked>��� <input type="radio" name="blnU" value="0">������</td>
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
<%Set clsedms = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	
	