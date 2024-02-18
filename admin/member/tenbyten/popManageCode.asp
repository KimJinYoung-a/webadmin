<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.03.17
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/event/eventCommonCls.asp"-->
<%
Dim arrList,intLoop
Dim selCodeType
Dim sCodeType,iCodeValue, sCodeDesc, iCodeSort, blnUsing, sCodeDispYN
Dim clsCode, sMode

iCodeValue  = requestCheckVar(Request("iCV"),10)	
selCodeType = requestCheckVar(Request("selCT"),20)
sCodeType   = requestCheckVar(Request("sCT"),20)
blnUsing = "Y"
sCodeDispYN ="Y"
sMode ="I"

IF selCodeType = "" THEN selCodeType = "jobkind"
 Set clsCode = new CEventCommonCode  	
 	IF iCodeValue <> "" THEN
 		sMode ="U"
 		clsCode.FCodeType  = sCodeType 
 		clsCode.FCodeValue = iCodeValue
 		clsCode.fnGetEventCodeCont 		
 		sCodeDesc = clsCode.FCodeDesc
 		iCodeSort = clsCode.FCodeSort
 		blnUsing  = clsCode.FCodeUsing
 		sCodeDispYN=clsCode.FCodeDispYN
   	END IF
 		 
   	clsCode.FCodeType = selCodeType
   	arrList = clsCode.fnGetEventCodeList
 Set clsCode = nothing 
IF isnull(iCodeSort) or iCodeSort = "" THEN iCodeSort = 0


Sub sbOptCodeType2(ByVal selCodeType)
    Dim arrSelCode(19), i
    arrSelCode(0)	= Split("jobkind|����","|")
    arrSelCode(1)	= Split("placekind|�ٹ���","|")
    for i=0 to ubound(arrSelCode)
        if isArray(arrSelCode(i)) then
        Response.Write "<option value=""" & arrSelCode(i)(0) &""" " & chkIIF(Cstr(selCodeType)=arrSelCode(i)(0),"selected","") & ">" & arrSelCode(i)(1) & "</option>" & vbCrlF
        end if
    next
End Sub
%>
<script language="javascript">
<!--
	// �ڵ�Ÿ�� �����̵�
	function jsSetCode(iCodeValue,selCodeType){	
		self.location.href = "PopManageCode.asp?iCV="+iCodeValue+"&sCT="+selCodeType+"&selCT="+selCodeType;
	}
	
	//�ڵ� �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.iCV.value) {
			alert("�ڵ尪�� �Է��� �ּ���");
			frm.iCV.focus();
			return false;
		}
			 
		if(!frm.sCD.value) {
			alert("�ڵ���� �Է��� �ּ���");
			frm.sCD.focus();
			return false;
		}
			
		return true;
	}
	
//-->
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->	
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="procCode.asp" onSubmit="return jsRegCode();">	
		<input type="hidden" name="sM" value="<%=sMode%>">			  
		<tr>			
			<td>	+ �ڵ� ��� �� ����</td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">										
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ�Ÿ��</td>
					<td bgcolor="#FFFFFF">
						<select name="sCT" class="select">						
						<% sbOptCodeType2(sCodeType)%>					
						</select>				
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ尪</td>
					<td bgcolor="#FFFFFF"><%IF iCodeValue ="" THEN%><input type="text" size="4" maxlength="10" name="iCV">
						<%ELSE%><%=iCodeValue%><input type="hidden" size="4" maxlength="10" name="iCV" value="<%=iCodeValue%>">
						<%END IF%>
						
					</td>
				</tr>					
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ��</td>
					<td bgcolor="#FFFFFF"><input type="text" size="15" maxlength="16" name="sCD" value="<%=sCodeDesc%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ� ���ļ���</td>
					<td bgcolor="#FFFFFF"><input type="text" size="4" maxlength="10" name="iCS" value="<%=iCodeSort%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ� ���ÿ���</td>
					<td bgcolor="#FFFFFF"><input type="radio" value="Y" name="rdoD" <%IF sCodeDispYN ="Y" THEN%>checked<%END IF%>>���� 
					<input type="radio" value="N" name="rdoD" <%IF  sCodeDispYN ="N" THEN%>checked<%END IF%>>���þ��� </td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">��뿩��</td>
					<td bgcolor="#FFFFFF"><input type="radio" value="Y" name="rdoU" <%IF blnUsing ="Y" THEN%>checked<%END IF%>>��� 
					<input type="radio" value="N" name="rdoU" <%IF  blnUsing ="N" THEN%>checked<%END IF%>>������ </td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"> 
				<a href="javascript:jsSetCode('','<%=selCodeType%>')"><img src="/images/icon_cancel.gif" border="0"></a></td>
		</tr>	
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<form name="frmSearch" method="post" action="PopManageCode.asp">
	<td colspan="2">+ �ڵ� ����Ʈ</td>
</tr>	
<tr>
	<td>�ڵ�Ÿ�� :
					<select name="selCT" onChange="jsSearch();" class="select">						
					<% sbOptCodeType2(selCodeType)%>					
					</select>	
	</td>
	<td align="right"><a href="javascript:jsSetCode('','<%=selCodeType%>');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">	
		<div id="divList" style="height:305px;overflow-y:scroll;">	
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
		<tr bgcolor="#EFEFEF">			
			<td  align="center" width="50">�ڵ尪</td>
			<td  align="center">�ڵ��</td>
			<td  align="center">���ļ���</td>
			<td  align="center">���ÿ���</td>
			<td  align="center">��뿩��</td>
			<td  align="center">ó��</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
		<tr bgcolor="#FFFFFF">			
			<td  align="center"><%=arrList(1,intLoop)%></td>
			<td  align="center"><%=arrList(2,intLoop)%></td> 
			<td  align="center"><%=arrList(4,intLoop)%></td>
			<td  align="center"><font color="<%=ynColor(arrList(5,intLoop))%>"><%=arrList(5,intLoop)%></font></td>
			<td  align="center"><font color="<%=ynColor(arrList(3,intLoop))%>"><%=arrList(3,intLoop)%></font></td>
			<td  align="center">
				<input type="button" value="����" onClick="javascript:jsSetCode('<%=arrList(1,intLoop)%>','<%=arrList(0,intLoop)%>');" class="input_b">				
			</td>
		</tr>
			<%Next%>
		<%ELSE%>	
		<tr bgcolor="#FFFFFF">			
			<td colspan="6" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>	
		<%End if%>		
		</table>
		</div>
	</td>
	</form>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->