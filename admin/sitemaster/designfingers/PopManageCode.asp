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
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
Dim arrList,intLoop
Dim selPCodeSeq
Dim iPCodeSeq, iCodeSeq, sCodeDesc, iCodeSort, blnUsing
Dim clsDFCode

iCodeSeq  =  requestCheckVar(Request("iCS"),10)	
selPCodeSeq = requestCheckVar(Request("sPCS"),10)	
blnUsing = True

IF selPCodeSeq= "" THEN selPCodeSeq = 0

 Set clsDFCode = new CDesignFingersCode  	
 	IF iCodeSeq <> "" THEN
 		clsDFCode.fnGetCodeCont(iCodeSeq)   
 		iPCodeSeq = clsDFCode.FPCodeSeq
 		sCodeDesc = clsDFCode.FCodeDesc
 		iCodeSort = clsDFCode.FCodeSort
 		blnUsing  = clsDFCode.FIsUsing
   	END IF	
   	arrList = clsDFCode.fnGetCommCode(selPCodeSeq)   
 Set clsDFCode = nothing
 
%>
<script language="javascript">
<!--
 	// �ڵ�Ÿ�� �����̵�
	function jsSetCode(iCodeSeq,selPCodeSeq){	
		self.location.href = "PopManageCode.asp?iCS="+iCodeSeq+"&sPCS="+selPCodeSeq;
	}
	
	//�ڵ� �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.sCD.value) {
			alert("�ڵ���� �Է��� �ּ���");
			frm.sCD.focus();
			return false;
		}
			
		return true;
	}
	
//-->
</script>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->	
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="procDF.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="sM" value="C">
		<input type="hidden" name="iCS" value="<%=iCodeSeq%>">
		<tr>			
			<td>	+ �ڵ� ��� �� ����</td>
		</tr>	
		<tr>
			<td>	
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">		
				<% IF iCodeSeq <> "" THEN%>	
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ��ȣ</td>
					<td bgcolor="#FFFFFF"><%=iCodeSeq%></td>
				</tr>	
				<%END IF%>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�ڵ� �з�</td>
					<td bgcolor="#FFFFFF">
						<select name="selPCS">
						<option value="0">-�ֻ���-</option>
						<%sbOptCommCode 0, iPCodeSeq%>				
						</select>				
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ��</td>
					<td bgcolor="#FFFFFF"><input type="text" size="15" maxlength="22" name="sCD" value="<%=sCodeDesc%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�ڵ� ���ļ���</td>
					<td bgcolor="#FFFFFF"><input type="text" size="4" maxlength="10" name="iCSort" value="<%=iCodeSort%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">��뿩��</td>
					<td bgcolor="#FFFFFF"><input type="radio" value="1" name="rdoU" <%IF blnUsing THEN%>checked<%END IF%>>��� 
					<input type="radio" value="0" name="rdoU" <%IF not blnUsing THEN%>checked<%END IF%>>������ </td>
				</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"> 
				<a href="javascript:jsSetCode('','<%=selPCodeSeq%>')"><img src="/images/icon_cancel.gif" border="0"></a></td>
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
	<td>�ڵ�з� : <select name="sPCS" onChange="jsSearch();">
						<option value="0" <%if selPCodeSeq ="0" THEN%>selected<%END IF%>>-�ֻ���-</option>
						<%sbOptCommCode 0, selPCodeSeq%>				
						</select>	
	</td>
	<td align="right"><a href="javascript:jsSetCode('','<%=selPCodeSeq%>');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">	
		<div id="divList" style="height:345px;overflow-y:scroll;">	
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
		<tr bgcolor="#EFEFEF">			
			<td  align="center" width="50">�ڵ��ȣ</td>
			<td  align="center">�ڵ��</td>
			<td  align="center">���ļ���</td>
			<td  align="center">��뿩��</td>
			<td  align="center">ó��</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
		<tr bgcolor="#FFFFFF">			
			<td  align="center"><%=arrList(0,intLoop)%></td>
			<td  align="center"><%=arrList(1,intLoop)%></td>
			<td  align="center"><%=arrList(3,intLoop)%></td>
			<td  align="center"><%IF arrList(4,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
			<td  align="center">
				<input type="button" value="����" onClick="javascript:jsSetCode('<%=arrList(0,intLoop)%>','<%=selPCodeSeq%>');" class="input_b">				
			</td>
		</tr>
			<%Next%>
		<%ELSE%>	
		<tr bgcolor="#FFFFFF">			
			<td colspan="5" align="center">��ϵ� ������ �����ϴ�.</td>
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