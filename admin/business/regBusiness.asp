<%@ language=vbscript %>
<% option explicit  %> 
<%
'###########################################################
' Description : ������  ����� ���� ����Ʈ
' History : 2011.09.26 ������  ����
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/Business/BusinessInfoCls.asp"-->
<%
Dim clsBusi, iBusIdx,sMode
Dim userid,busiNo,busiName,busiCEOName,busiAddr,busiType,busiItem,repName,repEmail,repTel,confirmYn,regdate,delYn,guestOrderserial,useType	
Dim  arrBNo,bN1,bN2,bN3
	iBusIdx = requestCheckvar(Request("iBI"),10) 
	
	sMode ="I"
IF 	iBusIdx <> "" THEN
	sMode = "U" 
Set clsBusi = new CBsuiness  
	clsBusi.FBusiIdx = iBusIdx 
	clsBusi.fnGetBusinessData 
	userid		        =clsBusi.Fuserid		          
	busiNo			    =clsBusi.FbusiNo
	IF busiNo <> "" THEN
		arrBNo = split(busiNo,"-")
		bN1 = arrBNo(0)
		bN2 = arrBNo(1)
		bN3 = arrBNo(2)
	END IF			          
	busiName            =clsBusi.FbusiName                 
	busiCEOName	        =clsBusi.FbusiCEOName	          
	busiAddr            =clsBusi.FbusiAddr                 
	busiType            =clsBusi.FbusiType                 
	busiItem            =clsBusi.FbusiItem                 
	repName			    =clsBusi.FrepName			
	repEmail            =clsBusi.FrepEmail        
	repTel              =clsBusi.FrepTel          
	confirmYn           =clsBusi.FconfirmYn       
	regdate             =clsBusi.Fregdate         
	delYn               =clsBusi.FdelYn           
	guestOrderserial    =clsBusi.FguestOrderserial
	useType             =clsBusi.FuseType         

Set clsBusi = nothing
END IF
%>
<script language="javascript">
<!--
	function jsSubmit(){
		if(jsChkBlank(document.frmReg.sBNa.value)){
 		alert("��ü����  �Է����ּ���");
 		document.frmReg.sBNa.focus();
 		return;
 		}
 		 
 		if(!chkNumeric(document.frmReg.sbN1.value))
		{
			document.frmReg.sbN1.focus();
			return;
		}
		if(document.frmReg.sbN1.value.length<3)
		{
			alert("����ڵ�Ϲ�ȣ 1��° �ڸ��� 3�ڸ� �����Դϴ�.");
			document.frmReg.sbN1.focus();
			return;
		}

		if(!chkNumeric(document.frmReg.sbN2.value))
		{
			document.frmReg.sbN2.focus();
			return;
		}
		if(document.frmReg.sbN2.value.length<1)
		{
			alert("����ڵ�Ϲ�ȣ 2��° �ڸ��� 2�ڸ� �����Դϴ�.");
			document.frmReg.sbN2.focus();
			return;
		}

		if(!chkNumeric(document.frmReg.sbN3.value))
		{
			document.frmReg.sbN3.focus();
			return;
		}
		if(document.frmReg.sbN3.value.length<5)
		{
			alert("����ڵ�Ϲ�ȣ 3��° �ڸ��� 5�ڸ� �����Դϴ�.");
			document.frmReg.sbN3.focus();
			return;
		}
		if(!check_bN(document.frmReg.sbN1.value + document.frmReg.sbN2.value + document.frmReg.sbN3.value))
		{
			alert("�ùٸ� ����ڵ�Ϲ�ȣ�� �ƴմϴ�.\n��Ȯ�� ����ڵ�Ϲ�ȣ�� �Է����ֽʽÿ�.");
			document.frmReg.sbN1.focus();
			return;
		}

 		
 		if(jsChkBlank(document.frmReg.sRN.value)){
 		alert("����ڸ� �Է����ּ���");
 		document.frmReg.sRN.focus();
 		return;
 		}
		document.frmReg.submit();
	}
	// �����Է� �˻�
	function chkNumeric(strNum)
	{
		var chk=0;
		if(!strNum)
		{
			alert("����ڵ�Ϲ�ȣ�� �Է����ֽʽÿ�.");
			return false;
		}
		else
		{
			for (var i = 0; i < strNum.length; i++) {
				ret = strNum.charCodeAt(i);
				if (!((ret > 47) && (ret < 58)))  {
					chk++;
				}
			}
			if(chk>0)
			{
				alert("���ڸ��� �Է����ֽʽÿ�.");
				return false;
			}
			else
				return true;
		}
	}

	// ����ڵ�Ϲ�ȣ üũ
	function check_bN(vencod) {
	        var sum = 0;
	        var getlist =new Array(10);
	        var chkvalue =new Array("1","3","7","1","3","7","1","3","5");
	        for(var i=0; i<10; i++) { getlist[i] = vencod.substring(i, i+1); }
	        for(var i=0; i<9; i++) { sum += getlist[i]*chkvalue[i]; }
	        sum = sum + parseInt((getlist[8]*5)/10);
	        sidliy = sum % 10;
	        sidchk = 0;
	        if(sidliy != 0) { sidchk = 10 - sidliy; }
	        else { sidchk = 0; }
	        if(sidchk != getlist[9]) { return false; }
	        return true;
	}
	
	function jsCancel(){
		location.href = "popBusiness.asp";
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" >  
<tr>
	<td>��ü����<br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frmReg" method="post" action="procBusiness.asp">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="hidBI" value="<%=iBusIdx%>">
		<input type="hidden" name="sUT" value="2">
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">��ü��</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBNa" value="<%=busiName%>" size="20"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ڵ�Ϲ�ȣ</td>
			<td bgcolor="#FFFFFF">
			<input name="sbN1" maxlength="3" type="text" style="width:50px;height:20px;" value="<%=bN1%>">
			-
			<input name="sbN2" maxlength="2" type="text" style="width:30px;height:20px;" value="<%=bN2%>">
			-
			<input name="sbN3" maxlength="5" type="text" style="width:80px;height:20px;" value="<%=bN3%>"></td>
		</tr>
		<tr>  
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">��ǥ��</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sCeo" value="<%=busiCEOName%>" size="10"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRN" value="<%=repName%>" size="10"></td>
		</tr> 
		<tr>  
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����ó</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRT" value="<%=repTel%>" size="15"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">�̸���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRE" value="<%=repEmail%>" size="30"></td> 
		</tr> 
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">������ּ�</td>
			<td  colspan="3" bgcolor="#FFFFFF"><input type="text" name="sBA" value="<%=busiAddr%>" size="60"></td>
		</tr>
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBT" value="<%=busiType%>" size="20"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBI" value="<%=busiItem%>" size="20"></td>
		</tr>  
		<%IF sMode="U" THEN%>
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">��뿩��</td>
			<td  colspan="3" bgcolor="#FFFFFF"><input type="radio" name="rdoD" value="N" <%IF delYN ="" or delYN="N" THEN%>checked<%END IF%>>��� 
			<input type="radio" name="rdoD" value="Y" <%IF delYN="Y" THEN%>checked<%END IF%>>������</td>
		</tr>
		<%END IF%>
		</table>
	</td>
</tr>
<tr>
	<td align="center"><input type="button" class="button_s" value="���" onClick="jsSubmit();">&nbsp;<input type="button" class="button_s" value="���" onClick="jsCancel();"></td>
</tr>
</form>
</table>
</body>
</html>	 
