<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : ������  ����
' History : 2011.05.30 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<%
Dim sMode
Dim clsPart, arrType , clsAccount, arrAccount 
Dim iOpExpPartIdx, iPartTypeIdx, sOpExpPartName, blnUsing,arrPartsn, intLoop, iPartsn
Dim sPartTypeName
Dim intY, dYear, intM, dMonth
iOpExpPartIdx = requestCheckvar(Request("hidOEP"),10) 
sMode ="I"

  '���а� ��������
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList 
Set clsPart = nothing

set clsAccount = new COpExpAccount
	arrAccount = clsAccount.fnGetAccountAll
set clsAccount = nothing  
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<script language="javascript">	
<!--
 	//���
 	function jsPartSubmit(){
 		if(document.frmReg.selPT.value==0 && document.frmReg.sPTN.value==""){
 		alert("���и��� ������ּ���");
 		return;
 		}
 		
 		if( document.frmReg.sPN.value==""){
 		alert("��� ���������� �Է����ּ���");
 		return;
 		}
 		
 		document.frmReg.submit();
 	}
	  
	  //���� 
	  function jsChPT(iValue){
	  if (iValue==0){
	  	document.all.divPT.style.display = "";
	  	}else{
	  	document.all.divPT.style.display = "none";
	  	}
	  }
	  
	  //�μ� �߰�
	  function jsAddPart(){
	    var winPart = window.open("popAddPart.asp","popPart","width=600, height=600");
	    winPart.focus();
	  }
	  
	  //���úμ� ����
	  function jsDelPart(iValue){
	    var arrValue = document.frmReg.hidPsn.value.split(",");  
	    if(typeof(arrValue.length)=="undefined"){
	    	document.frmReg.hidPsn.value  = ""
	    }else{
	    	if(arrValue[0] == iValue){
	    		document.frmReg.hidPsn.value  = document.frmReg.hidPsn.value.replace(iValue,"");	
	    	}else{
	    	 	document.frmReg.hidPsn.value  = document.frmReg.hidPsn.value.replace(","+iValue,"");
	    	}
	    } 
	  	eval("document.all.dP"+iValue).outerHTML = "";
	  }
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>��� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popAccount.asp"> 
			<input type="hidden" name="iCP" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
				<td align="left">
					 ��¥ :
					 <select name="selY">
					 <%For intY = Year(date()) To 2011 STEP -1%>
					 <option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					 </select>��
					  <select name="selM">
					 <%For intM = 1 To 12%>
					 <option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
					 </select>��
					 &nbsp;&nbsp;&nbsp;
					 ��������:
					 <select name="selPT">
					 <option value="">--����--</option>
					 <% sbOptPartType arrType,ipartTypeIdx%>
					 </select>
					 <select name="selP">
					 </select> 
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
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
		<form name="frmReg" method="post" action="procPart.asp"> 
		<input type="hidden" name="hidM" value="<%=sMode%>">  
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
		 	<td>��¥</td>  
			<td>����</td>  
			<td>��ü��</td>  
			<td>�Ա�</td>  
			<td>���</td>  
			<td>����(�󼼳���)</td>   
			<td>ó��</td>  	  
		</tr> 
		<tr bgcolor="#FFFFFF"  align="center">
		 	<td><input type="text" name="iD" size="2"></td>  
			<td><select name="selA">
				<% sbOptAccount arrAccount, ""%>
				</select></td>  
			<td><input type="text" name="sO" size="20"></td>  
			<td><input type="text" name="mIn" size="10"></td>  
			<td><input type="text" name="mOut" size="10"></td> 
			<td><input type="text" name="sDC" size="40" maxlength="200"></td> 
			<td><input type="button" class="button" value="�߰�"></td>  	   	  
		</tr> 
		</form>
		</table>	
	</td> 
</tr>  
</table>
<!-- ������ �� -->
</body>
</html>
 



	