<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'##############################################
' History: 2008.03.18 ����
' Description: ������ �ΰŽ� �ֱ� 3���� �� ����Ʈ �ڸ�Ʈ
'##############################################
 Dim clsDF,clsDFCode
 Dim arrList, intLoop
 Dim arrCode
  
 	
'//����Ʈ ��������	
 set clsDF = new CDesignFingers
 	arrList = clsDF.fnGetBestComment 	
 set clsDF = nothing
 
 '//�ΰŽ�����(10)�� �ش��ϴ� �ڵ峻�� �迭�� �ֱ�
 set clsDFCode = new CDesignFingersCode
 	arrCode = clsDFCode.fnGetCommCode(10)	
 set clsDFCode =nothing

 	
%>
<script language="javascript">
<!--
	function jsSearch(){
		document.frmSearch.submit();
	}
	
		
	function jsPopCode(){
		var winCode;
		winCode = window.open('popManageCode.asp','popCode','width=400,height=600');
		winCode.focus();
	}
	
 	function jsSetFile(iDFS){   
 	 var winfile = window.open('','setfile','width=1,height=1');	
 	 	 document.frmFile.iDFS.value = iDFS;
		 document.frmFile.target 	= "setfile";
		 document.frmFile.action 	= "<%=uploadUrl%>/chtml/make_designfingers_FlashText.asp";
		 document.frmFile.submit(); 
		
	 winfile.focus();			 
	}
	
	 //�̹���÷��
 function jsPopAddImg(sFolder,sImgID){
 document.domain ="10x10.co.kr";	
 	var chkIcon = 0;
 	var winImg;
 	var sImgURL;
 	 	 	
 		sImgURL = eval("document.frmBest.img"+sFolder+sImgID).value; 	 	
 		winImg = window.open('popAddImage.asp?sF='+sFolder+'&sID='+sImgID+'&chkI='+chkIcon+'&sIU='+sImgURL,'popImg','width=380,height=150');
 		winImg.focus();
 }
 
 	//����Ʈ ����
 	function jsSetBest(){
 		var frm = document.frmBest;
 		var arrDFS = "";
 		
 		if(typeof(frm.chkID) == "undefined"){
 			alert("���õ� ID�� �����ϴ�.");
 			return;
 		}
 		 		
 		if(typeof(frm.chkID.length) == "undefined"){ 		
 			if(frm.chkID.checked){  				
	 			arrDFS = frm.chkID.value;	 	
	 		}
	 	}else{			 	
	 		for(i=0;i<frm.chkID.length;i++){
	 			if(frm.chkID[i].checked){ 
	 				if(arrDFS ==""){
	 					arrDFS = frm.chkID[i].value;
	 				}else{
	 					arrDFS = arrDFS +"," +frm.chkID[i].value;
	 				}
	 			}
	 		}	 		
	 			
	 	}	
 		
 		if(arrDFS==""){
 		 alert("ID�� ������ �ּ���");	
 		 return;
 		}
 		
 		 var winfile = window.open('','setfile','width=1,height=1');	 	 	
		 document.frmBest.target 	= "setfile";
		 document.frmBest.action 	= "<%=uploadUrl%>/chtml/make_designfingers_BestJS.asp?menupos=<%=menupos%>&arrDFS="+arrDFS;
		 document.frmBest.submit(); 
		
	 	winfile.focus();	
 	}
//-->
</script>
 
<table width="800" border="0" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td colspan="2"> + ���� ����Ʈ ����Ʈ<p>
		<script language="JavaScript" src="<%=staticImgUrl%>/chtml/js/designfingers_Best.js"></script>
	</td>	
</tr>
<tr>
	<td colspan="2"><hr width="100%"></td>
</tR>
<tr>
	<td>+ �ֱ� 3���� ����Ʈ, ���ڸ�Ʈ�� �� ����
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
	    <tr height="40" valign="bottom">       
	        <td align="left">
				<input type="button" value="����ID ����Ʈ����" class="button" onClick="jsSetBest();">
			</td>
			<td align="right">	
				<input type="button" class="button" value="�ΰŽ�����Ʈ" onClick="location.href='listDF.asp?menupos=<%= menupos %>'">
				<% if C_ADMIN_AUTH then %><input type="button" class="button" value="�ڵ����" onclick="javascript:jsPopCode();"><%END IF%>				
			</td>
		</tr>			
		</table>
	</td>
</tr>
<tr>
	<td colspan="2"> 
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a"  bgcolor="#CCCCCC">	
		<form name="frmBest" method="post" action="">			
		<tr bgcolor="#EFEFEF">
			<td width="40" align="center" nowrap>����</td>
			<td width="40" align="center" nowrap>ID	</td>
			<td width="60" align="center" nowrap>����</td>			
			<td align="center">����</td>
			<td width="60" align="center" nowrap>��÷��ǥ��</td>
			<td width="60" align="center" nowrap>�����</td>
			<td width="60" align="center" nowrap>���ڸ�Ʈ��</td>
			<td width="150" align="center" nowrap>���</td>
		</tr>
		<%IF isArray(arrList) THEN%>
		<% For intLoop =0 To UBound(arrList,2) %>	
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="chkID" value="<%=arrList(0,intLoop)%>"></td>
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=fnGetCodeArrDesc(arrCode,arrList(1,intLoop))%></td>			
			<td align="left" ><a href="regDF.asp?iDFS=<%=arrList(0,intLoop)%>&menupos=<%= menupos %>"><%=arrList(2,intLoop)%></a></td>
			<td align="center" ><%=arrList(3,intLoop)%></td>
			<td align="center"><%=FormatDate(arrList(5,intLoop),"0000.00.00")%></td>
			<td align="center" ><%=arrList(7,intLoop)%></td>
			<td align="center"><%IF arrList(6,intLoop) <> "" THEN%><img src="<%=arrList(6,intLoop)%>" width="150"><%END IF%>
			<input type="button" value="���" class="button" onClick="javascript:jsPopAddImg('banner',<%=arrList(0,intLoop)%>);">
			<input type="hidden" name="imgbanner<%=arrList(0,intLoop)%>" value="<%=arrList(6,intLoop)%>">
			</td>
		</tr> 
		<% Next%>
		<%ELSE%>
		<tr bgcolor="#FFFFFF">
			<td colspan="8" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>
		<%END IF%>	
		</form>
		</table>
	</td>		
</tr>
</table>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->