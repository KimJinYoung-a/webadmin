<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : 운영비관리 - 부서 선택
' History : 2011.06.02 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/PartInfoCls.asp"-->
<%
Dim clsPart, arrList, intLoop 
  '구분값 가져오기
Set clsPart = new CPart 
	arrList = clsPart.fnGetPartInfoList 	 
Set clsPart = nothing 
	 
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<script language="javascript">	
<!--
	 
 function jsAddPart(){ 
 	 var oldValue=0;
	 var chkValue="";
	 var strValue="";
	 var arrValue = "";
	 
	 if(opener.document.frm.hidPsn.value!=""){
	 	arrValue = opener.document.frm.hidPsn.value.split(",");
	 	 oldValue = arrValue.length;
	 }
	
	 if(typeof(document.frmReg.chkpsn)=="undefined"){  
	 	return;
	 }
	 
	 if(typeof(document.frmReg.chkpsn.length)=="undefined"){  
	  	if(document.frmReg.chkpsn.checked){ 
		  	if(oldValue == 1){
		  		if(arrValue==document.frmReg.chkpsn.value){
		  			return;
		  		}
		  	}else if(oldValue > 1){
		  		for(j=0;j<oldValue;j++){
		  		if(arrValue[j]==document.frmReg.chkpsn.value){
		  			return;
		  		}
		  		}
		  	} 
	  	 chkValue = document.frmReg.chkpsn.value;
	  	 strValue ="<div id='dP"+document.frmReg.chkpsn.value+"'><input type='hidden' name='hidPSn' value='"+document.frmReg.chkpsn.value+"'>"+document.frmReg.hidPName.value+" <a href='javascript:jsDelPart("+document.frmReg.chkpsn.value+")'>[X]</a></div>"; 
	  	}
	  } 
	 
	 for(i=0;i<document.frmReg.chkpsn.length;i++){
	  var chkReturn=0;
	  if(document.frmReg.chkpsn[i].checked){
	  	if(oldValue == 1){
	  		if(arrValue==document.frmReg.chkpsn[i].value){
	  			chkReturn = 1;
	  		}
	  	}else if(oldValue > 1){
	  		for(j=0;j<oldValue;j++){
	  		if(arrValue[j]==document.frmReg.chkpsn[i].value){
	  			chkReturn = 1;
	  		}
	  		}
	  	} 
	  	if(chkReturn==0){
		   if(chkValue==""){ 
		    chkValue = document.frmReg.chkpsn[i].value;
		   	strValue = "<div id='dP"+document.frmReg.chkpsn[i].value+"'>"+document.frmReg.hidPName[i].value+" <a href='javascript:jsDelPart("+document.frmReg.chkpsn[i].value+")'>[X]</a></div>"
		   }else{
		    chkValue = chkValue +","+ document.frmReg.chkpsn[i].value;
		  	strValue = strValue + "<div id='dP"+document.frmReg.chkpsn[i].value+"'>"+document.frmReg.hidPName[i].value+" <a href='javascript:jsDelPart("+document.frmReg.chkpsn[i].value+")'>[X]</a></div>";
		  	}
		  }	
	  	}
	 }
	 
	 if(chkValue==""){
	 	alert("추가하실 부서를 선택해주세요");
		 return;
	 }
	 
	 if(opener.document.frm.hidPsn.value ==""){
	 opener.document.frm.hidPsn.value = chkValue;
	 }else{
	 opener.document.frm.hidPsn.value = opener.document.frm.hidPsn.value +","+chkValue;
	 }
	 opener.document.all.divPart.innerHTML = opener.document.all.divPart.innerHTML  + strValue; 	
	  self.close();
	 } 
	 
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>부서 선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<tD align="right"><input type="button" value="선택 추가" class="button" onClick="jsAddPart();"></td>
</tr>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frmReg" method="post" action="procPart.asp">  
		<tr height="25" bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td width="100">선택</td>
		 	<td width="100">부서번호</td>  
			<td width="400">부서명</td>  
		</tr>
		<tr>
			<td colspan="3" align="center"  bgcolor="#FFFFFF" >
				<div style="height:450;overflow:scroll;">
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				%>
		<tr  bgcolor="#FFFFFF" align="center">	
		 	 <td width="100"><input type="checkbox" name="chkpsn" value="<%=arrList(0,intLoop)%>"></td>	
		 	 <td width="100"><%=arrList(0,intLoop)%></td>  
		 	 <td width="400"><%=arrList(1,intLoop)%><input type="hidden" name="hidPName" value="<%=arrList(1,intLoop)%>"></td>
		</tr> 
		 	<%	Next
		 	END IF%>
		 		</table>
		 		</div>
		 	</td>	
		</tr>	 
		</form>
		</table>	
	</td> 
</tr> 	 
</table>
<!-- 페이지 끝 -->
</body>
</html>
 



	