<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 공통코드 등록 
' History : 2011.03.09 정윤정  생성
'			2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
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
	//등록전 필드 체크
	function jsSubmit(){
	 if(document.frmReg.sCN.value==""){
	 alert("코드명을 입력해주세요");
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
	<td><strong>공통코드등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">코드 IDX</td>
			<td bgcolor="#FFFFFF" width="180"><%=icomm_cd%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">그룹명</td>
			<td bgcolor="#FFFFFF" width="180"> 
			<select name="selPK">
			<option value="0">--최상위--</option>
			<%	clsComm.FRectParentKey = iparentkey
				clsComm.sbOptCommCDGroup%>
			</select> 
			</td>
		</tr>
		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">코드명</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="sCN" size="30" maxlength="32" value="<%= ReplaceBracket(scomm_name) %>"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">추가코드</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="iEC" size="5" maxlength="10" value="<%=ierpCode%>" style="text-align:right;"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">코드설명</td>
			<td bgcolor="#FFFFFF" width="180"><textarea name="sCD" cols="60" rows="3"><%= ReplaceBracket(scomm_desc) %></textarea></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">표시순서</td>
			<td bgcolor="#FFFFFF" width="180"><input type="text" name="iDN" size="3"  value="<%=idispnum%>" style="text-align:right;"></td>
		</tr>
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사용유무</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="blnAYN" value="1" checked>사용 <input type="radio" name="blnAYN" value="0">사용안함</td>
		</tr>	
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="등록" class="button"></td>
</tr>
</table>
</form>
</body>
</html> 
<%Set clscomm= nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 	
	