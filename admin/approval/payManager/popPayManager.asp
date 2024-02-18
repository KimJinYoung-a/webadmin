<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : manager regist
' History : 2011.03.26 정윤정  생성
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
//아이디 등록
	function jsRegID(iMode){  
		var winRI = window.open('/admin/approval/eapp/popSetID.asp?iM='+iMode+'&part_sn=8' ,'popAL','width=600, height=400, resizable=yes, scrollbars=yes');
		winRI.focus();
	} 
	
	//등록전 필드 체크
	function jsSubmit(){
	 if(document.frm.sALN.value==""){
	 alert("담당자을 입력해주세요");
	 document.frm.sALN.focus();
	 return false;
	 }
	  
	 return true;
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>결제요청처리 담당자 등록</strong><br><hr width="100%"></td>
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
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">코드 IDX</td>
			<td bgcolor="#FFFFFF" width="180"><%=ipaymanageridx%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">담당업무</td>
			<td bgcolor="#FFFFFF" width="180"> 
			<select name="selPMT">
			<option value="1" <%IF ipaymanagertype="1" THEN%>selected<%END IF%>>최종승인</option>
			<option value="2"  <%IF ipaymanagertype="2" THEN%>selected<%END IF%>>재무회계담당</option>
			</select> 
			</td>
		</tr> 		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">담당자</td>
			<td bgcolor="#FFFFFF" width="180">
				<input type="hidden" name="hidAI" value="<%=trim(suserid)%>">
				<input type="hidden" name="hidAJ" value="<%=sjob_name%>">
				<input type="text" name="sALN" size="30" maxlength="32" value="<%=susername&" "&sjob_name%>" readonly style="border:0;" > &nbsp;<input type="button" name="btnID" value="담당자 등록" onClick="jsRegID(3);" class="button">
			</td>
		</tr> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">기본담당자</td>
			<td bgcolor="#FFFFFF" width="180">
				<input type="checkbox" name="chkD" value="1" <%If blnDef THEN%>checked<%END IF%>> 설정
			</td>
		</tr> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사용유무</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="rdoU" value="1" checked>사용 <input type="radio" name="rdoU" value="0">사용안함</td>
		</tr>	
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="등록" class="button"></td>
</tr>
</form>
</table>
</body>
</html> 

	