<%@ language=vbscript %>
<% option explicit  %> 
<%
'###########################################################
' Description : 운영비관리  사업자 정보 리스트
' History : 2011.09.26 정윤정  생성
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
 		alert("업체명을  입력해주세요");
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
			alert("사업자등록번호 1번째 자리는 3자리 숫자입니다.");
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
			alert("사업자등록번호 2번째 자리는 2자리 숫자입니다.");
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
			alert("사업자등록번호 3번째 자리는 5자리 숫자입니다.");
			document.frmReg.sbN3.focus();
			return;
		}
		if(!check_bN(document.frmReg.sbN1.value + document.frmReg.sbN2.value + document.frmReg.sbN3.value))
		{
			alert("올바른 사업자등록번호가 아닙니다.\n정확한 사업자등록번호를 입력해주십시오.");
			document.frmReg.sbN1.focus();
			return;
		}

 		
 		if(jsChkBlank(document.frmReg.sRN.value)){
 		alert("담당자를 입력해주세요");
 		document.frmReg.sRN.focus();
 		return;
 		}
		document.frmReg.submit();
	}
	// 숫자입력 검사
	function chkNumeric(strNum)
	{
		var chk=0;
		if(!strNum)
		{
			alert("사업자등록번호를 입력해주십시오.");
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
				alert("숫자만을 입력해주십시오.");
				return false;
			}
			else
				return true;
		}
	}

	// 사업자등록번호 체크
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
	<td>업체정보<br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frmReg" method="post" action="procBusiness.asp">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="hidBI" value="<%=iBusIdx%>">
		<input type="hidden" name="sUT" value="2">
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">업체명</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBNa" value="<%=busiName%>" size="20"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사업자등록번호</td>
			<td bgcolor="#FFFFFF">
			<input name="sbN1" maxlength="3" type="text" style="width:50px;height:20px;" value="<%=bN1%>">
			-
			<input name="sbN2" maxlength="2" type="text" style="width:30px;height:20px;" value="<%=bN2%>">
			-
			<input name="sbN3" maxlength="5" type="text" style="width:80px;height:20px;" value="<%=bN3%>"></td>
		</tr>
		<tr>  
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">대표자</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sCeo" value="<%=busiCEOName%>" size="10"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">담당자</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRN" value="<%=repName%>" size="10"></td>
		</tr> 
		<tr>  
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">연락처</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRT" value="<%=repTel%>" size="15"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">이메일</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sRE" value="<%=repEmail%>" size="30"></td> 
		</tr> 
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사업장주소</td>
			<td  colspan="3" bgcolor="#FFFFFF"><input type="text" name="sBA" value="<%=busiAddr%>" size="60"></td>
		</tr>
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">업태</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBT" value="<%=busiType%>" size="20"></td>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">종목</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sBI" value="<%=busiItem%>" size="20"></td>
		</tr>  
		<%IF sMode="U" THEN%>
		<tr> 
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">사용여부</td>
			<td  colspan="3" bgcolor="#FFFFFF"><input type="radio" name="rdoD" value="N" <%IF delYN ="" or delYN="N" THEN%>checked<%END IF%>>사용 
			<input type="radio" name="rdoD" value="Y" <%IF delYN="Y" THEN%>checked<%END IF%>>사용안함</td>
		</tr>
		<%END IF%>
		</table>
	</td>
</tr>
<tr>
	<td align="center"><input type="button" class="button_s" value="등록" onClick="jsSubmit();">&nbsp;<input type="button" class="button_s" value="취소" onClick="jsCancel();"></td>
</tr>
</form>
</table>
</body>
</html>	 
