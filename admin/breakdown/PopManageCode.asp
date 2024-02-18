<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->
<%
Dim arrList,intLoop
Dim selCodeType
Dim sCodeType,iCodeValue, sCodeDesc, iCodeSort, blnUsing, sCodeComp, sCodeProd
Dim clsCode, sMode

iCodeValue  = requestCheckVar(Request("iCV"),10)	
selCodeType = requestCheckVar(Request("selCT"),20)
sCodeType   = requestCheckVar(Request("sCT"),20)
blnUsing = "Y"
sMode ="I"


 Set clsCode = new CBreakCommonCode  	
 	IF iCodeValue <> "" THEN
 		sMode ="U"
 		clsCode.FCodeType  = sCodeType 
 		clsCode.FCodeValue = iCodeValue
 		clsCode.fnGetBreakCodeCont
 		sCodeComp = clsCode.FCodeComp
 		sCodeProd = clsCode.FCodeProd
 		sCodeDesc = clsCode.FCodeDesc
 		iCodeSort = clsCode.FCodeSort
 		blnUsing  = clsCode.FCodeUsing
   	END IF 		 
   	clsCode.FCodeType = selCodeType
   	arrList = clsCode.fnGetBreakCodeList
 Set clsCode = nothing 

%>
<script language="javascript">
<!--
	// 코드타입 변경이동
	function jsSetCode(iCodeValue,selCodeType){	
		self.location.href = "PopManageCode.asp?iCV="+iCodeValue+"&sCT="+selCodeType+"&selCT="+selCodeType;
	}
	
	//코드 검색
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.sCT.value) {
			alert("대상타입을 선택해 주세요");
			frm.sCT.focus();
			return false;
		}
		
		if(!frm.iCV.value) {
			alert("코드값을 입력해 주세요");
			frm.iCV.focus();
			return false;
		}

		return true;
	}
	
//-->
</script>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//코드 등록 및 수정-->	
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="procCode.asp" onSubmit="return jsRegCode();">	
		<input type="hidden" name="sM" value="<%=sMode%>">
		<input type="hidden" name="sCTO" value="<%=sCodeType%>">
		<tr>
			<td>	+ 코드 등록 및 수정</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">대상타입</td>
					<td bgcolor="#FFFFFF">
						<select name="sCT">
						<option value="">-선택-</option>
						<% sbOptCodeType (sCodeType)%>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">코드값</td>
					<td bgcolor="#FFFFFF"><%IF iCodeValue ="" THEN%><input type="text" size="4" maxlength="10" name="iCV">
						<%ELSE%><%=iCodeValue%><input type="hidden" size="4" maxlength="10" name="iCV" value="<%=iCodeValue%>">
						<%END IF%>
						
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">제조사</td>
					<td bgcolor="#FFFFFF"><input type="text" size="35" name="sCC" value="<%=sCodeComp%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">제품</td>
					<td bgcolor="#FFFFFF"><input type="text" size="35" name="sCP" value="<%=sCodeProd%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">기타</td>
					<td bgcolor="#FFFFFF">
						<textarea name="sCD" cols="34" rows="3"><%=sCodeDesc%></textarea>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">코드 정렬순서</td>
					<td bgcolor="#FFFFFF"><input type="text" size="4" maxlength="10" name="iCS" value="<%=iCodeSort%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">사용여부</td>
					<td bgcolor="#FFFFFF"><input type="radio" value="Y" name="rdoU" <%IF blnUsing ="Y" THEN%>checked<%END IF%>>사용 
					<input type="radio" value="N" name="rdoU" <%IF  blnUsing ="N" THEN%>checked<%END IF%>>사용안함 </td>
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
	<td colspan="2">+ 코드 리스트</td>
</tr>	
<tr>
	<td>대상타입 :
					<select name="selCT" onChange="jsSearch();">
					<option value="">-선택-</option>
					<% sbOptCodeType (selCodeType)%>
					</select>
	</td>
	<td align="right"><a href="javascript:jsSetCode('','<%=selCodeType%>');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">
		<div id="divList" style="height:200px;overflow-y:scroll;">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
			<tr bgcolor="#FFFFFF">
				<td>
					코드 : <%=arrList(1,intLoop)%><br>
					제조사 : <%=arrList(2,intLoop)%><br>
					제품 : <%=arrList(3,intLoop)%><br>
					기타 : <%=arrList(4,intLoop)%>
				</td>
				<td  align="center">
					<input type="button" value="수정" onClick="javascript:jsSetCode('<%=arrList(1,intLoop)%>','<%=arrList(0,intLoop)%>');" class="input_b">
				</td>
			</tr>
			<%Next%>
		<%ELSE%>
		<tr bgcolor="#FFFFFF">
			<td colspan="5" align="center">등록된 내용이 없습니다.</td>
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