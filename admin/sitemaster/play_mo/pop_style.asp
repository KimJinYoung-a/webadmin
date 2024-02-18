<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/play/play_moCls.asp" -->

<%
Dim arrList,intLoop, clsCode, sMode, vType, vTypeName, vIsUsing
vType = requestCheckVar(request("playtype"),2)

Set clsCode = new CPlayMoContents
	clsCode.FRectType = vType
	
	arrList = clsCode.fnGetStyleCodeList
	
	vTypeName = clsCode.FOneItem.Ftypename
	vIsUsing = clsCode.FOneItem.Fisusing
Set clsCode = nothing

If vIsUsing = "" Then
	vIsUsing = "Y"
End IF
%>
<script>
function jsSetCode(type){
	self.location.href = "pop_style.asp?playtype="+type+"";
}

function jsRegCode(){
	var frm = document.frmReg;
 
	if(!frm.playtypename.value) {
		alert("분류명을 입력해 주세요");
		frm.playtypename.focus();
		return false;
	}
		
	return true;
}
</script>

<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2">
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="styleProc.asp" onSubmit="return jsRegCode();">	
		<input type="hidden" name="playtype" value="<%=vType%>">
		<tr>
			<td>	+ 분류 등록 및 수정</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#EFEFEF"   align="center">분류명</td>
					<td bgcolor="#FFFFFF"><input type="text" size="15" name="playtypename" value="<%=vTypeName%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">사용여부</td>
					<td bgcolor="#FFFFFF"><input type="radio" value="Y" name="isusing" <%IF vIsUsing ="Y" THEN%>checked<%END IF%>>사용 
					<input type="radio" value="N" name="isusing" <%IF  vIsUsing ="N" THEN%>checked<%END IF%>>사용안함 </td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif"> 
				<a href="javascript:jsSetCode('')"><img src="/images/icon_cancel.gif" border="0"></a></td>
		</tr>	
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<form name="frmSearch" method="post" action="typeProc.asp">
	<td colspan="2">+ style코드 리스트</td>
</tr>	
<tr>
	<td></td>
	<td align="right"><a href="javascript:jsSetCode('');"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr>
	<td colspan="2">
		<div id="divList" style="height:305px;overflow-y:scroll;">	
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF">
			<td  align="center" width="50">분류코드</td>
			<td  align="center">분류명</td>
			<td  align="center">사용여부</td>
			<td  align="center">처리</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
		<tr bgcolor="#FFFFFF">
			<td  align="center"><%=arrList(0,intLoop)%></td>
			<td  align="center"><%=arrList(1,intLoop)%></td>
			<td  align="center"><%=arrList(2,intLoop)%></td>
			<td  align="center">
				<input type="button" value="수정" onClick="javascript:jsSetCode('<%=arrList(0,intLoop)%>');" class="input_b">
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