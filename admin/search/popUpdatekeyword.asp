<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<%
Dim arritemid, arritemCnt
arritemid	= request("arritemid")
If Right(arritemid,1) = "," Then
	arritemid	= Left(arritemid, Len(arritemid) - 1)
End If
arritemCnt	= Ubound(Split(arritemid, ",")) + 1
%>
<style>
input:-ms-input-placeholder { color: #ADADAD; }
input::-webkit-input-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
input::-moz-placeholder { color: #ADADAD; }
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
function closePop(){
	if (confirm('변경정보를 저장하지 않고 취소하시겠습니까?')){
		self.close();
	}
}
function AllconfirmProcess(){
	if ($("#nextkeyword").val() == ""){
		alert("변경 키워드를 입력하세요");
		return false;
	}
	if ($("#etc").val() == ""){
		alert("비고를 입력하세요");
		return false;
	}

	if( $("#mode").val() == "U") {
		if ($("#prekeyword").val() == ""){
			alert("수정할 키워드를 입력하세요");
			return false;
		}

		if( $("#prekeyword").val().indexOf(",")  > 0) {
			alert("수정할 키워드에 ,는 입력할 수 없습니다.");
			$("#prekeyword").val("");
			return false;
		}
	}

	if( $("#nextkeyword").val().indexOf(",")  > 0) {
		alert("변경 키워드에 ,는 입력할 수 없습니다.");
		$("#nextkeyword").val("");
		return false;
	}

	if (confirm('<%=arritemCnt%>개의 상품 키워드 변경을 일괄적용하시겠습니까?')){
		document.frm.action = "/admin/search/keywordProc.asp"
		document.frm.submit();
	}
}
function chgSelectSH(v){
	if(v == 'U'){
		$("#prekeyword").show();
	}else{
		$("#prekeyword").hide();
		$("#prekeyword").val("");
	}
}
</script>
<table width="100%">
<form name="frm" method="POST">
<input type="hidden" name="cmdparam" value="allchk">
<input type="hidden" name="cksel" value="<%= arritemid %>">
<tr>
	<td align="LEFT"><strong>키워드 변경 정보</strong></td>
</tr>
<tr>
	<td colspan="2">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">변경 구분</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<select name="mode" class="select" id="mode" onchange="chgSelectSH(this.value);">
					<option value="I">등록</option>
					<option value="U">수정</option>
					<option value="D" selected>삭제</option>
				</select>
			</td>
		</tr>
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">변경 키워드</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<input type="text" size="25" class="text" id="prekeyword" name="prekeyword" placeholder="수정할 키워드 입력" style="display:none;">
				<input type="text" size="25" class="text" id="nextkeyword" name="nextkeyword" placeholder="변경할 최종 키워드 입력">
			</td>
		</tr>
		<tr align="LEFT" bgcolor="<%= adminColor("tabletop") %>" height="30">
			<td width="20%">비고</td>
			<td bgcolor="#FFFFFF" align="LEFT">
				<input type="text" size="70" class="text" name="etc" id="etc" placeholder="이력 정보를 간단하게 알 수 있도록 비고 입력">
			</td>
		</tr>
	</td>
</tr>	
</form>
</table>
<br/>
<table width="100%">
<tr>
	<td align="LEFT">* 키워드는 “,”를 제외한 한가지 키워드를 입력해주세요.</td>
</tr>
<tr>
	<td align="center">
		<input type="button" class="button" value="일괄 적용" onclick="AllconfirmProcess();">&nbsp;
		<input type="button" class="button" value="취소" onclick="closePop();">
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->