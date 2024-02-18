<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 비트윈
' History : 2014.10.02 원승현 생성
'			2015.08.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/betweenItemcls.asp"-->
<%
Response.CharSet = "euc-kr"

Dim cDisp, vDepth, vCateCode, vParentCateCode, vCateName, vCateName_E, vUseYN, vSortNo, vResultCount, vdispyn
vDepth			= Request("depth")
vCateCode 		= Request("catecode_s")
vParentCateCode	= Request("parentcatecode")

SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()
	
	vCateName		= cDisp.FCateName
	vUseYN			= cDisp.FUseYN
	vSortNo			= cDisp.FSortNo
	vdispyn	= cDisp.fdispyn
	vResultCount	= cDisp.FResultCount
SET cDisp = Nothing

If vUseYN = "" Then vUseYN = "Y" End If
If vdispyn = "" Then vdispyn = "N" End If
If vSortNo = "" Then vSortNo = "99" End If
%>
<script>
$(function() {
  	$(".rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});
</script>
<input type="hidden" name="parentcatecode" value="<%=vParentCateCode%>">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
<input type="hidden" name="depth" value="<%=vDepth%>">
<input type="hidden" name="completedel" id="completedel" value="">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr>
	<td bgcolor="#F3F3FF" width="70" height="30"></td>
	<td bgcolor="#FFFFFF" align="center"><b>카테고리 <%=CHKIIF(vCateCode="","생성","수정")%></b></td>
</tr>
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">카테고리코드</td>
	<td bgcolor="#FFFFFF"><%=vCateCode%></td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">카테고리명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="catename" style="width:250px;" value="<%=vCateName%>"> (※ 가급적 <u>특수문자는 자제</u>해주시길 바랍니다. 특히 <u>쉼표(,) 홑따옴표(') 쌍따옴표(")</u>)</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" id="useyn_1" value="Y" <%=CHKIIF(vUseYN="Y","checked","")%> /><label for="useyn_1">사용</label>
		<input type="radio" name="useyn" id="useyn_2" value="N" <%=CHKIIF(vUseYN="N","checked","")%> /><label for="useyn_2">사용안함</label>
	</td>
</tr>
<tr>
	<td bgcolor="#F3F3FF" height="30">노출여부</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="dispyn" id="dispyn_1" value="Y" <%=CHKIIF(vdispyn="Y","checked","")%> /><label for="dispyn_1">Y</label>
		<input type="radio" name="dispyn" id="dispyn_2" value="N" <%=CHKIIF(vdispyn="N","checked","")%> /><label for="dispyn_2">N</label>
	</td>
</tr>
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">사용유무</td>
	<td bgcolor="#FFFFFF">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td><input type="button" class="button" value="완전삭제" onClick="jsCateCompleteDel()"><td>
			<td valign="top">&nbsp;※ 주의 : 따로 데이터 저장없이 <b>완전 삭제</b>(복구안됨). 카테고리내 상품 <b>모두 삭제</b>(복구안됨)</td>
		</tr>
		</table>
	</td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">정렬번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" style="width:70px;" value="<%=vSortNo%>"> (※ 숫자가 작을수록 좌측에 나타납니다.)</td>
</tr>
<tr>
	<td id="lyrSbmBtn" bgcolor="#FFFFFF" colspan="2">
		<table width="100%" class="a">
		<tr>
			<td></td>
			<td align="right"><input type="button" value="저  장" onClick="jsSaveDispCate()"></td>
		</tr>
		</table>
		<script>
			$("#lyrSbmBtn input").button();
		</script>
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->