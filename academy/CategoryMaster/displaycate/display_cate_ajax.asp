<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/CategoryMaster/displaycate/classes/displaycateCls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Response.CharSet = "euc-kr"
	
	Dim cDisp, vDepth, vCateCode, vParentCateCode, vCateName, vCateName_E, vUseYN, vSortNo, vResultCount, vJaehuname, vIsNew
	vDepth			= RequestCheckvar(Request("depth"),10)
	vCateCode 		= RequestCheckvar(Request("catecode_s"),10)
	vParentCateCode	= RequestCheckvar(Request("parentcatecode"),10)
	
	SET cDisp = New cDispCate
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateDetail()
	
	vCateName 	= cDisp.FCateName
	vCateName_E	= cDisp.FCateName_E
	vJaehuname = cDisp.FJaehuname
	vUseYN		= cDisp.FUseYN
	vSortNo		= cDisp.FSortNo
	vIsNew		= cDisp.FIsNew
	vResultCount = cDisp.FResultCount
	SET cDisp = Nothing
	
	If vUseYN = "" Then vUseYN = "Y" End If
	If vIsNew = "" Then vIsNew = "x" End If
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
<% If (session("ssBctID") <> "cogusdk") Then %>
<input type="hidden" name="jaehuname" id="completedel" value="<%=vJaehuname%>">
<% End If %>
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
<!--
<tr>
	<td bgcolor="#F3F3FF" height="30">카테고리명(영문)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="catename_e" style="width:250px;" value="<%=vCateName_E%>"> (※ 가급적 <u>특수문자는 자제</u>해주시길 바랍니다. 특히 <u>쉼표(,) 홑따옴표(') 쌍따옴표(")</u>)</td>
</tr>
//-->
<tr>
	<td bgcolor="#F3F3FF" height="30">사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="useyn" id="useyn_1" value="Y" <%=CHKIIF(vUseYN="Y","checked","")%> /><label for="useyn_1" style="cursor:pointer;">사용</label>
		<input type="radio" name="useyn" id="useyn_2" value="N" <%=CHKIIF(vUseYN="N","checked","")%> /><label for="useyn_2" style="cursor:pointer;">사용안함</label>
		&nbsp;※ 주의 : <%=vCateCode%> <b>하위 depth</b> 카테고리 <b>모두</b>, 선택한 값으로 <b>변경</b>됩니다.
	</td>
</tr>
<input type="hidden" name="isnew" value="x">
<!--
<tr>
	<td bgcolor="#F3F3FF" height="30"><img src="http://fiximage.10x10.co.kr/web2013/common/ico_new2.gif" /> 아이콘 사용유무</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isnew" id="isnew_1" value="o" <%=CHKIIF(vIsNew="o","checked","")%> /><label for="isnew_1" style="cursor:pointer;">사용</label>
		<input type="radio" name="isnew" id="isnew_2" value="x" <%=CHKIIF(vIsNew="x","checked","")%> /><label for="isnew_2" style="cursor:pointer;">사용안함</label>
	</td>
</tr>
//-->
<% If vCateCode <> "" Then %>
<tr>
	<td bgcolor="#F3F3FF" height="30">사용유무</td>
	<td bgcolor="#FFFFFF">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td valign="top"><input type="button" value="완전삭제" onClick="jsCateCompleteDel()"><td>
			<td valign="top">&nbsp;※ 주의 : 따로 데이터 저장없이 <b>완전 삭제</b>(복구안됨). 카테고리내 상품 <b>모두 삭제</b>(복구안됨). 브랜드 전시카테고리 <b>삭제</b>(복구안됨).<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>하위 depth의 카테고리</b>가 있을때는 상품을 다른 카테고리로 <b>이동</b>을 하거나 하위 depth 카테고리 <b>삭제</b> 후 실행하세요.
			</td>
		</tr>
		</table>
	</td>
</tr>
<% End If %>
<tr>
	<td bgcolor="#F3F3FF" height="30">정렬번호</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sortno" style="width:70px;" value="<%=vSortNo%>"> (※ 숫자가 작을수록 상단에 나타납니다.)</td>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->