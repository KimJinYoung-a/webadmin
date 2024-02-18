<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/categoryCls.asp"-->
<%
	'// 변수 선언 //
	dim CateDiv
	dim oCate, lp
	CateDiv = RequestCheckvar(request("CateDiv"),16)

	'// 클래스 선언
	set oCate = new CCate
	oCate.FCateDiv = "code_large"
	oCate.FCurrPage = 1
	oCate.FPageSize = 100
	oCate.GetLargeCateList		

%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
	  
	 <% if CateDiv <> "code_large" then %>
		var sel = document.getElementById("code_large");

		if (sel.selectedIndex == 0)
		{
			alert("대카테고리 코드를 입력해주십시오");
			frm.code_large.focus();
			return false;
		}
	 <% end if %>
	 
	
		if(frm.CateCd.value.length<2)
		{
			alert("코드를 입력해주십시오.\n\n※코드는 2자리입니다.");
			frm.CateCd.focus();
			return false;
		}

		if(!frm.Cate_Name.value)
		{
			alert("코드명을 입력해주십시오.");
			frm.Cate_Name.focus();
			return false;
		}

		// 폼 전송
		return true;
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCategory2012.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="CateDiv" value="<%=CateDiv%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b><%=chkiif(CateDiv="code_large","대","중")%>카테고리 코드 신규등록</b></td>
</tr>
<tr <% if CateDiv = "code_large" then Response.Write "style='display:none'" %>>
	<td align="center" width="120" bgcolor="#E8E8F1">대카테고리명</td>
	<td width="630" bgcolor="#FFFFFF">
		<select name="code_large" id="code_large">
			<option value="">--카테고리 선택--</option>
			<% If oCate.FResultCount > 1 then%>
			<% For lp = 0 To oCate.FResultCount - 1 %>
			<option value="<%= oCate.FCateList(lp).FCateCD %>"><%= db2html(oCate.FCateList(lp).FCateCD_Name) %></option>
			<% Next %>
			<% End If %>
	</td>
</tr>
<tr >
	<td align="center" width="120" bgcolor="#E8E8F1"><%=chkiif(CateDiv="code_large","대","중")%>카테고리 코드</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" name="CateCd" size="2" maxlength="2" value="">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_Name" size="20" maxlength="30"></td>
</tr>
<tr id="lyEngFrm" <% if CateDiv = "code_large" then Response.Write "style='display:none'" %>>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명(영문)</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_NameEng" size="30" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">정렬순서</td>
	<td bgcolor="#FDFDFD"><input type="text" name="orderno" size="2"></td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<%
		set oCate = Nothing
%>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->