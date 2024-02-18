<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<%
	Dim page, Suid, chkIni, chkCmt, CplDiv, lp, oWdw

	page = requestCheckvar(Request("page"),10)
	if page="" then page=1
	Suid = requestCheckvar(Request("Suid"),32)
	chkIni = requestCheckvar(Request("chkIni"),32)
	chkCmt = requestCheckvar(Request("chkCmt"),32)
	CplDiv = requestCheckvar(Request("CplDiv"),32)

	'// 클래스 선언
	set oWdw = new CwithDraw
	oWdw.FCurrPage = page
	oWdw.FPageSize = 20
	oWdw.FRectUserId = Suid
	oWdw.FRectChkInit = chkIni
	oWdw.FRectChkCmt = chkCmt
	oWdw.FRectCplDiv = CplDiv

	oWdw.GetUserList
%>
<script language="javascript">
<!--
	function searchUser()
	{
		if(frm.Suid.value!=""&&frm.Suid.value.length<2) {
			alert("아이디는 적어도 두글자 이상 입력해주세요.");
			return false;
		}
		else {
			return true;
		}
	}

	function goPage(pg)
	{
		frm.page.value= pg;
		frm.submit();
	}
//-->
</script>
<!-- 검색 시작 -->
<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm" method="get" onsubmit="return searchUser()" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>&nbsp;</td>
	<td align="left">
		회원아이디
		<input type="text" name="Suid" size="20" value="<%=Suid%>"> /
		<label><input type="checkbox" name="chkIni" value="on" <% if chkIni="on" then Response.Write "checked"%>>앞글자</label> /
		<label><input type="checkbox" name="chkCmt" value="on" <% if chkCmt="on" then Response.Write "checked"%>>탈퇴이유 기재</label> /
		<select name="CplDiv" class="select">
		<option value="">::탈퇴구분::</option>
		<option value="01">상품품질불만</option>
		<option value="02">이용빈도낮음</option>
		<option value="03">배송지연</option>
		<option value="04">개인정보유출우려</option>
		<option value="05">교환/환불/품질불만</option>
		<option value="06">기타</option>
		<option value="07">A/S불만</option>
		<option value="not">미지정</option>
		</select> &nbsp;
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
		<script type="text/javascript">
		frm.CplDiv.value="<%=CplDiv%>";
		</script>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 상단 띠 시작 -->
<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">&nbsp;</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 상단 띠 끝 -->
<!-- 메인 목록 시작 -->
<table width="98%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="100">탈퇴일시</td>
	<td width="110">아이디</td>
	<td width="80">생년월일</td>
	<td width="120">탈퇴구분</td>
	<td>탈퇴이유</td>
</tr>
<%
	if oWdw.FTotalCount>0 then
		for lp=0 to oWdw.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100"><%= Replace(oWdw.FItemList(lp).Fregdate," 오","<br>오") %></td>
	<td width="110"><%= LEFT(oWdw.FItemList(lp).Fuid,LEN(oWdw.FItemList(lp).Fuid)-2) %>**</td>
	<td width="80"><%= LEFT(oWdw.FItemList(lp).Fjumin1,2) %>****</td>
	<td width="120"><%= oWdw.FItemList(lp).FcomplainDiv %></td>
	<td align="left"><%= oWdw.FItemList(lp).FcomplainText %></td>
</tr>
<%
		next
	else
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">검색된 탈퇴정보가 없습니다.</td>
</tr>
<%	end if %>
</table>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="center" valign="bottom">
	<!-- 페이지 시작 -->
	<%
		if oWdw.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oWdw.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oWdw.StartScrollPage to oWdw.FScrollCount + oWdw.StartScrollPage - 1

			if lp>oWdw.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oWdw.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->