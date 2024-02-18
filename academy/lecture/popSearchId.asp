<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
dim searchKey, searchString
searchKey = RequestCheckvar(request("searchKey"),16)
searchString = request("searchString")
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if searchKey="" then searchKey = "userId"

dim olecture
set olecture = new CuserInfo

if searchString<>"" then
	olecture.FRectsearchKey = searchKey
	olecture.FRectsearchString = searchString

	olecture.GetUserList
end if

dim i
%>
<script language='javascript'>
<!--
	function chk_form(frm)
	{
		if(frm.searchString.value.length<2)
		{
			alert("2자 이상의 검색어를 입력해주십시요.");
			frm.searchString.focus();
			return false;
		}
	}

	function inputData(id)
	{
		var Op = opener.frmlec;
		var Nw = document.frmSearch;

		if(Nw.userid.length>1)
		{
			Op.buy_userid.value = Nw.userid[id].value;
			Op.buy_name.value = Nw.username[id].value;
			if(Nw.userphone[id].value)
			{
				Op.buy_phone1.value = Nw.userphone[id].value.split("-")[0];
				Op.buy_phone2.value = Nw.userphone[id].value.split("-")[1];
				Op.buy_phone3.value = Nw.userphone[id].value.split("-")[2];
			}
			if(Nw.usercell[id].value)
			{
				Op.buy_hp1.value = Nw.usercell[id].value.split("-")[0];
				Op.buy_hp2.value = Nw.usercell[id].value.split("-")[1];
				Op.buy_hp3.value = Nw.usercell[id].value.split("-")[2];
			}
			Op.buy_level.value = Nw.userlevel[id].value;
			Op.buy_email.value = Nw.usermail[id].value;
		}
		else
		{
			Op.buy_userid.value = Nw.userid.value;
			Op.buy_name.value = Nw.username.value;
			if(Nw.userphone.value)
			{
				Op.buy_phone1.value = Nw.userphone.value.split("-")[0];
				Op.buy_phone2.value = Nw.userphone.value.split("-")[1];
				Op.buy_phone3.value = Nw.userphone.value.split("-")[2];
			}
			if(Nw.usercell.value)
			{
				Op.buy_hp1.value = Nw.usercell.value.split("-")[0];
				Op.buy_hp2.value = Nw.usercell.value.split("-")[1];
				Op.buy_hp3.value = Nw.usercell.value.split("-")[2];
			}
			Op.buy_level.value = Nw.userlevel.value;
			Op.buy_email.value = Nw.usermail.value;
		}

		self.close();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b>회원 검색</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="frmSearch" method="POST" onSubmit="return chk_form(this)" action="popSearchId.asp">
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="5">
		<select name="searchKey">
			<option value="userId">회원ID</option>
			<option value="username">회원이름</option>
		</select>
		<input type="text" name="searchString" value="<%=searchString%>" size="12">
		<input type="image" src="/images/icon_search.gif" align="absmiddle">
		<script language="javascript">
			document.frmSearch.searchKey.value="<%=searchKey%>";
		</script>
	</td>
</tr>
<% if searchString<>"" then %>
<tr bgcolor="#EEEEFF" align="center">
	<td>아이디</td>
	<td>이름</td>
	<td>이메일</td>
	<td>가입일</td>
	<td>선택</td>
</tr>
<%	if olecture.FTotalCount=0 then %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="5" height="70">검색된 회원이 없습니다.</td>
</tr>
<%
	else
		for i=0 to olecture.FTotalCount-1
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%=olecture.FItemList(i).FuserId%></td>
	<td><%=olecture.FItemList(i).Fusername%></td>
	<td><%=olecture.FItemList(i).Fusermail%></td>
	<td><%=FormatDate(olecture.FItemList(i).Fregdate,"00.00.00")%></td>
	<td><img src="/images/icon_go.gif" onClick="inputData(<%=i%>)" style="cursor:pointer" align="absmiddle"></td>
</tr>
<input type="hidden" name="userid" value="<%=olecture.FItemList(i).FuserId%>">
<input type="hidden" name="username" value="<%=olecture.FItemList(i).Fusername%>">
<input type="hidden" name="userphone" value="<%=olecture.FItemList(i).Fuserphone%>">
<input type="hidden" name="usercell" value="<%=olecture.FItemList(i).Fusercell%>">
<input type="hidden" name="userlevel" value="<%=olecture.FItemList(i).Fuserlevel%>">
<input type="hidden" name="usermail" value="<%=olecture.FItemList(i).Fusermail%>">
<%		next
	end if
  end if
%>
</form>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="F4F4F4" height="22">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td valign="bottom" align="center" bgcolor="F4F4F4">
		<img src="/images/icon_cancel.gif" onClick="self.close()" style="cursor:pointer" align="absbottom">
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" bgcolor="F4F4F4" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<%
	set olecture = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->