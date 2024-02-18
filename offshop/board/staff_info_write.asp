<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshop_staffcls.asp" -->
<%
dim idx,nstaff,mode

mode = request("mode")

idx = request("idx")
if idx = "" then idx=0

set nstaff = new COffshopStaffDetail
nstaff.GetOffshopStaff idx

dim yyyy1,mm1,dd1,datearr

if mode = "add" then
yyyy1 = Cstr(Year(now()))
mm1 = Cstr(Month(now()))
dd1 = Cstr(day(now()))
else
datearr = split(left(nstaff.Fipsadate,10),"-")
yyyy1 = datearr(0)
mm1 = datearr(1)
dd1 = datearr(2)
end if

%>
<script language="JavaScript">
<!--

	function GoReplyWrite(){
	 var frm = document.boardfrm;
		if (frm.shopid.value == ""){
			alert("샵선택을 해주세요");
			frm.shopid.focus();
		}
		else if (frm.title.value == ""){
			alert("제목을 입력해주세요");
			frm.title.focus();
		}
		else if (frm.contents.value == ""){
			alert("내용을 입력해주세요");
			frm.contents.focus();
		}
		else{
			frm.submit();
		}
	}

//-->
</script>

<table border="0" cellpadding="0" cellspacing="1" width="700" bgcolor="#808080" class="a" align="center">
<form method="post" name="boardfrm" action="<%=uploadUrl%>/linkweb/dostaffwrite.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<tr>
	<td bgcolor="#FFFFFF" height="30" align="center">샵선택</td>
	<td bgcolor="#FFFFFF">&nbsp;
		<select name="shopid">
			<option value="">선택</option>
			<%Call fnOptShopName(nstaff.Fshopid)%>			
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">스텝이름</td>
	<td>&nbsp;<input type="text" name="username" size="50" class="input_b" value="<% = nstaff.Fusername %>"></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">직급</td>
	<td>&nbsp;<select name="slevel">
		<option value="">--선택--</option>
		<%Call fnOptCommonCode("stafflevel",nstaff.Flevel)%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">입사일</td>
	<td>&nbsp;<% DrawOneDateBox yyyy1,mm1,dd1 %></td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">내용</td>
	<td>&nbsp;<textarea name="contents" rows="20" cols="70" class="input_b"><% = nstaff.Fcontents %></textarea></td>
</tr>
<% if mode="add" then %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">첨부사진</td>
	<td>&nbsp;<input type="file" name="file1" size="50" class="input_b"></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">첨부사진</td>
	<td>
	&nbsp;<input type="file" name="file1" size="50" class="input_b"><br>
	&nbsp;<input type="checkbox" name="dl_file1">파일삭제 <img src="<% = nstaff.Ficon1 %>" width="50" height="60" border="0">
	</td>
</tr>
<% end if %>
<% if mode="edit" then %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">사용유무</td>
	<td>
		<input type="radio" name="isusing" value="Y" <% if nstaff.Fisusing = "Y" then response.write "checked" %>>Y <input type="radio" name="isusing" value="N" <% if nstaff.Fisusing = "N" then response.write "checked" %>>N
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" height="30">
	<td align="right" colspan="2"><a href="javascript:GoReplyWrite();"><font color="red">글쓰기</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</form>
</table>

<% set nstaff = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
