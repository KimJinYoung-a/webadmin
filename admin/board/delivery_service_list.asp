<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  택배업체관리
' History : 2007.10.29 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<Script language="javascript">
function chkEditFrm(frm){
	if (frm.divname.value==''){
		alert('배송업체를 입력해 주세요');
		frm.divname.focus();

	}
	frm.submit();

}
function chkAddFrm(frm){
	if (frm.divcd.value==''){

		alert('번호를 입력해 주세요');
		frm.divcd.focus();
		return false;

	}

	if (eval('document.editFrm_' + frm.divcd.value)!=null) {
		alert('중복된 번호는 사용할수 없습니다');
		return false;
	}


	if (frm.divname.value==''){
		alert('배송업체를 입력해 주세요');
		frm.divname.focus();
		return false;
	}

	frm.submit();
}
</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="90" align="center">번호</td>
		<td width="150" align="center">배송업체</td>
		<td width="150" align="center">대표전화번호</td>
		<td align="center">배송조회 URL</td>
		<td align="center">반품접수 URL</td>
		<td width="80" align="center">사용유무</td>
		<td width="80" align="center">10X10사용</td>
		<td width="50" align="center">수정</td>
	</tr>

<%
dim sql
sql = " SELECT divcd,divname,findurl, returnURL, isUsing, isTenUsing,tel " &_
			" FROM db_order.[dbo].tbl_songjang_div " &_
			" ORDER BY isTenUsing desc ,divcd "

rsget.open sql,dbget,1

if not (rsget.eof or rsget.bof) then
	do until rsget.eof

	dim defBgColor
	if rsget("isUsing") = "Y" then
		if rsget("isTenUsing") ="Y" then
		defBgColor 	=	"#FFFFFF"
		else
		defBgColor	=	"#FFFFFF"
		end if

	else
		defBgColor="#CCCCCC"
	end if
	%>
	<tr align="center" bgcolor="<%= defBgColor %>">
	<form name="editFrm_<%= rsget("divcd") %>" method="post" target="subFrame" action="delivery_service_process.asp">
	<input type="hidden" name="mode" value="edit" />
	<input type="hidden" name="divcd" value="<%= rsget("divcd") %>" />
		<td><%= rsget("divcd") %></td>
		<td><input type="text" class="text" name="divname" size="15" value="<%= db2html(rsget("divname")) %>"></td>
		<td><input type="text" class="text" name="tel" size="15" value="<%= db2html(rsget("tel")) %>">
		</td>
		<td align="left"><input type="text" class="text" name="findurl" size="70" value="<%= db2html(rsget("findurl")) %>"></td>
		<td align="left"><input type="text" class="text" name="returnURL" size="70" value="<%= db2html(rsget("returnURL")) %>"></td>
		<td>
			<select class="select" name="isusing">
				<option value="Y" <% if rsget("isUsing") = "Y" then response.write "selected" %>>사용함 </option>
				<option value="N" <% if rsget("isUsing") = "N" then response.write "selected" %>>사용안함</option>
			</select>
		</td>
		<td>
			<select class="select" name="isTenUsing">
				<option value="Y" <% if rsget("isTenUsing") = "Y" then response.write "selected" %>>사용함 </option>
				<option value="N" <% if rsget("isTenUsing") = "N" then response.write "selected" %>>사용안함</option>
			</select>
		</td>
		<td align="center">
			<input type="button" class="button" value="수정" onclick="chkEditFrm(this.form);">
			</td>
	</form>
	</tr>

<%
	rsget.movenext
	loop
end if

rsget.close
%>

</table>
<br />
<!-- 신규입력 테이블 -->
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>신규 입력</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60" align="center">번호</td>
		<td width="150" align="center">배송업체</td>
		<td width="150" align="center">대표전화번호</td>
		<td align="center">배송조회 URL</td>
		<td align="center">반품접수 URL</td>
		<td width="50" align="center"></td>
	</tr>
	<form name="addFrm" method="post" target="subFrame" action="delivery_service_process.asp">
	<input type="hidden" name="mode" value="add" />
	<tr bgcolor="#FFFFFF">
		<td align="center"><input type="text" name="divcd" value="" size="4" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="text" name="divname" size="15" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="text" name="tel" size="15" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="left"><input type="text" name="findurl" size="70" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="left"><input type="text" name="returnURL" size="70" value="" style="border:1px solid #CCCCCC;" /></td>
		<td align="center"><input type="button" class="button" value="입력" onclick="chkAddFrm(this.form);"></td>
	</tr>
	</form>
</table>
<iframe src="" name="subFrame" frameborder="0" width="0" height="0"></iframe>
<br /><br /><br /><br /><br /><br />

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
