<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/CategoryCls.asp"-->
<%
'###############################################
' PageName : popEditCate.asp
' Discription : 카테고리 수정 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
' History : 2012.08.16 이종화 : 이전 Admin에서 이전/수정
'###############################################

dim cdl, cdm, cds, mode, name, name_eng, orderNo, copy_kor, copy_eng
dim sqlstr
dim display_yn

cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)
display_yn = RequestCheckvar(request("display_yn"),2)

mode = RequestCheckvar(request("mode"),16)

name = trim(html2db(RequestCheckvar(request("name"),64)))
name_eng = trim(html2db(RequestCheckvar(request("name_eng"),64)))
copy_kor = trim(html2db(RequestCheckvar(request("copy_kor"),64)))
copy_eng = trim(html2db(RequestCheckvar(request("copy_eng"),64)))
orderno=RequestCheckvar(request("orderno"),10)


if mode="editmid" then
	sqlstr = "update [db_academy].dbo.tbl_diy_item_Cate_mid"
	sqlstr = sqlstr + " set code_nm='" + name + "'"
	sqlstr = sqlstr + " ,orderNo='" + orderno + "'"
	sqlstr = sqlstr + " ,display_yn='" + display_yn + "'"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
end if


if mode="editsmall" Then
	sqlstr = "update [db_academy].dbo.tbl_diy_item_Cate_small"
	sqlstr = sqlstr + " set code_nm='" + name + "'"
	sqlstr = sqlstr + " ,orderNo='" + orderno + "'"
	sqlstr = sqlstr + " ,display_yn='" + display_yn + "'"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	sqlstr = sqlstr + " and code_small='" + Cstr(cds) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	'카테고리 중분류 상품목록의 소분류 아이콘 업데이트(2009.07.06; 허진원)
	'dbACADEMYget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")
end if

if mode<>"" then
	'부모페이지 새로고침
	response.write "<script>opener.document.location.reload();</script>"
end if


dim oLcate, currposStr
set oLcate = new CCate

if cdl<>"" then
	currposStr = oLcate.GetNewCateCurrentPos(cdl,cdm,cds)

	'// 내용 접수
	if cds<>"" then
		sqlstr = "select top 1 code_nm, orderNo, display_yn From [db_academy].dbo.tbl_diy_item_Cate_small where code_large='" & cdl & "' and code_mid='" & cdm & "' and code_small='" & cds & "'"

		if sqlstr<>"" then
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				name = rsACADEMYget("code_nm")
				orderNo = rsACADEMYget("orderNo")
				display_yn = rsACADEMYget("display_yn")
			end if
			rsACADEMYget.Close
		end if

	elseif cdm<>"" then
		sqlstr = "select top 1 code_nm, orderNo, display_yn From [db_academy].dbo.tbl_diy_item_Cate_mid where code_large='" & cdl & "' and code_mid='" & cdm & "'"

		if sqlstr<>"" then
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				name = rsACADEMYget("code_nm")
				orderNo = rsACADEMYget("orderNo")
				display_yn = rsACADEMYget("display_yn")
			end if
			rsACADEMYget.Close
		end if

	end if
end if
%>
<script language='javascript'>
function EditCate(frm){

	if (frm.name.value.length<1){
		alert('카테고리명을 입력하세요.');
		frm.name.focus;
		return;
	}

	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table border=0 cellspacing=1 cellpadding=3 width=280 class=a bgcolor="#808080">
<form name=frmadd method=post >
<input type=hidden name=cdl value="<%= cdl %>">
<input type=hidden name=cdm value="<%= cdm %>">
<input type=hidden name=cds value="<%= cds %>">
<tr bgcolor="#FFFFFF">
	<td colspan="2">현재위치: <%= currposStr %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	<% if cdl<>"" And cdm<>"" And cds="" then %>
	<b>중분류수정</b>
	<input type=hidden name=mode value="editmid">
	<% elseif cdl<>"" And cdm<>"" And cds<>"" then %>
	<b>소분류수정</b>
	<input type=hidden name=mode value="editsmall">
	<% end if %>
	</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td width=100>분류코드</td>
	<td align="left"><% If cdl <> "" Then %>대[<%= cdl %>]<% End If %>&nbsp;<% If cdm <> "" Then %>중[<%= cdm %>]<% End If %>&nbsp;<% If cds <> "" Then %>소[<%= cds %>]<% End If %></td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>카테고리명</td>
	<td align="left"><input type="text" name="name" value="<%=name%>" size="20" maxlength="32"></td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>정렬 순서</td>
	<td align="left"><input type="text" name="orderNo" value="<%=orderNo%>" size="5" maxlength="4"></td>
</tr>
<% if (cdm<>"") then %>
<tr align=center bgcolor="#FFFFFF">
	<td>전시 YN</td>
	<td align="left">
	<input type="radio" name="display_yn" value="Y" <%= CHKIIF(display_yn="Y","checked","") %> >Y
	<input type="radio" name="display_yn" value="N" <%= CHKIIF(display_yn="N","checked","") %> >N
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align=center><input type=button value="저장" onclick="EditCate(frmadd);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->