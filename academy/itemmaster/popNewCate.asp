<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYCategoryCls.asp"-->
<%
'###############################################
' PageName : popnewcate.asp
' Discription : 신규 카테고리 추가 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'###############################################

dim cdl, cdm, mode, cd, nm
dim sqlstr

cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cd  = RequestCheckvar(request("cd"),10)
nm  = trim(html2db(RequestCheckvar(request("nm"),64)))

mode = RequestCheckvar(request("mode"),16)

if mode="addlarge" then
	'중복여부 검사
	sqlstr = "select count(*) "
	sqlstr = sqlstr + " From [db_academy].dbo.tbl_diy_item_cate_large"
	sqlstr = sqlstr + " where code_large='" + cd + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	if rsACADEMYget(0)>0 then
		response.write "<script>alert('이미 존재하는 코드입니다.\n확인하고 다시 시도해주세요.');history.back();</script>"
		rsACADEMYget.close: dbACADEMYget.close
		response.end
	end if
	rsACADEMYget.Close

	'저장
	sqlstr = "insert into [db_academy].dbo.tbl_diy_item_cate_large"
	sqlstr = sqlstr + " (code_large, code_nm)"
	sqlstr = sqlstr + " values('" + cd + "'"
	sqlstr = sqlstr + " ,'" + nm + "')"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	response.write "<script>opener.document.location.reload();</script>"
end if


if mode="addmid" then
	'중복여부 검사
	sqlstr = "select count(*) "
	sqlstr = sqlstr + " From [db_academy].dbo.tbl_diy_item_cate_mid"
	sqlstr = sqlstr + " where code_large='" + cdl + "'"
	sqlstr = sqlstr + " and code_mid='" + cd + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	if rsACADEMYget(0)>0 then
		response.write "<script>alert('이미 존재하는 코드입니다.\n확인하고 다시 시도해주세요.');history.back();</script>"
		rsACADEMYget.close: dbACADEMYget.close
		response.end
	end if
	rsACADEMYget.Close

	'저장
	sqlstr = "insert into [db_academy].dbo.tbl_diy_item_cate_mid"
	sqlstr = sqlstr + " (code_large, code_mid, code_nm)"
	sqlstr = sqlstr + " values("
	sqlstr = sqlstr + " '" + cdl + "'"
	sqlstr = sqlstr + " ,'" + cd + "'"
	sqlstr = sqlstr + " ,'" + nm + "')"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	response.write "<script>opener.document.location.reload();</script>"
end if


if mode="addsmall" then
	'중복여부 검사
	sqlstr = "select count(*) "
	sqlstr = sqlstr + " From [db_academy].dbo.tbl_diy_item_cate_small"
	sqlstr = sqlstr + " where code_large='" + cdl + "'"
	sqlstr = sqlstr + " and code_mid='" + cdm + "'"
	sqlstr = sqlstr + " and code_small='" + cd + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
	if rsACADEMYget(0)>0 then
		response.write "<script>alert('이미 존재하는 코드입니다.\n확인하고 다시 시도해주세요.');history.back();</script>"
		rsACADEMYget.close: dbACADEMYget.close
		response.end
	end if
	rsACADEMYget.Close

	'저장
	sqlstr = "insert into [db_academy].dbo.tbl_diy_item_cate_small"
	sqlstr = sqlstr + " (code_large, code_mid, code_small, code_nm)"
	sqlstr = sqlstr + " values("
	sqlstr = sqlstr + " '" + cdl + "'"
	sqlstr = sqlstr + " ,'" + cdm + "'"
	sqlstr = sqlstr + " ,'" + cd + "'"
	sqlstr = sqlstr + " ,'" + nm + "')"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1

	response.write "<script>opener.document.location.reload();</script>"
end if


dim oLcate, currposStr
set oLcate = new CCatemanager

if cdl<>"" then
	currposStr = oLcate.GetNewCateCurrentPos(cdl,cdm,"")
end if

%>
<script language='javascript'>
function AddCate(frm){
	if (frm.cd.value.length!=3){
		alert('분류코드는 숫자 세자리입니다.');
		frm.cd.focus;
		return;
	}

	if (frm.nm.value.length<1){
		alert('카테고리명을 입력하세요.');
		frm.nm.focus;
		return;
	}

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table border=1 cellspacing=0 cellpadding=0 width=280 class=a>
<form name=frmadd method=post action="popNewCate.asp">
<input type=hidden name=cdl value="<%= cdl %>">
<input type=hidden name=cdm value="<%= cdm %>">
<tr>
	<td colspan=2>현재위치: <%= currposStr %></td>
</tr>
<tr>
	<td colspan="2">
	<% if cdl="" then %>
	대분류추가
	<input type=hidden name=mode value="addlarge">
	<% elseif cdm="" then %>
	중분류추가
	<input type=hidden name=mode value="addmid">
	<% else %>
	소분류추가
	<input type=hidden name=mode value="addsmall">
	<% end if %>
	</td>
</tr>
<tr align=center>
	<td width=100>분류코드</td>
	<td>카테고리명</td>
</tr>
<tr align=center>
	<td width=100><input type="text" name="cd" value="" size="3" maxlength="3"></td>
	<td><input type="text" name="nm" value="" size="16" maxlength="30"></td>
</tr>
<tr>
	<td colspan=2 align=center><input type=button value="저장" onclick="AddCate(frmadd);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->