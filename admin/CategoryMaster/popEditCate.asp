<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderutf8.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : popEditCate.asp
' Discription : 카테고리 수정 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'			2017.07.31 한용민 수정(utf8로 변경)
'###############################################

dim cdl, cdm, cds, mode, name, name_eng, orderNo, copy_kor, copy_eng, name_cn_gan, name_cn_bun
dim sqlstr
dim display_yn

cdl = requestCheckvar(request("cdl"),3)
cdm = requestCheckvar(request("cdm"),3)
cds = requestCheckvar(request("cds"),3)
display_yn = requestCheckvar(request("display_yn"),3)

mode = requestCheckvar(request("mode"),32)

name = requestCheckvar(trim(html2db(request("name"))),64)
name_eng = requestCheckvar(trim(html2db(request("name_eng"))),64)
copy_kor = requestCheckvar(trim(html2db(request("copy_kor"))),64)
copy_eng = requestCheckvar(trim(html2db(request("copy_eng"))),64)
orderno=requestCheckvar(request("orderno"),10)

name_cn_gan = requestCheckvar(trim(html2db(request("name_cn_gan"))),64)
name_cn_bun = requestCheckvar(trim(html2db(request("name_cn_bun"))),64)


if mode="editmid" then
	sqlstr = "update [db_item].dbo.tbl_Cate_mid" &VBCRLF
	sqlstr = sqlstr + " set code_nm='" + name + "'" &VBCRLF
	sqlstr = sqlstr + " ,code_nm_eng='" + name_eng + "'" &VBCRLF
	sqlstr = sqlstr + " ,copy_nm='" + copy_kor + "'" &VBCRLF
	sqlstr = sqlstr + " ,copy_nm_eng='" + copy_eng + "'" &VBCRLF
	sqlstr = sqlstr + " ,orderNo='" + orderno + "'" &VBCRLF
	sqlstr = sqlstr + " ,display_yn='" + display_yn + "'" &VBCRLF
	sqlstr = sqlstr + " ,code_nm_cn_gan=N'" + name_cn_gan + "'" &VBCRLF  ''2017/06/22 추가
	sqlstr = sqlstr + " ,code_nm_cn_bun=N'" + name_cn_bun + "'" &VBCRLF  ''2017/06/22 추가
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'" &VBCRLF
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'" &VBCRLF
	dbget.Execute sqlstr
end if


if mode="editsmall" Then
	sqlstr = "update [db_item].dbo.tbl_Cate_small" &VBCRLF
	sqlstr = sqlstr + " set code_nm='" + name + "'" &VBCRLF
	sqlstr = sqlstr + " ,code_nm_eng='" + name_eng + "'" &VBCRLF
	sqlstr = sqlstr + " ,copy_nm='" + copy_kor + "'" &VBCRLF
	sqlstr = sqlstr + " ,copy_nm_eng='" + copy_eng + "'" &VBCRLF
	sqlstr = sqlstr + " ,orderNo='" + orderno + "'" &VBCRLF
	sqlstr = sqlstr + " ,display_yn='" + display_yn + "'" &VBCRLF
	sqlstr = sqlstr + " ,code_nm_cn_gan=N'" + name_cn_gan + "'" &VBCRLF  ''2017/06/22 추가
	sqlstr = sqlstr + " ,code_nm_cn_bun=N'" + name_cn_bun + "'" &VBCRLF  ''2017/06/22 추가
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'" &VBCRLF
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'" &VBCRLF
	sqlstr = sqlstr + " and code_small='" + Cstr(cds) + "'" &VBCRLF
	dbget.Execute sqlstr
	'카테고리 중분류 상품목록의 소분류 아이콘 업데이트(2009.07.06; 허진원) =>주석처리 2017/06/22 않쓰이는듯? 필요시 야간 배치
	''dbget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")
end if

if mode<>"" then
	'부모페이지 새로고침
	response.write "<script>opener.document.location.reload();</script>"
end if


dim oLcate, currposStr
set oLcate = new CCatemanager

if cdl<>"" then
	currposStr = oLcate.GetNewCateCurrentPos(cdl,cdm,cds)

	'// 내용 접수
	if cds<>"" then
		sqlstr = "select top 1 code_nm, code_nm_eng, copy_nm, copy_nm_eng, orderNo, display_yn, code_nm_cn_gan, code_nm_cn_bun From [db_item].dbo.tbl_Cate_small where code_large='" & cdl & "' and code_mid='" & cdm & "' and code_small='" & cds & "'"

		if sqlstr<>"" then
			rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			if Not(rsget.EOF or rsget.BOF) then
				name = rsget("code_nm")
				name_eng = rsget("code_nm_eng")
				copy_kor = rsget("copy_nm")
				copy_eng = rsget("copy_nm_eng")
				orderNo = rsget("orderNo")
				display_yn = rsget("display_yn")
				name_cn_gan = rsget("code_nm_cn_gan")
				name_cn_bun = rsget("code_nm_cn_bun")
			end if
			rsget.Close
		end if

	elseif cdm<>"" then
		sqlstr = "select top 1 code_nm, code_nm_eng, copy_nm, copy_nm_eng, orderNo, display_yn, code_nm_cn_gan, code_nm_cn_bun From [db_item].dbo.tbl_Cate_mid where code_large='" & cdl & "' and code_mid='" & cdm & "'"

		if sqlstr<>"" then
			rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			if Not(rsget.EOF or rsget.BOF) then
				name = rsget("code_nm")
				name_eng = rsget("code_nm_eng")
				copy_kor = rsget("copy_nm")
				copy_eng = rsget("copy_nm_eng")
				orderNo = rsget("orderNo")
				display_yn = rsget("display_yn")
				name_cn_gan = rsget("code_nm_cn_gan")
				name_cn_bun = rsget("code_nm_cn_bun")
			end if
			rsget.Close
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
	<td>영문명</td>
	<td align="left"><input type="text" name="name_eng" value="<%=name_eng%>" size="22" maxlength="32">
		※ 해외배송시 상품종류를 나타내는 이름입니다. 브랜드명을 입력하지 말아주세요.
	</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>중문(간자체)</td>
	<td align="left"><input type="text" name="name_cn_gan" value="<%=name_cn_gan%>" size="22" maxlength="32">
	</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>중문(번자체)</td>
	<td align="left"><input type="text" name="name_cn_bun" value="<%=name_cn_bun%>" size="22" maxlength="32">
	</td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>카피(한글)</td>
	<td align="left"><input type="text" name="copy_kor" value="<%=copy_kor%>" size="22" maxlength="32"></td>
</tr>
<tr align=center bgcolor="#FFFFFF">
	<td>카피(영문)</td>
	<td align="left"><input type="text" name="copy_eng" value="<%=copy_eng%>" size="22" maxlength="32"></td>
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
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->