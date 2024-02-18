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
' PageName : popDelCate.asp
' Discription : 카테고리 삭제처리 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
' History : 2012.08.17 이종화 : 이전 Admin에서 이전/수정
'###############################################

dim cdl, cdm, cds, mode
dim sqlstr
Dim FTotalcnt, acode

cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)

mode = RequestCheckvar(request("mode"),16)

If mode="mdel" Then

	sqlstr = "select count(code_large) as cnt from [db_academy].dbo.tbl_diy_item_Cate_small"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If Not rsACADEMYget.Eof Then
			FTotalcnt = rsACADEMYget("cnt")
		End If
	rsACADEMYget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('소 카테고리를 삭제 하셔야만\n중 카테고리를 삭제 할 수 있습니다.');self.close();</script>"
		dbACADEMYget.close()	:	response.End
	Else
		 '// 중카테고리 삭제
		 sqlstr = "Delete from [db_academy].dbo.tbl_diy_item_Cate_mid" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"
		 dbACADEMYget.Execute(sqlStr)

		 response.write "<script>alert('중 카테고리를 삭제 하였습니다.');opener.document.location.reload();self.close();</script>"
	End If
End If

if mode="sdel" Then

	'상품테이블에서 확인(기본 카테고리)
	sqlstr = "select count(itemid) as cnt from [db_academy].dbo.tbl_diy_item" &_
			" where cate_large='" + Cstr(cdl) + "'" &_
			" and cate_mid='" + Cstr(cdm) + "'" &_
			" and cate_small='" + Cstr(cds) + "'"
	rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If Not rsACADEMYget.Eof Then
			FTotalcnt = rsACADEMYget("cnt")
		End If
	rsACADEMYget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('이동하지않은 상품이 있습니다.\n기본 카테고리 내에 상품이 없어야만 카테고리를 삭제 할 수 있습니다.');self.close();</script>"
		dbACADEMYget.close()	:	response.End
	Else
		 '// 추가 카테고리에 등록된 상품 삭제
		 sqlstr = "Delete from [db_academy].dbo.tbl_diy_item_category " &_
		 		" where code_div='A' " &_
		 		" and code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'" & vbCrLf

		 '// 소카테고리 삭제
		 sqlstr = sqlstr & "Delete from [db_academy].dbo.tbl_diy_item_Cate_small" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"
		 dbACADEMYget.Execute(sqlStr)

		'카테고리 중분류 상품목록의 소분류 아이콘 업데이트(2009.07.06; 허진원)
		'dbACADEMYget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")

		 response.write "<script>alert('소 카테고리를 삭제 하였습니다.');opener.document.location.reload();self.close();</script>"
	End If
end if

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->