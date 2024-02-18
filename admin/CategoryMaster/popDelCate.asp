<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : popDelCate.asp
' Discription : 카테고리 삭제처리 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'###############################################

dim cdl, cdm, cds, mode
dim sqlstr
Dim FTotalcnt, acode

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

mode = request("mode")

If mode="mdel" Then

	sqlstr = "select count(code_large) as cnt from [db_item].dbo.tbl_Cate_small"
	sqlstr = sqlstr + " where code_large='" + Cstr(cdl) + "'"
	sqlstr = sqlstr + " and code_mid='" + Cstr(cdm) + "'"
	rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof Then
			FTotalcnt = rsget("cnt")
		End If
	rsget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('소 카테고리를 삭제 하셔야만\n중 카테고리를 삭제 할 수 있습니다.');self.close();</script>"
		dbget.close()	:	response.End
	Else
		 '// 카테고리 속성 삭제
		 sqlstr = "declare @acode varchar(4000);" & vbCrLf &_
				" Select @acode = coalesce(@acode + ',', '') + convert(varchar, attrib_Code) from [db_item].dbo.tbl_Cate_Attrib_div" & vbCrLf &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "';" & vbCrLf &_
				" Delete from [db_item].dbo.tbl_Item_Attribute" &_
				" where attrib_Code in (@acode);" & vbCrLf &_
				" Delete from [db_item].dbo.tbl_Cate_Attrib_item" &_
				" where attrib_Code in (@acode);"
		 dbget.Execute(sqlStr)
		 	
		 '속성 옵션 및 상품지정옵션 삭제
		 sqlstr = "Delete from [db_item].dbo.tbl_Cate_Attrib_div" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"

		'카테고리 관련 링크정보 삭제
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_RelateLink" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"

		 dbget.Execute(sqlStr)

		 '// 중카테고리 삭제
		 sqlstr = "Delete from [db_item].dbo.tbl_Cate_mid" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'"
		 dbget.Execute(sqlStr)

		 response.write "<script>alert('중 카테고리를 삭제 하였습니다.');opener.document.location.reload();self.close();</script>"
	End If
End If

if mode="sdel" Then

	'상품테이블에서 확인(기본 카테고리)
	sqlstr = "select count(itemid) as cnt from [db_item].dbo.tbl_item" &_
			" where cate_large='" + Cstr(cdl) + "'" &_
			" and cate_mid='" + Cstr(cdm) + "'" &_
			" and cate_small='" + Cstr(cds) + "'"
	rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof Then
			FTotalcnt = rsget("cnt")
		End If
	rsget.close

	If FTotalcnt > 0 Then
		response.write "<script>alert('이동하지않은 상품이 있습니다.\n기본 카테고리 내에 상품이 없어야만 카테고리를 삭제 할 수 있습니다.');self.close();</script>"
		dbget.close()	:	response.End
	Else
		 '// 추가 카테고리에 등록된 상품 삭제
		 sqlstr = "Delete from db_item.dbo.tbl_Item_category " &_
		 		" where code_div='A' " &_
		 		" and code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'" & vbCrLf

		'// 카테고리 관련 링크정보 삭제
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_RelateLink" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"

		 '// 소카테고리 삭제
		 sqlstr = sqlstr & "Delete from [db_item].dbo.tbl_Cate_small" &_
		 		" where code_large='" + Cstr(cdl) + "'" &_
		 		" and code_mid='" + Cstr(cdm) + "'" &_
		 		" and code_small='" + Cstr(cds) + "'"
		 dbget.Execute(sqlStr)

		'카테고리 중분류 상품목록의 소분류 아이콘 업데이트(2009.07.06; 허진원)
		dbget.execute("exec db_const.dbo.sp_Ten_MakeCategorySmallIconList")

		 response.write "<script>alert('소 카테고리를 삭제 하였습니다.');opener.document.location.reload();self.close();</script>"
	End If
end if

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->