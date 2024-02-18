<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 통합 게시판
' History : 2010.06.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls.asp"-->

<%
Dim strSql, vWorkerTemp, vWorkerViewTemp, vFileTemp, i, page
Dim iDoc_Idx, iAns_Idx, sDoc_Id, sDoc_Content, sAns_Type, sDoc_RegistId
	iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
	iAns_Idx		= NullFillWith(requestCheckVar(Request("aidx"),10),"")
	page 	= NullFillWith(requestCheckVar(Request("page"),10),1)
	sDoc_Id			= session("ssBctId")
	sAns_Type		= "1"
	sDoc_Content	= html2db(Request("ans_content"))
	sDoc_RegistId	= requestCheckVar(NullFillWith(Request("registid"),""),32)

If iAns_Idx = "" Then

	'####### 답변 저장 #######
	strSql = " INSERT INTO db_shop.dbo.tbl_offshop_board_ans " & _
			 "		(doc_idx, id, ans_type, ans_content) " & _
			 "	VALUES " & _
			 "		('" & iDoc_Idx & "', '" & sDoc_Id & "', '" & sAns_Type & "', '" & html2db(sDoc_Content) & "') " & _
			 " UPDATE db_shop.dbo.tbl_offshop_board_document SET doc_ans_ox = 'o' WHERE doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	dbget.execute strSql

Else

	If requestCheckVar(Request("del"),1) = "o" Then
		
		'####### 답변 삭제 #######
		strSql = " UPDATE db_shop.dbo.tbl_offshop_board_ans SET " & _
				 "		ans_useyn = 'N' " & _
				 "	WHERE " & _
				 "		ans_idx = '" & iAns_Idx & "' "
		
		'response.write strSql &"<br>"
		dbget.execute strSql

		strSql = " if exists(select ans_idx from db_shop.dbo.tbl_offshop_board_ans where ans_useyn='Y' and ans_idx = '" & iAns_Idx & "' and doc_idx = '" & iDoc_Idx & "')" & _
				 "		begin " & _
				 "		update db_shop.dbo.tbl_offshop_board_document set doc_ans_ox = 'o' where doc_idx = '" & iDoc_Idx & "'" & _
				 "		end " & _
				 "	else" & _
				 "		begin " & _
				 "		update db_shop.dbo.tbl_offshop_board_document set doc_ans_ox = 'x' where doc_idx = '" & iDoc_Idx & "'" & _
				 "		end "

		'response.write strSql &"<br>"
		dbget.execute strSql

	Else
	
		'####### 답변 저장 #######
		strSql = " UPDATE db_shop.dbo.tbl_offshop_board_ans SET " & _
				 "		ans_content = '" & html2db(sDoc_Content) & "' " & _
				 "	WHERE " & _
				 "		ans_idx = '" & iAns_Idx & "' "

		'response.write strSql &"<br>"
		dbget.execute strSql

		strSql = " if exists(select ans_idx from db_shop.dbo.tbl_offshop_board_ans where ans_useyn='Y' and ans_idx = '" & iAns_Idx & "' and doc_idx = '" & iDoc_Idx & "')" & _
				 "		begin " & _
				 "		update db_shop.dbo.tbl_offshop_board_document set doc_ans_ox = 'o' where doc_idx = '" & iDoc_Idx & "'" & _
				 "		end " & _
				 "	else" & _
				 "		begin " & _
				 "		update db_shop.dbo.tbl_offshop_board_document set doc_ans_ox = 'x' where doc_idx = '" & iDoc_Idx & "'" & _
				 "		end "

		'response.write strSql &"<br>"
		dbget.execute strSql

	End IF

End If


Response.Write "<script>"
response.write "	location.href='iframe_board_ans.asp?didx="&iDoc_Idx&"&page="&page&"&registid="&sDoc_RegistId&"';"
response.write "</script>"
response.end
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->