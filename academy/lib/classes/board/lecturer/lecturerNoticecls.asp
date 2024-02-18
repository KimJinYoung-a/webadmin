<%
Class ClecturerItem
	Public fpart_sn
	Public fdoc_idx
	Public fdoc_subject
	Public fdoc_type
	Public fdoc_important
	Public fdoc_difficult
	Public fdoc_status
	Public fdoc_regdate
	Public fcompany_name
	Public fdoc_ans_ox
	Public fdoc_type_nm
	Public fdoc_important_nm
	Public fdoc_status_nm
	Public FDoc_Id
	Public FDoc_Name
	Public FDoc_Import
	Public FDoc_Diffi
	Public FDoc_Subj
	Public FDoc_Content
	Public FDoc_UseYN
	Public fans_idx
	Public fans_type
	Public fans_content
	Public fans_regdate
	Public fid
	Public fans_count
	Public fadmin_usingyn

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class ClecturerList
	Public FItemList()
	Public FTotalCount
	Public FOneItem
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount
	Public FrectDoc_Status
	Public FrectDoc_Type
	Public FrectDoc_AnsOX
	Public FrectDoc_Idx
	Public tmp_tbl
	Public FrectAns_Idx
	Public frectsearchKey
	Public frectsearchString
	Public FRECTAdmin_UsingNInclude
	Public FRectMakerID

	Private Sub Class_Initialize()
		ReDim FItemList(0)
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
 		If application("Svr_Info")="Dev" Then 
 			tmp_tbl = "[TEndB]."
 		Else
 			tmp_tbl = "[TEndB]."
 		End If
	End Sub

	Private Sub Class_Terminate()

	End Sub


	Public Sub fnGetlecturerList()
		Dim sqlStr,i , strSubSql
		
		'//요청구분
		If FrectDoc_Type<>"" Then
			strSubSql = strSubSql & " and a.doc_type='" & FrectDoc_Type & "' "
		End If
		'//처리구분
		If FrectDoc_Status <> "" Then
			strSubSql = strSubSql & " and a.doc_status='" & FrectDoc_Status & "' "
		End If
		'//답변여부
		If FrectDoc_AnsOX <> "" Then
			strSubSql = strSubSql & " and a.doc_ans_ox='" & FrectDoc_AnsOX & "' "
		End If
		'//상세검색
		If FRectsearchString <> "" And FRectsearchKey <> "" Then
			strSubSql = strSubSql & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		End If		

        ''업체는 자기 글 or 공지글 만 보이게.
        If (FRectMakerID<>"") Then
            strSubSql = strSubSql & " and ((A.id='"& FRectMakerID &"') or (A.doc_type='G010'))"
        End If
         
        If (FRECTAdmin_UsingNInclude="on") Then
            
        Else
            strSubSql = strSubSql & " and A.admin_usingyn='Y'"  '' 관리자가 N로 지정한것은 안보이게 ''2016/08/25
        End If
        
		'총 갯수 구하기
		sqlStr = "select count(a.doc_idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM db_academy.dbo.tbl_lecturer_board_document AS A" + vbcrlf
		sqlStr = sqlStr & " join "&tmp_tbl&"db_partner.dbo.tbl_partner c" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = C.id" + vbcrlf
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr & " 	and t.part_sn = 16" + vbcrlf
		sqlStr = sqlStr & " where A.doc_useyn = 'Y'" & strSubSql

		'response.write sqlStr &"<br>"		
		rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " A.doc_idx, A.doc_subject, A.doc_type, A.doc_important, A.doc_difficult, A.doc_status, A.doc_regdate" + vbcrlf
		sqlStr = sqlStr & " , C.company_name, A.doc_ans_ox, A.id, t.part_sn, A.admin_usingyn" + vbcrlf
		sqlStr = sqlStr & " ,(select commNm" + vbcrlf
		sqlStr = sqlStr & " 	from db_academy.dbo.tbl_commCd" + vbcrlf
		sqlStr = sqlStr & " 	where commCd = a.doc_type) as doc_type_nm" + vbcrlf
		sqlStr = sqlStr & " ,(select commNm" + vbcrlf
		sqlStr = sqlStr & " 	from db_academy.dbo.tbl_commCd" + vbcrlf
		sqlStr = sqlStr & " 	where commCd = a.doc_important) as doc_important_nm" + vbcrlf
		sqlStr = sqlStr & " ,(select commNm" + vbcrlf
		sqlStr = sqlStr & " 	from db_academy.dbo.tbl_commCd" + vbcrlf
		sqlStr = sqlStr & " 	where commCd = a.doc_status) as doc_status_nm" + vbcrlf
		sqlStr = sqlStr & " ,(select count(ans_idx)" + vbcrlf
		sqlStr = sqlStr & " 	from db_academy.dbo.tbl_lecturer_board_ans" + vbcrlf
		sqlStr = sqlStr & " 	where doc_idx = a.doc_idx and ans_useyn='Y') as ans_count" + vbcrlf							
		sqlStr = sqlStr & " FROM db_academy.dbo.tbl_lecturer_board_document AS A" + vbcrlf
		sqlStr = sqlStr & " join "&tmp_tbl&"db_partner.dbo.tbl_partner c" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = C.id" + vbcrlf
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr & " where A.doc_useyn = 'Y'" & strSubSql
		sqlStr = sqlStr & " ORDER BY A.doc_idx DESC"
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
        rsACADEMYget.CursorLocation = adUseClient
        rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
        
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1

		ReDim Preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		If Not rsACADEMYget.EOF  Then
			rsACADEMYget.absolutepage = FCurrPage
			Do Until rsACADEMYget.EOF
				Set FItemList(i) = New ClecturerItem
				FItemList(i).fpart_sn = rsACADEMYget("part_sn")
				FItemList(i).fans_count = rsACADEMYget("ans_count")
				FItemList(i).fdoc_idx = rsACADEMYget("doc_idx")
				FItemList(i).fdoc_id = rsACADEMYget("id")
				FItemList(i).fdoc_subject = db2html(rsACADEMYget("doc_subject"))
				FItemList(i).fdoc_type = rsACADEMYget("doc_type")
				FItemList(i).fdoc_important = rsACADEMYget("doc_important")
				FItemList(i).fdoc_difficult = rsACADEMYget("doc_difficult")
				FItemList(i).fdoc_status = rsACADEMYget("doc_status")
				FItemList(i).fdoc_regdate = rsACADEMYget("doc_regdate")
				FItemList(i).fcompany_name = db2html(rsACADEMYget("company_name"))
				FItemList(i).fdoc_ans_ox = rsACADEMYget("doc_ans_ox")
				FItemList(i).fdoc_important_nm = rsACADEMYget("doc_important_nm")
				FItemList(i).fdoc_type_nm = rsACADEMYget("doc_type_nm")
				FItemList(i).fdoc_status_nm = rsACADEMYget("doc_status_nm")
				FItemList(i).fadmin_usingyn = rsACADEMYget("admin_usingyn")
				rsACADEMYget.movenext
				i=i+1
			Loop
		End If
		rsACADEMYget.Close
	End Sub

    public Sub fnGetlecturerView()
        dim sqlStr
        
		sqlStr = " SELECT A.id, B.company_name, A.doc_regdate, A.doc_status, A.doc_type, A.doc_important" & vbCRLF
		sqlStr = sqlStr & " , A.doc_difficult, A.doc_subject, A.doc_content, A.doc_useyn, t.part_sn, A.admin_usingyn" & vbCRLF
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lecturer_board_document AS A " & vbCRLF
		sqlStr = sqlStr & " INNER JOIN "&tmp_tbl&"[db_partner].[dbo].tbl_partner AS B " & vbCRLF
		sqlStr = sqlStr & " 	ON A.id = B.id " & vbCRLF
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr &" WHERE A.doc_idx = " & FrectDoc_Idx & vbCRLF
        
        ''업체는 자기 글 or 공지글 만 보이게.
        if (FRectMakerID<>"") then
            sqlStr = sqlStr & " and ((A.id='" & FRectMakerID & "') or (A.doc_type='G010'))"
        end if
        
        if (FRECTAdmin_UsingNInclude="on") then
            
        else
            sqlStr = sqlStr & " and A.admin_usingyn='Y'"  '' 관리자가 N로 지정한것은 안보이게 ''2016/08/25
        end if
        
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        Set FOneItem = New ClecturerItem
        
        if Not rsACADEMYget.Eof then
			FOneItem.fpart_sn 		= rsACADEMYget("part_sn")
			FOneItem.FDoc_Id 		= rsACADEMYget("id")
			FOneItem.FDoc_Name		= db2html(rsACADEMYget("company_name"))
			FOneItem.FDoc_Status	= rsACADEMYget("doc_status")
			FOneItem.FDoc_Type		= rsACADEMYget("doc_type")
			FOneItem.FDoc_Import	= rsACADEMYget("doc_important")
			FOneItem.FDoc_Diffi		= rsACADEMYget("doc_difficult")
			FOneItem.FDoc_Subj		= db2html(rsACADEMYget("doc_subject"))
			FOneItem.FDoc_Content	= db2html(rsACADEMYget("doc_content"))
			FOneItem.FDoc_UseYN		= rsACADEMYget("doc_useyn")
			FOneItem.FDoc_Regdate	= rsACADEMYget("doc_regdate")
			FOneItem.fadmin_usingyn = rsACADEMYget("admin_usingyn")
        end if
        rsACADEMYget.Close
    end Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

End Class
%>