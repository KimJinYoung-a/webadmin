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

	Public Function GetTypeName() 
		'isSoldOut = (FSellYn="N")
		IF FDoc_Type="G010" Then
			GetTypeName = "핑거스공지"
		ElseIf FDoc_Type="G020" Then
			GetTypeName = "업무협조"
		ElseIf FDoc_Type="G030" Then
			GetTypeName = "문의사항"
		Else
			GetTypeName = "기타문의"
		End If
	End Function

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
			If FrectDoc_Type="G010" Then
				strSubSql = strSubSql & " and a.doc_type='" & FrectDoc_Type & "' "
			Else
				strSubSql = strSubSql & " and a.doc_type<>'G010'"
			End If
		End If

		''업체는 자기 글 or 공지글 만 보이게.
        If (FRectMakerID<>"") Then
            strSubSql = strSubSql & " and ((A.id='"& FRectMakerID &"') or (A.doc_type='G010'))"
        End If

		strSubSql = strSubSql & " and A.admin_usingyn='Y'"  '' 관리자가 N로 지정한것은 안보이게 ''2016/08/25
        
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
		sqlStr = sqlStr & " , A.doc_difficult, A.doc_subject, A.doc_content, A.doc_useyn, t.part_sn, A.admin_usingyn, A.doc_ans_ox, A.doc_idx" & vbCRLF
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
			FOneItem.fdoc_idx		= rsACADEMYget("doc_idx")
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
			FOneItem.fdoc_ans_ox = rsACADEMYget("doc_ans_ox")
        end if
        rsACADEMYget.Close
    end Sub

	public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_name " & VbCRLF
		strSql = strSql&"		FROM [db_academy].[dbo].tbl_lecturer_board_file " & VbCRLF
		strSql = strSql&"	WHERE doc_idx = '" & FrectDoc_Idx & "' " & VbCRLF
        strSql = strSql&"	ORDER BY file_idx ASC "
        
		rsACADEMYget.Open strSql,dbACADEMYget,1
		'response.write strSql
		IF not rsACADEMYget.EOF THEN
			fnGetFileList = rsACADEMYget.getRows() 
		END IF	
		rsACADEMYget.close
	End Function

	public sub fnGetolectList()
		dim sqlStr,i , strSubSql

		strSubSql = " AND A.doc_idx = '" & FrectDoc_Idx & "' "

		'총 갯수 구하기
		sqlStr = "select count(A.ans_idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lecturer_board_ans AS A" + vbcrlf
		sqlStr = sqlStr & " JOIN "&tmp_tbl&"[db_partner].[dbo].tbl_partner AS B " + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = B.id" + vbcrlf
		sqlStr = sqlStr & " where A.ans_useyn = 'Y'" & strSubSql
			
		'response.write sqlStr &"<br>"					
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " A.ans_idx, A.ans_type, A.ans_content, A.ans_regdate, B.company_name, A.id, A.doc_idx" + vbcrlf
		sqlStr = sqlStr & " , t.part_sn" + vbcrlf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lecturer_board_ans AS A" + vbcrlf
		sqlStr = sqlStr & " JOIN "&tmp_tbl&"[db_partner].[dbo].tbl_partner AS B " + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = B.id" + vbcrlf
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr & " where A.ans_useyn = 'Y'" & strSubSql
		sqlStr = sqlStr & " ORDER BY A.ans_idx asc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.EOF
				set FItemList(i) = new ClecturerItem
				FItemList(i).fpart_sn = rsACADEMYget("part_sn")
				FItemList(i).fdoc_idx = rsACADEMYget("doc_idx")
				FItemList(i).fans_idx = rsACADEMYget("ans_idx")
				FItemList(i).fans_type = rsACADEMYget("ans_type")
				FItemList(i).fans_content = db2html(rsACADEMYget("ans_content"))
				FItemList(i).fans_regdate = rsACADEMYget("ans_regdate")
				FItemList(i).fcompany_name = db2html(rsACADEMYget("company_name"))
				FItemList(i).fid = rsACADEMYget("id")
											
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

    public Sub fnGetolectView()
        dim sqlStr
		sqlStr = " SELECT top 1 ans_content" & _
				" FROM [db_academy].[dbo].tbl_lecturer_board_ans " & _
				" WHERE  ans_useyn = 'Y' AND ans_idx = '" & FrectAns_Idx & "' AND id = '" & FRectMakerID & "' "

        'response.write sqlStr&"<br>"
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new ClecturerItem
        
        if Not rsACADEMYget.Eof then
    			
			FOneItem.FAns_Content = rsACADEMYget("ans_content")
			           
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

Function getthefingers_staff(staff, part_sn, username)
	Dim tmpusername

	If part_sn=16 Then
		tmpusername="더핑거스"
	Else
		tmpusername=username
	End If

	getthefingers_staff = tmpusername
End Function

Function fingmaster(userid)
	If userid = "tozzinet" Or userid = "thefingers01" Or userid = "nownhere21" Or userid = "frudia79" Then
		fingmaster = True
	Else
		fingmaster = False
	End If
End Function

'####### 코드매니저 페이지 외 하나씩 불러쓰는 공통 코드 관리. write 용, view 용. #######
Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode)
	Dim strSql, sBody, i
	sBody = ""
	i = 0
	
	'### sUse = "w" write 용
	If sUse = "w" Then
		strSql = " SELECT commcd, commnm "
		strSql = strSql & " From db_academy.dbo.tbl_commCd "
		strSql = strSql & " WHERE groupcd ='"&sType&"' AND isusing = 'Y' ORDER BY commcd ASC"
		
		If sType = "doc_status" AND sCode = "" Then
			sBody = "<input type='hidden' name='doc_status' value='K001'>작성"
		Else
			
			'response.write strSql &"<br>"
			rsACADEMYget.Open strSql,dbACADEMYget,1
			Do Until rsACADEMYget.Eof				
				If sType = "K000" Then
					If i = 0 Then
						sBody = "<select name='"&sType&"' class='select'>"
						If GetFileName() = "lecturer"Then
							sBody = sBody & "<option value=''>-선택-</option> "
						End IF
					End IF
					sBody = sBody & "<option value='" & rsACADEMYget("commcd") & "' "
					If CStr(sCode) = CStr(rsACADEMYget("commcd")) Then
						sBody = sBody & "selected"
					End If
					sBody = sBody & ">" & rsACADEMYget("commnm") & "</option>"
					If i = rsACADEMYget.RecordCount-1 Then
						sBody = sBody & "</select>"
					End IF
					
				elseIf sType = "G000" Then
					If i = 0 Then
						sBody = "<select name='"&sType&"' class='select'>"
						sBody = sBody & "<option value='' selected>구분선택</option> "
					End IF
					'//강사일경우 핑거스 공지사항은 노출안함
					if fnlecturer(requestCheckVar(request.cookies("partner")("userid"),32)) then
						if rsACADEMYget("commcd") <> "G010" then
							sBody = sBody & "<option value='" & rsACADEMYget("commcd") & "' "
							sBody = sBody & ">" & rsACADEMYget("commnm") & "</option>"
						end if	

					'//핑거스아카데미 일경우 다 출력
					else
						sBody = sBody & "<option value='" & rsACADEMYget("commcd") & "' "
						sBody = sBody & ">" & rsACADEMYget("commnm") & "</option>"													
					end if
					If i = rsACADEMYget.RecordCount-1 Then
						sBody = sBody & "</select>"
					End IF							
				Else
					sBody = sBody & "<label id='" & sType & rsACADEMYget("commcd") & "'>" & _
									"<input type='radio' name='" & sType & "' id='" & sType & rsACADEMYget("commcd") & "' value='" & rsACADEMYget("commcd") & "' "
					If CStr(sCode) = CStr(rsACADEMYget("commcd")) Then
						sBody = sBody & "checked"
					End If
					sBody = sBody & ">" & rsACADEMYget("commnm") & "</label>&nbsp;&nbsp;"
				End If
			rsACADEMYget.MoveNext
			i = i + 1
			Loop
			rsACADEMYget.Close
		End If
	Else
	'### sUse = "v" view 용
		strSql = " SELECT commnm From db_academy.dbo.tbl_commCd"
		strSql = strSql & " WHERE groupcd ='"&sType&"' AND commcd = '" & sCode & "' AND isusing = 'Y'"
		rsACADEMYget.Open strSql,dbACADEMYget
		If Not rsACADEMYget.Eof Then
			sBody = rsACADEMYget(0)
		End If
		rsACADEMYget.Close
	End If
	CommonCode = sBody
End Function

' 현재 페이지 URL에서 파일명 뽑기
Function GetFileName()
	On Error Resume Next
	Dim vUrl			'/소스 경로저장 변수
	Dim FullFilename		'파일이름
	Dim strName			'확장자를 제외한 파일이름

	vUrl = Request.ServerVariables("SCRIPT_NAME")
	FullFilename = mid(vUrl,instrrev(vUrl,"/")+1)
	strName = Mid(FullFilename, 1, Instr(FullFilename, ".") - 1)

	GetFileName = strName
End Function

'//강사인지 체크한다
public Function fnlecturer(userid)
dim sql , tmp

sql = "select top 1 * "
sql = sql + " from [db_user].[dbo].tbl_user_c"
sql = sql + " where userid = '" + userid + "'" + vbCrlf

rsget.Open sql,dbget,1
if  not rsget.EOF  then
	tmp = rsget("userdiv")
end if
rsget.close

if tmp = "14" then
	fnlecturer = true
else
	fnlecturer = false
end if

End Function
%>