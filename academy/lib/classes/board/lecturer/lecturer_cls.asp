<%
'###########################################################
' Description :  핑거스 강사 게시판
' History : 2010.03.29 한용민 생성
'###########################################################
%>
<%
Class clecturer_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fpart_sn
	public fdoc_idx
	public fdoc_subject
	public fdoc_type
	public fdoc_important
	public fdoc_difficult
	public fdoc_status
	public fdoc_regdate
	public fcompany_name
	public fdoc_ans_ox
	public fdoc_type_nm
	public fdoc_important_nm
	public fdoc_status_nm
	public FDoc_Id
	public FDoc_Name
	public FDoc_Import
	public FDoc_Diffi
	public FDoc_Subj
	public FDoc_Content
	public FDoc_UseYN
	public fans_idx
	public fans_type
	public fans_content
	public fans_regdate
	public fid
	public fans_count
	public fadmin_usingyn
	public fpushsenddate
	
	''푸시 발송 게시글인지 여부.
	public function IsPushSendReqNotice()
	    Dim ret 
	    ret = (fdoc_type="G010")
	    ret = ret AND (FDoc_UseYN<>"N")
	    
	    if (application("Svr_Info")="Dev") then
	        ret = ret AND (fdoc_regdate>=dateAdd("d",-120,now()))
	    else
	        ret = ret AND (fdoc_regdate>=dateAdd("d",-10,now()))
	    end if
	    IsPushSendReqNotice = ret
    end function

	public function IsPushSended()
	    IsPushSended = NOT (isNULL(fpushsenddate) or (fpushsenddate=""))
    end function
end class

class clecturer_list
	public FItemList()
	public FTotalCount
	public FOneItem
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FrectDoc_Status
	public FrectDoc_Type
	public FrectDoc_AnsOX
	public FrectDoc_Idx
	public tmp_tbl
	public FrectAns_Idx
	public frectsearchKey
	public frectsearchString
	
	public FRECTAdmin_UsingNInclude
	
	''/academy/board/lecturer/lecturer.asp
	public sub fnGetlecturerList()
		dim sqlStr,i , strSubSql
		dim isUpcheSsn : isUpcheSsn =  (CLNG(session("ssBctDiv"))>10) ''(session("ssBctDiv")>"10")
		
		'//요청구분
		if FrectDoc_Type<>"" Then
			If FrectDoc_Type="G010" Then
				strSubSql = strSubSql & " and a.doc_type='" & FrectDoc_Type & "' "
			Else
				strSubSql = strSubSql & " and a.doc_type<>'G010'"
			End If
		end if
		'//처리구분
		if FrectDoc_Status <> "" then
			strSubSql = strSubSql & " and a.doc_status='" & FrectDoc_Status & "' "
		end if
		'//답변여부
		if FrectDoc_AnsOX <> "" then
			strSubSql = strSubSql & " and a.doc_ans_ox='" & FrectDoc_AnsOX & "' "
		end if
		'//상세검색
		if FRectsearchString <> "" and FRectsearchKey <> "" then
			strSubSql = strSubSql & " and " & FRectsearchKey & " like '%" & FRectsearchString & "%' "
		end if		

        ''업체는 자기 글 or 공지글 만 보이게.
        if (isUpcheSsn) then
            strSubSql = strSubSql & " and ((A.id='"&session("ssBctID")&"') or (A.doc_type='G010'))"
        end if
         
        if (FRECTAdmin_UsingNInclude="on") then
            
        else
            strSubSql = strSubSql & " and A.admin_usingyn='Y'"  '' 관리자가 N로 지정한것은 안보이게 ''2016/08/25
        end if
        
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
		sqlStr = sqlStr & " ,A.pushsenddate" + vbcrlf		''2016/11/30 추가
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
				set FItemList(i) = new clecturer_item
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
				FItemList(i).fpushsenddate = rsACADEMYget("pushsenddate")
				rsACADEMYget.movenext
				i=i+1
			loop
		end if
		rsACADEMYget.Close
	end sub

	'//academy/board/lecturer/lecturer_read.asp '/academy/board/lecturer/lecturer_write.asp
    public Sub fnGetlecturerView()
        dim sqlStr
        dim isUpcheSsn : isUpcheSsn =  (CLNG(session("ssBctDiv"))>10) ''(session("ssBctDiv")>"10")
        
		sqlStr = " SELECT A.id, B.company_name, A.doc_regdate, A.doc_status, A.doc_type, A.doc_important" & vbCRLF
		sqlStr = sqlStr & " , A.doc_difficult, A.doc_subject, A.doc_content, A.doc_useyn, t.part_sn, A.admin_usingyn, A.pushsenddate" & vbCRLF
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lecturer_board_document AS A " & vbCRLF
		sqlStr = sqlStr & " INNER JOIN "&tmp_tbl&"[db_partner].[dbo].tbl_partner AS B " & vbCRLF
		sqlStr = sqlStr & " 	ON A.id = B.id " & vbCRLF
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr &" WHERE A.doc_idx = " & FrectDoc_Idx & vbCRLF
        
        ''업체는 자기 글 or 공지글 만 보이게.
        if (isUpcheSsn) then
            sqlStr = sqlStr & " and ((A.id='"&session("ssBctID")&"') or (A.doc_type='G010'))"
        end if
        
        if (FRECTAdmin_UsingNInclude="on") then
            
        else
            sqlStr = sqlStr & " and A.admin_usingyn='Y'"  '' 관리자가 N로 지정한것은 안보이게 ''2016/08/25
        end if
        
        'response.write sqlStr&"<br>"
        rsACADEMYget.Open SqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new clecturer_item
        
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
			FOneItem.fpushsenddate = rsACADEMYget("pushsenddate")
        end if
        rsACADEMYget.Close
    end Sub

	'/academy/board/lecturer/lecturer_write.asp '//academy/board/lecturer/lecturer_read.asp
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

	''/academy/board/lecturer/iframe_lecurer_ans.asp
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
		sqlStr = sqlStr & " A.ans_idx, A.ans_type, A.ans_content, A.ans_regdate, B.company_name, A.id" + vbcrlf
		sqlStr = sqlStr & " , t.part_sn" + vbcrlf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lecturer_board_ans AS A" + vbcrlf
		sqlStr = sqlStr & " JOIN "&tmp_tbl&"[db_partner].[dbo].tbl_partner AS B " + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = B.id" + vbcrlf
		sqlStr = sqlStr & " left join "&tmp_tbl&"db_partner.dbo.tbl_user_tenbyten t" + vbcrlf
		sqlStr = sqlStr & " 	ON A.id = t.userid" + vbcrlf
		sqlStr = sqlStr & " where A.ans_useyn = 'Y'" & strSubSql
		sqlStr = sqlStr & " ORDER BY A.ans_idx DESC" + vbcrlf
		
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
				set FItemList(i) = new clecturer_item
				FItemList(i).fpart_sn = rsACADEMYget("part_sn")
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

	''/academy/board/lecturer/iframe_lecurer_ans.asp
    public Sub fnGetolectView()
        dim sqlStr
		sqlStr = " SELECT top 1 ans_content" & _
				" FROM [db_academy].[dbo].tbl_lecturer_board_ans " & _
				" WHERE  ans_useyn = 'Y' AND ans_idx = '" & FrectAns_Idx & "' AND id = '" & session("ssBctId") & "' "

        'response.write sqlStr&"<br>"
        rsACADEMYget.Open sqlStr, dbACADEMYget, 1
        FResultCount = rsACADEMYget.RecordCount
        
        set FOneItem = new clecturer_item
        
        if Not rsACADEMYget.Eof then
    			
			FOneItem.FAns_Content = rsACADEMYget("ans_content")
			           
        end if
        rsACADEMYget.Close
    end Sub
        

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
 		IF application("Svr_Info")="Dev" THEN 
 			tmp_tbl = "[TENDB]."
 		else
 			tmp_tbl = "[TENDB]."
 		end if			
	End Sub
	Private Sub Class_Terminate()

	End Sub
	
	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

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
						If GetFileName() = "lecturer"Then
							sBody = sBody & "<option value=''>-선택-</option> "
						End IF
					End IF
					'//강사일경우 핑거스 공지사항은 노출안함
					if fnlecturer(session("ssBctId")) then
						if rsACADEMYget("commcd") <> "G010" then
							sBody = sBody & "<option value='" & rsACADEMYget("commcd") & "' "
							If CStr(sCode) = CStr(rsACADEMYget("commcd")) Then
								sBody = sBody & "selected"
							End If
							sBody = sBody & ">" & rsACADEMYget("commnm") & "</option>"
						end if	

					'//핑거스아카데미 일경우 다 출력
					else
						sBody = sBody & "<option value='" & rsACADEMYget("commcd") & "' "
						If CStr(sCode) = CStr(rsACADEMYget("commcd")) Then
							sBody = sBody & "selected"
						End If
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

function DrawMainPosCodeCombo(selectBoxName,selectedId)
%>
<select name="<%=selectBoxName%>">
	<option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
	<option value="a.doc_idx" <%if selectedId="a.doc_idx" then response.write " selected"%>>번호</option>
	<option value="a.doc_subject" <%if selectedId="a.doc_subject" then response.write " selected"%>>제목</option>
	<option value="a.doc_content" <%if selectedId="a.doc_content" then response.write " selected"%>>내용</option>
	<option value="a.id" <%if selectedId="a.id" then response.write " selected"%>>ID</option>
	<option value="c.company_name" <%if selectedId="c.company_name" then response.write " selected"%>>이름</option>		
</select>
<%   
end function

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
 
function fingmaster()
	if session("ssBctId") = "tozzinet" or session("ssBctId") = "thefingers01" or session("ssBctId") = "nownhere21" or session("ssBctId") = "frudia79" then
		fingmaster = true
	else
		fingmaster = false
	end if	
end function

function getthefingers_staff(staff, part_sn, username)
	dim tmpusername

	if part_sn=16 then
		tmpusername="더핑거스"
	else
		tmpusername=username
	end if

	getthefingers_staff = tmpusername
end function
%>
