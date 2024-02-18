<%
'###########################################################
' Description :  오프라인 게시판
' History : 2010.06.18 한용민 생성
'###########################################################
%>
<%
Class clecturer_item
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

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
	public fshopidcount
	public fdispshop_nm
	public fusername
	public fread_count
	public fdispshopall
	public fdispshopdiv
	public fshopname
	public fshopid
	public fregdate
	public fdoc_kind
	public fdoc_kind_nm
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
	public FrectAns_Idx
	public frectsearchKey
	public frectsearchString
	public frectdispshop
	public frectshopid
	public frectuserid
	public FDoc_WorkerName
	public FDoc_WorkerViewdate
	public frectdoc_kind
	
	''/admin/offshop/board/offshop_board.asp
	public sub fnGetboardList()
		dim sqlStr,i , strSubSql
		
		'//종류
		if frectdoc_kind<>"" then
			strSubSql = strSubSql & " and a.doc_kind='" & frectdoc_kind & "' "
		end if		
		'//요청구분
		if FrectDoc_Type<>"" then
			strSubSql = strSubSql & " and a.doc_type='" & FrectDoc_Type & "' "
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

		'//권한별 조건
		'//매장권한(99본사권한)
		if frectdispshop <> "99" then
			strSubSql = strSubSql & " and ("
			strSubSql = strSubSql & "	(a.dispshopall = 'Y' and ("&frectdispshop&" = '1' or "&frectdispshop&" = '3' or "&frectdispshop&" = '7'))"	'모든매장권한
			strSubSql = strSubSql & "	or a.dispshopdiv='"&frectdispshop&"'"		'매장별 접속한 아이디별 권한
			strSubSql = strSubSql & "	or replace(a.dispshopdiv,'90','1')='"&frectdispshop&"'"		'직영점+가맹점처리
			strSubSql = strSubSql & "	or replace(a.dispshopdiv,'90','3')='"&frectdispshop&"'"		'직영점+가맹점처리
						
			if frectshopid <> "" then
				strSubSql = strSubSql & " 	or a.doc_idx in ("
				strSubSql = strSubSql & " 		select distinct doc_idx"
				strSubSql = strSubSql & " 		from db_shop.dbo.tbl_offshop_board_shop"
				strSubSql = strSubSql & " 		where shopid = '"&frectshopid&"'"
				strSubSql = strSubSql & " 	)"		'특정매장 권한
			end if
			
			strSubSql = strSubSql & " 	or a.id = '"&frectuserid&"'"	'작성한 매장 자기 매장글 보기
			strSubSql = strSubSql & " )"
		end if
		
		'총 갯수 구하기
		sqlStr = "select count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/'"&FPageSize&"' ) as totPg" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_offshop_board_document AS A" + vbcrlf
		sqlStr = sqlStr & " join db_partner.dbo.tbl_partner c"
		sqlStr = sqlStr & "		ON A.id = C.id"
		'sqlStr = sqlStr & "		and c.isusing='Y'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten ut"
		sqlStr = sqlStr & "		ON A.id = ut.userid"
		'sqlStr = sqlStr & "		and ut.isusing=1"
		sqlStr = sqlStr & " where A.doc_useyn = 'Y'" & strSubSql
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit sub
		end if
	
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " A.doc_idx, A.doc_subject, A.doc_type, A.doc_important, A.doc_difficult"
		sqlStr = sqlStr & " , A.doc_status, A.doc_regdate, ut.username, A.doc_ans_ox, A.id"
		sqlStr = sqlStr & " ,a.dispshopall ,a.dispshopdiv ,a.doc_kind ,c.company_name"
		sqlStr = sqlStr & " ,(select codename"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_commoncode"
		sqlStr = sqlStr & " 	where codeid = a.doc_type and codekind = 'G000') as doc_type_nm"
		sqlStr = sqlStr & " ,(select codename"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_commoncode"
		sqlStr = sqlStr & " 	where codeid = a.doc_important and codekind = 'L000') as doc_important_nm"
		sqlStr = sqlStr & " ,(select codename"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_commoncode"
		sqlStr = sqlStr & " 	where codeid = a.doc_status and codekind = 'K000') as doc_status_nm"
		sqlStr = sqlStr & " ,(select codename"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_commoncode"
		sqlStr = sqlStr & " 	where codeid = a.dispshopdiv and codekind = 'A000') as dispshop_nm"
		sqlStr = sqlStr & " ,(select count(*)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_board_ans"
		sqlStr = sqlStr & " 	where doc_idx = a.doc_idx and ans_useyn='Y') as ans_count"
		sqlStr = sqlStr & " ,(select count(*)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_board_read"
		sqlStr = sqlStr & " 	where doc_idx = a.doc_idx) as read_count"
		sqlStr = sqlStr & " ,(select count(*)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_board_shop"
		sqlStr = sqlStr & " 	where doc_idx = a.doc_idx) as shopidcount"
		sqlStr = sqlStr & " ,(select codename"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_commoncode"
		sqlStr = sqlStr & " 	where codeid = a.doc_kind and codekind = 'doc_kind') as doc_kind_nm"		
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_offshop_board_document AS A"
		sqlStr = sqlStr & " join db_partner.dbo.tbl_partner c"
		sqlStr = sqlStr & "		ON A.id = C.id"
		'sqlStr = sqlStr & "		and c.isusing='Y'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten ut"
		sqlStr = sqlStr & "		ON A.id = ut.userid"
		'sqlStr = sqlStr & "		and ut.isusing=1"
		sqlStr = sqlStr & " where A.doc_useyn = 'Y'" & strSubSql
		sqlStr = sqlStr & " ORDER BY A.doc_idx DESC"
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new clecturer_item
								
				FItemList(i).fdoc_kind_nm = rsget("doc_kind_nm")
				FItemList(i).fdoc_kind = rsget("doc_kind")
				FItemList(i).fread_count = rsget("read_count")
				FItemList(i).fans_count = rsget("ans_count")
				FItemList(i).fdoc_idx = rsget("doc_idx")
				FItemList(i).fdoc_id = rsget("id")
				FItemList(i).fdoc_subject = db2html(rsget("doc_subject"))
				FItemList(i).fdoc_type = rsget("doc_type")
				FItemList(i).fdoc_important = rsget("doc_important")
				FItemList(i).fdoc_difficult = rsget("doc_difficult")
				FItemList(i).fdoc_status = rsget("doc_status")
				FItemList(i).fdoc_regdate = rsget("doc_regdate")
				FItemList(i).fusername = db2html(rsget("username"))
				
				if FItemList(i).fusername = "" then
					FItemList(i).fusername = rsget("company_name")
				end if
				
				FItemList(i).fdoc_ans_ox = rsget("doc_ans_ox")
				FItemList(i).fdoc_important_nm = rsget("doc_important_nm")
				FItemList(i).fdoc_type_nm = rsget("doc_type_nm")
				FItemList(i).fdoc_status_nm = rsget("doc_status_nm")
				FItemList(i).fshopidcount = rsget("shopidcount")
				FItemList(i).fdispshop_nm = rsget("dispshop_nm")
				FItemList(i).fdispshopall = rsget("dispshopall")
				FItemList(i).fdispshopdiv = rsget("dispshopdiv")
											
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'/admin/offshop/board/offshop_board_write.asp '/admin/offshop/board/offshop_board_read.asp
    public Sub fnGetlecturerread()
        dim sqlStr

		SqlStr = "SELECT"
		SqlStr = SqlStr & " A.worker_id, isnull(B.username,c.company_name) as username, Convert(varchar(20),A.worker_viewdate,120) AS worker_viewdate"
		SqlStr = SqlStr & " FROM db_shop.dbo.tbl_offshop_board_read A"
		SqlStr = SqlStr & " join db_partner.dbo.tbl_partner c"
		SqlStr = SqlStr & " 		ON A.worker_id = C.id"
		'SqlStr = SqlStr & "			and c.isusing='Y'"	
		SqlStr = SqlStr & " left JOIN [db_partner].[dbo].tbl_user_tenbyten AS B"
		SqlStr = SqlStr & " 		ON A.worker_id = B.userid"
		SqlStr = SqlStr & " WHERE doc_idx = '" & FrectDoc_Idx & "'"
		SqlStr = SqlStr & " ORDER BY A.idx ASC"
				
		'response.write SqlStr&"<br>"		
		rsget.Open SqlStr,dbget,1
		IF not rsget.EOF THEN
			Do Until rsget.Eof
				
				FDoc_WorkerName = FDoc_WorkerName & rsget("username") & ","
				FDoc_WorkerViewdate = FDoc_WorkerViewdate & rsget("worker_viewdate") & ","
				
			rsget.MoveNext
			Loop
		END IF
		rsget.close
		
		if FDoc_WorkerName <> "" then
			FDoc_WorkerName = Left(FDoc_WorkerName,Len(FDoc_WorkerName)-1)
			FDoc_WorkerViewdate = Left(FDoc_WorkerViewdate,Len(FDoc_WorkerViewdate)-1)
		end if	        
    end Sub
    
	'/admin/offshop/board/offshop_board_write.asp '/admin/offshop/board/offshop_board_read.asp
    public Sub fnGetlecturerView()
        dim sqlStr
		sqlStr = " SELECT A.id, ut.username, A.doc_regdate, A.doc_status"
		sqlStr = sqlStr & " ,A.doc_type, A.doc_important, A.doc_difficult, A.doc_subject, A.doc_content, A.doc_useyn"
		sqlStr = sqlStr & " ,A.dispshopall ,A.dispshopdiv ,a.doc_kind ,c.company_name"
		sqlStr = sqlStr & " ,(select count(*)"
		sqlStr = sqlStr & " 	from db_shop.dbo.tbl_offshop_board_shop"
		sqlStr = sqlStr & " 	where doc_idx = a.doc_idx) as shopidcount"
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_offshop_board_document AS A"
		sqlStr = sqlStr & " join db_partner.dbo.tbl_partner c"
		sqlStr = sqlStr & "		ON A.id = C.id"
		'sqlStr = sqlStr & "		and c.isusing='Y'"
		sqlStr = sqlStr & " left join db_partner.dbo.tbl_user_tenbyten ut"
		sqlStr = sqlStr & "		ON A.id = ut.userid"
		'sqlStr = sqlStr & "		and ut.isusing=1"
		sqlStr = sqlStr & " WHERE A.doc_idx = " & FrectDoc_Idx

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new clecturer_item
        
        if Not rsget.Eof then
    			
    		FOneItem.fdoc_kind = rsget("doc_kind")	
			FOneItem.FDoc_Id = rsget("id")
			FOneItem.fusername = db2html(rsget("username"))
			
			if FOneItem.fusername = "" then
				FOneItem.fusername = db2html(rsget("company_name"))
			end if
			
			FOneItem.FDoc_Status = rsget("doc_status")
			FOneItem.FDoc_Type = rsget("doc_type")
			FOneItem.FDoc_Import = rsget("doc_important")
			FOneItem.FDoc_Diffi = rsget("doc_difficult")
			FOneItem.FDoc_Subj = db2html(rsget("doc_subject"))
			FOneItem.FDoc_Content = db2html(rsget("doc_content"))
			FOneItem.FDoc_UseYN = rsget("doc_useyn")
			FOneItem.FDoc_Regdate = rsget("doc_regdate")
			FOneItem.fshopidcount = rsget("shopidcount")
			FOneItem.fdispshopall = rsget("dispshopall")
			FOneItem.fdispshopdiv = rsget("dispshopdiv")
			           
        end if
        rsget.Close      
    end Sub

	'/admin/offshop/board/offshop_board_write.asp '/admin/offshop/board/offshop_board_read.asp
    public Sub getShopList()
        dim sqlStr, i
        
        sqlStr = "select a.doc_idx ,a.shopid ,a.regdate ,u.shopname"
        sqlStr = sqlStr& " from db_shop.dbo.tbl_offshop_board_shop a"
        sqlStr = sqlStr& " join db_shop.dbo.tbl_shop_user u"
        sqlStr = sqlStr& "		on a.shopid = u.userid"
        sqlStr = sqlStr& " where a.doc_idx="&FrectDoc_Idx
        sqlStr = sqlStr& " order by a.shopid asc"
        
        'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.RecordCount
        
        
        if  not rsget.EOF  then
			redim preserve FItemList(FResultCount)
			
			do until rsget.EOF
				set FItemList(i) = new clecturer_item
				
				FItemList(i).fshopname  = db2html(rsget("shopname"))
				FItemList(i).fdoc_idx  = rsget("doc_idx")
				FItemList(i).fshopid    = rsget("shopid")
				FItemList(i).fregdate  = rsget("regdate")

				rsget.movenext
				i=i+1
			loop
		end if
        rsget.Close
    end sub
    
	'/admin/offshop/board/offshop_board_write.asp '/admin/offshop/board/offshop_board_read.asp
	public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_name " & _
				"		FROM db_shop.dbo.tbl_offshop_board_file " & _
				"	WHERE doc_idx = '" & FrectDoc_Idx & "' " & _
				"	ORDER BY file_idx ASC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetFileList = rsget.getRows() 
		END IF	
		rsget.close
	End Function

	''/admin/offshop/board/iframe_board_ans.asp
	public sub fnGetoboardList()
		dim sqlStr,i , strSubSql

		strSubSql = " AND A.doc_idx = '" & FrectDoc_Idx & "' "

		'총 갯수 구하기
		sqlStr = "select count(A.ans_idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_offshop_board_ans AS A" + vbcrlf
		sqlStr = sqlStr & " join db_partner.dbo.tbl_partner c"
		sqlStr = sqlStr & "		ON A.id = C.id"
		'sqlStr = sqlStr & "		and c.isusing='Y'"
		sqlStr = sqlStr & " Left join db_partner.dbo.tbl_user_tenbyten ut"
		sqlStr = sqlStr & "		ON A.id = ut.userid"
		'sqlStr = sqlStr & "		and ut.isusing=1"
		sqlStr = sqlStr & " where A.ans_useyn = 'Y'" & strSubSql
			
		'response.write sqlStr &"<br>"					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " A.ans_idx, A.ans_type, A.ans_content, A.ans_regdate, ut.username, A.id,c.company_name" + vbcrlf
		sqlStr = sqlStr & " FROM db_shop.dbo.tbl_offshop_board_ans AS A" + vbcrlf
		sqlStr = sqlStr & " join db_partner.dbo.tbl_partner c"
		sqlStr = sqlStr & "		ON A.id = C.id"
		'sqlStr = sqlStr & "		and c.isusing='Y'"
		sqlStr = sqlStr & " Left join db_partner.dbo.tbl_user_tenbyten ut"
		sqlStr = sqlStr & "		ON A.id = ut.userid"
		'sqlStr = sqlStr & "		and ut.isusing=1"
		sqlStr = sqlStr & " where A.ans_useyn = 'Y'" & strSubSql
		sqlStr = sqlStr & " ORDER BY A.ans_idx DESC" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

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
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new clecturer_item
				
				FItemList(i).fcompany_name = db2html(rsget("company_name"))
				FItemList(i).fans_idx = rsget("ans_idx")
				FItemList(i).fans_type = rsget("ans_type")
				FItemList(i).fans_content = db2html(rsget("ans_content"))
				FItemList(i).fans_regdate = rsget("ans_regdate")
				FItemList(i).fusername = db2html(rsget("username"))
				FItemList(i).fid = rsget("id")
											
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''/admin/offshop/board/iframe_board_ans.asp
    public Sub fnGetoboardView()
        dim sqlStr
		sqlStr = " SELECT top 1 ans_content" & _
				" FROM db_shop.dbo.tbl_offshop_board_ans" & _
				" WHERE ans_useyn = 'Y'" & _
				" AND ans_idx = '" & FrectAns_Idx & "'" & _
				" AND id = '" & session("ssBctId") & "'"

        'response.write sqlStr&"<br>"
        rsget.Open sqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new clecturer_item
        
        if Not rsget.Eof then
    			
			FOneItem.FAns_Content = rsget("ans_content")
			           
        end if
        rsget.Close
    end Sub
        

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0		
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

'####### 코드매니저 페이지 외 하나씩 불러쓰는 공통 코드 관리 selectbox. write 용, view 용 #######
Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode ,C_ADMIN_USER ,style)
	Dim strSql, sBody, i
	sBody = ""
	i = 0
	if C_ADMIN_USER = "" then C_ADMIN_USER = false
		
	'### sUse = "w" write 용
	If sUse = "w" Then
		strSql = " SELECT codeid, codename "
		strSql = strSql & " From db_shop.dbo.tbl_offshop_commoncode "
		strSql = strSql & " WHERE codekind ='"&sType&"' AND useyn = 'Y' ORDER BY codeid ASC"
		
		If sType = "doc_status" AND sCode = "" Then
			sBody = "<input type='hidden' name='doc_status' value='01'>작성"
		Else
			
			'response.write strSql &"<br>"
			rsget.Open strSql,dbget,1
			
			Do Until rsget.Eof				
				If sType = "K000" Then
					If i = 0 Then
						sBody = "<select name='"&sType&"' class='select' "&style&">"
						If GetFileName() = "offshop_board" Then
							sBody = sBody & "<option value=''>-선택-</option> "
						End IF
					End IF
					sBody = sBody & "<option value='" & rsget("codeid") & "' "
					If CStr(sCode) = CStr(rsget("codeid")) Then
						sBody = sBody & "selected"
					End If
					sBody = sBody & ">" & rsget("codename") & "</option>"
					If i = rsget.RecordCount-1 Then
						sBody = sBody & "</select>"
					End IF

				elseIf sType = "doc_kind" Then
					If i = 0 Then
						sBody = "<select name='"&sType&"' class='select' "&style&">"
							sBody = sBody & "<option value=''>-선택-</option> "
					End IF

					sBody = sBody & "<option value='" & rsget("codeid") & "' "
					If CStr(sCode) = CStr(rsget("codeid")) Then
						sBody = sBody & "selected"
					End If
					sBody = sBody & ">" & rsget("codename") & "</option>"
					If i = rsget.RecordCount-1 Then
						sBody = sBody & "</select>"
					End IF
					
				elseIf sType = "A000" Then
					'If i = 0 Then
					'	sBody = "<select name='"&sType&"' class='select' "&style&">"						
					'		sBody = sBody & "<option value=''>-선택-</option> "
					'End IF
					'sBody = sBody & "<option value='" & rsget("codeid") & "' "
					'If CStr(sCode) = CStr(rsget("codeid")) Then
					'	sBody = sBody & "selected"
					'End If
					'sBody = sBody & ">" & rsget("codename") & "</option>"
					'If i = rsget.RecordCount-1 Then
					'	sBody = sBody & "</select>"
					'End IF
					
					sBody = sBody & "<input type='radio' name='" & sType & "' id='" & sType & rsget("codeid") & "' value='" & rsget("codeid") & "' "
					If sCode = rsget("codeid") Then
						sBody = sBody & "checked"
					End If
					sBody = sBody & ">" & rsget("codename") & "&nbsp;&nbsp;"
				
				elseIf sType = "G000" Then
					If i = 0 Then
						sBody = "<select name='"&sType&"' class='select' "&style&">"
							sBody = sBody & "<option value=''>-선택-</option> "
					End IF
					
					If GetFileName() = "offshop_board" Then
						sBody = sBody & "<option value='" & rsget("codeid") & "' "
						If CStr(sCode) = CStr(rsget("codeid")) Then
							sBody = sBody & "selected"
						End If
						sBody = sBody & ">" & rsget("codename") & "</option>"
												
					'//본사 직원일 경우
					elseif C_ADMIN_USER then
						if rsget("codeid") = "01" then
							sBody = sBody & "<option value='" & rsget("codeid") & "' "
							If CStr(sCode) = CStr(rsget("codeid")) Then
								sBody = sBody & "selected"
							End If
							sBody = sBody & ">" & rsget("codename") & "</option>"													
						end if
					else
						if rsget("codeid") <> "01" then
							sBody = sBody & "<option value='" & rsget("codeid") & "' "
							If CStr(sCode) = CStr(rsget("codeid")) Then
								sBody = sBody & "selected"
							End If
							sBody = sBody & ">" & rsget("codename") & "</option>"													
						end if					
					end if
					
					If i = rsget.RecordCount-1 Then
						sBody = sBody & "</select>"
					End IF							
				Else
					sBody = sBody & "<label id='" & sType & rsget("codeid") & "'>" & _
									"<input type='radio' name='" & sType & "' id='" & sType & rsget("codeid") & "' value='" & rsget("codeid") & "' "
					If CStr(sCode) = CStr(rsget("codeid")) Then
						sBody = sBody & "checked"
					End If
					sBody = sBody & ">" & rsget("codename") & "</label>&nbsp;&nbsp;"
				End If
			rsget.MoveNext
			i = i + 1
			Loop
			rsget.Close
		End If
	Else
	'### sUse = "v" view 용
		strSql = " SELECT codename From db_shop.dbo.tbl_offshop_commoncode"
		strSql = strSql & " WHERE codekind ='"&sType&"' AND codeid = '" & sCode & "' AND useyn = 'Y'"
		
		rsget.Open strSql,dbget
		If Not rsget.Eof Then
			sBody = rsget(0)
		End If
		rsget.Close
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
	<option value="ut.username" <%if selectedId="ut.username" then response.write " selected"%>>이름</option>		
</select>
<%   
end function

'글 읽은 시간 저장
Public Function WorkerView(iDoc_idx)
Dim strSql

	if iDoc_idx = "" then exit Function

	strSql = " IF EXISTS(" &_
			 " 		SELECT worker_viewdate FROM db_shop.dbo.tbl_offshop_board_read" & _
			 " 		WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "'" & _
			 " )" & _
			 "		UPDATE db_shop.dbo.tbl_offshop_board_read SET" & _
			 " 		lastupdate = getdate()" & _
			 " 		WHERE doc_idx = '" & iDoc_idx & "' AND worker_id = '" & session("ssBctId") & "'" & _
			 " else" &_
			 " 		insert into db_shop.dbo.tbl_offshop_board_read (doc_idx ,worker_id ,worker_viewdate ,lastupdate)" &_
			 "		values("&iDoc_idx&",'" & session("ssBctId") & "',getdate(),getdate())"
	
	'response.write strSql &"<Br>"
	dbget.execute strSql
End Function
%>
