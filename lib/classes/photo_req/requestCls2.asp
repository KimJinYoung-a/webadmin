<%
	Class Photoreq
		public FCPage	'Set 현재 페이지
		public FPSize	'Set 페이지 사이즈
		public FTotCnt
		public FComeCnt
		public FSendCnt
		Public Fbrd_sn
		public FTeam
		public FDoc_Idx
		public FDoc_Subj
		public FDoc_Id
		public FDoc_Regdate
		public FDoc_Start
		public FDoc_Status
		public FDoc_End
		public FDoc_Name
		public FDoc_Type
		public FDoc_Import
		public FDoc_Diffi
		public FDoc_Content
		public FDoc_UseYN
		public FDoc_WorkerName
		public FDoc_Worker
		public FDoc_WorkerViewdate
		public FDoc_Refer
		public FDoc_ReferName
		public FDoc_AnsOX
		public FDoc_MineOX

		Public Smakerid
		Public Sreq_use
		Public Ss_type
		Public Sitemid
		Public Sreq_status_type
		Public Srequest_name
		Public Sreq_photo_user

		public FAns_Idx
		public FAns_Content

		'####### 촬영요청리스트 #######
		public Function fnGetPhotoreqList
			Dim strSql,iDelCnt, strSubSql


			strSql = " SELECT COUNT(DISTINCT req_no) From " & _
					" 		[db_partner].[dbo].tbl_photo_req where req_use <> '' and use_yn = 'Y' "

			'브랜드 검색
			If Smakerid <> "" Then
				strsql = strsql + " and brand_id = '"&Smakerid&"'"
			End If
			'촬영용도 검색
			If Sreq_use <> "" Then
				strsql = strsql + " and req_use = '"&Sreq_use&"'"
			End If
			'no/상품명 검색
			If Ss_type <> "" Then
				If Ss_type = "req_no" Then
					strsql = strsql + " and req_use = '"&Sitemid&"'"
				ElseIf Ss_type = "prd_name" Then
					strsql = strsql + " and prd_name = '"&Sitemid&"'"
				End If
			End If
			'진행상태 검색
			If Sreq_status_type <> "" Then
				strsql = strsql + " and req_status = '"&Sreq_status_type&"'"
			End If
			'촬영요청자 검색
			If Srequest_name <> "" Then
				strsql = strsql + " and req_name = '"&Srequest_name&"'"
			End If
			'진행상태 검색
			If Sreq_photo_user <> "" Then
				strsql = strsql + " and photo_id = '"&Sreq_photo_user&"'"
			End If

			rsget.Open strSql,dbget
			'response.write strSql&"<BR>"
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
				iDelCnt =  ((FCPage - 1) * FPSize )

				strSql = "	SELECT DISTINCT TOP "&FPSize&" " & _
						"		req_no, req_status, req_use, req_use_detail, prd_name, req_category, brand_id, photo_id, req_name, req_department"&_
						"	FROM [db_partner].[dbo].tbl_photo_req where req_use <> '' and use_yn = 'Y' "


				'브랜드 검색
				If Smakerid <> "" Then
					strsql = strsql + " and brand_id = '"&Smakerid&"'"
				End If
				'촬영용도 검색
				If Sreq_use <> "" Then
					strsql = strsql + " and req_use = '"&Sreq_use&"'"
				End If
				'no/상품명 검색
				If Ss_type <> "" Then
					If Ss_type = "req_no" Then
						strsql = strsql + " and req_no = '"&Sitemid&"'"
					ElseIf Ss_type = "prd_name" Then
						strsql = strsql + " and prd_name = '"&Sitemid&"'"
					End If
				End If
				'진행상태 검색
				If Sreq_status_type <> "" Then
					strsql = strsql + " and req_status = '"&Sreq_status_type&"'"
				End If
				'촬영요청자 검색
				If Srequest_name <> "" Then
					strsql = strsql + " and req_name = '"&Srequest_name&"'"
				End If
				'진행상태 검색
				If Sreq_photo_user <> "" Then
					strsql = strsql + " and photo_id = '"&Sreq_photo_user&"'"
				End If
					strsql = strsql + " order by req_no desc"

				'response.write strSql

				rsget.Open strSql,dbget

				IF not rsget.EOF THEN
					fnGetPhotoreqList = rsget.getRows()
				END IF
				rsget.close

			END IF
			'response.write strsql
		End Function


		'####### 직원리스트 #######
		public Function fnGetMemberList
			Dim strSql
			strSql = "	SELECT A.id, B.part_name, C.posit_name, D.username AS company_name, A.part_sn, " & _
					"		isNull((select Convert(varchar(20),worker_viewdate,120) from [db_partner].[dbo].tbl_cooperate_worker where doc_idx = '" & FDoc_Idx & "' and worker_id = A.id),'x') AS worker_viewdate " & _
					"		, isNull(E.halfgubun,'') AS vacation " & _
					"		FROM [db_partner].[dbo].tbl_partner AS A " & _
					"		INNER JOIN [db_partner].[dbo].tbl_partInfo AS B ON A.part_sn = B.part_sn " & _
					"		INNER JOIN [db_partner].[dbo].tbl_positInfo AS C ON A.posit_sn = C.posit_sn " & _
					"		INNER JOIN [db_partner].[dbo].tbl_user_tenbyten AS D ON A.id = D.userid " & _
					"		LEFT JOIN ( " & _
					"			SELECT mm.userid, dd.halfgubun FROM [db_partner].[dbo].tbl_vacation_master AS mm " & _
					"			INNER JOIN [db_partner].[dbo].tbl_vacation_detail AS dd on mm.idx = dd.masteridx " & _
					"			WHERE mm.deleteyn <> 'Y' AND dd.deleteyn <> 'Y' AND dd.statedivcd IN ('R','A') AND ('" & date() & "' between convert(varchar(10),dd.startday,120) and convert(varchar(10),dd.endday,120)) " & _
					"		) AS E ON A.id = E.userid " & _
					"	WHERE A.isusing = 'Y' AND A.userdiv < 999 AND A.id <> '' AND Left(A.id,10) <> 'streetshop' AND B.part_name <> '오프라인 - 취화선' AND C.posit_name <> '임시직' " & _
					"			AND A.part_sn IN(" & FTeam & ") " & _
					"	ORDER BY A.part_sn ASC, A.posit_sn ASC, A.regdate ASC "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				fnGetMemberList = rsget.getRows()
			END IF
			rsget.close
		End Function

		Public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_no, file_name, file_regdate " & _
				"		FROM [db_partner].[dbo].tbl_photo_file " & _
				"	WHERE req_no = '" & Fbrd_sn & "' " & _
				"	ORDER BY file_no ASC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetFileList = rsget.getRows()
		END IF
		rsget.close
	End Function

	End Class


	'####### 코드 매니저 페이지용. 리스트. #######
	Class CCoopCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort

		'####### 공통코드 리스트 #######
		public Function fnGetCoopCodeList
			IF FCodeType = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT code_type, code_value, code_name, code_useyn, code_sort "&_
					" From [db_partner].[dbo].[tbl_photo_code] "&_
					" WHERE code_type = '"&FCodeType&"' ORDER BY code_sort ASC "
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetCoopCodeList = rsget.getRows()
			End IF
			rsget.Close
		End Function

		'####### 선택한 코드 내용 가져오기 #######
		public Function fnGetCoopCodeCont
			IF FCodeValue = "" or FCodeType = ""  THEN Exit Function
			Dim strSql
			strSql = " SELECT code_type, code_value, code_name, code_useyn, code_sort "&_
					" From  [db_partner].[dbo].[tbl_photo_code] "&_
					" WHERE code_value = "&FCodeValue&" and code_type ='"&FCodeType&"'"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FCodeType 	= rsget("code_type")
				FCodeValue 	= rsget("code_value")
				FCodeDesc 	= rsget("code_name")
				FCodeUsing 	= rsget("code_useyn")
				FCodeSort	= rsget("code_sort")
			End IF
			rsget.Close
		End Function
	End Class

	'####### 작업자리스트(촬영요청페이지). #######
	Class CCoopUserList
	Public FUserList()
	Public FUser_no
	public FUserType
	public FUser_id
	public FUser_name
	public FUser_useyn
	Public FCodeType
	Public FMode
	Public FResultCount
	Public FTotalCount
	Public FPageCount
	Public FCurrPage

	'작업자 리스트
	public Sub fnGetCoopUserList
		Dim strSql, i
		strSql = ""
		strSql = strSql &  "SELECT user_no, user_type, user_id, user_name, user_useyn " & vbcrlf
		strSql = strSql &  "From [db_partner].[dbo].[tbl_photo_user] " & vbcrlf

		If FMode = "I" Then
			strSql = strSql &  "WHERE user_type = '"&FCodeType&"' ORDER BY user_no ASC "
		ElseIf FMode = "U" Then
			strSql = strSql &  "WHERE user_no = '"&FUser_no&"' "
		ElseIf FMode = "BB" Then
			strSql = strSql &  "WHERE user_id = '"&session("ssBctID")&"' "
		End If

		rsget.Open strSql,dbget,1

		FResultCount = rsget.recordcount
		FTotalCount = FResultCount

		Redim preserve FUserList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			Do until rsget.EOF
				set FUserList(i) = new CCoopUserList
				FUserList(i).FUser_no		= rsget("user_no")
				FUserList(i).FUserType 		= rsget("user_type")
				FUserList(i).FUser_id 		= rsget("user_id")
				FUserList(i).FUser_name 	= rsget("user_name")
				FUserList(i).FUser_useyn 	= rsget("user_useyn")
				rsget.Movenext
				i = i + 1
			Loop
		End If
		rsget.Close

	End Sub

	End Class


	'####### 코드 select 박스 #######
	Sub sbOptCodeType(ByVal selCodeType)
%>
		<option value="doc_status" <%IF Cstr(selCodeType)="doc_status" THEN%>selected<%END IF%>>촬영용도구분</option>
		<option value="doc_status_detail" <%IF Cstr(selCodeType)="doc_status_detail" THEN%>selected<%END IF%>>촬영용도구분_상세</option>
		<option value="doc_use_type" <%IF Cstr(selCodeType)="doc_use_type" THEN%>selected<%END IF%>>필요촬영군</option>
		<option value="doc_use_concept" <%IF Cstr(selCodeType)="doc_use_concept" THEN%>selected<%END IF%>>메인 촬영 컨셉</option>
		<option value="doc_ing_type" <%IF Cstr(selCodeType)="doc_ing_type" THEN%>selected<%END IF%>>진행상태</option>
<%
	End Sub


	'####### 코드매니저 페이지 외 하나씩 불러쓰는 공통 코드 관리. write 용, view 용. #######
	Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode)
		Dim strSql, sBody, i
		sBody = ""
		i = 0

		'### sUse = "w" write 용
		If sUse = "w" Then
			strSql = " SELECT code_value, code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_useyn = 'Y' ORDER BY code_sort ASC"

			If sType = "doc_status" AND sCode = "" Then
				sBody = "<input type='hidden' name='doc_status' value='1'>협조문 작성"
			Else
				rsget.Open strSql,dbget,1
				Do Until rsget.Eof
					'####### 업무구분은 select박스 형으로. 나머지는 radio 로. #######
					If sType = "doc_type" Then
						If i = 0 Then
							sBody = "<select name='doc_type' class='select'>"
							If GetFileName() = "index" OR GetFileName() = "my_cooperate" Then
								sBody = sBody & "<option value=''>-요청구분-</option> "
							End IF
						End IF
						sBody = sBody & "<option value='" & rsget("code_value") & "' "
						If CStr(sCode) = CStr(rsget("code_value")) Then
							sBody = sBody & "selected"
						End If
						sBody = sBody & ">" & rsget("code_desc") & "</option>"
						If i = rsget.RecordCount-1 Then
							sBody = sBody & "</select>"
						End IF
					Else
						If Left(sCode,1) = "s" Then
							'####### index 페이지에 현재상태 검색용. select박스. #######
							If i = 0 Then
								sBody = "<select name='doc_status' class='select'><option value='x'>-처리상태-</option><option value='0' "
								If CStr(Replace(sCode,"s","")) = CStr(0) Then
									sBody = sBody & "selected"
								End If
								sBody = sBody & ">미처리 전체</option>"
							End IF
							sBody = sBody & "<option value='" & rsget("code_value") & "' "
							If CStr(Replace(sCode,"s","")) = CStr(rsget("code_value")) Then
								sBody = sBody & "selected"
							End If
							sBody = sBody & ">" & rsget("code_desc") & "</option>"
							If i = rsget.RecordCount-1 Then
								sBody = sBody & "</select>"
							End IF
						Else
							sBody = sBody & "<label id='" & sType & rsget("code_value") & "'>" & _
											"<input type='radio' name='" & sType & "' id='" & sType & rsget("code_value") & "' value='" & rsget("code_value") & "' "
							If CStr(sCode) = CStr(rsget("code_value")) Then
								sBody = sBody & "checked"
							End If
							sBody = sBody & ">" & rsget("code_desc") & "</label>&nbsp;&nbsp;"
						End If
					End If
				rsget.MoveNext
				i = i + 1
				Loop
				rsget.Close
			End If
		Else
		'### sUse = "v" view 용
			strSql = " SELECT code_desc From [db_partner].[dbo].[tbl_cooperate_comCode] WHERE code_type ='"&sType&"' AND code_value = '" & sCode & "' AND code_useyn = 'Y'"
			rsget.Open strSql,dbget
			If Not rsget.Eof Then
				sBody = rsget(0)
			End If
			rsget.Close
		End If
		CommonCode = sBody
	End Function


	'####### 직원 연락처 Get. SMS 발송용. #######
	public Function fnGetMemberHp(id)
		Dim strSql
		strSql = "	SELECT isNull(usercell,'0') AS manager_hp FROM [db_partner].[dbo].tbl_user_tenbyten WHERE userid = '" & id & "' "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			If rsget("manager_hp") = "" Then
				fnGetMemberHp = "0"
			Else
				fnGetMemberHp = rsget("manager_hp")
			End If
		Else
			fnGetMemberHp = "0"
		END IF
		rsget.close
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
%>
