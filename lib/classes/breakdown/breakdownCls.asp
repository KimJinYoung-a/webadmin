<%
	Class CBreakdown
		public FCurrPage	'Set 현재 페이지
		public FPageSize	'Set 페이지 사이즈
		public FTotalCount
		public FComeCnt
		public FSendCnt
		public FResultCount
		public FTotalPage
		public FItemList()
		public FScrollCount

		public FTeam

		public FReqIdx
		public FReqDIdx
		public FReqEquipment
		public FReqEquipmentName
		public FReqPartSn
		public FWorkPartSn
		public FWorkType
		public FWorkTarget
		public FRequserid
		public FReqComment
		public FReqCapImage1
		public FReqSDate
		public FReqEDate
		public FReqState
		public FReqName
		public FReqDate
		public FRectMyOnly
		public FRectUserName

		'####### 시스템 장애 신청리스트 #######
		public Function fnGetBreakdownList
			Dim strSql,iDelCnt, strSubSql

			If FReqPartSn <> "" Then
				strSubSql = strSubSql & " AND r.req_part_sn IN(" & FReqPartSn & ") "
			End IF

			If FWorkPartSn <> "" Then
				strSubSql = strSubSql & " AND d.work_part_sn IN(" & FWorkPartSn & ") "
			End IF

			If FWorkType <> "" Then
				strSubSql = strSubSql & " AND d.work_type = '" & FWorkType & "' "
			End IF

			If FWorkTarget <> "" Then
				strSubSql = strSubSql & " AND d.work_target = '" & FWorkTarget & "' "
			End IF

			If FReqSDate <> "" Then
				strSubSql = strSubSql & " AND d.work_state = '5' AND Convert(varchar(10),d.work_lastupdate,120) >= '" & FReqSDate & "' "
			End IF

			If FReqEDate <> "" Then
				strSubSql = strSubSql & " AND d.work_state = '5' AND Convert(varchar(10),d.work_lastupdate,120) <= '" & FReqEDate & "' "
			End IF

			If FReqState <> "" Then
				If (FReqState = "N") Then
					strSubSql = strSubSql & " AND d.work_state < '5' "
				Else
					strSubSql = strSubSql & " AND d.work_state = '" & FReqState & "' "
				End If
			End IF

			If FRectMyOnly <> "" Then
				strSubSql = strSubSql & " AND (r.req_part_sn = " & session("ssAdminPsn") & " or d.work_part_sn = " & session("ssAdminPsn") & ")"
			End If

			If FRectUserName <> "" Then
				strSubSql = strSubSql & " AND t.username = '" & FRectUserName & "'"
			End IF

			strSql = " SELECT COUNT(r.idx) " & vbCrLf
			strSql = strSql & "	FROM [db_temp].[dbo].[tbl_breakdown_request] AS r " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_temp].[dbo].[tbl_breakdown_request_detail] AS d ON r.idx = d.req_idx " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS t ON r.req_userid = t.userid " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_partInfo] AS i ON t.part_sn = i.part_sn " & vbCrLf
			strSql = strSql & "	WHERE " & vbCrLf
			strSql = strSql & "		1=1 and d.isusing = 'Y' " & strSubSql & vbCrLf

			''response.write strSql
			''response.end
			rsget.Open strSql,dbget

			IF not rsget.EOF THEN
				FTotalCount = rsget(0)
			END IF
			rsget.close

			IF FTotalCount > 0 THEN
				strSql = "	SELECT top " + CStr(FPageSize*FCurrPage) + " " & vbCrLf
				strSql = strSql & "		r.idx, i.part_name, t.username, d.work_type, d.work_target, d.req_equipment, " & vbCrLf
				strSql = strSql & "		d.req_comment, d.work_comment, d.work_lastupdate, " & vbCrLf
				strSql = strSql & "		CASE d.work_state WHEN '1' THEN '신청' WHEN '3' THEN '작업중' WHEN '5' THEN '작업완료' END AS work_state " & vbCrLf
				strSql = strSql & "		, r.req_userid, d.work_state, d.idx, d.now_worker, isNull(d.req_captimage,'') AS req_captimage, ii.part_name, d.work_part_sn, wc.work_type_name, wc.work_target_name, tt.username, r.req_regdate, d.work_startdate " & vbCrLf
				strSql = strSql & "	FROM [db_temp].[dbo].[tbl_breakdown_request] AS r " & vbCrLf
				strSql = strSql & "		INNER JOIN [db_temp].[dbo].[tbl_breakdown_request_detail] AS d ON r.idx = d.req_idx " & vbCrLf
				strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS t ON r.req_userid = t.userid " & vbCrLf
				strSql = strSql & "		LEFT JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS tt ON d.now_worker = tt.userid " & vbCrLf
				strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_partInfo] AS i ON t.part_sn = i.part_sn " & vbCrLf
				strSql = strSql & "		LEFT JOIN [db_partner].[dbo].[tbl_partInfo] AS ii ON d.work_part_sn = ii.part_sn " & vbCrLf
				strSql = strSql & "		left join [db_temp].[dbo].[tbl_breakdown_work_code] wc on d.work_part_sn = wc.part_sn and d.work_type = wc.work_type and d.work_target = wc.work_target " & vbCrLf
				strSql = strSql & "	WHERE " & vbCrLf
				strSql = strSql & "		1=1 and d.isusing = 'Y' " & strSubSql & vbCrLf
				strSql = strSql & "	ORDER BY r.idx DESC "
				''response.write strSql
				''response.end

				rsget.pagesize = FPageSize
				rsget.Open strSql,dbget,1

				FtotalPage =  CLng(FTotalCount\FPageSize)
				if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
					FtotalPage = FtotalPage +1
				end if
				FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

				if FResultCount<1 then FResultCount=0

				IF not rsget.EOF THEN
					rsget.absolutepage = FCurrPage

					fnGetBreakdownList = rsget.getRows()
				END IF
				rsget.close
			END IF

		End Function


		'####### 시스템 장애 신청 보기 #######
		public Function fnGetBreakdownView
			Dim strSql
			strSql = "	SELECT r.req_userid, d.req_idx, d.work_type, d.work_target, d.req_equipment, d.req_comment, isNull(d.req_captimage,'') AS req_captimage, d.work_part_sn " & _
					"		FROM [db_temp].[dbo].[tbl_breakdown_request] AS r " & _
					"		INNER JOIN [db_temp].[dbo].[tbl_breakdown_request_detail] AS d ON r.idx = d.req_idx " & _
					"	WHERE d.idx = '" & FReqDIdx & "' "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				 FRequserid 	= rsget("req_userid")
				 FReqIdx 		= rsget("req_idx")
				 FWorkType		= rsget("work_type")
				 FWorkTarget	= rsget("work_target")
				 FReqEquipment	= rsget("req_equipment")
				 FReqComment	= db2html(rsget("req_comment"))
				 FReqCapImage1	= rsget("req_captimage")
				 FWorkPartSn	= rsget("work_part_sn")
			END IF
			rsget.close
		End Function


		'####### 시스템 장애 신청 보기_모바일용 #######
		public Function fnGetBreakdownMobileView
			Dim strSql

			response.Write "에러 : 시스템팀 문의!!"
			response.End

			strSql = "	SELECT " & vbCrLf
			strSql = strSql & "		r.idx, i.part_name, t.username, d.work_type, d.work_target, d.req_equipment, " & vbCrLf
			strSql = strSql & "		d.req_comment, d.work_comment, d.work_lastupdate, r.req_regdate, " & vbCrLf
			strSql = strSql & "		r.req_userid, d.work_state, d.idx, d.now_worker, isNull(d.req_captimage,'') AS req_captimage " & vbCrLf
			strSql = strSql & "		FROM [db_temp].[dbo].[tbl_breakdown_request] AS r " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_temp].[dbo].[tbl_breakdown_request_detail] AS d ON r.idx = d.req_idx " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_user_tenbyten] AS t ON r.req_userid = t.userid " & vbCrLf
			strSql = strSql & "		INNER JOIN [db_partner].[dbo].[tbl_partInfo] AS i ON t.part_sn = i.part_sn " & vbCrLf
			strSql = strSql & "	WHERE r.idx = '" & FReqIdx & "' "
			rsget.Open strSql,dbget,1
			'response.write strSql
			IF not rsget.EOF THEN
				 FReqIdx 		= rsget("idx")
				 FWorkType		= rsget("work_type")
				 FWorkTarget	= rsget("work_target")
				 FReqEquipment	= rsget("req_equipment")
				 FReqComment	= db2html(rsget("req_comment"))
				 FReqCapImage1	= rsget("req_captimage")
				 FTeam			= rsget("part_name")
				 FReqName		= rsget("username")
				 FReqDate		= rsget("req_regdate")
				 FReqState		= rsget("work_state")

			END IF
			rsget.close
		End Function

		Public Function HasPreScroll()
			HasPreScroll = StartScrollPage > 1
		End Function

		Public Function HasNextScroll()
			HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
		End Function

		Public Function StartScrollPage()
			StartScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
		End Function

	    Private Sub Class_Initialize()
			redim  FItemList(0)
			FScrollCount = 10
		End Sub

		Private Sub Class_Terminate()
	    End Sub
	End Class


	'####### 코드 매니저 페이지용. 리스트. #######
	Class CBreakCommonCode
	public FCodeType
	public FCodeValue
	public FCodeDesc
	public FCodeUsing
	public FCodeSort
	public FCodeComp
	public FCodeProd
	public FCodeGubun

		'####### 공통코드 리스트 #######
		public Function fnGetBreakCodeList
			IF FCodeType = "" THEN Exit Function
			Dim strSql
			strSql = "SELECT code_type, code_value, code_comp, code_prod, code_desc, code_useyn, code_sort "&_
					" From [db_temp].[dbo].[tbl_breakdown_comCode] "&_
					" WHERE code_type = '"&FCodeType&"' ORDER BY code_sort ASC "
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				fnGetBreakCodeList = rsget.getRows()
			End IF
			rsget.Close
		End Function

		'####### 선택한 코드 내용 가져오기 #######
		public Function fnGetBreakCodeCont
			IF FCodeValue = "" or FCodeType = ""  THEN Exit Function
			Dim strSql
			strSql = " SELECT code_type, code_value, code_comp, code_prod, code_desc, code_useyn, code_sort "&_
					" From  [db_temp].[dbo].[tbl_breakdown_comCode] "&_
					" WHERE code_value = "&FCodeValue&" and code_type ='"&FCodeType&"'"
			rsget.Open strSql,dbget
			IF not rsget.EOF THEN
				FCodeType 	= rsget("code_type")
				FCodeValue 	= rsget("code_value")
				FCodeComp 	= rsget("code_comp")
				FCodeProd 	= rsget("code_prod")
				FCodeDesc 	= rsget("code_desc")
				FCodeUsing 	= rsget("code_useyn")
				FCodeSort	= rsget("code_sort")
			End IF
			rsget.Close
		End Function


	End Class

	'####### 코드 select 박스 #######
	Sub sbOptCodeType(ByVal selCodeType)
%>
		<option value="pc_list" <%IF Cstr(selCodeType)="pc_list" THEN%>selected<%END IF%>>PC List</option>
		<option value="pos_list" <%IF Cstr(selCodeType)="pos_list" THEN%>selected<%END IF%>>POS List</option>
		<option value="moni_list" <%IF Cstr(selCodeType)="moni_list" THEN%>selected<%END IF%>>Moniter List</option>
		<option value="">-----------------</option>
		<option value="pc_list_break" <%IF Cstr(selCodeType)="pc_list_break" THEN%>selected<%END IF%>>PC 장애 List</option>
		<option value="pos_list_break" <%IF Cstr(selCodeType)="pos_list_break" THEN%>selected<%END IF%>>POS 장애 List</option>
		<option value="moni_list_break" <%IF Cstr(selCodeType)="moni_list_break" THEN%>selected<%END IF%>>모니터 장애 List</option>
		<option value="">-----------------</option>
		<option value="etc" <%IF Cstr(selCodeType)="etc" THEN%>selected<%END IF%>>기타 장애 List</option>
<%
	End Sub


	Function fnWorkType(code)
		Select Case code
			Case "1"
				fnWorkType = "신규"
			Case "2"
				fnWorkType = "교체"
			Case "3"
				fnWorkType = "장애처리"
			Case Else
				fnWorkType = ""
		End Select
	End Function

	Function fnWorkTargetName(code)
		Dim vTemp
		vTemp = Replace(code,"_break","")
		Select Case vTemp
			Case "pc_list"
				fnWorkTargetName = "PC"
			Case "pos_list"
				fnWorkTargetName = "POS"
			Case "moni_list"
				fnWorkTargetName = "모니터"
			Case "etc"
				fnWorkTargetName = "기타"
			Case Else
				fnWorkTargetName = ""
		End Select
	End Function

	Function fnWorkTargetCode3(code1,code2)
		If code1 = "3" Then
			Select Case code2
				Case "pc_list_break"
					fnWorkTargetCode3 = "PC 장애 List"
				Case "pos_list_break"
					fnWorkTargetCode3 = "POS 장애 List"
				Case "moni_list_break"
					fnWorkTargetCode3 = "모니터 장애 List"
				Case "etc"
					fnWorkTargetCode3 = "기타 장애 List"
				Case Else
					fnWorkTargetCode3 = ""
			End Select
		Else
			Select Case code2
				Case "pc_list"
					fnWorkTargetCode3 = "PC List"
				Case "pos_list"
					fnWorkTargetCode3 = "POS List"
				Case "moni_list"
					fnWorkTargetCode3 = "모니터 List"
				Case "etc"
					fnWorkTargetCode3 = "기타"
				Case Else
					fnWorkTargetCode3 = ""
			End Select
		End If
	End Function

	Function fnWorkState(code)
		Select Case code
			Case "1"
				fnWorkState = "신청"
			Case "3"
				fnWorkState = "작업중"
			Case "5"
				fnWorkState = "완료"
			Case Else
				fnWorkState = ""
		End Select
	End Function

	Function fnWorkStateNext(code)
		Select Case code
			Case "1"
				fnWorkStateNext = "3"
			Case "3"
				fnWorkStateNext = "5"
			Case Else
				fnWorkStateNext = ""
		End Select
	End Function

	Function fnWorkStateTRColor(code)
		Select Case code
			Case "1"
				fnWorkStateTRColor = "#FFFFFF"
			Case "3"
				fnWorkStateTRColor = "#FFDB57"
			Case "5"
				fnWorkStateTRColor = "silver"
			Case Else
				fnWorkStateTRColor = "#FFFFFF"
		End Select
	End Function


	Function NowWorkerName(userid)
		Dim strSql
		strSql = " SELECT username From [db_partner].[dbo].[tbl_user_tenbyten] WHERE userid = '" & userid & "'"
		rsget.Open strSql,dbget
		If Not rsget.Eof Then
			NowWorkerName = rsget(0)
		End If
		rsget.Close
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


	'####### 코드매니저 페이지 외 하나씩 불러쓰는 공통 코드 관리. write 용, view 용. #######
	Public Function CommonCode(ByVal sUse, ByVal sType, ByVal sCode)
		Dim strSql, sBody, i
		sBody = ""
		i = 0

		'### sUse = "w" write 용
		If sUse = "w" Then
			strSql = " SELECT code_value, code_desc From [db_temp].[dbo].[tbl_breakdown_comCode] WHERE code_type ='"&sType&"' AND code_useyn = 'Y' ORDER BY code_sort ASC"

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
			strSql = " SELECT code_desc From [db_temp].[dbo].[tbl_breakdown_comCode] WHERE code_type ='"&sType&"' AND code_value = '" & sCode & "' AND code_useyn = 'Y'"
			'response.write strSql
			rsget.Open strSql,dbget
			If Not rsget.Eof Then
				sBody = rsget(0)
			End If
			rsget.Close
		End If
		CommonCode = sBody
	End Function


	Function AgentGubun()
		Dim userAgent, userBrowser
		userAgent = Request.ServerVariables("HTTP_USER_AGENT")

		If inStr(userAgent, "MSIE") > 0 then
			userBrowser = "IE"
		else
			userBrowser = "ETC"
		End if

		AgentGubun = userBrowser
	End Function
%>
