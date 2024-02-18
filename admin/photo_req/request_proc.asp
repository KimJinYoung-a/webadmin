<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  촬영요청 프로세스
' History : 2011.03.13 김진영 생성
'			2015.07.28 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
Dim menupos, i, j, k
	menupos = requestcheckvar(request("menupos"),10)

Dim cPhotoreq, PhotoCnt
set cPhotoreq = new Photoreq
	PhotoCnt = cPhotoreq.fnGetPhotoUser
set cPhotoreq = nothing

Dim sMode, strSql, strsql1, strsql2, sReq_no, sMode2, sReq_gubun, sReq_use, sPrd_type, sPrd_type2, sReq_date
Dim sReq_use_detail, sPrd_name, sPrd_price, sImport_level, sReq_department, sReq_category, sMakerid, sItemid
dim req_cdl_disp, sReq_use_type, sReq_use_concept, sReq_etc1, sReq_url, sReq_etc2, sfontColor, comment, FC, sReq_defaultOpt
Dim sReq_gubunS, sReq_status, Req_photo, Req_stylist, sReq_comment, vTemp, sReq_day, sReq_day_start
Dim sReq_day_end, sReq_day_start_time, sReq_day_end_time, Req_start_date, Req_end_date, Fusername
Dim sDoc_File, vFileTemp, vRFileTemp, sDoc_RealFile, udate, suse_yn, sReq_MDid, Sreq_SMS, sReq_id, vopenidxcount, vopenidx, vopenurl
Dim isdefaultCnt : isdefaultCnt = 0
'sMode				: 모드 *
'sReq_gubun			: 촬영구분 *
'sReq_use			: 촬영용도 구분 *
'sReq_use_detail	: 기본 상세페이지 선택
'sPrd_name 			: 상품명
'sPrd_type 			: 상품군 *
'sPrd_type2 		: 상품종 *
'sPrd_price			: 판매가
'sImport_level		: 중요도
'sReq_department	: 요청부서
'sReq_category		: 카테고리
'sMakerid			: 브랜드ID
'sItemid			: 상품코드
'sReq_date 			: 희망 촬영일자 *
'sDoc_use_type		: 필요촬영군
'sDoc_use_concept	: 메인촬영컨셉
'sReq_etc1 			: 상품 특징 *
'sReq_url			: URL
'sReq_etc2			: 촬영시 유의사항

sMode	 			= Request("mode")
sMode2	 			= Request("mode2")
sReq_no				= Request("req_no")
sReq_gubun			= Request("req_gubun")
sReq_use			= Request("req_use")
sReq_use_detail		= Request("req_use_detail")
sPrd_name			= Request("prd_name")
sPrd_type			= Request("prd_type")
sPrd_type2			= getNumeric(Request("prd_type2"))
sPrd_price			= Request("prd_price")
sImport_level		= Request("import_level")
sReq_department		= Request("req_department")
req_cdl_disp		= requestcheckvar(Request("req_cdl_disp"),10)
sReq_category		= Request("req_category")
sMakerid			= Request("makerid")
sItemid				= Request("itemid")
sReq_date			= Request("req_date")
sReq_use_type		= Request("req_use_type")
sReq_use_concept	= Request("req_use_concept")
sReq_defaultOpt		= Request("defaultOpt")
sReq_etc1			= Request("req_etc1")
sReq_url			= Request("req_url")
sReq_etc2			= Request("req_etc2")
sReq_MDid			= Request("MDid")
sReq_status 		= Request("req_status")
Req_photo			= Request("req_photo")
Req_stylist		= Request("req_stylist")
sReq_comment		= Request("req_comment")
sReq_gubunS 		= Request("req_gubunS")
sfontColor			= Request("fontColor")
Sreq_SMS			= Request("req_SMS")
sReq_id				= Request("req_id")
suse_yn				= Request("use_yn")
vTemp				= Request("lineCnt")
sDoc_File			= NullFillWith(Request("doc_file"),"")
sDoc_RealFile		= NullFillWith(Request("doc_realfile"),"")
udate = Request("udate")

If vTemp = "0" Then vTemp = "1"

If sMode = "I" or sMode2 = "I" Then
	strSql = ""
	strSql = strSql & "Insert into [db_partner].[dbo].tbl_photo_req" & vbCrLf
	strSql = strSql & "(req_gubun, req_use, req_use_detail, prd_name, prd_type, prd_type2, prd_price, " & vbCrLf
	strSql = strSql & " import_level, req_department, req_category, req_cdl_disp, makerid, itemid, " & vbCrLf
	strSql = strSql & " req_date, req_etc1, req_url, req_etc2, req_regdate, use_yn, req_name, MDid)" & vbCrLf
	strSql = strSql & " Values " & vbCrLf
	strSql = strSql & "('"&sReq_gubun&"', '"&sReq_use&"', '"&sReq_use_detail&"', '"&html2db(sPrd_name)&"', '"&html2db(sPrd_type)&"', '"&html2db(sPrd_type2)&"', '"&sPrd_price&"', " & vbCrLf
	strSql = strSql & " '"&sImport_level&"', '"&sReq_department&"', '"&sReq_category&"', '"& req_cdl_disp &"', '"&html2db(sMakerid)&"', '"&sItemid&"', " & vbCrLf
	strSql = strSql & " '"&sReq_date&"', '"&html2db(sReq_etc1)&"', '"&sReq_url&"', '"&html2db(sReq_etc2)&"', getdate(), 'Y', '"&session("ssBctid")&"', '"&sReq_MDid&"')" & vbCrLf
	'response.write strSql & "<br>"
	dbget.execute strSql

	Dim currentNo
	strSql = ""
	strSql = strSql & "SELECT max(req_no) as max_no from [db_partner].[dbo].tbl_photo_req " & vbCrLf

	rsget.Open strSql, dbget
	IF not rsget.EOF THEN
		currentNo = rsget("max_no")
	End IF
	rsget.Close

	If sMode2 = "I" Then
		Dim max_no
		strSql = ""
		strSql = strSql & "SELECT max(req_no) as max_no from [db_partner].[dbo].tbl_photo_req " & vbCrLf

		rsget.Open strSql, dbget
		IF not rsget.EOF THEN
			max_no = rsget("max_no")
		End IF
		rsget.Close

		strSql = ""
		strSql = strSql & " Update [db_partner].[dbo].tbl_photo_req set " & vbCrLf
		strsql = strsql & " load_req = '"&sReq_no&"'" & vbCrLf
		strSql = strSql & " where req_no = '"& max_no &"' "
		dbget.execute strSql

		If sReq_gubunS = "4" Then
			strSql = ""
			strSql = strSql & " Update [db_partner].[dbo].tbl_photo_req set " & vbCrLf
			strsql = strsql & " req_status = '"&sReq_gubunS&"'" & vbCrLf
			strSql = strSql & " where req_no = '"& max_no &"' "
			dbget.execute strSql
		End If
	End If

	If Isempty(sReq_use_type) <> "True" Then
		sReq_use_type			= Split(sReq_use_type,",")
		For i = 0 to Ubound(sReq_use_type)
			strSql = ""
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_req_use_type " & vbcrlf
			strSql = strSql & " (req_no, req_use_type) " & vbcrlf
			strSql = strSql & "	VALUES " & vbcrlf
			strSql = strSql & "	('" & currentNo & "', '" & Trim(sReq_use_type(i)) & "')" & vbcrlf
			dbget.execute strSql
			If Trim(sReq_use_type(i)) = "11" Then
				isdefaultCnt = isdefaultCnt + 1
			End If
		Next
		
		If isdefaultCnt > 0 Then
			sReq_defaultOpt	= Split(sReq_defaultOpt,",")
			For k = 0 to Ubound(sReq_defaultOpt)
				strSql = ""
				strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_req_concept " & vbcrlf
				strSql = strSql & " (req_no, req_use_concept) " & vbcrlf
				strSql = strSql & "	VALUES " & vbcrlf
				strSql = strSql & "	('" & currentNo & "', '" & Trim(sReq_defaultOpt(k)) & "')" & vbcrlf
				dbget.execute strSql
			Next
		End If
	End If

	If Isempty(sReq_use_concept) <> "True" Then
		sReq_use_concept			= Split(sReq_use_concept,",")
		For j = 0 to Ubound(sReq_use_concept)
			strSql = ""
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_req_concept " & vbcrlf
			strSql = strSql & " (req_no, req_use_concept) " & vbcrlf
			strSql = strSql & "	VALUES " & vbcrlf
			strSql = strSql & "	('" & currentNo & "', '" & Trim(sReq_use_concept(j)) & "')" & vbcrlf
			dbget.execute strSql
		Next
	End If

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_file " & _
							  "		(req_no, file_name, real_name, file_regdate) " & _
							  "	VALUES " & _
							  "		('"&currentNo&"', '" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "',  getdate()) " & vbCrLf
		Next
			dbget.execute strSql
	End If

	Response.Write "<script>alert('저장되었습니다.');location.href='/admin/photo_req/request_list.asp?menupos="&menupos&"';</script>"
	dbget.close()
	Response.End

ElseIf sMode = "U" or sMode2 = "" Then
	If sReq_use <> "기본상세페이지" Then
		sReq_use_detail = ""
	End If

	If Request("userFont") = "R" Then
		FC = "G"
	End If

	strSql = ""
	strSql = strSql & " Update [db_partner].[dbo].tbl_photo_req" & vbCrLf
	strsql = strsql & " set req_gubun = '"&sReq_gubun&"', req_use = '"&sReq_use&"',req_use_detail = '"&sReq_use_detail&"',prd_name = '"&html2db(sPrd_name)&"'" & vbCrLf
	strsql = strsql & " ,prd_type = '"&html2db(sPrd_type)&"',prd_type2 = '"&html2db(sPrd_type2)&"',prd_price = '"&sPrd_price&"', import_level = '"&sImport_level&"'" & vbCrLf
	strsql = strsql & " , req_department = '"&sReq_department&"',req_category = '"&sReq_category&"', req_cdl_disp = '"&req_cdl_disp&"' ,makerid = '"&sMakerid&"'" & vbCrLf
	strsql = strsql & " ,itemid = '"&sItemid&"', req_etc1 = '"&html2db(sReq_etc1)&"',req_url = '"&sReq_url&"',req_etc2 = '"&html2db(sReq_etc2)&"', fontColor='"&FC&"'" & vbCrLf
	strsql = strsql & " , use_yn = '"&suse_yn&"', MDid = '"&sReq_MDid&"' where" & vbCrLf
	strSql = strSql & " req_no = '"& sReq_no &"' "
	
	'response.write strSql &"<br>"
	dbget.execute strSql

	strSql = ""
	strSql = strSql & " DELETE FROM [db_partner].[dbo].tbl_photo_req_use_type where"
	strSql = strSql & " req_no = '" & sReq_no & "'"
	
	'response.write strSql &"<br>"
	dbget.execute strSql

	If Isempty(sReq_use_type) <> "True" Then
		sReq_use_type			= Split(sReq_use_type,",")
		For i = 0 to Ubound(sReq_use_type)
			strSql = ""
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_req_use_type" & vbcrlf
			strSql = strSql & " (req_no, req_use_type) " & vbcrlf
			strSql = strSql & "	VALUES " & vbcrlf
			strSql = strSql & "	('" & sReq_no & "', '" & Trim(sReq_use_type(i)) & "')" & vbcrlf
			
			'response.write strSql &"<br>"
			dbget.execute strSql
			If Trim(sReq_use_type(i)) = "11" Then
				isdefaultCnt = isdefaultCnt + 1
			End If
		Next
	End If

	strSql = "DELETE FROM [db_partner].[dbo].tbl_photo_req_concept where"
	strSql = strSql & " req_no = '" & sReq_no & "'"

	'response.write strSql &"<br>"
	dbget.execute strSql

	If isdefaultCnt > 0 Then
		sReq_defaultOpt	= Split(sReq_defaultOpt,",")
		For k = 0 to Ubound(sReq_defaultOpt)
			strSql = ""
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_req_concept " & vbcrlf
			strSql = strSql & " (req_no, req_use_concept) " & vbcrlf
			strSql = strSql & "	VALUES " & vbcrlf
			strSql = strSql & "	('" & sReq_no & "', '" & Trim(sReq_defaultOpt(k)) & "')" & vbcrlf
			dbget.execute strSql
		Next
	End If

	If Isempty(sReq_use_concept) <> "True" Then
		sReq_use_concept			= Split(sReq_use_concept,",")
		For j = 0 to Ubound(sReq_use_concept)
			strSql = "INSERT INTO [db_partner].[dbo].tbl_photo_req_concept " & vbcrlf
			strSql = strSql & " (req_no, req_use_concept) " & vbcrlf
			strSql = strSql & "	VALUES " & vbcrlf
			strSql = strSql & "	('" & sReq_no & "', '" & Trim(sReq_use_concept(j)) & "')" & vbcrlf

			'response.write strSql &"<br>"
			dbget.execute strSql
		Next
	End If

	If PhotoCnt > 0 OR session("ssBctId") = "eoslove" OR session("ssBctId") = "sss162000" OR session("ssBctId") = "dhalsdud57" OR session("ssBctId") = "tozzinet" OR session("ssBctId") = "hrkang97" Then
		strSql = "Update [db_partner].[dbo].tbl_photo_req set " & vbCrLf
		strsql = strsql & " req_status = '"&sReq_status&"', req_comment = '"&sReq_comment&"', fontColor='"&sfontColor&"' " & vbCrLf
		strSql = strSql & " where req_no = '"& sReq_no &"' "

		'response.write strSql &"<br>"
		dbget.execute strSql

		strSql2 = "Update [db_partner].[dbo].tbl_photo_schedule set " & vbCrLf
		strSql2 = strSql2 & " status = '"&sReq_status&"'" & vbCrLf
		strSql2 = strSql2 & " where req_no = '"& sReq_no &"' "

		'response.write strSql &"<br>"
		dbget.execute strSql2

		If Sreq_SMS = "Y" Then
				strSql = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) " & vbcrlf
				strSql = strSql & " select usercell, '1644-6030' as tran_callback, '1' as tran_status, getdate() as tran_date, '[촬영요청서 No."&sReq_no&"]번 요청서 확인부탁드립니다.' as tran_msg from " & vbcrlf
				strSql = strSql & "	db_user.dbo.tbl_user_n " & vbcrlf
			If sReq_MDid <> "00" Then
				strSql = strSql & " where userid = '"&sReq_MDid&"'" & vbcrlf
			ElseIf sReq_MDid = "00" Then
				strSql = strSql & " where userid = '"&sReq_id&"'" & vbcrlf
			End If

			'response.write strSql &"<br>"
			dbget.execute strSql
		End If

		If udate <> "A" Then

			vTemp = Request.Form("yyyy").count

			strSql = "DELETE FROM [db_partner].[dbo].tbl_photo_schedule where" & vbcrlf
			strSql = strSql & " req_no = '" & sReq_no & "'"
			
			'response.write strSql & "<Br>"
			dbget.execute strSql
	
			For j = 1 To vTemp
				If trim(Request("yyyy")(j)) = "" or trim(Request("mm")(j)) = "" or trim(Request("dd")(j)) = "" or trim(Request("Req_day_start")(j)) = "" or trim(Request("Req_day_end")(j)) = "" Then
					response.write "<script>alert('스케쥴 날짜가 잘못 되었습니다.');history.back(-1);</script>"
					response.end
				End If
				If trim(Request("Req_day_start")(j)) = "" or trim(Request("Req_day_end")(j)) = "" Then
					response.write "<script>alert('스케쥴 시간이 잘못 되었습니다.');history.back(-1);</script>"
					response.end
				End If

				sReq_day		= Request("yyyy")(j) & "-" & Request("mm")(j) & "-" & Request("dd")(j)
				sReq_day_start	= Request("Req_day_start")(j)
				sReq_day_end	= Request("Req_day_end")(j)
				comment = Request("comment")(j)
				req_photo = Request("req_photo")(j)
				req_Stylist = Request("req_Stylist")(j)

				If sReq_day_start = "" or sReq_day_end = "" Then
					response.write "<script>alert('스케쥴 시간이 잘 못 되었습니다.');history.back(-1);</script>"
					response.end
				End If

				Select Case sReq_day_start
					Case "8"	sReq_day_start_time = "10:00"
					Case "9"	sReq_day_start_time = "10:30"
					Case "10"	sReq_day_start_time = "11:00"
					Case "11"	sReq_day_start_time = "11:30"
					Case "12"	sReq_day_start_time = "12:00"
					Case "13"	sReq_day_start_time = "12:30"
					Case "14"	sReq_day_start_time = "13:00"
					Case "15"	sReq_day_start_time = "13:30"
					Case "16"	sReq_day_start_time = "14:00"
					Case "17"	sReq_day_start_time = "14:30"
					Case "18"	sReq_day_start_time = "15:00"
					Case "19"	sReq_day_start_time = "15:30"
					Case "20"	sReq_day_start_time = "16:00"
					Case "21"	sReq_day_start_time = "16:30"
					Case "22"	sReq_day_start_time = "17:00"
					Case "23"	sReq_day_start_time = "17:30"
					Case "24"	sReq_day_start_time = "18:00"
					Case "25"	sReq_day_start_time = "18:30"
				End Select

				Select Case sReq_day_end
					Case "8"	sReq_day_end_time = "10:00"
					Case "9"	sReq_day_end_time = "10:30"
					Case "10"	sReq_day_end_time = "11:00"
					Case "11"	sReq_day_end_time = "11:30"
					Case "12"	sReq_day_end_time = "12:00"
					Case "13"	sReq_day_end_time = "12:30"
					Case "14"	sReq_day_end_time = "13:00"
					Case "15"	sReq_day_end_time = "13:30"
					Case "16"	sReq_day_end_time = "14:00"
					Case "17"	sReq_day_end_time = "14:30"
					Case "18"	sReq_day_end_time = "15:00"
					Case "19"	sReq_day_end_time = "15:30"
					Case "20"	sReq_day_end_time = "16:00"
					Case "21"	sReq_day_end_time = "16:30"
					Case "22"	sReq_day_end_time = "17:00"
					Case "23"	sReq_day_end_time = "17:30"
					Case "24"	sReq_day_end_time = "18:00"
					Case "25"	sReq_day_end_time = "18:30"
				End Select

				If sReq_day <> "" then
					Req_start_date	= sReq_day&" "&sReq_day_start_time
					Req_end_date 	= sReq_day&" "&sReq_day_end_time

					strsql1 = "select req_name from [db_partner].[dbo].[tbl_photo_req] a " & vbcrlf
					strsql1 = strsql1 &" left join [db_partner].[dbo].[tbl_photo_schedule] b " & vbcrlf
					strsql1 = strsql1 &" on a.req_no = b.req_no " & vbcrlf
					strsql1 = strsql1 &" left join [db_partner].[dbo].[tbl_photo_user] c " & vbcrlf
					strsql1 = strsql1 &" on b.req_photo = c.user_id " & vbcrlf
					strsql1 = strsql1 &" where b.start_date < '"&Req_end_date&"' " & vbcrlf
					strsql1 = strsql1 &" and b.end_date > '"&Req_start_date&"' " & vbcrlf
					strsql1 = strsql1 &" and c.user_id = '"&Req_photo&"' and c.user_type = '1' and a.use_yn = 'Y' " & vbcrlf

					'response.write strSql & "<Br>"
					rsget.Open strsql1,dbget,1
					If  not rsget.EOF  then
						Fusername   =  rsget("req_name")
					End If
					rsget.close

					If Fusername <> "" Then
						response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
						dbget.close()	:	response.End
					end if

					strsql2 = "insert into [db_partner].[dbo].tbl_photo_schedule "& vbCrLf
					strsql2 = strsql2 &" (req_no, start_date, end_date, schedule_regdate, status, comment, req_photo, req_stylist) " & vbCrLf
					strsql2 = strsql2 &" values (" & vbCrLf
					strsql2 = strsql2 &" '"&sReq_no&"', '"&Req_start_date&"', '"&Req_end_date&"', getdate(), '"&sReq_status&"', '"& html2db(comment) &"'" & vbCrLf
					strsql2 = strsql2 &" ,'"&Req_photo&"', '"&Req_stylist&"'" & vbCrLf
					strsql2 = strsql2 &" )" & vbCrLf

					'response.write strSql & "<Br>"
					dbget.execute strSql2
				End If
			Next
		End If
	End If

	' 최종 오픈 정보
	strsql2 = "delete from db_partner.dbo.tbl_photo_opendata where req_no = "& sReq_no &""& vbCrLf

	'response.write strsql2 & "<Br>"
	dbget.execute strSql2

	vopenidxcount = Request.Form("openidx").count

	For j = 1 To vopenidxcount
		vopenidx = Request("openidx")(j)
		vopenurl = Request("openurl")(j)

		strsql2 = "insert into db_partner.dbo.tbl_photo_opendata (req_no,openurl) values ("&sReq_no&",'"& html2db(trim(vopenurl)) &"')"& vbCrLf

		'response.write strsql2 & "<Br>"
		dbget.execute strSql2
	Next

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If sReq_no <> "" Then
			strSql = " DELETE [db_partner].[dbo].tbl_photo_file WHERE req_no = '" & sReq_no & "' "
		End If
		vFileTemp 	= Split(sDoc_File, ",")
		vRFileTemp	= Split(sDoc_RealFile, ",")
		'response.write UBOUND(vFileTemp)
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO [db_partner].[dbo].tbl_photo_file " & _
							  "		(file_name, real_name, req_no, file_regdate) " & _
							  "	VALUES " & _
							  "		('" & Trim(vFileTemp(i)) & "', '" & Trim(vRFileTemp(i)) & "', '" & sReq_no & "', getdate()) " & vbCrLf
		Next
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE [db_partner].[dbo].tbl_photo_file WHERE req_no = '" & sReq_no & "' "
		End If
	End If

	Response.Write "<script>alert('수정 되었습니다.');location.href='/admin/photo_req/request_list.asp?menupos="&menupos&"';</script>"
	dbget.close()
	Response.End

End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->