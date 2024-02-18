<%
'####################################################
' Description : 촬영 요청 클래스
' History : 2012.03.13 김진영 생성
'####################################################

Class Photoreq_Item
	public fopencount
	public fcomment
	Public FReq_no
	Public FReq_status
	Public FReq_use
	Public FReq_use_detail
	Public FReq_prd_name
	Public FReq_codenm
	Public FReq_makerid
	Public FReq_regdate
	Public FReq_name
	Public FReq_id
	Public FReq_Photo
	Public FReq_Photoname
	Public FLoad_req
	Public FFontColor
	Public FStylistname
	Public FReq_gubun
	Public FPrd_name
	Public FPrd_type
	Public FPrd_type2
	Public FPrd_price
	Public FImport_level
	Public FReq_department
	Public FReq_category
	public freq_cdl_disp
	Public FMakerid
	Public Fitemid
	Public FReq_date
	Public FReq_etc1
	Public FReq_url
	Public FReq_etc2
	Public FReq_stylist
	Public FReq_comment
	Public FMDid
	Public FMDname
	Public FStart_date
	Public FEnd_date
	Public FUse_yn
	public fconfirmdate
	public fopenidx
	public fopenurl

End Class

Class Photoreq
	Public FPhotoreqList()
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount

	Public FMakerid
	Public FCdl
	Public FReq_use
	Public FS_type
	Public FNum_Name
	Public FReq_status_type
	Public FRequest_name
	Public FReq_photo_user
	Public FReq_stylist

	Public FReq_gubun
	Public FReq_use_detail
	Public FPrd_name
	Public FPrd_type
	Public FPrd_type2
	Public FPrd_price
	Public FImport_level
	Public FReq_department
	Public FReq_category
	Public FItemid
	Public FReq_date
	Public FReq_etc1
	Public FReq_url
	Public FReq_etc2

	Public FReq_status
	Public FReq_photo
	Public FReq_comment

	Public Freq_no
	Public Sub fnReqno
		Dim strSql
		strSql = "select max(req_no) as req_no from [db_partner].[dbo].tbl_photo_req"
		rsget.open strSql, dbget, 1

		If not rsget.EOF Then
			Freq_no = rsget("req_no")
		End If

		If IsNull(rsget("req_no")) Then
			Freq_no = "0"
		End If

		rsget.close
	End Sub

	'//admin/photo_req/request_list.asp
	Public Function fnPhotoreqlist
		Dim strSql, i, where, wSQL

		If FMakerid <> "" Then
			where = where & " and R.makerid= '"&FMakerid&"' "
		End If

		If FCdl <> "" Then
			where = where & " and R.req_cdl_disp= '"&FCdl&"' "
		End If

		If FReq_use <> "" Then
			where = where & " and R.req_use= '"&FReq_use&"' "
		End If

		If FNum_Name <> "" Then
			If FS_type = "1" Then
				where = where & " and R.req_no = '"&FNum_Name&"' "
			ElseIf FS_type = "2" Then
				where = where & " and R.prd_name like '%"&FNum_Name&"%' "
			End If
		End If

		If FReq_status_type <> "" Then
			where = where & " and R.req_status= '"&FReq_status_type&"' "
		End If

		If FRequest_name <> "" Then
			where = where & " and T.username= '"&FRequest_name&"' "
		End If

		If FReq_photo_user <> "" Then
			Dim photoID
			wSQL = ""
			wSQL = wSQL & " SELECT TOP 1 userid FROM db_partner.dbo.tbl_user_tenbyten WHERE username = '"&FReq_photo_user&"' and part_sn='23' and isnull(userid, '') <> '' "
			
			'response.write wSQL & "<br>"
			rsget.Open wSQL,dbget,1
			If not rsget.EOF Then
				photoID = rsget("userid")
			End If
			rsget.Close
		End If

		If FReq_stylist <> "" Then
			Dim styleID
			wSQL = ""
			wSQL = wSQL & " SELECT TOP 1 userid FROM db_partner.dbo.tbl_user_tenbyten WHERE username = '"&FReq_stylist&"' and part_sn='23' and isnull(userid, '') <> '' "
			
			'response.write wSQL & "<br>"
			rsget.Open wSQL,dbget,1
			If not rsget.EOF Then
				styleID = rsget("userid")
			End If
			rsget.Close
		End If

		strSql = "select count(R.req_no) as cnt " + vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_req as R" & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T" & vbcrlf 
		strSql = strSql & "  	on R.req_name = T.userid " & vbcrlf

		If FReq_photo_user <> "" Then
			strSql = strSql & " Join [db_partner].dbo.tbl_photo_schedule as S1" & vbcrlf
			strSql = strSql & " 	on R.req_no = S1.req_no " & vbcrlf

			If photoID <> "" Then
				strSql = strSql & " 	and s1.req_photo = '"&photoID&"' "
			Else
				strSql = strSql & " 	and s1.req_photo = '"&FReq_photo_user&"' "
			End If
		End If
		If FReq_stylist <> "" Then
			strSql = strSql & " Join [db_partner].dbo.tbl_photo_schedule as S2" & vbcrlf
			strSql = strSql & " 	on R.req_no = S2.req_no " & vbcrlf

			If styleID <> "" Then
				strSql = strSql & " 	and s2.req_stylist = '"&styleID&"' "
			Else
				strSql = strSql & " 	and s2.req_stylist = '"&FReq_stylist&"' "
			End If
		End If

		strSql = strSql & " where R.use_yn = 'Y' " & where

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close

		strSql = "select top "& Cstr(FPageSize * FCurrPage)
		strSql = strSql & " R.req_no, R.req_status, R.req_use, R.req_use_detail, R.prd_name"
		strSql = strSql & " ,(case when r.req_cdl_disp='999999999' then '선택안함'"
		strSql = strSql & " 	when r.req_cdl_disp='999999998' then 'PLAY' else l.catename end) as catename"
		strSql = strSql & " , R.makerid, R.req_regdate, T.username as req_name"
		'strSql = strSql & " ,s.req_photo,(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TT where s.req_photo = TT.userid) as req_photoname" & vbcrlf
		strSql = strSql & " , R.fontColor, r.Import_level" & vbcrlf
		strSql = strSql & " ,(case when isnull(R.MDid,'')<>'' then"
		strSql = strSql & " 		(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TTT where R.MDid = TTT.userid)"
		strSql = strSql & " 	else '미지정' end) as MDid"
		'strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TTTT where s.req_stylist = TTTT.userid) as req_stylist " & vbcrlf
		strSql = strSql & " ,substring(STUFF((" & vbcrlf
		strSql = strSql & " 	SELECT Top 30 '|^|' + cast(convert(varchar(20),ss.start_date,121) as varchar(20)) + '|*|' + cast(convert(varchar(20),ss.end_date,121) as varchar(20))" & vbcrlf
		strSql = strSql & " 	+ '|*|' + isnull(ut1.username,'') + '|*|' + isnull(ut2.username,'')" & vbcrlf
		strSql = strSql & " 	FROM db_partner.dbo.tbl_photo_schedule ss" & vbcrlf
		strSql = strSql & " 	left join db_partner.dbo.tbl_user_tenbyten as ut1" & vbcrlf
		strSql = strSql & " 		on ss.req_photo = ut1.userid" & vbcrlf
		strSql = strSql & " 	left join db_partner.dbo.tbl_user_tenbyten as ut2" & vbcrlf
		strSql = strSql & " 		on ss.req_stylist = ut2.userid" & vbcrlf
		strSql = strSql & " 	WHERE R.req_no = ss.req_no" & vbcrlf
		strSql = strSql & " 	order by ss.schedule_no asc" & vbcrlf
		strSql = strSql & " 	FOR XML PATH('')), 1, 1, ''),3,800) as confirmdate" & vbcrlf
		strSql = strSql & " , (select count(openidx) from db_partner.dbo.tbl_photo_opendata where r.req_no = req_no) as opencount" & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_req as R " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on R.req_name = T.userid " & vbcrlf
		strSql = strSql & " Left Join db_item.dbo.tbl_display_cate as L" & vbcrlf 
		strSql = strSql & " 	on R.req_cdl_disp = L.catecode " & vbcrlf
		strSql = strSql & " 	and L.useyn='Y' and L.depth=1 " & vbcrlf

		If FReq_photo_user <> "" Then
			strSql = strSql & " Join [db_partner].dbo.tbl_photo_schedule as S1" & vbcrlf
			strSql = strSql & " 	on R.req_no = S1.req_no " & vbcrlf

			If photoID <> "" Then
				strSql = strSql & " 	and s1.req_photo = '"&photoID&"' "
			Else
				strSql = strSql & " 	and s1.req_photo = '"&FReq_photo_user&"' "
			End If
		End If
		If FReq_stylist <> "" Then
			strSql = strSql & " Join [db_partner].dbo.tbl_photo_schedule as S2" & vbcrlf
			strSql = strSql & " 	on R.req_no = S2.req_no " & vbcrlf

			If styleID <> "" Then
				strSql = strSql & " 	and s2.req_stylist = '"&styleID&"' "
			Else
				strSql = strSql & " 	and s2.req_stylist = '"&FReq_stylist&"' "
			End If
		End If

		strSql = strSql & " where 1=1 "& where &" and R.use_yn = 'Y' " & vbcrlf
		strSql = strSql & " order by R.req_no desc"

		'response.write strSql & "<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FPhotoreqList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FPhotoreqList(i) = new Photoreq_Item

				FPhotoreqList(i).FReq_no 			= rsget("req_no")
				FPhotoreqList(i).FReq_status		= rsget("req_status")
				FPhotoreqList(i).FReq_use			= rsget("req_use")
				FPhotoreqList(i).FReq_use_detail	= rsget("req_use_detail")
				FPhotoreqList(i).FReq_prd_name		= rsget("prd_name")
				FPhotoreqList(i).FReq_codenm 		= rsget("catename")
				FPhotoreqList(i).FReq_makerid 		= rsget("makerid")
				FPhotoreqList(i).FReq_regdate 		= rsget("req_regdate")
				FPhotoreqList(i).FReq_name 			= rsget("req_name")
				'FPhotoreqList(i).FReq_Photo 		= rsget("req_photo")
				'FPhotoreqList(i).FReq_Photoname		= rsget("req_photoname")
				FPhotoreqList(i).FFontColor			= rsget("fontColor")
				FPhotoreqList(i).FMDid				= rsget("MDid")
				'FPhotoreqList(i).FReq_stylist		= rsget("req_stylist")
				FPhotoreqList(i).fconfirmdate 			= rsget("confirmdate")
				FPhotoreqList(i).FImport_level 			= rsget("Import_level")
				FPhotoreqList(i).fopencount 			= rsget("opencount")

				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
		'게시글 리스트 구하기 끝'
	End Function

	'/사용금지(2017.01.12 한용민)
	Public Function fnGetSchedulelist
		Dim strSql
		strSql = "select req_no, start_date, end_date from db_partner.dbo.tbl_photo_schedule" + vbcrlf
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			fnGetSchedulelist = rsget.getRows()
		END IF
		rsget.close
	End Function

	Public Function fnGetFileList
		Dim strSql
		strSql = "	SELECT file_no, file_name, file_regdate, real_name " & _
				"		FROM [db_partner].[dbo].tbl_photo_file " & _
				"	WHERE req_no = '" & FReq_no & "' " & _
				"	ORDER BY file_no ASC "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnGetFileList = rsget.getRows()
		END IF
		rsget.close
	End Function

	'//admin/photo_req/request_modi.asp
	Public Function fnPhotoreqUpdate
		Dim strSql, i

		strSql = "select count(*) as cnt " + vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_req as R " & vbcrlf
		strSql = strSql & " Left Join [db_partner].dbo.tbl_photo_schedule as S on R.req_no = S.req_no " & vbcrlf
		strSql = strSql & " where R.req_no = '"& FReq_no &"'" & vbcrlf
		
		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		If not rsget.EOF Then
			FTotalCount = rsget("cnt")
		Else
			FTotalCount = 0
		End If
		rsget.Close

		If FTotalCount = 0 Then
			Call Alert_move("해당 정보가 없습니다","request_list.asp")
		End If

		strSql = "select R.req_no, R.req_gubun, R.req_use, R.req_use_detail, R.prd_name, R.prd_type, R.prd_type2, R.prd_price, R.import_level"
		strSql = strSql & " , R.req_department, isnull(R.req_category,'') as req_category, isnull(r.req_cdl_disp,'') as req_cdl_disp, R.makerid"
		strSql = strSql & " , R.req_date, R.itemid, R.req_regdate, R.req_etc1, R.req_url, R.req_etc2, R.req_status, s.req_photo, s.req_stylist , R.req_comment, R.MDid"
		strSql = strSql & " , S.start_date, S.end_date, s.comment, L.code_nm " & vbcrlf
		strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TT where s.req_photo = TT.userid) as req_photoname " & vbcrlf
		strSql = strSql & " , T.username as req_name, R.req_name as req_id" & vbcrlf
		strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TTT where s.req_stylist = TTT.userid) as stylistname " & vbcrlf
		strSql = strSql & " , R.load_req, R.fontColor, R.use_yn " & vbcrlf
		strSql = strSql & " ,(select top 1 username from db_partner.dbo.tbl_user_tenbyten as TTTT where R.MDid = TTTT.userid) as MDname " & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_req as R " & vbcrlf
		strSql = strSql & " Left Join db_item.dbo.tbl_Cate_large as L on R.req_category = L.code_large  " & vbcrlf
		strSql = strSql & " Inner Join db_partner.dbo.tbl_user_tenbyten as T on R.req_name = T.userid " & vbcrlf
		strSql = strSql & " Left Join [db_partner].dbo.tbl_photo_schedule as S on R.req_no = S.req_no " & vbcrlf
		strSql = strSql & " where R.req_no = '"& FReq_no &"'" & vbcrlf

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		If (FCurrPage * FPageSize < FTotalCount) Then
			FResultCount = FPageSize
		Else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		End If

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FPhotoreqList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FPhotoreqList(i) = new Photoreq_Item

				FPhotoreqList(i).FReq_gubun				= rsget("req_gubun")
				FPhotoreqList(i).FReq_use				= rsget("req_use")
				FPhotoreqList(i).FReq_use_detail		= rsget("req_use_detail")
				FPhotoreqList(i).FPrd_name				= rsget("prd_name")
				FPhotoreqList(i).FPrd_type				= rsget("prd_type")
				FPhotoreqList(i).FPrd_type2				= rsget("prd_type2")
				FPhotoreqList(i).FPrd_price				= rsget("prd_price")
				FPhotoreqList(i).FImport_level 			= rsget("import_level")
				FPhotoreqList(i).FReq_department 		= rsget("req_department")
				FPhotoreqList(i).FReq_category 			= rsget("req_category")
				FPhotoreqList(i).freq_cdl_disp 			= rsget("req_cdl_disp")
				FPhotoreqList(i).FMakerid 				= rsget("makerid")
				FPhotoreqList(i).FReq_date 				= rsget("req_date")
				FPhotoreqList(i).FItemid 				= rsget("itemid")
				FPhotoreqList(i).FReq_regdate 			= rsget("req_regdate")
				FPhotoreqList(i).FReq_etc1				= rsget("req_etc1")
				FPhotoreqList(i).FReq_url				= rsget("req_url")
				FPhotoreqList(i).FReq_etc2				= rsget("req_etc2")
				FPhotoreqList(i).FReq_status			= rsget("req_status")
				FPhotoreqList(i).FReq_photo				= rsget("req_photo")
				FPhotoreqList(i).FReq_stylist			= rsget("req_stylist")
				FPhotoreqList(i).FReq_comment			= rsget("req_comment")
				FPhotoreqList(i).FStart_date			= rsget("start_date")
				FPhotoreqList(i).FEnd_date				= rsget("end_date")

				FPhotoreqList(i).FReq_codenm			= rsget("code_nm")
				FPhotoreqList(i).FReq_Photo 			= rsget("req_photo")
				FPhotoreqList(i).FReq_Photoname			= rsget("req_photoname")
				FPhotoreqList(i).FReq_name				= rsget("req_name")
				FPhotoreqList(i).FReq_id				= rsget("req_id")
				FPhotoreqList(i).FStylistname			= rsget("stylistname")
				FPhotoreqList(i).FLoad_req				= rsget("load_req")
				FPhotoreqList(i).FFontColor				= rsget("fontColor")
				FPhotoreqList(i).FUse_yn				= rsget("use_yn")
				FPhotoreqList(i).FMDid					= rsget("MDid")
				FPhotoreqList(i).FMDname				= rsget("MDname")
				FPhotoreqList(i).fcomment				= db2html(rsget("comment"))

				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function

	'//admin/photo_req/request_modi.asp
	Public Function fnphoto_opendata
		Dim strSql, i

		if FReq_no="" then exit Function

		strSql = "select o.req_no, o.openidx, o.openurl" & vbcrlf
		strSql = strSql & " from db_partner.dbo.tbl_photo_opendata as o" & vbcrlf
		strSql = strSql & " where o.req_no = '"& FReq_no &"'" & vbcrlf

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1
		FResultCount = rsget.recordcount
		ftotalcount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		If (FTotalPage<>FTotalCount/FPageSize) Then FTotalPage = FTotalPage +1
		Redim preserve FPhotoreqList(FResultCount)
		FPageCount = FCurrPage - 1

		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FPhotoreqList(i) = new Photoreq_Item

				FPhotoreqList(i).FReq_no				= rsget("req_no")
				FPhotoreqList(i).fopenidx				= rsget("openidx")
				FPhotoreqList(i).fopenurl				= db2html(rsget("openurl"))

				rsget.movenext
				i = i + 1
			Loop
		End if
		rsget.Close
	End Function

	Public Function getdefaultOpt(irno)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT * FROM [db_partner].[dbo].tbl_photo_req_concept " & vbcrlf
		strSql = strSql & "	WHERE req_no = '"&irno&"' "
		rsget.Open strSql,dbget,1
		IF not rsget.EOF THEN
			getdefaultOpt = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetPhotoUser
		Dim strSql
		strSql = "select count(*) as cnt from [db_partner].[dbo].tbl_photo_user "&_
				" where user_type='1' and user_useyn = 'Y' and user_id = '"&session("ssBctID")&"' "
		rsget.Open strSql, dbget

		IF not rsget.EOF THEN
			fnGetPhotoUser = rsget("cnt")
		End IF

		rsget.Close
	End Function

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

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

Function DrawPicGubun(selectBoxName,selectedId, ggg)
	Dim tmp_str, strSql, j

	If ggg = "1" Then
		strSql = "select code_no, code_name from [db_partner].[dbo].tbl_photo_code  where code_useyn='Y' and code_type = '"& selectedId &"' order by code_sort "
	ElseIf ggg= "2" Then
		strSql = "select code_no, code_name from [db_partner].[dbo].tbl_photo_code  where code_useyn='Y' and code_type = 'doc_status' "
	End If

	rsget.Open strSql,dbget,1
%>
	<select name="<%=selectBoxName%>" <% If selectedId = "doc_status" and ggg = "1" Then %>  onchange="jsChkSubj(this.selectedIndex);" <% End If %> class="select">
<%
	If selectedId = "doc_status" or ggg="2" Then
		response.write("<option value='' selected>-- 촬영용도선택 --</option>")
	Else
		response.write("<option value='' selected>-- 기본 상세페이지 선택 --</option>")
	End If

	If not rsget.EOF Then
		j = 1
		Do Until rsget.EOF
			If rsget("code_name") = selectedId Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("code_name")& "' "&tmp_str&">" & rsget("code_name") & "" & "</option>")
			tmp_str = ""
			rsget.MoveNext
		j = j + 1
		Loop
	End If
	rsget.close
	response.write("</select>")
End Function

Function DrawPicGubun2(selectBoxName,selectedId, ggg)
	Dim tmp_str, strSql, j
	strSql = "select code_no, code_name from [db_partner].[dbo].tbl_photo_code  where code_useyn='Y' and code_type = '"& selectedId &"' order by code_sort "
	rsget.Open strSql,dbget,1
%>
	<select name="<%=selectBoxName%>" <% If selectedId = "doc_status" Then %> onchange="jsChkSubj(this.selectedIndex);" <% End If %> class="select">
<%
	If selectedId = "doc_status" Then
		response.write("<option value='' selected>-- 촬영용도선택 --</option>")
	Else
		response.write("<option value='' selected>-- 기본 상세페이지 선택 --</option>")
	End If

	If not rsget.EOF Then
		j = 1
		Do Until rsget.EOF
			If rsget("code_name") = ggg Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("code_name")& "' "&tmp_str&">" & rsget("code_name") & "" & "</option>")
			tmp_str = ""
			rsget.MoveNext
		j = j + 1
		Loop
	End If
	rsget.close
	response.write("</select>")
End Function

Sub DrawCategoryLarge(byval selectBoxName,selectedId)
   Dim tmp_str,query1
%>
	<select class='select' name="<%=selectBoxName%>">
		<option value="" <% if selectedId="" then response.write " selected"%>>선택</option>
<%
   query1 = " select code_large, code_nm from [db_item].[dbo].tbl_Cate_large "
   query1 = query1 + " where display_yn = 'Y'"
   query1 = query1 + " order by code_large Asc"
   rsget.Open query1,dbget,1

	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			If Cstr(selectedId) = Cstr(rsget("code_large")) Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("code_large")&"' "&tmp_str&">"& db2html(rsget("code_nm")) &"</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End Sub

'/ 사용중지
function DrawCategoryLarge_disp(byval selectBoxName, byval selectedId)
   Dim tmp_str,query1
%>
	<select class='select' name="<%=selectBoxName%>">
		<option value="" <% if selectedId="" then response.write " selected"%>>선택</option>
		<option value="999999999" <% if selectedId="999999999" then response.write " selected"%>>선택안함</option>
		<option value="999999998" <% if selectedId="999999998" then response.write " selected"%>>PLAY</option>
<%
   query1 = " select catecode, catename"
   query1 = query1 & " from db_item.dbo.tbl_display_cate"
   query1 = query1 & " where useyn='Y' and depth=1"
   query1 = query1 & " order by sortno asc"
   
   response.write query1 & "<br>"
   rsget.Open query1,dbget,1

	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			if selectedId<>"" and rsget("catecode")<>"" then
				If Cstr(selectedId) = Cstr(rsget("catecode")) Then
					tmp_str = " selected"
				End If
			End If
			response.write("<option value='"&rsget("catecode")&"' "&tmp_str&">"& db2html(rsget("catename")) &"</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End function

'// 사용중지
Sub DrawMDList(byval selectBoxName,selectedId)
   Dim tmp_str,query1
%>
	<select class='select' name="<%=selectBoxName%>">
		<option value="00" <% if selectedId="" or selectedId="00" then response.write " selected"%>>--담당 MD--</option>
<%
	query1 = " select userid, username from db_partner.dbo.tbl_user_tenbyten "
	query1 = query1 + " where part_sn in('11','21')" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	query1 = query1 & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	query1 = query1 & " and posit_sn < '9' and posit_sn <> '2' and isnull(userid,'')<>''" & vbcrlf
	query1 = query1 + " order by posit_sn asc, regdate desc "
   
   'response.write query1 & "<Br>"
   rsget.Open query1,dbget,1

	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			If cstr(selectedId) = cstr(rsget("userid")) Then
				tmp_str = " selected"
			End If
			response.write("<option value='"&rsget("userid")&"' "&tmp_str&">"& db2html(rsget("username")) &"</option>")
			tmp_str = ""
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End Sub

Sub CheckBoxUseType(byval chkName, sn, gugu)
	Dim query1, ckname, AA, BB, tmp_str, i

	If chkName = "doc_use_type" Then
		ckname = "req_use_type"
	Else
		ckname = "req_use_concept"
	End If

	If gugu = "" Then
		query1 = " select code_no, code_name from [db_partner].[dbo].tbl_photo_code"
		query1 = query1 + " where code_useyn='Y' and code_type = '"&chkName&"'"
		query1 = query1 &  " order by code_sort asc "
		rsget.Open query1,dbget,1

		If  not rsget.EOF  Then
		   rsget.Movefirst

		   Do until rsget.EOF
				If ckname = "req_use_type" AND rsget("code_no") = "11" Then
					response.write("<input type='checkbox' name='"&ckname&"' value='"&rsget("code_no")& "' id='dopt' onclick=jsDefaultOpt();>" & rsget("code_name") & "" & "")
				Else
					response.write("<input type='checkbox' name='"&ckname&"' value='"&rsget("code_no")& "'>" & rsget("code_name") & "" & "")
				End If

				If rsget("code_name") = "스토리" or rsget("code_name") = "오리엔탈" or rsget("code_name") = "글로시" Then
		       		response.write "<br>"	
				End If
		       rsget.MoveNext
		   Loop
		End if
	Else
		If gugu = "1" or gugu = "3" Then
			AA = "type"
			BB = "use_type"
		ElseIf gugu = "2" or gugu = "4" Then
			AA = "concept"
			BB = "concept"
		End If

		query1 = query1 &  " select C.req_use_"&AA&", T.code_no, T.code_name, T.code_sort from " & vbcrlf
		query1 = query1 &  " db_partner.dbo.tbl_photo_req_"&BB&" as C " & vbcrlf
		query1 = query1 &  " Right Join [db_partner].[dbo].tbl_photo_code as T " & vbcrlf
		query1 = query1 &  " on T.code_no = C.req_use_"&AA&" and C.req_no = '"&sn&"'  " & vbcrlf
		query1 = query1 &  " where T.code_type = '"&chkName&"'"
		query1 = query1 &  " and T.code_useyn='Y' "
		query1 = query1 &  " group by C.req_use_"&AA&", T.code_no, T.code_name, T.code_sort "
		query1 = query1 &  " order by T.code_sort asc "
		
		rsget.Open query1,dbget,1

		Dim jjj, qqq, pr_cdname
		Dim rr, tt
		If gugu = "3" or gugu = "4" Then
			qqq = 0
			If not rsget.EOF  Then
				For jjj = 0 to rsget.RecordCount - 1
					If db2html((rsget("req_use_"&AA&""))) = db2html((rsget("code_no"))) Then
						pr_cdname = pr_cdname & rsget("code_name")&", "
					End If
					rsget.MoveNext
				Next
				rr = split(pr_cdname,",")
				For tt=0 to Ubound(rr)-1
					response.write rr(tt)
					If tt < Ubound(rr)-1 Then
						response.write ","
					End If
				Next
			End if
		Else
			If not rsget.EOF  Then
				Do until rsget.EOF
					If db2html((rsget("req_use_"&AA&""))) = db2html((rsget("code_no"))) Then
						tmp_str = " checked "
					Else
						tmp_str = ""
					End If

					If ckname = "req_use_type" AND rsget("code_no") = "11" Then
						response.write("<input type='checkbox' name='"&ckname&"' value='"&rsget("code_no")& "' "& tmp_str &" id='dopt' onclick=jsDefaultOpt();>" & rsget("code_name") & "" & "")
					Else
						response.write("<input type='checkbox' name='"&ckname&"' value='"&rsget("code_no")& "' "& tmp_str &" >" & rsget("code_name") & "" & "")
					End If

					If rsget("code_name") = "스토리" or rsget("code_name") = "오리엔탈" or rsget("code_name") = "글로시" Then
			       		response.write "<br>"	
					End If
					rsget.MoveNext
				Loop
			End if
		End if
	End If
	rsget.close
End Sub

Sub SelectUser(AA, BB, CC)
	Dim query1
	query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user"
	query1 = query1 + " where user_type='"&AA&"' and user_useyn = 'Y'"
	rsget.Open query1,dbget,1
%>
	<select class="select" name='<%=BB%>'>
		<%= chkIIF(AA = "1","<option value='0'>-- 포토그래퍼 선택 --</option>","<option value=0>-- 스타일리스트 선택 --</option>") %>
<%
	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			response.write("<option value='"&rsget("user_id")& "' "& chkIIF(CC = rsget("user_id"),"selected","") &">" & rsget("user_name") & "" & "</option>")
			rsget.MoveNext
		Loop
	End If
	rsget.close
	response.write("</select>")
End Sub


Sub SelectUser2(AA, BB)
	Dim query1
	query1 = " select user_no, user_id, user_name from [db_partner].[dbo].tbl_photo_user"
	query1 = query1 + " where user_type='"&AA&"' and user_useyn = 'Y' and user_id = '"&BB&"'"
	rsget.Open query1,dbget,1

	If not rsget.EOF Then
		rsget.Movefirst
		Do until rsget.EOF
			response.write rsget("user_name")
			rsget.MoveNext
		Loop
	End If
	rsget.close
End Sub
%>