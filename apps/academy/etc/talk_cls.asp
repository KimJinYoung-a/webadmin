<%
Class cCornerItem
	Public FLecturer_id
	Public FLecturer_name
	Public FHistory
	Public FHistory_act
	Public FCatecd2
	Public FSocname
	Public FSocname_kor
	Public FImage_profile
	Public FImage_top
	Public FImage_list
	Public FRegdate
	Public FIsusing
	Public FHomepage
	Public FEvalsum
	Public FNowlec
	Public FTwitter
	Public FAvgevalPoint
	Public FIsBestCnt
	Public FFavorOnCnt
	Public FOnesentence
	Public FEvtCode
	Public FContentsCode
	Public FEvt_startdate
	Public FEvt_enddate
	Public FEvt_name
	Public FFirstRegdate

	Public FIdx
	Public FGubun
	Public FParamid
	Public FReply_group_idx
	Public FReply_depth
	Public FReply_num
	Public FUserid
	Public FComment
	Public FDevice
	Public FItemID
	Public FTotalPoint
	Public Flinkimg1
	Public Flinkimg2
	Public FItemname
	Public FBrandname
	Public FOrderserial
	Public FItemoption

	Public FLecIdx
	Public FLec_title
	Public FLec_cost
	Public FLec_count
	Public FLec_period
	Public FLec_startday1
	Public FLimit_count
	Public FLimit_sold
	Public FMainimg
	Public FStoryimg
	Public FSmallimg
	Public FIcon1
	Public FOblongImg2
	Public FOblongImg3
	Public Fmorollingimg1
	Public FMat_cost
	Public Fmatinclude_yn
	Public FReg_yn
	Public FReg_startday
	Public FReg_endday
	Public FLecturer_regdate
	Public FLecLimitCount
	Public FLecLimitSold
	Public FKeyword
	Public FCode_nm

	Public FPoint1
	Public FPoint2
	Public FPoint3
	Public FPoint4
	Public FTitle
	Public FContents
	Public FFile1only
	Public FFile2only
	Public FLecTitle
	Public FFile1
	Public FFile2
	
	'신규강사 여부
	Public Function isNewLecturer()
		If Datediff("m", FFirstRegdate, date()) < 1 Then
			isNewLecturer = True
		Else
			isNewLecturer = False
		End If
	End Function

	'베스트 강좌 여부
	Public Function isBestLecture()
		If FIsBestCnt > 0 Then
			isBestLecture = True
		Else
			isBestLecture = False
		End If
	End Function

	'후기 이미지
	Public Function ImageYN()
		If Ffile1only = "" and Ffile2only = "" Then
			ImageYN = ""
		ElseIf IsNull(Ffile1only) and IsNull(Ffile2only) Then
			ImageYN = ""
		Else
			ImageYN = "<img src='http://image.thefingers.co.kr/academy2012/common/photo.gif' alt='사진등록' class='icoPhoto' />"
		End If
	End Function

	public Function getLinkImage1()					
		getLinkImage1 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg1		
	end function 
	
	public Function getLinkImage2()
		getLinkImage2 = fingersImgUrl + "/contents/academy_diyevaluate/" + GetImageSubFolderByItemid(Fitemid) + "/" + Flinkimg2
	end function 

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class cCorner
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectLecturer_id
	Public FRectUserid
	Public FRectGubun
	Public FRectMyjob
	Public FRectParamid
	Public FRectIdx
	Public FRectSearchDIV
	Public FRectSearchTXT

	'탭별 카운트
	Public Function fnTabGubunCount(byval iGubun, byval iparamid, byref icheerTabCnt, byref iparamActCnt, byref ivalCnt)
		Dim sqlStr, addSqlcheerTab

		addSqlcheerTab = addSqlcheerTab & " and gubun = '"&iGubun&"' "
		addSqlcheerTab = addSqlcheerTab & " and paramid = '"&iparamid&"' "

		'응원톡 카운트
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cheerTabCnt FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and isusing = 'Y' "
		sqlStr = sqlStr & addSqlcheerTab
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			icheerTabCnt = rsACADEMYget("cheerTabCnt")
		rsACADEMYget.Close

		If iGubun = "L" Then			'강사 프로필
			'개설강좌 카운트
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(i.idx) as iparamActCnt "
			sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_item as i "
			sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
			sqlStr = sqlStr & " WHERE 1 = 1 "
			sqlStr = sqlStr & " and i.isusing = 'Y' "
			sqlStr = sqlStr & " and i.disp_yn = 'Y' "
			sqlStr = sqlStr & " and i.lecturer_id = '"&iparamid&"' "
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
				iparamActCnt = rsACADEMYget("iparamActCnt")
			rsACADEMYget.Close

			'수강후기 카운트
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(v.idx) as valCnt "
			sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_valuation v "
			sqlStr = sqlStr & " JOIN [db_academy].[dbo].tbl_lec_item i on v.masteridx = i.idx "
			sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
			sqlStr = sqlStr & " and v.isusing = 'Y' "
			sqlStr = sqlStr & " and i.lecturer_id = '"&iparamid&"' "
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
				ivalCnt = rsACADEMYget("valCnt")
			rsACADEMYget.Close
		Else
			'판매작품 카운트
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(i.itemid) as iparamActCnt "
			sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_diy_item] as i "
			sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as c on i.dispcate1 = c.catecode and c.useyn = 'Y' "
			sqlStr = sqlStr & " WHERE 1 = 1 "
			sqlStr = sqlStr & " and i.isusing = 'Y' "
			sqlStr = sqlStr & " and i.sellyn = 'Y' "
			sqlStr = sqlStr & " and ((i.limityn = 'N') or ((i.limityn = 'Y') and (i.limitno - i.limitsold > 0))) "
			sqlStr = sqlStr & " and i.makerid = '"&iparamid&"' "
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
				iparamActCnt = rsACADEMYget("iparamActCnt")
			rsACADEMYget.Close

			'구매후기 카운트
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(v.itemid) as valCnt "
			sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_diy_item_Evaluate] v "
			sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_diy_item] i on v.itemid = i.itemid "
			sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as c on i.dispcate1 = c.catecode and c.useyn = 'Y' "
			sqlStr = sqlStr & " and v.isusing = 'Y' "
			sqlStr = sqlStr & " and i.makerid = '"&iparamid&"' "
			rsACADEMYget.Open sqlStr, dbACADEMYget, 1
				ivalCnt = rsACADEMYget("valCnt")
			rsACADEMYget.Close
		End If
	End Function

	'강사 프로필
	Public Sub getlecturerItem()
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " EXEC [db_academy].[dbo].[academy_corner_Lecturer] '" & FRectLecturer_id & "', '"&FRectUserid&"', '"&FRectMyjob&"' "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.Locktype = adLockReadOnly
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FTotalCount = rsACADEMYget.RecordCount
		Set FOneItem = new cCornerItem
		If Not rsACADEMYget.Eof Then
			FOneItem.FLecturer_id	= rsACADEMYget("lecturer_id")
			FOneItem.FLecturer_name	= db2html(rsACADEMYget("lecturer_name"))
			FOneItem.FHistory		= db2html(rsACADEMYget("history"))
			FOneItem.FHistory_act	= db2html(rsACADEMYget("history_act"))
			FOneItem.FCatecd2		= rsACADEMYget("catecd2")
			FOneItem.FSocname		= db2html(rsACADEMYget("socname"))
			FOneItem.FSocname_kor	= db2html(rsACADEMYget("socname_kor"))
			If db2html(rsACADEMYget("newImage_profile")) <> "" Then
				FOneItem.FImage_profile	= fingersImgUrl & "/corner/newImage_profile/" & db2html(rsACADEMYget("newImage_profile"))
			Else
				FOneItem.FImage_profile	= ""
			End If
			FOneItem.FImage_top		= db2html(rsACADEMYget("image_top"))
			FOneItem.FImage_list	= db2html(rsACADEMYget("image_list"))
			FOneItem.FRegdate		= rsACADEMYget("regdate")
			FOneItem.FIsusing		= rsACADEMYget("isusing")
			FOneItem.FHomepage		= db2html(rsACADEMYget("homepage"))
			FOneItem.FTwitter		= rsACADEMYget("twitter")
			FOneItem.FAvgevalPoint	= rsACADEMYget("avgevalPoint")
			FOneItem.FIsBestCnt		= rsACADEMYget("isBestCnt")
			FOneItem.FFavorOnCnt	= rsACADEMYget("favorOnCnt")
			FOneItem.FOnesentence	= db2html(rsACADEMYget("onesentence"))
			FOneItem.FEvtCode		= rsACADEMYget("evt_code")
			FOneItem.FContentsCode	= rsACADEMYget("contentsCode")
			FOneItem.FEvt_startdate	= rsACADEMYget("evt_startdate")
			FOneItem.FEvt_enddate	= rsACADEMYget("evt_enddate")
			FOneItem.FEvt_name		= db2html(rsACADEMYget("evt_name"))
			FOneItem.FFirstRegdate	= rsACADEMYget("firstRegdate")
		End If
		rsACADEMYget.Close
	End Sub

	'작가프로필 하단 이미지List
	Public Function fnGetImageItem
		Dim sqlStr
		sqlStr = "SELECT TOP 15 " & vbcrlf
		sqlStr = sqlStr & " idx ,lecturer_id ,image_400x400 ,image_50x50 , image_80x80 ,regdate ,isusing" & vbcrlf
		sqlStr = sqlStr & " from db_academy.dbo.tbl_corner_good_item" & vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y' and lecturer_id = '"&FRectLecturer_id&"'" & vbcrlf
		sqlStr = sqlStr & " and isnull(image_80x80, '') <> ''  " & vbcrlf
		sqlStr = sqlStr & " order by idx desc" & vbcrlf
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If not rsACADEMYget.EOF Then
			fnGetImageItem = rsACADEMYget.getRows()
		End If
		rsACADEMYget.Close
    End Function

	'응원톡 List
	Public Sub getCheerTalkCommentList()
		Dim sqlStr, i, addSql

		If FRectGubun <> "" Then
			addSql = addSql & " and gubun = '"&FRectGubun&"' "
		End If

		If FRectParamid <> "" Then
			addSql = addSql & " and paramid = '"&FRectParamid&"' "
		End If

		If FRectSearchDIV <> "" Then
			If FRectSearchDIV = "1" Then '//1:작성자 , 2:작성내용
				addSql = addSql & " and userid like '%"&FRectSearchTXT&"%' "
			Else
				addSql = addSql & " and comment like '%"&FRectSearchTXT&"%' "
			End If 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and isusing = 'Y'  "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage) & " idx, gubun, paramid "
		sqlStr = sqlStr & "	,reply_group_idx, reply_depth, reply_num, userid, comment, device, isusing, regdate "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
		sqlStr = sqlStr & " WHERE 1 = 1  "
		sqlStr = sqlStr & " and isusing = 'Y'  "
		sqlStr = sqlStr & addSql
	    sqlStr = sqlStr & " ORDER BY reply_group_idx DESC, reply_num ASC  "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cCornerItem
					FItemList(i).FIdx				= rsACADEMYget("idx")
					FItemList(i).FGubun				= rsACADEMYget("gubun")
					FItemList(i).FParamid			= rsACADEMYget("paramid")
					FItemList(i).FReply_group_idx	= rsACADEMYget("reply_group_idx")
					FItemList(i).FReply_depth		= rsACADEMYget("reply_depth")
					FItemList(i).FReply_num			= rsACADEMYget("reply_num")
					FItemList(i).FUserid			= rsACADEMYget("userid")
					FItemList(i).FComment			= db2html(rsACADEMYget("comment"))
					FItemList(i).FDevice			= rsACADEMYget("device")
					FItemList(i).FIsusing			= rsACADEMYget("isusing")
					FItemList(i).FRegdate			= rsACADEMYget("regdate")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub

	'내가 쓴 응원톡
	Public Sub getMycommRead
		Dim sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, gubun, paramid, reply_group_idx, reply_depth, reply_num, userid, comment "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_academy_cheertalk_comments] "
		sqlStr = sqlStr & " WHERE isusing = 'Y' "
		sqlStr = sqlStr & " and idx = '"&FRectIdx&"' "
		sqlStr = sqlStr & " and userid = '"&FRectuserid&"'"
		sqlStr = sqlStr & " and paramid = '"&FRectParamid&"' "
		rsACADEMYget.Open SqlStr, dbACADEMYget, 1
		FTotalCount = rsACADEMYget.RecordCount
		Set FOneItem = new cCornerItem

		If Not rsACADEMYget.Eof Then
			FOneItem.Fidx				= rsACADEMYget("idx")
			FOneItem.FGubun				= rsACADEMYget("gubun")
			FOneItem.FParamid			= rsACADEMYget("paramid")
			FOneItem.FReply_group_idx	= rsACADEMYget("reply_group_idx")
			FOneItem.FReply_depth		= rsACADEMYget("reply_depth")
			FOneItem.FReply_num			= rsACADEMYget("reply_num")
			FOneItem.Fuserid			= rsACADEMYget("userid")
			FOneItem.FComment			= rsACADEMYget("comment")
		End If
		rsACADEMYget.Close
	End Sub

	'구매 후기 리스트
	Public Sub getDiyvaluationList
		Dim i, sqlStr, addSql

		If FRectLecturer_id <> "" then
			addSql = addSql & " and i.makerid = '" & FRectLecturer_id & "'" 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(v.idx) as cnt, CEILING(CAST(Count(v.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_diy_Item_Evaluate] as v "
		sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_diy_item as i on v.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as d on i.dispcate1 = d.catecode and d.useyn = 'Y' "
		sqlStr = sqlStr & " and v.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " v.idx, v.UserID, v.itemid, v.TotalPoint, v.Contents, v.Regdate, v.File1, v.File2, v.orderserial, v.itemoption, i.itemname, i.brandname "
		sqlStr = sqlStr & " FROM db_academy.[dbo].[tbl_diy_Item_Evaluate] as v "
		sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_diy_item as i on v.itemid = i.itemid "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as d on i.dispcate1 = d.catecode and d.useyn = 'Y' "
		sqlStr = sqlStr & " and v.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY v.idx DESC"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cCornerItem
					FItemList(i).FIdx			= rsACADEMYget("idx")
					FItemList(i).FUserid		= rsACADEMYget("userid")
					FItemList(i).FItemID		= rsACADEMYget("ItemID")
					FItemList(i).FTotalPoint	= rsACADEMYget("TotalPoint")
					FItemList(i).FContents	 	= db2html(rsACADEMYget("contents"))
					FItemList(i).FRegdate		= rsACADEMYget("regdate")
					FItemList(i).Flinkimg1		= rsACADEMYget("file1")
					FItemList(i).Flinkimg2		= rsACADEMYget("file2")
					FItemList(i).FItemname		= db2html(rsACADEMYget("itemname"))
					FItemList(i).FBrandname		= db2html(rsACADEMYget("brandname"))
					FItemList(i).FOrderserial	= rsACADEMYget("orderserial")
					FItemList(i).FItemoption	= rsACADEMYget("itemoption")
				rsACADEMYget.MoveNext
				i = i + 1
			Loop
		End if
		rsACADEMYget.close
	End Sub


	'수강 후기 리스트
	Public Sub getValuationList()
		Dim i, sqlStr, addSql

		If FRectLecturer_id <> "" then
			addSql = addSql & " and i.lecturer_id = '" & FRectLecturer_id & "'" 
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(v.idx) as cnt, CEILING(CAST(Count(v.idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_valuation v "
		sqlStr = sqlStr & " JOIN [db_academy].[dbo].tbl_lec_item i on v.masteridx = i.idx "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
		sqlStr = sqlStr & " and v.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage) & " v.idx, v.userid "
		sqlStr = sqlStr & " ,v.point1, v.point2, v.point3, v.point4, v.title, v.contents, v.file1, v.file2, v.regdate, i.lec_title, i.idx as lecidx "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].tbl_lec_valuation v "
		sqlStr = sqlStr & " JOIN [db_academy].[dbo].tbl_lec_item i on v.masteridx = i.idx "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_lec_Cate_large] as l on i.newCate_large=L.code_large and l.display_yn = 'Y' and l.code_large > 70 "
		sqlStr = sqlStr & " and v.isusing = 'Y' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY v.idx DESC"
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cCornerItem
					FItemList(i).FIdx			= rsACADEMYget("idx")
					FItemList(i).FUserid		= rsACADEMYget("userid")
					FItemList(i).FPoint1		= rsACADEMYget("point1")
					FItemList(i).FPoint2		= rsACADEMYget("point2")
					FItemList(i).FPoint3		= rsACADEMYget("point3")
					FItemList(i).FPoint4		= rsACADEMYget("point4")
					IF len(rsACADEMYget("title")) < 2 Then
						FItemList(i).FTitle		= "..."
					Else
						FItemList(i).FTitle		= db2html(rsACADEMYget("title"))
					End If
					FItemList(i).FContents		= db2html(rsACADEMYget("contents"))
					FItemList(i).FRegdate		= rsACADEMYget("regdate")
					FItemList(i).FFile1only		= rsACADEMYget("file1")
					FItemList(i).FFile2only		= rsACADEMYget("file2")
					If rsACADEMYget("file1") <> "" Then
						If Left(FItemList(i).FRegdate,10) > "2016-09-04" then
							FItemList(i).FFile1		= fingersImgUrl & "/contents/academy_evaluate/" & rsACADEMYget("file1")
						Else
							FItemList(i).FFile1		= "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" & rsACADEMYget("file1")
						End If 
					End If
					If rsACADEMYget("file2") <> "" Then
						If Left(FItemList(i).FRegdate,10) > "2016-09-04" then
							FItemList(i).FFile2		= fingersImgUrl & "/contents/academy_evaluate/" & rsACADEMYget("file2")
						Else
							FItemList(i).FFile2		= "http://imgstatic.10x10.co.kr/contents/academy_evaluate/" & rsACADEMYget("file2")
						End If 
					End if
					FItemList(i).FLecTitle		= db2html(rsACADEMYget("lec_title"))
					FItemList(i).FLecIdx		= rsACADEMYget("lecidx")
				rsACADEMYget.MoveNext
				i = i + 1
			Loop
		End if
		rsACADEMYget.close
	End Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
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
End Class

Function fnWhatIsMyJob(ilectureId)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 "
	strSql = strSql & " 	Case WHEN lec_yn = 'Y' and diy_yn = 'N' THEN 'L' "
	strSql = strSql & " 		 WHEN ((lec_yn = 'Y' and diy_yn = 'Y') OR (lec_yn = 'N' and diy_yn = 'Y')) THEN 'D' Else 'X' End as gubun "
	strSql = strSql & " FROM [db_academy].[dbo].tbl_lec_user "
	strSql = strSql & " WHERE lecturer_id = '"&ilectureId&"' "
	rsACADEMYget.Open strSql,dbACADEMYget,1
	If not rsACADEMYget.EOF Then
		fnWhatIsMyJob = rsACADEMYget("gubun")
	Else
		fnWhatIsMyJob = "X"
	End If
	rsACADEMYget.Close
End Function
%>