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
	'구매 후기 리스트
	Public Sub getDiyvaluationList
		Dim i, sqlStr, addSql

		addSql = addSql & " and i.makerid = '" & FRectLecturer_id & "'" 
		If FRectSearchDIV<>0 Then
			If FRectSearchDIV=1 Then
				addSql = addSql & " and i.itemid like '%" & FRectSearchTXT & "%'"
			ElseIf FRectSearchDIV=2 Then
				addSql = addSql & " and i.itemname like '%" & FRectSearchTXT & "%'"
			ElseIf FRectSearchDIV=3 Then
				addSql = addSql & " and v.userid like '%" & FRectSearchTXT & "%'"
			ElseIf FRectSearchDIV=4 Then
				addSql = addSql & " and v.contents like '%" & FRectSearchTXT & "%'"
			End If
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

%>