<%
class cUploadImageOneItem
	public FUseYN
	public FSortNo
	public FItemID
	public FItemName
	public FMakerID
	public FBrandName
	public FListImage100
	public FSiteGubun
	public FIdx
	public FOptInt
	public FOptChar
	public FRegImgCnt
	public FRegUserID
	public FRegdate
	
end Class

Class cUploadImage
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FRectUseYN
	public FRectSortNo
	public FRectItemID
	public FRectMakerId
	public FRectItemName
	public FRectSiteGubun
	public FRectNoUp
	public FUseYN
	public FItemID
	public FItemName
	
	
	Public Sub sbUploadImageMngList()
		Dim sqlStr, i, addsql
		
		If FRectSiteGubun <> "" Then
			addsql = addsql & " and u.sitegubun = '" & FRectSiteGubun & "' "
		End If
		
		If FRectItemID <> "" Then
			addsql = addsql & " and u.opt_int in(" & FRectItemID & ") "
		End If
		
		If FRectItemName <> "" Then
			addsql = addsql & " and i.itemname like '%" & FRectItemName & "%' "
		End If
		
		If FRectMakerId <> "" Then
			addsql = addsql & " and i.makerid = '" & FRectMakerId & "' "
		End If
		
		If FRectNoUp <> "" Then
			addsql = addsql & " and u.regimgcnt < 1 "
		End If
		
		sqlStr = "SELECT count(u.idx) FROM [db_sitemaster].[dbo].[tbl_common_upload_image] as u "
		
		If FRectSiteGubun = "china" Then
			sqlStr = sqlStr & "INNER JOIN [db_item].[dbo].[tbl_item] as i ON u.opt_int = i.itemid "
		End IF
		
		sqlStr = sqlStr & "WHERE 1=1" & addsql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
		FTotalCount = rsget.RecordCount
		rsget.close
		
		If FTotalCount > 0 Then
			sqlStr = "SELECT u.idx, u.sitegubun, u.opt_int, u.opt_char, u.regimgcnt, u.reguserid, u.regdate "
			sqlStr = sqlStr & ", i.itemname, i.listimage, i.makerid, i.brandname "
			sqlStr = sqlStr & "FROM [db_sitemaster].[dbo].[tbl_common_upload_image] as u "
			
			If FRectSiteGubun = "china" Then
				sqlStr = sqlStr & "INNER JOIN [db_item].[dbo].[tbl_item] as i ON u.opt_int = i.itemid "
			End IF
			
			sqlStr = sqlStr & "WHERE 1=1" & addsql & " ORDER BY "
			
			If FRectSortNo = "" Then
			 	sqlStr = sqlStr & " u.idx DESC"
			Else
				If FRectSortNo = "1" Then		'이미지등록 최신순
					sqlStr = sqlStr & " u.idx DESC"
				ElseIf FRectSortNo = "2" Then		'이미지등록 오래된순
					sqlStr = sqlStr & " u.idx ASC"
				ElseIf FRectSortNo = "3" Then		'상품코드 최신순
					sqlStr = sqlStr & " i.itemid DESC"
				ElseIf FRectSortNo = "4" Then		'상품코드 오래된순
					sqlStr = sqlStr & " i.itemid ASC"
				End If
			End If
			rsget.pagesize = FPageSize
			rsget.Open sqlStr,dbget,1
			FResultCount = rsget.RecordCount
			FTotalCount = FResultCount
			
			Redim preserve FItemList(FResultCount)
			i = 0
			If not rsget.EOF Then
				rsget.absolutepage = FCurrPage
				Do until rsget.EOF
					Set FItemList(i) = new cUploadImageOneItem
						FItemList(i).FIdx 		= rsget("idx")
						FItemList(i).FSiteGubun	= rsget("sitegubun")
						FItemList(i).FOptInt 		= rsget("opt_int")
						FItemList(i).FOptChar 		= rsget("opt_char")
						FItemList(i).FRegImgCnt 	= rsget("regimgcnt")
						FItemList(i).FRegUserID 	= rsget("reguserid")
						FItemList(i).FRegdate 		= rsget("regdate")
						FItemList(i).FItemName 	= db2html(rsget("itemname"))
						FItemList(i).FMakerID		= rsget("makerid")
						FItemList(i).FBrandName	= db2html(rsget("brandname"))
						FItemList(i).FListImage100	= "http://webimage.10x10.co.kr/image/List/" & GetImageSubFolderByItemid(rsget("opt_int")) & "/" & rsget("listimage")

					i = i + 1
					rsget.moveNext
				Loop
			End If
			rsget.Close
		End If
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

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FScrollCount = 10
	End Sub

	Private Sub Class_Terminate()
    End Sub
end Class
%>