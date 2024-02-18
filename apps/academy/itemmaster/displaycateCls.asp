<%
class cDispCateOneItem
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FMakerID
	public FSmallImage
	
end Class

Class cDispCate
	Public FItemList()
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FResultCount
	Public FScrollCount
	public FRectCateCode
	public FRectDepth
	public FRectCateName
	public FRectUseYN
	public FRectSortNo
	public FRectItemID
	public FRectIsDefault
	public FRectCDL
	public FRectCDM
	public FRectCDS
	public FRectMakerId
	public FRectItemName
	public FRectKeyword
	public FRectSellYN
	public FRectIsUsing
	public FRectDanjongyn
	public FRectLimityn
	public FRectSailYn
	public FRectDeliveryType
	public FRectSortDiv
	public FRectNotCateReg
	public FRectSiteGubun
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
	public FJaehuname
	public FCateFullName
	public FUseYN
	public FSortNo
	public FItemID
	public FIsDefault
	public FCDL
	public FCDM
	public FCDS
	public FItemName
	public FCateNameTitle
	public FSearchDispCate
	public FIsNew

	
	Public Sub GetDispCateList()
		Dim sqlStr, i, addsql

		sqlStr = sqlStr & "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	c.catecode, c.depth, c.catename, c.useyn, c.sortNo " & vbCrLf
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS c " & vbCrLf
		sqlStr = sqlStr & " 	WHERE c.depth = '" & FRectDepth & "' " & vbCrLf
		
		If FRectUseYN <> "" Then
			sqlStr = sqlStr & " AND c.useyn = '" & FRectUseYN & "' " & vbCrLf
		End IF

		if FRectSiteGubun="upche" Then
			'업체뷰에서는 스페셜 카테고리 제외
			sqlStr = sqlStr & " AND c.catecode <> '101' " & vbCrLf
		end if

		sqlStr = sqlStr & " AND Left(c.catecode,3) = '" & Left(FRectCateCode,3) & "' " & vbCrLf
		sqlStr = sqlStr & "ORDER BY c.depth ASC, c.sortNo ASC, c.catecode ASC" & vbCrLf
'Response.write sqlStr
'response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new cDispCateOneItem
					FItemList(i).FCateCode 		= rsACADEMYget("catecode")
					FItemList(i).FDepth 		= rsACADEMYget("depth")
					FItemList(i).FCateName 		= db2html(rsACADEMYget("catename"))
					FItemList(i).FUseYN 		= rsACADEMYget("useyn")
					FItemList(i).FSortNo 		= rsACADEMYget("sortNo")
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
		
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
End Class

%>