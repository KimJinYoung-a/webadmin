<%
class cDispCateMenuOneItem
	public FIdx
	public Ftype
	public Fsubject
	public Fitemid
	public Fimgurl
	public FimgurlReal
	public Flinkurl
	public Fstartdate
	public Fenddate
	public Fuseyn
	public Freguserid
	public Fregusername
	public Fregdate
	public Fsortno
	public Fcatename
	public Fdisp1
	
end Class

Class cDispCateMenu
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
	public FCateCode
	public FDepth
	public FCateName
	public FCateName_E
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
	public FDisp1
	public FType
	public FOrderBy

	
	Public Function GetDispCateMenuList()
		Dim sqlStr, i, addsql

		sqlStr = sqlStr & "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	m.type, m.number, m.value, " & vbCrLf
		sqlStr = sqlStr & " 	CASE WHEN type = 'category' " & vbCrLf
		sqlStr = sqlStr & " 		THEN (select catename from [db_item].[dbo].[tbl_display_cate] where catecode = convert(bigint,m.value)) " & vbCrLf
		sqlStr = sqlStr & " 		ELSE m.value " & vbCrLf
		sqlStr = sqlStr & " 	END AS codename " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate_menu] AS m " & vbCrLf
		sqlStr = sqlStr & " 	WHERE m.catecode = '" & FRectCateCode & "' AND m.useyn = 'y' " & vbCrLf
		sqlStr = sqlStr & "ORDER BY m.number ASC" & vbCrLf
'rw sqlStr
		rsget.Open sqlStr,dbget,1

		IF not rsget.EOF THEN
			GetDispCateMenuList = rsget.getRows()
		END IF
		rsget.Close
		
	End Function
	
	
	Public Function GetDispCate1Depth()
		Dim sqlStr, i, addsql

		sqlStr = sqlStr & "SELECT " & vbCrLf
		sqlStr = sqlStr & " 	catecode, catename " & vbCrLf
		sqlStr = sqlStr & " FROM [db_item].[dbo].[tbl_display_cate] " & vbCrLf
		sqlStr = sqlStr & " 	WHERE depth = '1' " & vbCrLf
		sqlStr = sqlStr & "ORDER BY sortno ASC" & vbCrLf
'rw sqlStr
		rsget.Open sqlStr,dbget,1

		IF not rsget.EOF THEN
			GetDispCate1Depth = rsget.getRows()
		END IF
		rsget.Close
		
	End Function
	
	
	Public Sub GetCateMainIssueList()
		Dim sqlStr, i, addsql
		
		If FDisp1 <> "" Then
			addsql = addsql & " and a.disp1 = '" & FDisp1 & "' "
		End IF
		
		If FType <> "" Then
			addsql = addsql & " and a.type = '" & FType & "' "
		Else
			addsql = addsql & " and a.type in('issue_text','issue_image') "
		End IF
		
		If FUseYN <> "" Then
			addsql = addsql & " and a.useyn = '" & FUseYN & "' "
		End IF
		
		sqlStr = sqlStr & "select count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/20) AS totPg from db_item.dbo.tbl_display_cate_menu_top as a where 1=1 " & vbCrLf
		sqlStr = sqlStr & " " & addsql & " "
'rw sqlStr
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close
		
		sqlStr = "SELECT Top " & CStr(20*FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	a.idx, a.disp1, a.type, a.subject, a.linkurl, a.itemid, a.imgurl, a.sortno, a.sdate, a.edate, a.useyn, a.regdate, a.reguserid, " & vbCrLf
		sqlStr = sqlStr & " 	(select username from db_partner.dbo.tbl_user_tenbyten where userid = a.reguserid) as regusername, " & vbCrLf
		sqlStr = sqlStr & " 	(select listimage120 from db_item.dbo.tbl_item where itemid = a.itemid) as imagename, " & vbCrLf
		sqlStr = sqlStr & " 	(select top 1 catename from db_item.dbo.tbl_display_cate where catecode = a.disp1 and depth = 1) as catename " & vbCrLf
		sqlStr = sqlStr & " from db_item.dbo.tbl_display_cate_menu_top as a where 1=1 " & addsql & " " & vbCrLf
		sqlStr = sqlStr & "order by " & FOrderBy & " " & vbCrLf
'rw sqlStr
		rsget.pagesize = 20
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(20*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new cDispCateMenuOneItem
					FItemList(i).FIdx 			= rsget("idx")
					FItemList(i).Ftype 			= rsget("type")
					FItemList(i).Fsubject		= db2html(rsget("subject"))
					FItemList(i).Fitemid		= rsget("itemid")
					FItemList(i).Fimgurl		= "http://webimage.10x10.co.kr/image/List120/" & GetImageSubFolderByItemid(rsget("itemid")) & "/" & rsget("imagename")
					FItemList(i).FimgurlReal	= rsget("imgurl")
					FItemList(i).Flinkurl		= rsget("linkurl")
					FItemList(i).Fstartdate 	= rsget("sdate")
					FItemList(i).Fenddate 		= rsget("edate")
					FItemList(i).Fuseyn 		= rsget("useyn")
					FItemList(i).Fsortno		= rsget("sortno")
					FItemList(i).Freguserid		= rsget("reguserid")
					FItemList(i).Fregusername	= rsget("regusername")
					FItemList(i).Fregdate 		= rsget("regdate")
					FItemList(i).Fcatename 		= rsget("catename")
					FItemList(i).Fdisp1			= rsget("disp1")
					
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
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