<%
Class CWriterItem
	Public FIdx
	Public FGubun
	Public FWritername
	Public FBunya
	Public FZipcode
	Public FAddress1
	Public FAddress2
	Public FUsercell
	Public FUserphone
	Public FUsermail
	Public FHomepage
	Public FIntroduce
	Public FEtc
	Public FWritefile
	Public FDeleteyn
	Public FConfirmyn
	Public FConfirmMemo
	Public FRegdate

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CWriter
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	
	Public FRectIdx
	Public FRectSearchConfirm
	Public FRectSearchKey
	Public FRectsearchString
	Public upfolder

	Public Sub getWriterRegedItemList
		Dim i, sqlStr, addSql
		
		If FRectSearchConfirm <> "" Then
			addSql = addSql & " and confirmyn = '"&FRectSearchConfirm&"' "
		End If
		
		If FRectSearchKey <> "" AND FRectsearchString <> "" Then
			If FRectSearchKey = "writername" Then
				addSql = addSql & " and writername = '"&FRectsearchString&"' "
			ElseIf FRectSearchKey = "bunya" Then
				addSql = addSql & " and bunya like '%"&FRectsearchString&"%' "
			End If
		End If
		
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(idx) as cnt, CEILING(CAST(Count(idx) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_partner_writer] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and deleteyn = 'N' "
		sqlStr = sqlStr & addSql
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close


		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, gubun, writername, bunya, zipcode, address1, address2, usercell, userphone, usermail, homepage, introduce, etc, writefile, deleteyn, confirmyn, regdate "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_partner_writer] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and deleteyn = 'N' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				Set FItemList(i) = new CWriterItem
					FItemList(i).FIdx			= rsACADEMYget("idx")
					FItemList(i).FGubun			= rsACADEMYget("gubun")
					FItemList(i).FWritername	= db2html(rsACADEMYget("writername"))
					FItemList(i).FBunya			= db2html(rsACADEMYget("bunya"))
					FItemList(i).FZipcode		= db2html(rsACADEMYget("zipcode"))
					FItemList(i).FAddress1		= db2html(rsACADEMYget("address1"))
					FItemList(i).FAddress2		= db2html(rsACADEMYget("address2"))
					FItemList(i).FUsercell		= db2html(rsACADEMYget("usercell"))
					FItemList(i).FUserphone		= db2html(rsACADEMYget("userphone"))
					FItemList(i).FUsermail		= db2html(rsACADEMYget("usermail"))
					FItemList(i).FHomepage		= db2html(rsACADEMYget("homepage"))
					FItemList(i).FIntroduce		= db2html(rsACADEMYget("introduce"))
					FItemList(i).FEtc			= db2html(rsACADEMYget("etc"))
					FItemList(i).FWritefile		= db2html(rsACADEMYget("writefile"))
					FItemList(i).FDeleteyn		= db2html(rsACADEMYget("deleteyn"))
					FItemList(i).FConfirmyn		= db2html(rsACADEMYget("confirmyn"))
					FItemList(i).FRegdate		= db2html(rsACADEMYget("regdate"))
				i = i + 1
				rsACADEMYget.moveNext
			Loop
		End If
		rsACADEMYget.Close
	End Sub
	
	Public Sub getWriterViewOneitem
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 "
		sqlStr = sqlStr & " idx, gubun, writername, bunya, zipcode, address1, address2, usercell, userphone, usermail, homepage, introduce, etc, writefile, deleteyn, confirmyn, confirmMemo, regdate "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_partner_writer] "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & " and idx = '"& FRectIdx &"' "
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If not rsACADEMYget.EOF Then
			Set FOneItem = new CWriterItem
				FOneItem.FIdx				= rsACADEMYget("idx")
				FOneItem.FGubun				= rsACADEMYget("gubun")
				FOneItem.FWritername		= db2html(rsACADEMYget("writername"))
				FOneItem.FBunya				= db2html(rsACADEMYget("bunya"))
				FOneItem.FZipcode			= db2html(rsACADEMYget("zipcode"))
				FOneItem.FAddress1			= db2html(rsACADEMYget("address1"))
				FOneItem.FAddress2			= db2html(rsACADEMYget("address2"))
				FOneItem.FUsercell			= db2html(rsACADEMYget("usercell"))
				FOneItem.FUserphone			= db2html(rsACADEMYget("userphone"))
				FOneItem.FUsermail			= db2html(rsACADEMYget("usermail"))
				FOneItem.FHomepage			= db2html(rsACADEMYget("homepage"))
				FOneItem.FIntroduce			= db2html(rsACADEMYget("introduce"))
				FOneItem.FEtc				= db2html(rsACADEMYget("etc"))
				FOneItem.FWritefile			= db2html(rsACADEMYget("writefile"))
				FOneItem.FDeleteyn			= db2html(rsACADEMYget("deleteyn"))
				FOneItem.FConfirmyn			= db2html(rsACADEMYget("confirmyn"))
				FOneItem.FConfirmMemo		= db2html(rsACADEMYget("confirmMemo"))
				FOneItem.FRegdate			= db2html(rsACADEMYget("regdate"))
		End If
		rsACADEMYget.close
	End Sub

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		upfolder = "/contents/partnership/"		'업로드 폴더
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