<%
Class CCommonItem
	Public FItemid
	Public FItemname
	Public FSocname_kor
	Public FMakerid

	Public FIdx
	Public FStartDate
	Public FEndDate
	Public FMargin
	Public FIsusing
	Public FRegdate
	Public FKeywords
	Public FBigo

	Public FId
	Public FSourceArea

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CCommon
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectMallGubun
	Public FRectMakerID
	Public FRectItemID
	Public FRectSItemid
	Public FRectItemName
	Public FRectIdx
	Public FRectIsusing
	Public FRectBigo
	Public FRectBigoText
	Public FRectSourceArea

	Public FRectoutmallorderserial
	Public FRectIsSongjang

	Public Sub getTargetMall_Not_In_makerid_List
		Dim sqlStr, i, addsql

		If FRectMallGubun <> "" Then
			addsql = addsql & " and ti.mallgubun = '"& FRectMallGubun &"' "
		End If

		If FRectMakerID <> "" Then
			addsql = addsql & " and c.userid = '"& FRectMakerID &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_targetMall_Not_in_makerid as ti "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on ti.makerid = c.userid "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " ti.makerid, c.socname_kor "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_targetMall_Not_in_makerid as ti "
		sqlStr = sqlStr & " JOIN db_user.dbo.tbl_user_c as c on ti.makerid = c.userid "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		sqlStr = sqlStr & " ORDER BY ti.regdate DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCommonItem
	                FItemList(i).FMakerid		= rsget("makerid")
	                FItemList(i).FSocname_kor	= rsget("socname_kor")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getTargetMall_Not_In_itemid_List
		Dim sqlStr, i, addsql

		If FRectMallGubun <> "" Then
			addsql = addsql & " and ti.mallgubun = '"& FRectMallGubun &"' "
		End If

		If FRectMakerID <> "" Then
			addsql = addsql & " and i.makerid = '"& FRectMakerID &"' "
		End If

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            End If
        End If

		If FRectItemName <> "" Then
			addsql = addsql & " and i.itemname Like '%" & FRectItemName & "%' "
		End If

		Select Case FRectBigo
			Case "Y"
				addSql = addSql & " AND isNull(ti.bigo, '') <> ''  "
			Case "N"
				addSql = addSql & " AND isNull(ti.bigo, '') = ''  "
		End Select

		If FRectBigo = "Y" and FRectBigoText <> "" Then
			addSql = addSql & " AND isNull(ti.bigo, '') like '%"& FRectBigoText &"%'  "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_targetMall_Not_in_itemid as ti "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on ti.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top " + CStr(FPageSize*FCurrPage) + " ti.itemid, i.itemname, ti.bigo "
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_targetMall_Not_in_itemid as ti "
		sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on ti.itemid = i.itemid "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		sqlStr = sqlStr & " ORDER BY ti.regdate DESC "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CCommonItem
	                FItemList(i).FItemid	= rsget("itemid")
	                FItemList(i).FItemname	= rsget("itemname")
					FItemList(i).FBigo	= rsget("bigo")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getMarginCateOneItem
	    Dim i, sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 idx, startDate, endDate, margin, isusing "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginCate_master] "
	    sqlStr = sqlStr & " WHERE idx = " & CStr(FRectIdx)
		sqlStr = sqlStr & " and mallid = '"&FRectMallGubun&"' "
		rsget.Open sqlStr,dbget,1
		FResultCount = rsget.RecordCount
		set FOneItem = new CCommonItem
		If not rsget.EOF Then
			FOneItem.FIdx			= rsget("idx")
			FOneItem.FStartDate		= rsget("startDate")
			FOneItem.FEndDate		= rsget("endDate")
			FOneItem.FMargin		= rsget("margin")
			FOneItem.FIsusing		= rsget("isusing")
		End If
		rsget.Close
	End Sub

	Public Sub getMarginCateList
		Dim sqlStr, addSql, i
		If FRectIsusing <> "" Then
			addSql = addSql & " and isusing = '"& FRectIsusing &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginCate_master] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " and mallid = '"&FRectMallGubun&"' "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " idx, startDate, endDate, margin, isusing, regdate "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginCate_master] "
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " and mallid = '"&FRectMallGubun&"' "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CCommonItem
				    FItemList(i).FIdx			= rsget("idx")
					FItemList(i).FStartDate		= rsget("startDate")
					FItemList(i).FEndDate		= rsget("endDate")
					FItemList(i).FMargin		= rsget("margin")
					FItemList(i).FIsusing		= rsget("isusing")
					FItemList(i).FRegdate		= rsget("regdate")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getTenKeyWordsList
		Dim sqlStr, addSql, add

		'상품코드 검색
        If (FRectItemid <> "") then
            If Right(Trim(FRectItemid) ,1) = "," Then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and c.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            Else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and c.itemid in (" + FRectItemid + ")"
            End If
		Else
			Exit Function
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 20 "
		sqlStr = sqlStr & " c.itemid, c.keywords, isnull(k.keywords, '') as chgKeywords "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_contents as c "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_keywords] as k on c.itemid = k.itemid and mallid = '" & FRectMallGubun & "'"
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY c.itemid DESC "
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			getTenKeyWordsList = rsget.getRows()
		End If
		rsget.Close
	End Function

	Public Sub getOutmallKeyWordsList
		Dim sqlStr, addSql
        If (FRectMallGubun <> "") then
            addSql = addSql & " and mallid = '" & FRectMallGubun & "'"
		Else
			Exit Sub
        End If

		'상품코드 검색
        If (FRectSItemid <> "") then
            If Right(Trim(FRectSItemid) ,1) = "," Then
            	FRectSItemid = Replace(FRectSItemid,",,",",")
            	addSql = addSql & " and itemid in (" + Left(FRectSItemid,Len(FRectSItemid)-1) + ")"
            Else
				FRectSItemid = Replace(FRectSItemid,",,",",")
            	addSql = addSql & " and itemid in (" + FRectSItemid + ")"
            End If
        End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_keywords] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " itemid, keywords "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_keywords] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY IsNull(lastupdate, regdate) DESC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CCommonItem
				    FItemList(i).FItemid			= rsget("itemid")
					FItemList(i).FKeywords		= rsget("keywords")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Sub getOutmallSSGSourceAreaMappList
		Dim sqlStr, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] m  "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_ssg_sourceAreaCode] o on m.sourcearea = o.sourcearea "
		sqlStr = sqlStr & " WHERE isnull(o.id, '') = '' "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " m.id, m.sourceArea "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] m  "
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.[dbo].[tbl_ssg_sourceAreaCode] o on m.sourcearea = o.sourcearea "
		sqlStr = sqlStr & " WHERE isnull(o.id, '') = '' "
		sqlStr = sqlStr & " ORDER BY m.id ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CCommonItem
				    FItemList(i).FId			= rsget("id")
					FItemList(i).FSourceArea	= rsget("sourceArea")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub


	Public Sub getOutmallOrgSSGSourceAreaList
		Dim sqlStr, addSql

		If FRectSourceArea <> "" Then
			addSql = addSql & " and sourcearea = '"& FRectSourceArea &"' "
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg  " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCode] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage)
		sqlStr = sqlStr & " id, sourceArea "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCode] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY sourceArea ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CCommonItem
				    FItemList(i).FId			= rsget("id")
					FItemList(i).FSourceArea	= rsget("sourceArea")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getOutmallNotKeyWordsList
		Dim sqlStr, addSql, add
		sqlStr = ""
		sqlStr = sqlStr & " SELECT idx, mallid, keywords "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_notKeywords] "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY idx "
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			getOutmallNotKeyWordsList = rsget.getRows()
		End If
		rsget.Close
	End Function

	Public Function getgs25SongjangList
		Dim sqlStr, addSql, arrRows

		If FRectOutMallOrderSerial <> "" Then
			addSql = addSql & " and T.outmallorderserial in ("& FRectOutMallOrderSerial &") "
		End If

		If FRectIsSongjang <> "" Then
			Select Case FRectIsSongjang
				Case "Y"
					addSql = addSql & " and isNull(D.songjangNo, '') <> '' "
				Case "N"
					addSql = addSql & " and isNull(D.songjangNo, '') = '' "
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT T.OutMallOrderSerial, T.OrgDetailKey, T.outMallGoodsNo, T.orderItemName "
		sqlStr = sqlStr & " , 'hyundai' as songjangDiv, isNull(D.songjangNo, '') as songjangNo "
		sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder T WITH(NOLOCK) "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_master M WITH(NOLOCK) on T.orderserial=M.orderserial"
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_order_detail D WITH(NOLOCK) on T.orderserial=D.orderserial "
		sqlStr = sqlStr & " 	and IsNull(T.changeitemid, T.matchitemid)=D.itemid "
		sqlStr = sqlStr & " 	and IsNull(T.changeitemoption, T.matchitemoption)=D.itemoption "
		sqlStr = sqlStr & " 	and D.currstate=7 "
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div V WITH(NOLOCK) on D.songjangDiv=V.divcd "
		sqlStr = sqlStr & " WHERE 1=1 "
		sqlStr = sqlStr & " and T.sellsite in ('GS25')"
		sqlStr = sqlStr & " and T.matchState not in ('R','D','B') "
		sqlStr = sqlStr & " and IsNULL(T.sendState,0) = 0 "
		sqlStr = sqlStr & addSql
		sqlStr = sqlStr & " ORDER BY T.regdate desc"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			getgs25SongjangList = rsget.getRows()
		End If
		rsget.Close
	End Function

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