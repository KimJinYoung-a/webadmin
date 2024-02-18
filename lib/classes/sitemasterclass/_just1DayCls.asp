<%
Class Cjust1DayItem
	public FJustDate
	public Fitemid
	public Fitemname
	public FsmallImage
	public ForgPrice
	public FjustSalePrice
	public FsaleSuplyCash
	public FjustDesc
	public Fsale_code
	public FlimitNo
	public FsellYn
	public Fimg1
	public Fimg2
	public Fimg3
	public Fimg4
	public Fregdate
	public FCateName
	public Fitemdiv

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class Cjust1Day
	public FItemList()

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectDate
	public FRectItemId
	public FRectSdt
	public FRectEdt
	public FRectDispCate
	Public FRectItemName

	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function Getjust1DayList()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectDate<>"" then
			addSql = addSql & " and t1.justDate='" & FRectDate & "'"
		end if
		if FRectItemId<>"" then
			addSql = addSql & " and t1.itemid=" & FRectItemId
		end If
		if FRectItemName<>"" then
			addSql = addSql & " and t2.itemname like '%" & FRectItemName & "%'"
		end if
		if Not(FRectSdt="" or FRectEdt="") then
			addSql = addSql & " and t1.justDate between '" & FRectSdt & "' and '" & FRectEdt & "'"
		end if
		if FRectDispCate<>"" then
			addSql = addSql + " and t1.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if
		'카운트
		sqlStr = "select count(t1.justDate) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_just1day_temp as t1" + vbcrlf
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item as t2 " + vbcrlf
		sqlStr = sqlStr + " on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash "
		sqlStr = sqlStr + " , t1.sale_code, t1.limitNo, t1.regdate" + vbcrlf
		sqlStr = sqlStr + " , t2.itemname, t2.sellYn, t2.smallImage " + vbcrlf
		sqlStr = sqlStr + " , isNull(t1.img1,'') as img1, isNull(t1.img2,'') as img2, isNull(t1.img3,'') as img3, isNull(t1.img4,'') as img4, " + vbcrlf
		sqlStr = sqlStr + " 	STUFF(( " + vbCrLf
		sqlStr = sqlStr + " 		SELECT '|^|' + convert(varchar,dci2.catecode) + '$' + ([db_item].[dbo].[getCateCodeFullDepthName](dci2.catecode)) " + vbCrLf
		sqlStr = sqlStr + " 		+ case when dci2.isDefault = 'y' then ' [기본]' else ' [추가]' end " + vbCrLf
		sqlStr = sqlStr + " 		FROM [db_item].[dbo].[tbl_display_cate] AS dc " + vbCrLf
		sqlStr = sqlStr + " 			INNER JOIN [db_item].[dbo].[tbl_display_cate_item] AS dci2 on dc.catecode = dci2.catecode " + vbCrLf
		sqlStr = sqlStr + " 		WHERE dci2.itemid = t1.itemid " + vbCrLf
		sqlStr = sqlStr + " 		ORDER BY dci2.isDefault DESC " + vbCrLf
		sqlStr = sqlStr + " 	FOR XML PATH('') " + vbCrLf
		sqlStr = sqlStr + " 	), 1, 3, '') AS catename "
		sqlStr = sqlStr + " 	, (select justDesc from [db_sitemaster].[dbo].tbl_just1day_temp where JustDate=t1.JustDate) as justDesc, t2.itemdiv"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_just1day_temp as t1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_item].[dbo].tbl_item as t2 " + vbcrlf
		sqlStr = sqlStr + "			on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " 	LEFT JOIN [db_item].[dbo].[tbl_display_cate_item] AS i4 on t1.itemid = i4.itemid " + vbCrLf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " group by t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash , t1.sale_code, t1.limitNo, t1.regdate , t2.itemname, t2.sellYn, t2.smallImage,t1.img1 ,t1.img2,t1.img3,t1.img4, t2.itemdiv "
		sqlStr = sqlStr + " order by t1.justdate desc "
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cjust1DayItem

				FItemList(i).FJustDate		= rsget("JustDate")
				FItemList(i).FjustDesc		= db2html(rsget("justDesc"))
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemname		= rsget("itemname")
				If rsget("itemdiv")="21" Then
				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + rsget("smallImage")
				Else
				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				End If
				FItemList(i).ForgPrice		= rsget("orgPrice")
				FItemList(i).FjustSalePrice	= rsget("justSalePrice")
				FItemList(i).FsaleSuplyCash	= rsget("saleSuplyCash")
				FItemList(i).Fsale_code		= rsget("sale_code")
				FItemList(i).FlimitNo		= rsget("limitNo")
				FItemList(i).FsellYn		= rsget("sellYn")
				FItemList(i).Fimg1			= rsget("img1")
				FItemList(i).Fimg2			= rsget("img2")
				FItemList(i).Fimg3			= rsget("img3")
				FItemList(i).Fimg4			= rsget("img4")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).FCateName 		= db2html(rsget("catename"))
				If FItemList(i).FCateName = "" Then
					FItemList(i).FCateName = "<center>없음</center>"
				End If
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function Getjust1Daymodify()
		dim sqlStr, addSql, i

		'추가 조건절
		if FRectDate<>"" then
			addSql = addSql & " and t1.justDate='" & FRectDate & "'"
		end if
		if FRectItemId<>"" then
			addSql = addSql & " and t1.itemid=" & itemid
		end if
		if Not(FRectSdt="" or FRectEdt="") then
			addSql = addSql & " and t1.justDate between '" & FRectSdt & "' and '" & FRectEdt & "'"
		end if

		'카운트
		sqlStr = "select count(t1.justDate) as cnt"
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_just1day_temp as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsget.Open sqlStr,dbget,1
		FTotalCount = rsget("cnt")
		rsget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash "
		sqlStr = sqlStr + " , t1.justDesc, t1.sale_code, t1.limitNo, t1.regdate " + vbcrlf
		sqlStr = sqlStr + " , t2.itemname, t2.sellYn, t2.smallImage " + vbcrlf
		sqlStr = sqlStr + " , isNull(t1.img1,'') as img1, isNull(t1.img2,'') as img2, isNull(t1.img3,'') as img3, isNull(t1.img4,'') as img4, t2.itemdiv " + vbcrlf
		sqlStr = sqlStr + " from [db_sitemaster].[dbo].tbl_just1day_temp as t1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_item].[dbo].tbl_item as t2 " + vbcrlf
		sqlStr = sqlStr + "			on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by t1.justdate desc "

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new Cjust1DayItem

				FItemList(i).FJustDate		= rsget("JustDate")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemname		= rsget("itemname")
				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallImage")
				FItemList(i).ForgPrice		= rsget("orgPrice")
				FItemList(i).FjustSalePrice	= rsget("justSalePrice")
				FItemList(i).FsaleSuplyCash	= rsget("saleSuplyCash")
				FItemList(i).FjustDesc		= db2html(rsget("justDesc"))
				FItemList(i).Fsale_code		= rsget("sale_code")
				FItemList(i).FlimitNo		= rsget("limitNo")
				FItemList(i).FsellYn		= rsget("sellYn")
				FItemList(i).Fimg1			= rsget("img1")
				FItemList(i).Fimg2			= rsget("img2")
				FItemList(i).Fimg3			= rsget("img3")
				FItemList(i).Fimg4			= rsget("img4")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fitemdiv		= rsget("itemdiv")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

Function fnCateCodeNameSplit(n,itemid)
	Dim i, arr, vBody
	If n <> "" AND n <> "<center>없음</center>" Then
		arr = Split(n,"|^|")
		For i = LBound(arr) To UBound(arr)
			vBody = vBody & Split(arr(i),"$")(1)
			If i <> UBound(arr) Then
				vBody = vBody & "<br>"
			End If
		Next
	Else
		vBody = vBody & "<center>없음</center>"
	End IF
	vBody = Replace(vBody,"^^","-")
	fnCateCodeNameSplit = vBody
End Function
%>