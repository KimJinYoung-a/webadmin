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
	public Fidx

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
			addSql = addSql & " and t1.itemid=" & itemid
		end if
		if Not(FRectSdt="" or FRectEdt="") then
			addSql = addSql & " and t1.justDate between '" & FRectSdt & "' and '" & FRectEdt & "'"
		end if
		if FRectDispCate<>"" then
			addSql = addSql + " and t1.itemid in (select itemid from [db_academy].[dbo].[tbl_display_cate_item_Academy] where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

		'카운트
		sqlStr = "select count(t1.justDate) as cnt"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_just1day as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t1.idx, t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash "
		sqlStr = sqlStr + " , t1.sale_code, t1.limitNo, t1.regdate " + vbcrlf
		sqlStr = sqlStr + " , t2.itemname, t2.sellYn, t2.smallImage " + vbcrlf
		sqlStr = sqlStr + " , isNull(t1.img1,'') as img1, isNull(t1.img2,'') as img2, isNull(t1.img3,'') as img3, isNull(t1.img4,'') as img4, " + vbcrlf
		sqlStr = sqlStr + " 	STUFF(( " + vbCrLf
		sqlStr = sqlStr + " 		SELECT '|^|' + convert(varchar,dci2.catecode) + '$' + ([db_academy].[dbo].[getCateCodeFullDepthName_Academy](dci2.catecode)) " + vbCrLf
		sqlStr = sqlStr + " 		+ case when dci2.isDefault = 'y' then ' [기본]' else ' [추가]' end " + vbCrLf
		sqlStr = sqlStr + " 		FROM [db_academy].[dbo].[tbl_display_cate_Academy] AS dc " + vbCrLf
		sqlStr = sqlStr + " 			INNER JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] AS dci2 on dc.catecode = dci2.catecode " + vbCrLf
		sqlStr = sqlStr + " 		WHERE dci2.itemid = t1.itemid " + vbCrLf
		sqlStr = sqlStr + " 		ORDER BY dci2.isDefault DESC " + vbCrLf
		sqlStr = sqlStr + " 	FOR XML PATH('') " + vbCrLf
		sqlStr = sqlStr + " 	), 1, 3, '') AS catename "
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_just1day as t1 " + vbcrlf
		sqlStr = sqlStr + "		Join [db_academy].[dbo].[tbl_diy_item] as t2 " + vbcrlf
		sqlStr = sqlStr + "			on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " 	LEFT JOIN [db_academy].[dbo].[tbl_display_cate_item_Academy] AS i4 on t1.itemid = i4.itemid " + vbCrLf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " group by t1.idx, t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash , t1.sale_code, t1.limitNo, t1.regdate , t2.itemname, t2.sellYn, t2.smallImage,t1.img1 ,t1.img2,t1.img3,t1.img4 "
		sqlStr = sqlStr + " order by t1.justdate desc "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new Cjust1DayItem

				FItemList(i).Fidx		= rsACADEMYget("idx")
				FItemList(i).FJustDate		= rsACADEMYget("JustDate")
				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).Fitemname		= rsACADEMYget("itemname")
'				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallImage")
				FItemList(i).FsmallImage	= imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/"  + rsACADEMYget("smallImage")
				FItemList(i).ForgPrice		= rsACADEMYget("orgPrice")
				FItemList(i).FjustSalePrice	= rsACADEMYget("justSalePrice")
				FItemList(i).FsaleSuplyCash	= rsACADEMYget("saleSuplyCash")
				FItemList(i).Fsale_code		= rsACADEMYget("sale_code")
				FItemList(i).FlimitNo		= rsACADEMYget("limitNo")
				FItemList(i).FsellYn		= rsACADEMYget("sellYn")
				FItemList(i).Fimg1			= rsACADEMYget("img1")
				FItemList(i).Fimg2			= rsACADEMYget("img2")
				FItemList(i).Fimg3			= rsACADEMYget("img3")
				FItemList(i).Fimg4			= rsACADEMYget("img4")
				FItemList(i).Fregdate		= rsACADEMYget("regdate")
				FItemList(i).FCateName 		= db2html(rsACADEMYget("catename"))
				If FItemList(i).FCateName = "" Then
					FItemList(i).FCateName = "<center>없음</center>"
				End If
				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_just1day as t1" + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
'response.write sqlStr
'response.end
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'목록 접수
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + "" + vbcrlf
		sqlStr = sqlStr + " t1.idx, t1.JustDate, t1.itemid, t1.orgPrice, t1.justSalePrice, t1.saleSuplyCash "
		sqlStr = sqlStr + " , t1.justDesc, t1.sale_code, t1.limitNo, t1.regdate " + vbcrlf
		sqlStr = sqlStr + " , t2.itemname, t2.sellYn, t2.smallImage " + vbcrlf
		sqlStr = sqlStr + " , isNull(t1.img1,'') as img1, isNull(t1.img2,'') as img2, isNull(t1.img3,'') as img3, isNull(t1.img4,'') as img4 " + vbcrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_just1day as t1 " + vbcrlf
		sqlStr = sqlStr + "		Join db_academy.dbo.[tbl_diy_item] as t2 " + vbcrlf
		sqlStr = sqlStr + "			on t1.itemid=t2.itemid " + vbcrlf
		sqlStr = sqlStr + " where 1=1 " + addSql + vbcrlf
		sqlStr = sqlStr + " order by t1.justdate desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new Cjust1DayItem

				FItemList(i).Fidx		= rsACADEMYget("idx")
				FItemList(i).FJustDate		= rsACADEMYget("JustDate")
				FItemList(i).Fitemid		= rsACADEMYget("itemid")
				FItemList(i).Fitemname		= rsACADEMYget("itemname")
'				FItemList(i).FsmallImage	= "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/" + rsACADEMYget("smallImage")
				FItemList(i).FsmallImage	= imgFingers & "/diyItem/webimage/small/" + GetImageSubFolderByItemid(rsACADEMYget("itemid")) + "/"  + rsACADEMYget("smallImage")
				FItemList(i).ForgPrice		= rsACADEMYget("orgPrice")
				FItemList(i).FjustSalePrice	= rsACADEMYget("justSalePrice")
				FItemList(i).FsaleSuplyCash	= rsACADEMYget("saleSuplyCash")
				FItemList(i).FjustDesc		= db2html(rsACADEMYget("justDesc"))
				FItemList(i).Fsale_code		= rsACADEMYget("sale_code")
				FItemList(i).FlimitNo		= rsACADEMYget("limitNo")
				FItemList(i).FsellYn		= rsACADEMYget("sellYn")
				FItemList(i).Fimg1			= rsACADEMYget("img1")
				FItemList(i).Fimg2			= rsACADEMYget("img2")
				FItemList(i).Fimg3			= rsACADEMYget("img3")
				FItemList(i).Fimg4			= rsACADEMYget("img4")
				FItemList(i).Fregdate		= rsACADEMYget("regdate")

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close
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