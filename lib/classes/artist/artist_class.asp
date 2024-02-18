<%
Class cposcode_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fdesignerid
	public ffile1
	public ffile2
	public fposcode
	public fposname
	public fimagetype
	public fimagewidth
	public fimageheight
	public fisusing
	public fimagepath
	public flinkpath
	public fevt_code
	public fregdate
	public fsortNo
	public fmainHOT
	public fimagecount
	public fimage_order
	public fitemid
	public fimagepath2
	public fimagepath3
	public fsocname
	public fsocname_kor
	public fcomment
	public ficon2image
end class

class cposcode_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public FOneItem
	public Hotorder
	
	public FRectPoscode
	public FRectIsusing
	public FRectvaliddate
	public FRectIdx
	public FDesignerID
	public frecttoplimit
	public FGubun

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	'//admin/culturestation/imagemake_list.asp
	public sub fcontents_list()
		Dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_mainImageBanner " & vbcrlf
        sqlStr = sqlStr & " where gubun="&FGubun&" " & vbcrlf

		if FRectIsusing <> "" then
			sqlStr = sqlStr & " and isusing = '"& FRectIsusing &"'" & vbcrlf		
		end if

		if FDesignerID <> "" then
			sqlStr = sqlStr & " and designerid = '"& FDesignerID &"'" & vbcrlf		
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, imagepath, linkpath, isusing, regdate, image_order " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_mainImageBanner " & vbcrlf
		sqlStr = sqlStr & " where gubun="&FGubun&" " & vbcrlf
       	
       	If FGubun = 1 Then
			sqlStr = sqlStr & " order by image_order asc" + vbcrlf
		End If

       	If FGubun = 2 Then
			sqlStr = sqlStr & " order by idx desc" + vbcrlf
		End If

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cposcode_oneitem
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")		
				FItemList(i).fimage_order = rsget("image_order")													
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

'//admin/culturestation/imagemake_contents.asp
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " idx, imagepath, linkpath, isusing, regdate, image_order " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_mainImageBanner " & vbcrlf
        sqlStr = sqlStr & " where gubun="&FGubun&" " & vbcrlf
        sqlStr = sqlStr & " and idx = "& FRectIdx&""

        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cposcode_oneitem

        if Not rsget.Eof then
    		FOneItem.fidx = rsget("idx")
    		FOneItem.fimagepath = rsget("imagepath")
    		FOneItem.flinkpath = rsget("linkpath")
    		FOneItem.fisusing = rsget("isusing")
    		FOneItem.fregdate = rsget("regdate")
    		FOneItem.fimage_order = rsget("image_order")
        end if
        rsget.Close
    end Sub

	public sub FArtistBrandList()
		Dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_brand where 1=1  " & vbcrlf

		if FRectIsusing <> "" then
			sqlStr = sqlStr & " and isusing = '"& FRectIsusing &"'" & vbcrlf		
		end if

		if FDesignerID <> "" then
			sqlStr = sqlStr & " and designerid = '"& FDesignerID &"'" & vbcrlf		
		end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, designerid, file1, file2, isusing,regdate, sortNo, mainHOT " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_brand " & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf

		if FRectIsusing <> "" then
			sqlStr = sqlStr & " and isusing = '"& FRectIsusing &"'" & vbcrlf		
		end if

		if FDesignerID <> "" then
			sqlStr = sqlStr & " and designerid = '"& FDesignerID &"'" & vbcrlf		
		end if

		if Hotorder = "Y" then
			sqlStr = sqlStr & " order by mainHOT desc, sortNo asc, idx desc" + vbcrlf
		else
			sqlStr = sqlStr & " order by idx desc" + vbcrlf
		end if

		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cposcode_oneitem
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fdesignerid = rsget("designerid")
				FItemList(i).ffile1 = rsget("file1")
				FItemList(i).ffile2 = rsget("file2")
				FItemList(i).fisusing = rsget("isusing")		
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fsortNo = rsget("sortNO")
				FItemList(i).fmainHOT = rsget("mainHOT")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

    public Sub FArtistBrand_oneitem()
        dim sqlStr
        sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " B.idx, B.designerid, B.file1, B.file2, B.isusing, B.regdate, C.socname, C.socname_kor " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_brand as B " & vbcrlf
		sqlStr = sqlStr & " Inner Join db_user.dbo.tbl_user_c as C on B.designerid = C.userid " & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and B.idx = "& FRectIdx&""
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cposcode_oneitem

        if Not rsget.Eof then
    		FOneItem.fidx = rsget("idx")
    		FOneItem.fdesignerid = rsget("designerid")
    		FOneItem.ffile1 = rsget("file1")
    		FOneItem.ffile2 = rsget("file2")
    		FOneItem.fisusing = rsget("isusing")
    		FOneItem.fregdate = rsget("regdate")
    		FOneItem.fsocname = rsget("socname")
    		FOneItem.fsocname_kor = rsget("socname_kor")
        end if
        rsget.Close
    end Sub

    public Sub FArtistMonthItemList
		Dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_shop_MonthItem where isusing='Y' " & vbcrlf

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " M.idx, M.itemid, M.comment, M.sortNo, M.isusing, M.regdate, i.icon2image " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_shop_MonthItem as M " & vbcrlf
		sqlStr = sqlStr & " Inner Join db_item.dbo.tbl_item as i on M.itemid = i.itemid " & vbcrlf
        sqlStr = sqlStr & " where M.isusing='Y' " & vbcrlf
        sqlStr = sqlStr & " order by M.idx desc, M.sortNo asc " & vbcrlf

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cposcode_oneitem
				FItemList(i).fidx 			= rsget("idx")
				FItemList(i).fitemid 		= rsget("itemid")
				FItemList(i).fcomment 		= rsget("comment")
				FItemList(i).fsortNo 		= rsget("sortNo")
				FItemList(i).fisusing 		= rsget("isusing")
				FItemList(i).ficon2image	= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2image")
				FItemList(i).fregdate 		= rsget("regdate")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close

	end Sub

    public Sub FArtistMonthItem_one()
        dim sqlStr
        sqlStr = "select top 1 " & vbcrlf
		sqlStr = sqlStr & " idx, itemid, comment, sortNo, isusing, regdate " & vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_shop_MonthItem " & vbcrlf
        sqlStr = sqlStr & " where idx = "& FRectIdx&""
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new cposcode_oneitem

        if Not rsget.Eof then
			FOneItem.fidx 			= rsget("idx")
			FOneItem.fitemid 		= rsget("itemid")
			FOneItem.fcomment 		= rsget("comment")
			FOneItem.fsortNo 		= rsget("sortNo")
			FOneItem.fisusing 		= rsget("isusing")		
			FOneItem.fregdate 		= rsget("regdate")
        end if
        rsget.Close
    end Sub


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
%>