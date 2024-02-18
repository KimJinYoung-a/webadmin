<%

Class DiaryItemsCls
	public fidx
	public fbrandid
	public fmainbrandimg
	public fbrandtext
	public fbrandmovieurl
	public fitemid
	public fisusing
	public fsortnum
	public fregdate
	
	public fpcmainbrandtextimg
	public fmomainbrandimg
	public Fleftright
End Class

Class DiaryCls

	Public FItemList()
	public FOneItem
	Public FRectIdx
	Public FRectIsusing
	Public FRectbrandid
	
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount


    public Sub fcontents_oneitem()
        dim sqlStr, sqlsearch

		if FRectIdx <> "" then
			sqlsearch = sqlsearch & " and idx="& FRectIdx &""
		end If
		
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " idx, brandid, mainbrandimg, brandtext, brandmovieurl, isusing, sortnum, regdate, pcmainbrandtextimg, momainbrandimg, leftright " & vbcrlf
		sqlStr = sqlStr & " ,STUFF((  " & vbcrlf
		sqlStr = sqlStr & " SELECT ',' + cast(k.itemid as varchar(10))  " & vbcrlf
		sqlStr = sqlStr & " FROM db_diary2010.dbo.tbl_diaryspecial_brand_itemid as k " & vbcrlf
		sqlStr = sqlStr & " WHERE (k.vidx = b.idx)  " & vbcrlf
		sqlStr = sqlStr & " FOR XML PATH ('')) ,1,1,'') AS itemid  " & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diaryspecial_brand as b" & vbcrlf
		sqlStr = sqlStr & " where idx=" + CStr(FRectIdx)

'        response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new DiaryItemsCls

        if Not rsget.Eof then

			FOneItem.fidx				= rsget("idx")
			FOneItem.fbrandid			= rsget("brandid")
			FOneItem.fmainbrandimg	= rsget("mainbrandimg")
			FOneItem.fbrandtext		= db2html(rsget("brandtext"))
			FOneItem.fbrandmovieurl	= db2html(rsget("brandmovieurl"))
			FOneItem.fitemid			= rsget("itemid")
			FOneItem.fisusing			= rsget("isusing")
			FOneItem.fsortnum			= rsget("sortnum")
			FOneItem.fregdate			= rsget("regdate")

			FOneItem.fpcmainbrandtextimg	= rsget("pcmainbrandtextimg")
			FOneItem.fmomainbrandimg		= rsget("momainbrandimg")
			FOneItem.fleftright			= rsget("leftright")
        end if
        rsget.Close
    end Sub

''tbl_diaryspecial_brand
'idx
'brandid
'mainbrandimg
'brandtext
'brandmovieurl
'itemid

	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(*) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diaryspecial_brand" & vbcrlf
        sqlStr = sqlStr & " where 1=1 " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and isusing = '"& FRectIsusing &"'" & vbcrlf
			end if

			if FRectbrandid <> "" then
				sqlStr = sqlStr & " and brandid = "& FRectbrandid &"" & vbcrlf
			end if

		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		'데이터 리스트
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx, brandid, mainbrandimg, brandtext, brandmovieurl, isusing, sortnum, regdate" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_diaryspecial_brand a" & vbcrlf
        sqlStr = sqlStr & " where 1=1 " & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and isusing = '"& FRectIsusing &"'" & vbcrlf
			end if

			if FRectbrandid <> "" then
				sqlStr = sqlStr & " and brandid = "& FRectbrandid &"" & vbcrlf
			end if

		sqlStr = sqlStr & " order by idx desc, sortnum asc" + vbcrlf

'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new DiaryItemsCls

				FItemList(i).fidx				= rsget("idx")
				FItemList(i).fbrandid			= rsget("brandid")
				FItemList(i).fmainbrandimg	= rsget("mainbrandimg")
				FItemList(i).fbrandtext		= db2html(rsget("brandtext"))
				FItemList(i).fbrandmovieurl	= db2html(rsget("brandmovieurl"))
'				FItemList(i).fitemid			= rsget("itemid")
				FItemList(i).fisusing			= rsget("isusing")
				FItemList(i).fsortnum			= rsget("sortnum")
				FItemList(i).fregdate			= rsget("regdate")

'				FItemList(i).FImageList	= "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("ListImage"))
'				FItemList(i).FImageList120	= "http://webimage.10x10.co.kr/image/list120/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & db2html(rsget("ListImage120"))
'				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("smallImage"))
'				FItemList(i).FImageicon1 = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon1image")
'				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon2image")



				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
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