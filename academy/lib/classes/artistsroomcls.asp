<%
Class CArtistsRoomImage
	public Fimgidx
	public Flecuserid
	public Fimagetype
	public Fimagevalue
	public Fimageicon

	public FOrgImagevalue
	public FOrgImageicon
	public Fimgcontents

	public function GetimageName()
		select case Fimagetype
			case "10"
				GetimageName = "메인이미지"
			case "20"
				GetimageName = "공방이미지"
			case "50"
				GetimageName = "작품소개"
			case else
				GetimageName = Fimagetype
		end select
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CArtistsRoomMajotLecItem
	public Flecuserid
	public Flec_idx

	public Flec_title

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CArtistsRoomItem
	public Flecuserid
	public Fsummarycontents
	public Fsummaryimage
	public Fcontents1
	public Fregdate
	public Ftitle
	
	public Fsocname
	public Fsocname_kor
	public Fstreetusing


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CArtistsRoom
	public FOneItem
	public FItemList()

	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount

	public FRectUserid

	function getImageFolderByType(imagetype,isIcon)
		select case imagetype
			case "10"
				getImageFolderByType = "mainimage"
			case "20"
				getImageFolderByType = "gongbang"
			case "50"
				if (isicon) then
					getImageFolderByType = "articon"
				else
					getImageFolderByType = "artimage"
				end if
			case else
				getImageFolderByType = ""
		end select
	end function

	public function GetImageList()
		dim sqlStr,i
		sqlStr = "select top 300 m.*"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_artistRoom_image m"
		sqlStr = sqlStr + " where m.lecuserid='" + FRectUserid + "'"
		sqlStr = sqlStr + " order by m.imagetype, m.imgidx"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CArtistsRoomImage

				FItemList(i).Fimgidx		= rsACADEMYget("imgidx")
				FItemList(i).Flecuserid		= rsACADEMYget("lecuserid")
				FItemList(i).Fimagetype		= rsACADEMYget("imagetype")
				FItemList(i).FOrgImagevalue	= rsACADEMYget("imagevalue")
				FItemList(i).FOrgimageicon		= rsACADEMYget("imageicon")

				FItemList(i).FImagevalue	= imgFingers & "/contents/artistsroom/" + getImageFolderByType(FItemList(i).Fimagetype,false) + "/" + FItemList(i).FOrgImagevalue
				FItemList(i).Fimageicon	= imgFingers & "/contents/artistsroom/" + getImageFolderByType(FItemList(i).Fimagetype,true) + "/" + FItemList(i).FOrgimageicon

				FItemList(i).Fimgcontents	= db2html(rsACADEMYget("imgcontents"))

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.Close
	end function


	public function GetMajorLec()
		dim sqlStr,i
		sqlStr = "select top 100 m.lecuserid, m.lec_idx, i.lec_title"
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_artistRoom_majorlec m,"
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_lec_item i"
		sqlStr = sqlStr + " where m.lecuserid='" + FRectUserid + "'"
		sqlStr = sqlStr + " and m.lec_idx=i.idx"
		sqlStr = sqlStr + " order by m.lec_idx desc"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			do until rsACADEMYget.eof
				set FItemList(i) = new CArtistsRoomMajotLecItem

				FItemList(i).Flecuserid      = rsACADEMYget("lecuserid")
				FItemList(i).Flec_idx		=  rsACADEMYget("lec_idx")
				FItemList(i).Flec_title   = db2html(rsACADEMYget("lec_title"))

				i=i+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.Close
	end function

	public function GetOneArtistRoom()
		dim sqlStr,i
		sqlStr = "select top 1 c.userid, c.socname, c.socname_kor, c.streetusing, "
		sqlStr = sqlStr + " a.lecuserid, a.summarycontents, a.summaryimage, a.contents1, a.regdate, a.title"
		sqlStr = sqlStr + " from [db3_common].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " , [db_academy].[dbo].tbl_artistRoom a "
		sqlStr = sqlStr + " where c.userid=a.lecuserid"
		sqlStr = sqlStr + " and c.userid='" + FRectUserid + "'"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FResultCount = rsACADEMYget.RecordCount
		FTotalCount = FResultCount

		if  not rsACADEMYget.EOF  then
			set FOneItem = new CArtistsRoomItem
			FOneItem.Flecuserid      = rsACADEMYget("lecuserid")
			FOneItem.Fsummarycontents= db2html(rsACADEMYget("summarycontents"))
			FOneItem.Fsummaryimage   = rsACADEMYget("summaryimage")
			FOneItem.Fcontents1      = db2html(rsACADEMYget("contents1"))
			FOneItem.Fregdate        = rsACADEMYget("regdate")
			FOneItem.Ftitle        = rsACADEMYget("title")
			
			FOneItem.Fsocname        = db2html(rsACADEMYget("socname"))
			FOneItem.Fsocname_kor    = db2html(rsACADEMYget("socname_kor"))
			FOneItem.Fstreetusing    = rsACADEMYget("streetusing")

			FOneItem.Fsummaryimage = imgFingers & "/contents/artistsroom/summaryimage/" + FOneItem.Fsummaryimage

		end if

		rsACADEMYget.Close
	end function

	public function GetArtistRoomList()
		dim sqlStr,i
		sqlStr = "select count(*) as cnt "
		sqlStr = sqlStr + " from [db3_common].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " , [db_academy].[dbo].tbl_artistRoom a "
		sqlStr = sqlStr + " where c.userdiv='14'"
		sqlStr = sqlStr + " and c.userid=a.lecuserid"

		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " c.userid, c.socname, c.socname_kor, c.streetusing, "
		sqlStr = sqlStr + " a.lecuserid, a.summarycontents, a.summaryimage, a.contents1, a.regdate, a.title"
		sqlStr = sqlStr + " from [db3_common].[dbo].tbl_user_c c "
		sqlStr = sqlStr + " , [db_academy].[dbo].tbl_artistRoom a "
		sqlStr = sqlStr + " where c.userdiv='14'"
		sqlStr = sqlStr + " and c.userid=a.lecuserid"
		sqlStr = sqlStr + " order by a.regdate desc "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsACADEMYget.EOF  then
			i = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FItemList(i) = new CArtistsRoomItem
				FItemList(i).Flecuserid      = rsACADEMYget("lecuserid")
				FItemList(i).Fsummarycontents= db2html(rsACADEMYget("summarycontents"))
				FItemList(i).Fsummaryimage   = rsACADEMYget("summaryimage")
				FItemList(i).Fcontents1      = db2html(rsACADEMYget("contents1"))
				FItemList(i).Fregdate        = rsACADEMYget("regdate")
				FItemList(i).Ftitle        	= rsACADEMYget("title")

				FItemList(i).Fsocname        = db2html(rsACADEMYget("socname"))
				FItemList(i).Fsocname_kor    = db2html(rsACADEMYget("socname_kor"))
				FItemList(i).Fstreetusing    = rsACADEMYget("streetusing")

				FItemList(i).Fsummaryimage = imgFingers & "/contents/artistsroom/summaryimage/" + FItemList(i).Fsummaryimage
				rsACADEMYget.MoveNext
				i = i + 1
			loop
		end if
		rsACADEMYget.Close
	end function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class

%>