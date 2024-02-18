<%
'#######################################################
'	History	:  2009.04.09 한용민 2008프론트 이동/수정
'	Description : artist gallery
'#######################################################
%>
<%
Class Cinquiry_oneitem

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
	public fartist_name
	public fuser_name
	public faddress
	public ftel
	public fhp
	public fmail
	public flicense
	public fhomepage
	public fuser_info
	public fsell_count
	public fon_off_isusing
	public fitem_info
	public ffile1
	public fartist_idx
	public ftag
	public fregdate
	public fblog
	public fisusing
	public fwhyrecommend
	public fuserid

end Class

Class Cinquiry_list
	public FItemList()
	public FTotalCount
	public FResultCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount
	public FPageCount
	public foneitem

	public frectidx
	public frectartist_idx
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub
	
	''// 아티스트추천리스트 ///admin/artist/artist_recommend.asp
	public sub frecommend_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(artist_idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_recommend" + vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y'" + vbcrlf
							
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " artist_idx , artist_name, tag,regdate, homepage, blog" + vbcrlf
		sqlStr = sqlStr & " , isusing, whyrecommend , userid" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_recommend" + vbcrlf
		sqlStr = sqlStr & " where isusing = 'Y'" + vbcrlf
		sqlStr = sqlStr & " order by artist_idx desc" + vbcrlf
		
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
				set FItemList(i) = new cinquiry_oneitem

	
				FItemList(i).fartist_idx = rsget("artist_idx")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fwhyrecommend = rsget("whyrecommend")				
				FItemList(i).fartist_name = rsget("artist_name")
				FItemList(i).ftag = rsget("tag")
				FItemList(i).fhomepage = rsget("homepage")
				FItemList(i).fblog = rsget("blog")
				FItemList(i).fuserid = rsget("userid")
																													
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//입점문의 리스트 '///admin/artist/artist_inquiry.asp
	public sub finquiry_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_inquiry" + vbcrlf
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " idx,artist_name,user_name,address,tel,hp,mail,license,homepage," + vbcrlf
		sqlStr = sqlStr & " user_info,sell_count,on_off_isusing,item_info,file1" + vbcrlf		
		sqlStr = sqlStr & " from db_contents.dbo.tbl_artist_inquiry" + vbcrlf				
		sqlStr = sqlStr & " order by idx Desc" + vbcrlf

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
				set FItemList(i) = new cinquiry_oneitem
	
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fartist_name = db2html(rsget("artist_name"))
				FItemList(i).fuser_name = rsget("user_name")
				FItemList(i).faddress = db2html(rsget("address"))
				FItemList(i).ftel = db2html(rsget("tel"))
				FItemList(i).fhp = db2html(rsget("hp"))				
				FItemList(i).fmail = db2html(rsget("mail"))
				FItemList(i).flicense = db2html(rsget("license"))				
				FItemList(i).fhomepage = db2html(rsget("homepage"))
				FItemList(i).fuser_info = db2html(rsget("user_info"))				
				FItemList(i).fsell_count = rsget("sell_count")
				FItemList(i).fon_off_isusing = db2html(rsget("on_off_isusing"))				
				FItemList(i).fitem_info = db2html(rsget("item_info"))	
				FItemList(i).ffile1 = db2html(rsget("file1"))		
																							
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//아티스트추천 상세보기 '///admin/artist/artist_recommendview.asp
	public sub frecommend_oneitem()
		dim sqlStr	

		'데이터 리스트 
		sqlStr = "exec [db_contents].[dbo].[ten_artist_recommend_one] "&frectartist_idx&" "

		'Response.write sqlStr &"<br>"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic  		
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		if  not rsget.EOF  then
				set foneitem = new cinquiry_oneitem

				foneitem.fartist_idx = rsget("artist_idx")
				foneitem.fregdate = rsget("regdate")
				foneitem.fisusing = rsget("isusing")
				foneitem.fwhyrecommend = rsget("whyrecommend")				
				foneitem.fartist_name = rsget("artist_name")
				foneitem.ftag = rsget("tag")
				foneitem.fhomepage = rsget("homepage")
				foneitem.fblog = rsget("blog")
				foneitem.fuserid = rsget("userid")
																								
		end if
		rsget.Close
	end sub

	'//아티스트입점문의 '///admin/artist/artist_inquiry.asp
	public sub finquiry_oneitem()
		dim sqlStr	

		'데이터 리스트 
		sqlStr = "exec [db_contents].[dbo].[ten_artist_inquiry] "&frectidx&" "

		'Response.write sqlStr &"<br>"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic  
		rsget.Open sqlStr,dbget,1

		ftotalcount = rsget.recordcount

		if  not rsget.EOF  then
				set foneitem = new cinquiry_oneitem
	
				foneitem.fidx = rsget("idx")
				foneitem.fartist_name = db2html(rsget("artist_name"))
				foneitem.fuser_name = rsget("user_name")
				foneitem.faddress = db2html(rsget("address"))
				foneitem.ftel = db2html(rsget("tel"))
				foneitem.fhp = db2html(rsget("hp"))				
				foneitem.fmail = db2html(rsget("mail"))
				foneitem.flicense = db2html(rsget("license"))				
				foneitem.fhomepage = db2html(rsget("homepage"))
				foneitem.fuser_info = db2html(rsget("user_info"))				
				foneitem.fsell_count = rsget("sell_count")
				foneitem.fon_off_isusing = db2html(rsget("on_off_isusing"))				
				foneitem.fitem_info = db2html(rsget("item_info"))	
				foneitem.ffile1 = db2html(rsget("file1"))		
																							
		end if
		rsget.Close
	end sub

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

Class CGalleryItem
	public Fgal_sn
	public Fdesignerid
	public Fgal_div
	public Fgal_imgorg
	public Fgal_img400
	public Fgal_regdate
	public Fgal_isusing
	public Fgal_desc
	public Fgal_sortNo
	public Fsocname
	public Fsocname_kor
	public fidx
	public fgubun
	public fitemid
	public fregdate
	public fisusing
	public fitemname
	public fsellyn
	public fLimitYn
	public fLimitNo
	public fLimitSold
	public fdanjongyn
	public fsellcash
	public fbuycash
	public FImageMain
	public FImageList
	public FImageSmall
	public FImageBasic
	public Ficon1Image
	public flistimage120
	
	public Function getGalDivName()
		Select Case Fgal_div
			Case "W"
				getGalDivName = "Work"
			Case "D"
				getGalDivName = "Drawing"
			Case "P"
				getGalDivName = "Photo"
			Case Else
				getGalDivName = "N/A"
		End Select
	end Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class

Class CGallery
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

	public FRectGal_sn
	public FRectGal_div
	public FRectDesignerId
	public FRectIsusing
	public Frectcatecode
	public FrectstandardCateCode
	public Frectmduserid
	Public Frectbrandgubun
	
	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage =1
		FPageSize = 15
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'##### 아티스트 갤러리 상품목록 접수 #####  '///admin/artirst/artist_gallery.asp
	public sub getgalleryitem()
		dim sqlStr , i

		sqlStr = "exec [db_contents].[dbo].[ten_artist_banner] "
		
		'Response.write sqlStr &"<br>"

        rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic  
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		ftotalcount = rsget.recordcount
		redim preserve FItemList(ftotalcount)

		i=0
		if  not rsget.EOF  then
			do until rsget.EOF
				set FItemList(i) = new CGalleryItem

				FItemList(i).fidx		= rsget("idx")
				FItemList(i).fgal_div		= rsget("gal_div")
				FItemList(i).fgubun	= rsget("gubun")
				FItemList(i).fitemid	= rsget("itemid")
				FItemList(i).fregdate	= rsget("regdate")
				FItemList(i).fisusing	= rsget("isusing")				
				FItemList(i).fitemname	= db2html(rsget("itemname"))
				FItemList(i).fSellYn	= rsget("SellYn")
				FItemList(i).fLimitYn	= rsget("LimitYn")
				FItemList(i).fLimitNo	= rsget("LimitNo")
				FItemList(i).fLimitSold	= rsget("LimitSold")
				FItemList(i).fdanjongyn	= rsget("danjongyn")
				FItemList(i).fsellcash	= rsget("sellcash")
				FItemList(i).fbuycash	= rsget("buycash")
				FItemList(i).FImageMain = "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FItemList(i).FImageList = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("smallimage")
				FItemList(i).FImageBasic = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage")
				FItemList(i).Ficon1Image = "http://webimage.10x10.co.kr/image/icon1/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon1image")
				FItemList(i).flistimage120 = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage120")
				
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'##### 아티스트 갤러리 목록 접수 #####  '///admin/artirst/artist_gallery.asp
	public Sub GetGalleryList()
		dim SQL, AddSQL, i, strTemp

		If Frectcatecode <> "" Then
			AddSQL = AddSQL & " and t2.catecode = '"&Frectcatecode&"' " 
		End If
		If Frectstandardcatecode <> "" Then
			AddSQL = AddSQL & " and t2.standardcatecode = '"&Frectstandardcatecode&"' " 
		End If
		If Frectmduserid <> "" Then
			AddSQL = AddSQL & " and t2.mduserid = '"&Frectmduserid&"' " 
		End If
		If frectbrandgubun <> "" Then
			AddSQL = AddSQL & " and sm.brandgubun = '"&frectbrandgubun&"' " 
		End If
		
		if FRectGal_div<>"" then
			AddSQL = AddSQL & " and t1.gal_div='" & FRectGal_div & "' "
		end if
		if FRectDesignerId<>"" then
			AddSQL = AddSQL & " and t1.designerid='" & FRectDesignerId & "' "
		end if
		if FRectIsusing<>"" then
			AddSQL = AddSQL & " and t1.gal_isusing='" & FRectIsusing & "' "
		end if

		'// 개수 파악 //
		SQL =	"Select count(gal_sn), CEILING(CAST(Count(gal_sn) AS FLOAT)/" & FPageSize & ")"
		SQL = SQL & " From db_contents.dbo.tbl_artist_gallery as t1"
		SQL = SQL & " Join db_user.dbo.tbl_user_c as t2"
		SQL = SQL & " 	on t1.designerid=t2.userid"
		SQL = SQL & " left join db_brand.dbo.tbl_street_manager sm"		
		SQL = SQL & " 	on t1.designerid=sm.makerid"
		SQL = SQL & " Where 1=1 " & AddSQL
		
		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL = "Select top " & CStr(FPageSize*FCurrPage) & " t1.*, t2.socname, t2.socname_kor"
		SQL = SQL & " From db_contents.dbo.tbl_artist_gallery as t1"
		SQL = SQL & " Join db_user.dbo.tbl_user_c as t2"
		SQL = SQL & " 	on t1.designerid=t2.userid"
		SQL = SQL & " left join db_brand.dbo.tbl_street_manager sm"		
		SQL = SQL & " 	on t1.designerid=sm.makerid"		
		SQL = SQL & " Where 1=1 " & AddSQL
		SQL = SQL & " Order by gal_sn desc"
				
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CGalleryItem

				FItemList(i).Fgal_sn		= rsget("gal_sn")
				FItemList(i).Fdesignerid	= rsget("designerid")
				FItemList(i).Fgal_div		= rsget("gal_div")
				FItemList(i).Fgal_img400	= staticImgUrl & "/contents/artistGallery/" & rsget("gal_img400")
				FItemList(i).Fgal_regdate	= rsget("gal_regdate")
				FItemList(i).Fgal_isusing	= rsget("gal_isusing")
				FItemList(i).Fgal_desc		= rsget("gal_desc")
				FItemList(i).Fgal_sortNo	= rsget("gal_sortNo")
				FItemList(i).Fsocname		= rsget("socname")
				FItemList(i).Fsocname_kor	= rsget("socname_kor")

				rsget.moveNext
				i=i+1
			loop
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


	'##### 부서 내용 접수 #####
	public Sub GetGalleryInfo()
		dim SQL

		'// 목록 접수 //
		SQL =	"Select t1.*, t2.socname, t2.socname_kor " &_
				"From db_contents.dbo.tbl_artist_gallery as t1 " &_
				"	Join db_user.dbo.tbl_user_c as t2 " &_
				"		on t1.designerid=t2.userid " &_
				"Where gal_sn=" & FRectgal_sn
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then

			FResultCount = 1
			redim preserve FItemList(1)
			set FItemList(1) = new CGalleryItem

			FItemList(1).Fdesignerid	= rsget("designerid")
			FItemList(1).Fgal_div		= rsget("gal_div")
			FItemList(1).Fgal_imgorg	= staticImgUrl & "/contents/artistGallery/" & rsget("gal_imgorg")
			FItemList(1).Fgal_img400	= staticImgUrl & "/contents/artistGallery/" & rsget("gal_img400")
			FItemList(1).Fgal_regdate	= rsget("gal_regdate")
			FItemList(1).Fgal_isusing	= rsget("gal_isusing")
			FItemList(1).Fgal_desc		= rsget("gal_desc")
			FItemList(1).Fgal_sortNo	= rsget("gal_sortNo")
			FItemList(1).Fsocname		= rsget("socname")
			FItemList(1).Fsocname_kor	= rsget("socname_kor")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub
end Class

'// 브랜드 선택상자(데이터에 있을때만 접수) //
Sub DrawSelectBoxUseBrand(byval selectBoxName,selectedId)
	dim tmp_str,query1

	tmp_str = "<select name='" & selectBoxName & "'>" & vbCrLf
	tmp_str = tmp_str & "<option value=''"
	if selectedId="" then tmp_str = tmp_str & " selected"
	tmp_str = tmp_str & ">선택</option value=''>" & vbCrLf
   
   query1 = "Select distinct t1.designerid, t2.socname, t2.socname_kor, t1.regdate " &_
   			"From db_contents.dbo.tbl_artist_brand as t1" &_
			"	Join db_user.dbo.tbl_user_c as t2 " &_
			"		on t1.designerid=t2.userid " &_
   			" order by t1.regdate desc "
   rsget.Open query1,dbget,1

	if Not(rsget.EOF or rsget.BOF) then
		rsget.Movefirst

		do until rsget.EOF
			tmp_str = tmp_str & "<option value='" & rsget("designerid") & "' "
			if Cstr(selectedId) = Cstr(rsget("designerid")) then tmp_str = tmp_str & " selected"
			tmp_str = tmp_str & ">" & rsget("socname") & " (" & rsget("socname_kor") & ")</option>"
			rsget.MoveNext
		loop
	end if
	rsget.close
   
	tmp_str = tmp_str & "</select>"

	response.write(tmp_str)
End Sub
%>