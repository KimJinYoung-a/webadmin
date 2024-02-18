<%
'#############################################
' Description : DIY SHOP 클래스
' History : 2016.04.11 유태욱 생성
'#############################################
Class CdiyQnaItem
    public Fidx
    public Fqna
	public Fitemid
	Public Fuserid
	Public Fuserlevel
	Public Freplyuserid
	Public Ftitle
	public Fsmsok
	public Fsmsnum
	public Fregdate
	public Femailok
	public Fcomment
	public Freply_num
	Public Freply_depth
	public Freply_group_idx
	Public Fdevice
	Public Fisusing
	public FanswerYN
	public Fmakerid
	public FItemName
end Class

Class DiyItemCls

	Private SUB Class_initialize()
		FListDiv = "list"
		FResultCount = 0
		FTotalCount = 0
		FPageSize = 16
		FCurrPage = 1
'		FPageSize = 30
	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FOneItem
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FListDiv			'카테고리/검색 구분용
	dim FRectSearchTxt		'검색어
	dim FRectSearchItemDiv	'카테고리 검색 범위(D:기본카테고리,A:추가 카테고리)
	dim FRectSortMethod		'정렬방식
	dim FRectSearchFlag 	'세일&쿠폰상품 검색용
	dim FRectMakerid		'업체 아이디
	dim FRectCdL			'대카테고리코드
	dim FRectCdM			'중카테고리코드
	dim FRectCdS			'소카테고리코드
	dim FSellScope			'판매가능 상품검색 여부
	dim FRectitemid
	dim FRectuserid
	dim FRectgroupidx
	dim FRectmode
	dim FRectqnaidx

	dim FminPrice			'가격최소값
	dim FmaxPrice			'가격최대값
	dim FSalePercentHigh	'할인율 최대값
	dim FSalePercentLow		'할인율 최소값
	dim Fuserid			'리스트 마이위시 체크용 userid
	dim FRectcouponidx


	'####### 상품 검색 ######
	PUBLIC SUB getdiySearchList()

		dim strSQL, strTblNm, iLoopCnt
		IF FSalePercentHigh ="" THEN FSalePercentHigh = 0
		IF FSalePercentLow ="" THEN FSalePercentLow = 0

		Select Case FListDiv
			Case "list"
				strTblNm = "sp_academy_DIYItemList"
			Case "search"
				strTblNm = "sp_academy_DIYSearchList"
		End Select

		'// 결과 카운트
		strSQL =" EXECUTE db_academy.dbo." & strTblNm & "_Tcnt " &_
					" @cdL= '" &FRectCdL&"'" &_
					" ,@cdM='" &FRectCdM&"' " &_
					" ,@cdS='" &FRectCdS&"' " &_ 
					" ,@Makerid ='" &FRectMakerid& "' " &_
					" ,@PgSize='" &FPageSize&"' " &_
					" ,@CurrPg='" &FCurrPage&"' " &_
					" ,@WhereMtd='" &FRectSearchFlag&"' " &_
					" ,@SortMtd='" &FRectSortMethod&"' "&_
					" ,@SalePercentHigh = "&FSalePercentHigh&_
					" ,@SalePercentLow = "&FSalePercentLow&_
					" ,@SearchItemDiv = '"&FRectSearchItemDiv&"'"  &_
					" ,@SellScope = '"&FSellScope&"'" &_
					" ,@searchTxt = '"&FRectSearchTxt&"'"  &_
					" ,@couponidx = '"&FRectcouponidx&"'"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open strSQL, dbACADEMYget
'rw strSQL
		IF not rsACADEMYget.eof then
			FTotalCount= rsACADEMYget("Totalcnt")
			FTotalPage = rsACADEMYget("TotalPage")
		End if
		
		rsACADEMYget.close

		'// 내용 접수
		strSQL =" EXECUTE db_academy.dbo." & strTblNm &_
					" @cdL= '" &FRectCdL&"'" &_
					" ,@cdM='" &FRectCdM&"' " &_
					" ,@cdS='" &FRectCdS&"' " &_ 
					" ,@Makerid ='" &FRectMakerid& "' " &_
					" ,@PgSize='" &FPageSize&"' " &_
					" ,@CurrPg='" &FCurrPage&"' " &_
					" ,@WhereMtd='" &FRectSearchFlag&"' " &_
					" ,@SortMtd='" &FRectSortMethod&"' "&_
					" ,@SalePercentHigh = "&FSalePercentHigh&_
					" ,@SalePercentLow = "&FSalePercentLow&_
					" ,@SearchItemDiv = '"&FRectSearchItemDiv&"'" &_
					" ,@SellScope = '"&FSellScope&"'" &_
					" ,@searchTxt = '"&FRectSearchTxt&"'" &_
					" ,@couponidx = '"&FRectcouponidx&"'"  &_
					" ,@userid = '"&Frectuserid&"'"
'	response.write strSQL
'	response.end
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.PageSize=FPageSize
		rsACADEMYget.Open strSQL, dbACADEMYget
		'response.write strSQL
		FResultCount = rsACADEMYget.RecordCount-((FCurrPage-1)*FPageSize)
		
		if (FResultCount<1) then FResultCount=0
		
		redim FItemList(FResultCount)

		iLoopCnt=0
		if  not rsACADEMYget.EOF  then
			rsACADEMYget.absolutePage=FCurrPage
			do until rsACADEMYget.eof
				set FItemList(iLoopCnt) = new CCategoryPrdItem
				
				FItemList(iLoopCnt).FCdL			= rsACADEMYget("code_large")
				FItemList(iLoopCnt).FCdM			= rsACADEMYget("code_mid")
				FItemList(iLoopCnt).FCdS			= rsACADEMYget("code_small")
				FItemList(iLoopCnt).FItemid			= rsACADEMYget("Itemid")
				FItemList(iLoopCnt).FItemName		= db2html(rsACADEMYget("ItemName"))
				FItemList(iLoopCnt).FSellCash		= rsACADEMYget("SellCash")
				FItemList(iLoopCnt).FOrgPrice		= rsACADEMYget("OrgPrice")
				FItemList(iLoopCnt).FMakerId		= rsACADEMYget("MakerId")
				FItemList(iLoopCnt).FBrandName		= db2html(rsACADEMYget("BrandName"))
				FItemList(iLoopCnt).Fkeywords		= db2html(rsACADEMYget("keywords"))
				FItemList(iLoopCnt).FImageList		= fingersImgUrl & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" &db2html(rsACADEMYget("ListImage"))
				FItemList(iLoopCnt).FImageList120	= fingersImgUrl & "/diyitem/webimage/list120/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & db2html(rsACADEMYget("ListImage120"))
				FItemList(iLoopCnt).FImageSmall		= fingersImgUrl & "/diyitem/webimage/small/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" &db2html(rsACADEMYget("smallImage"))
				FItemList(iLoopCnt).FImageicon1 	= fingersImgUrl & "/diyitem/webimage/icon1/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & rsACADEMYget("icon1image")
				FItemList(iLoopCnt).FImageicon2 	= fingersImgUrl & "/diyitem/webimage/icon2/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & rsACADEMYget("icon2image")
				
				FItemList(iLoopCnt).Fchkfav			= rsACADEMYget("chkfav")	
				FItemList(iLoopCnt).FSellyn			= rsACADEMYget("sellYn")
				FItemList(iLoopCnt).FSaleyn			= rsACADEMYget("SaleYn")
				FItemList(iLoopCnt).FLimityn		= rsACADEMYget("LimitYn")
				FItemList(iLoopCnt).FRegdate		= rsACADEMYget("regdate")
				FItemList(iLoopCnt).FItemcouponyn	= rsACADEMYget("itemcouponYn")
				FItemList(iLoopCnt).FItemcouponvalue= rsACADEMYget("itemCouponValue")
				FItemList(iLoopCnt).FItemcoupontype	= rsACADEMYget("itemCouponType")
				FItemList(iLoopCnt).FEvalcnt		= rsACADEMYget("evalCnt")
				FItemList(iLoopCnt).FItemScore		= rsACADEMYget("itemScore")
				
				FItemList(iLoopCnt).FCurrItemCouponIdx		= rsACADEMYget("curritemcouponidx")
				
				FItemList(iLoopCnt).Flecturer_img		= rsACADEMYget("lecturer_img")
				if FItemList(iLoopCnt).Flecturer_img <> "" then
					FItemList(iLoopCnt).Flecturer_img		= webImgUrl & "/image/brandlogo/t1_"& rsACADEMYget("lecturer_img")
				end if

				iLoopCnt=iLoopCnt+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close

	End SUB

	PUBLIC FUNCTION getSearchGroupCount()
		dim strSQL, strTblNm, iLoopCnt

		'// 결과 카운트
		strSQL =" EXECUTE [dbo].[sp_academy_DIYSearchGroupCnt] '"&FRectSearchTxt&"', '"&FRectMakerid&"', '"&FRectSearchFlag&"' "
		'rw strSQL
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open strSQL, dbACADEMYget

		IF not rsACADEMYget.eof then
			getSearchGroupCount = rsACADEMYget.getRows()
		End if

		rsACADEMYget.Close

	End FUNCTION

	PUBLIC SUB getSearchTotalCount()

		dim strSQL, strTblNm, iLoopCnt

		Select Case FListDiv
			Case "list"
				strTblNm = "sp_academy_DIYItemList"
			Case "search"
				strTblNm = "sp_academy_DIYSearchList"
		End Select

		'// 결과 카운트
		strSQL =" EXECUTE db_academy.dbo." & strTblNm & "_Tcnt " &_
					" @cdL= '" &FRectCdL&"'" &_
					" ,@cdM='" &FRectCdM&"' " &_
					" ,@cdS='" &FRectCdS&"' " &_ 
					" ,@Makerid ='" &FRectMakerid& "' " &_
					" ,@PgSize='1' " &_
					" ,@CurrPg='1' " &_
					" ,@WhereMtd='" &FRectSearchFlag&"' " &_
					" ,@SortMtd='" &FRectSortMethod&"' "&_
					" ,@SalePercentHigh = '' " &_
					" ,@SalePercentLow = '' " &_
					" ,@SearchItemDiv = '"&FRectSearchItemDiv&"'"  &_
					" ,@SellScope = '"&FSellScope&"'" &_
					" ,@searchTxt = '"&FRectSearchTxt&"'"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.Open strSQL, dbACADEMYget

		IF not rsACADEMYget.eof then
			FTotalCount= rsACADEMYget("Totalcnt")
		End if

		rsACADEMYget.Close

	End SUB

	''---------------------------------------------------------------------------------
	'//corner/lectureDetail.asp 작가프로필..판매작품..2016-08-13 김진영 작성
	Public Sub getCornerDiyList()
		Dim sqlStr, i, addSql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(i.itemid) as cnt, CEILING(CAST(Count(i.itemid) AS FLOAT)/" & FPageSize & ") as totPg "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_diy_item] as i "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as d on i.dispcate1 = d.catecode and d.useyn = 'Y' "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_diy_item_Contents as k on i.itemid = k.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as c on i.makerid=c.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.sellyn = 'Y' "
		sqlStr = sqlStr & " and ((i.limityn = 'N') or ((i.limityn = 'Y') and (i.limitno - i.limitsold > 0))) "
		sqlStr = sqlStr & " and i.makerid = '"&FRectMakerid&"' "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
			FTotalPage = rsACADEMYget("totPg")
		rsACADEMYget.Close

		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " i.itemid, i.itemname, i.sellcash, i.orgPrice, i.Makerid, i.brandName, i.ListImage, i.ListImage120, i.SmallImage, icon1image, i.icon2image "
		sqlStr = sqlStr & " ,i.Sellyn, i.SaleYn, i.LimitYn, i.LimitNo, i.LimitSold, i.RegDate, i.ItemCouponYn, i.ItemCouponValue, i.ItemCouponType, i.evalCnt, i.ItemScore, i.itemDiv, optioncnt "
		sqlStr = sqlStr & " , isnull(c.newImage_profile, '') as newImage_profile, c.lecturer_name, k.keywords "
		sqlStr = sqlStr & " , (select TOP 1 userid from [db_academy].dbo.tbl_diy_myfavorite where itemid = i.itemid and userid = '"& Fuserid &"') as chkfav "
		sqlStr = sqlStr & " FROM [db_academy].[dbo].[tbl_diy_item] as i "
		sqlStr = sqlStr & " JOIN db_academy.[dbo].[tbl_display_cate_Academy] as d on i.dispcate1 = d.catecode and d.useyn = 'Y' "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_diy_item_Contents as k on i.itemid = k.itemid "
		sqlStr = sqlStr & " LEFT JOIN db_academy.dbo.tbl_corner_good as c on i.makerid=c.lecturer_id "
		sqlStr = sqlStr & " WHERE 1 = 1 "
		sqlStr = sqlStr & " and i.isusing = 'Y' "
		sqlStr = sqlStr & " and i.sellyn = 'Y' "
		sqlStr = sqlStr & " and ((i.limityn = 'N') or ((i.limityn = 'Y') and (i.limitno - i.limitsold > 0))) "
		sqlStr = sqlStr & " and i.makerid = '"&FRectMakerid&"' "
	    sqlStr = sqlStr & " ORDER BY i.itemid DESC "
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim FItemList(FResultCount)
		i = 0
		If not rsACADEMYget.EOF Then
			rsACADEMYget.absolutepage = FCurrPage
			Do until rsACADEMYget.EOF
				set FItemList(i) = new CCategoryPrdItem
					FItemList(i).FItemid			= rsACADEMYget("Itemid")
					FItemList(i).FItemName			= db2html(rsACADEMYget("ItemName"))
					FItemList(i).FSellCash			= rsACADEMYget("SellCash")
					FItemList(i).FOrgPrice			= rsACADEMYget("OrgPrice")
					FItemList(i).FMakerId			= rsACADEMYget("MakerId")
					FItemList(i).FBrandName			= db2html(rsACADEMYget("BrandName"))
					FItemList(i).FImageList			= fingersImgUrl & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" &db2html(rsACADEMYget("ListImage"))
					FItemList(i).FImageList120		= fingersImgUrl & "/diyitem/webimage/list120/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & db2html(rsACADEMYget("ListImage120"))
					FItemList(i).FImageSmall		= fingersImgUrl & "/diyitem/webimage/small/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" &db2html(rsACADEMYget("smallImage"))
					FItemList(i).FImageicon1 		= fingersImgUrl & "/diyitem/webimage/icon1/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & rsACADEMYget("icon1image")
					FItemList(i).FImageicon2 		= fingersImgUrl & "/diyitem/webimage/icon2/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & rsACADEMYget("icon2image")
					FItemList(i).FSellyn			= rsACADEMYget("sellYn")
					FItemList(i).FSaleyn			= rsACADEMYget("SaleYn")
					FItemList(i).FLimityn			= rsACADEMYget("LimitYn")
					FItemList(i).FRegdate			= rsACADEMYget("regdate")
					FItemList(i).FItemcouponyn		= rsACADEMYget("itemcouponYn")
					FItemList(i).FItemcouponvalue	= rsACADEMYget("itemCouponValue")
					FItemList(i).FItemcoupontype	= rsACADEMYget("itemCouponType")
					FItemList(i).FEvalcnt			= rsACADEMYget("evalCnt")
					FItemList(i).FItemScore			= rsACADEMYget("itemScore")
					FItemList(i).FOptioncnt			= rsACADEMYget("optioncnt")
					FItemList(i).Flecturer_name		= rsACADEMYget("lecturer_name")
				If rsACADEMYget("newImage_profile") <> "" Then
					FItemList(i).Flecturer_img		= fingersImgUrl & "/corner/newImage_profile/thumbimg3/t3_" & rsACADEMYget("newImage_profile")
				Else
					FItemList(i).Flecturer_img		= ""
				End If
					FItemList(i).Fkeywords			= db2html(rsACADEMYget("keywords"))
					FItemList(i).FChkfav			= rsACADEMYget("chkfav")
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    End Sub


	''---------------------------------------------------------------------------------
	'내 관심 상품 리스트
	public function GetMyDiyItemList()
		dim sqlStr, i, sqlsearch

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_diy_item as i"
		sqlStr = sqlStr & " join [db_academy].dbo.tbl_diy_myfavorite as f"
		sqlStr = sqlStr & " 	on i.itemid=f.itemid"
		sqlStr = sqlStr & " left Join db_academy.dbo.tbl_corner_good as C "
		sqlStr = sqlStr & " 		on i.makerid=c.lecturer_id "
		sqlStr = sqlStr & " left join [db_academy].[dbo].[tbl_diy_item_Contents] as s "
		sqlStr = sqlStr & " 	on i.itemid=s.itemid "
		sqlStr = sqlStr & " where i.isusing='Y'  and f.userid='" & Fuserid & "' "		''and  i.sellyn='Y'

'		response.write sqlStr &"<Br>"
'		response.end
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit function

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr + " i.itemid, i.itemname, i.sellcash, i.orgPrice, i.Makerid, i.brandName, i.ListImage, i.ListImage120, i.SmallImage, icon1image, i.icon2image, i.basicimage "
		sqlStr = sqlStr + " ,i.Sellyn, i.SaleYn, i.LimitYn, i.LimitNo, i.LimitSold, i.RegDate, i.ItemCouponYn, i.ItemCouponValue, i.ItemCouponType, i.evalCnt, i.ItemScore, i.itemDiv, optioncnt "
		sqlStr = sqlStr + " , c.image_profile_75x75, c.lecturer_name, c.newImage_profile, s.keywords "
		sqlStr = sqlStr & " , (select count(*) from [db_academy].dbo.tbl_diy_myfavorite where itemid = i.itemid and userid = '"& Fuserid &"') as chkfav "
		sqlStr = sqlStr & " from [db_academy].[dbo].tbl_diy_item as i"
		sqlStr = sqlStr & " join [db_academy].dbo.tbl_diy_myfavorite as f"
		sqlStr = sqlStr & " 	on i.itemid=f.itemid"
		sqlStr = sqlStr & " left Join db_academy.dbo.tbl_corner_good as C "
		sqlStr = sqlStr & " 		on i.makerid=c.lecturer_id "
		sqlStr = sqlStr & " left join [db_academy].[dbo].[tbl_diy_item_Contents] as s "
		sqlStr = sqlStr & " 	on i.itemid=s.itemid "
		sqlStr = sqlStr & " where i.isusing='Y'  and f.userid='" & Fuserid & "' "		''and  i.sellyn='Y'
		sqlStr = sqlStr & " order by f.regdate desc "	

'	response.write sqlStr &"<Br>"
'	response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
				set FItemList(i) = new CCategoryPrdItem
	
					FItemList(i).FItemid			= rsACADEMYget("Itemid")
					FItemList(i).Fkeywords			= rsACADEMYget("keywords")
					FItemList(i).FItemName		= db2html(rsACADEMYget("ItemName"))
					FItemList(i).FSellCash		= rsACADEMYget("SellCash")
					FItemList(i).FOrgPrice		= rsACADEMYget("OrgPrice")
					FItemList(i).FMakerId		= rsACADEMYget("MakerId")
					FItemList(i).FBrandName		= db2html(rsACADEMYget("BrandName"))
					FItemList(i).FImageList		= fingersImgUrl & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" &db2html(rsACADEMYget("ListImage"))
					FItemList(i).FImageList120	= fingersImgUrl & "/diyitem/webimage/list120/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & db2html(rsACADEMYget("ListImage120"))
					FItemList(i).FImageSmall		= fingersImgUrl & "/diyitem/webimage/small/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" &db2html(rsACADEMYget("smallImage"))
					FItemList(i).FImageicon1 	= fingersImgUrl & "/diyitem/webimage/icon1/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & rsACADEMYget("icon1image")
					FItemList(i).FImageicon2 	= fingersImgUrl & "/diyitem/webimage/icon2/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" & rsACADEMYget("icon2image")
					FItemList(i).FImageBasic 	= fingersImgUrl & "/diyitem/webimage/basic/" & GetImageSubFolderByItemid(FItemList(i).FItemid) & "/" &rsACADEMYget("basicimage")

					FItemList(i).FSellyn			= rsACADEMYget("sellYn")
					FItemList(i).FSaleyn			= rsACADEMYget("SaleYn")
					FItemList(i).FLimityn		= rsACADEMYget("LimitYn")
					FItemList(i).FRegdate		= rsACADEMYget("regdate")
					FItemList(i).FItemcouponyn	= rsACADEMYget("itemcouponYn")
					FItemList(i).FItemcouponvalue= rsACADEMYget("itemCouponValue")
					FItemList(i).FItemcoupontype	= rsACADEMYget("itemCouponType")
					FItemList(i).FEvalcnt		= rsACADEMYget("evalCnt")
					FItemList(i).FItemScore		= rsACADEMYget("itemScore")
					FItemList(i).FOptioncnt		= rsACADEMYget("optioncnt")
					
					FItemList(i).Flecturer_name		= rsACADEMYget("lecturer_name")
					FItemList(i).Flecturer_img		= rsACADEMYget("image_profile_75x75")

					If rsACADEMYget("newImage_profile") <> "" Then
						FItemList(i).FNewlectureimg		= fingersImgUrl & "/corner/newImage_profile/thumbimg3/t3_" & rsACADEMYget("newImage_profile")
					Else
						FItemList(i).FNewlectureimg		= ""
					End If
					
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Function

	''---------------------------------------------------------------------------------
	'Qna 전체 리스트
	public function GetDiyQnaList()
		dim sqlStr, i, sqlsearch

		if FRectmode = "list" then
			sqlsearch = sqlsearch & " and q.reply_depth = '0' "
		elseif FRectmode = "mylist" then
			sqlsearch = sqlsearch & " and q.reply_depth = '0' and q.userid = '"&FRectuserid&"' "
		end If

		if FRectgroupidx <> "" then		''idx(x) 에 속한 리플그룹 idx
			sqlsearch = sqlsearch & " and q.reply_group_idx = '"&FRectgroupidx&"'"
		end If

		if FRectitemid <> "" then		''상품상세 qna
			sqlsearch = sqlsearch & " and q.itemid = '"&FRectitemid&"'"
		end If

'		if FRectuserid <> "" then		''내 상품 qna
'			sqlsearch = sqlsearch & " and q.userid = '"&FRectuserid&"'"
'		end If

		'// 총 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_academy_qna_new] as q "
        sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_diy_item as i "
        sqlStr = sqlStr & "		on q.itemid=i.itemid "
        sqlStr = sqlStr & " where q.isusing='Y' " + sqlsearch
'        sqlStr = sqlStr & " where isusing='Y' and  itemid=" + CStr(FRectitemid) & sqlsearch

'		response.write sqlStr &"<Br>"
'		response.end
        rsACADEMYget.Open sqlStr,dbACADEMYget,1
            FTotalCount = rsACADEMYget("cnt")
        rsACADEMYget.Close

		if FTotalCount < 1 then exit function

		sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " q.idx, q.itemid, q.reply_group_idx, q.reply_depth, q.reply_num, q.userid, q.userlevel, q.replyuserid, q.answerYN, q.qna, q.title, q.comment, q.device, q.isusing, q.regdate, q.makerid, i.itemname"
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_academy_qna_new] as q "
		sqlStr = sqlStr & " JOIN db_academy.dbo.tbl_diy_item as i "
		sqlStr = sqlStr & "		on q.itemid=i.itemid "
		sqlStr = sqlStr & " where q.isusing='Y' " + sqlsearch
'		sqlStr = sqlStr & " where isusing='Y' and  itemid=" + CStr(FRectitemid) & sqlsearch

		if FRectmode = "reply" then
			sqlStr = sqlStr & " order by q.reply_num asc "
		else
			sqlStr = sqlStr & " order by case when (q.userid='"&FRectuserid&"') then 1 else 2 end, q.idx desc "
		end if 

'		response.write sqlStr &"<Br>"
'	response.end
		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		
		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

        if (FResultCount<1) then FResultCount=0

        redim FItemList(FResultCount)

        i=0
        if  not rsACADEMYget.EOF  then
            rsACADEMYget.absolutepage = FCurrPage
            do until rsACADEMYget.EOF
				set FItemList(i) = new CdiyQnaItem

					FItemList(i).Fidx				= rsACADEMYget("idx")
					FItemList(i).Fqna				= rsACADEMYget("qna")	'qna 구분
					FItemList(i).Fitemid				= rsACADEMYget("itemid")
					FItemList(i).FItemName			= db2html(rsACADEMYget("itemname"))
					FItemList(i).Fuserlevel			= rsACADEMYget("userlevel")
					FItemList(i).Freplyuserid		= rsACADEMYget("replyuserid")
					FItemList(i).Ftitle				= rsACADEMYget("title")
					FItemList(i).Fdevice				= rsACADEMYget("device")
					FItemList(i).Fuserid				= rsACADEMYget("userid")	
					FItemList(i).Fmakerid			= rsACADEMYget("makerid")
					FItemList(i).Fcomment			= rsACADEMYget("comment")
					FItemList(i).Fisusing			= rsACADEMYget("isusing")
					FItemList(i).Fregdate			= rsACADEMYget("regdate")
					FItemList(i).FanswerYN			= rsACADEMYget("answerYN")			'답변 여부
					FItemList(i).Freply_num			= rsACADEMYget("reply_num")			'리플 순서
					FItemList(i).Freply_depth		= rsACADEMYget("reply_depth")		'원글0,리플1 뎁스
					FItemList(i).Freply_group_idx	= rsACADEMYget("reply_group_idx")	'리플 그룹 idx
                rsACADEMYget.movenext
                i=i+1
            loop
        end if
        rsACADEMYget.Close
    end Function
	''---------------------------------------------------------------------------------
	'내 qna
	public Sub GetmyqnaRead()
		dim sqlStr, sqlsearch

		if FRectqnaidx <> "" then		'원 qna idx
			sqlsearch = sqlsearch & " and idx = '"&FRectqnaidx&"'"
		end If

		if FRectitemid <> "" then		''상품코드
			sqlsearch = sqlsearch & " and itemid = '"&FRectitemid&"'"
		end If

		if FRectgroupidx <> "" then		''idx(x) 에 속한 리플그룹 idx
			sqlsearch = sqlsearch & " and reply_group_idx = '"&FRectgroupidx&"'"
		end If

		if FRectuserid <> "" then
			sqlsearch = sqlsearch & " and userid = '"&FRectuserid&"'"
		end If

		
		sqlStr = " Select * "
		sqlStr = sqlStr & "  from [db_academy].[dbo].[tbl_academy_qna_new] "
		sqlStr = sqlStr & "  Where isusing = 'Y' "&sqlsearch

'		response.write sqlStr
'		response.end

		rsACADEMYget.Open SqlStr, dbACADEMYget, 1
		FResultCount = rsACADEMYget.RecordCount

		set FOneItem = new CdiyQnaItem

		if Not rsACADEMYget.Eof then
			FOneItem.Fidx					= rsACADEMYget("idx")
			FOneItem.Fitemid				= rsACADEMYget("itemid")
			FOneItem.Fuserid				= rsACADEMYget("userid")
			FOneItem.Fuserlevel			= rsACADEMYget("userlevel")
			FOneItem.Freplyuserid		= rsACADEMYget("replyuserid")
			FOneItem.Ftitle				= rsACADEMYget("title")
			FOneItem.Fsmsok				= rsACADEMYget("smsok")
			FOneItem.Fsmsnum				= rsACADEMYget("smsnum")
			FOneItem.Fregdate			= rsACADEMYget("regdate")
			FOneItem.Femailok			= rsACADEMYget("emailok")
			FOneItem.Fcomment			= rsACADEMYget("comment")
			FOneItem.Freply_num			= rsACADEMYget("reply_num")
			FOneItem.Freply_depth		= rsACADEMYget("reply_depth")
			FOneItem.Freply_group_idx	= rsACADEMYget("reply_group_idx")
		end if
		rsACADEMYget.Close

	end sub
	''---------------------------------------------------------------------------------

	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

end Class

																								''''''수정
'//강좌 종류(구분) name
function DrawdiyshopGubunName(cdl,cdm,cds)
	dim i, sqlStr, catecodename, sqlsearch, catecode
	catecode = cdl&cdm&cds
	catecodename = "카테고리선택"

	if cdl <> "" and cdm = "" and cds = "" then
		sqlsearch = sqlsearch & " and left(catecode,3) = '"&cdl&"'"
	elseif cdl <> "" and cdm <> "" and cds="" then
		sqlsearch = sqlsearch & " and left(catecode,6) = '"&cdl&cdm&"'"
	elseif cdl <> "" and cdm <> "" and cds <> "" then
		sqlsearch = sqlsearch & " and left(catecode,9) = '"&cdl&cdm&cds&"'"
	end if

	'// 본문 내용 접수
	sqlStr = "select top 1 catename"
	sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_display_cate_Academy]"
	sqlStr = sqlStr & " where 1=1 and useyn='Y' "&sqlsearch

'	response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	IF Not rsACADEMYget.EOF THEN
		catecodename = rsACADEMYget(0)
	END IF
	rsACADEMYget.Close

	if catecode = "" then
		response.write "카테고리선택"
	else
		response.write catecodename
	end if
end Function


'//diy shop 카테고리구분 종류 레이어
function DrawDiyCateMidGubun(cdl, cdm, cds, SortMet, ics, pgGubun, cpidx)
	dim FTotalCount, arrList, i, sqlStr, sqlsearch, catecode, depthsearch, nextdepthsearch
	sqlsearch=""
	catecode = cdl&cdm&cds
'	len(catecode)

	if catecode <> "" then
		if len(catecode) = 3 then
			sqlsearch = sqlsearch & " and left(catecode,3) = '"&catecode&"'"
		elseif len(catecode) = 6 then
			sqlsearch = sqlsearch & " and left(catecode,6) = '"&catecode&"'"
		elseif len(catecode) = 9 then
			sqlsearch = sqlsearch & " and left(catecode,9) = '"&catecode&"'"
		end if
	end if

	if len(catecode) = 3 then
		depthsearch = depthsearch & " and depth = '2' "
	elseif len(catecode) = 6 then
		depthsearch = depthsearch & " and depth = '3' "
	elseif len(catecode) = 9 then
		depthsearch = depthsearch & " and depth = '3' "
	else
		depthsearch = depthsearch & " and depth = '1' "
	end if

	if len(catecode) = 3 then
		nextdepthsearch = nextdepthsearch & " and depth = '1' "
	elseif len(catecode) = 6 then
		nextdepthsearch = nextdepthsearch & " and depth = '2' "
	elseif len(catecode) = 9 then
		nextdepthsearch = nextdepthsearch & " and depth = '2' "
	else
		nextdepthsearch = nextdepthsearch & " and depth = '1' "
	end if
	
	'// 결과수 카운트
	sqlStr = "select count(*) as cnt"
	sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_display_cate_Academy]"
	sqlStr = sqlStr & " where 1=1 and useyn='Y' "&sqlsearch&depthsearch

	'response.write sqlStr &"<Br>"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		FTotalCount = rsACADEMYget("cnt")
	rsACADEMYget.Close

'	if FTotalCount < 1 then exit function

	if FTotalCount > 0 then
		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " catecode, catename, depth"
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_display_cate_Academy]"
		sqlStr = sqlStr & " where 1=1 and useyn='Y' "&sqlsearch&depthsearch
		sqlStr = sqlStr & " order by sortNo asc "
	
'		response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		IF Not rsACADEMYget.EOF THEN
			arrList = rsACADEMYget.getRows()
		END IF
		rsACADEMYget.Close
	else
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_display_cate_Academy]"
		sqlStr = sqlStr & " where 1=1 and useyn='Y' "&sqlsearch&nextdepthsearch
	
		'response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.Close

		'// 본문 내용 접수
		sqlStr = "select "
		sqlStr = sqlStr & " catecode, catename, depth"
		sqlStr = sqlStr & " from [db_academy].[dbo].[tbl_display_cate_Academy]"
		sqlStr = sqlStr & " where 1=1 and useyn='Y' "&sqlsearch&nextdepthsearch
		sqlStr = sqlStr & " order by sortNo asc "
	
'		response.write sqlStr &"<Br>"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		IF Not rsACADEMYget.EOF THEN
			arrList = rsACADEMYget.getRows()
		END IF
		rsACADEMYget.Close
	end if

%>
	<div class="sortList" id="gubunselect" style="display:none">
		<ul>
			<% if cdl = "" then %>
				<li class="current"><a href="/myfingers/coupon/cpdiyList.asp?srm=<%=trim(SortMet)%>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">전체보기</a></li>
			<% end if %>

			<% for i = 0 to FTotalCount-1 %>
				<% if arrList(2,i) = "1" then %>
					<li <% if trim(cdl) = trim(arrList(0,i)) then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=<%=trim(SortMet)%>&cdl=<%=left(arrList(0,i),3)%>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>"><%= arrList(1,i) %></a></li>
				<% elseif arrList(2,i) = "2" then  %>
					<li <% if trim(cdm) = trim(mid(arrList(0,i),4,6)) then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=<%=trim(SortMet)%>&cdl=<%=left(arrList(0,i),3)%>&cdm=<%=mid(arrList(0,i),4,6)%>&cds=<%=mid(arrList(0,i),7,9)%>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>"><%= arrList(1,i) %></a></li>
				<% elseif arrList(2,i) = "3" then  %>
					<li <% if trim(cds) = trim(mid(arrList(0,i),7,9)) then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=<%=trim(SortMet)%>&cdl=<%=left(arrList(0,i),3)%>&cdm=<%=mid(arrList(0,i),4,6)%>&cds=<%=mid(arrList(0,i),7,9)%>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>"><%= arrList(1,i) %></a></li>
				<% end if %>
			<% next %>

		</ul>
	</div>
<%
end Function

'//강좌 정렬(신규,인기,마감,낮은가격,높은가격) 레이어
function DrawdiySort(cdl, cdm, cds, SortMet, ics, pgGubun, cpidx)
%>
	<div class="sortList" id="sortselect" style="display:none">
		<ul>
			<li <% if SortMet="ne" then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=ne&cdl=<%=cdl%>&cdm=<%= cdm %>&cds=<%= cds %>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">신규순</a></li>
			<li <% if SortMet="be" then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=be&cdl=<%=cdl%>&cdm=<%= cdm %>&cds=<%= cds %>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">인기순</a></li>
			<li <% if SortMet="lp" then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=lp&cdl=<%=cdl%>&cdm=<%= cdm %>&cds=<%= cds %>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">낮은가격순</a></li>
			<li <% if SortMet="hp" then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=hp&cdl=<%=cdl%>&cdm=<%= cdm %>&cds=<%= cds %>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">높은가격순</a></li>
			<li <% if SortMet="hs" then %>class="current"<% end if %>><a href="/myfingers/coupon/cpdiyList.asp?srm=hs&cdl=<%=cdl%>&cdm=<%= cdm %>&cds=<%= cds %>&icoSize=<%= ics %>&pgGubun=<%= pgGubun %>&cpidx=<%= cpidx %>">높은할인율순</a></li>
		</ul>
	</div>
<%
end Function

Function getMyphoneNumber(vuserid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 usercell FROM [db_user].[dbo].tbl_user_n " 
	strSql = strSql & " WHERE userid = '"&vuserid&"' " 
	rsUSERget.Open strSql, dbUSERget, 1
	If Not(rsUSERget.EOF or rsUSERget.BOF) Then
		getMyphoneNumber = rsUSERget("usercell")
	End If
	rsUSERget.Close
End Function
%>