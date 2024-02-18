<%
CLASS DIYItemPrdCls

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	dim Prd
	dim FADD
	dim FResultCount
	dim itEvtImg

	dim FRectCdL
	dim FRectCdM
	dim FRectCdS 
	dim FRectItemid
	dim FRectMakerid
	dim FPageSize
	dim FSellScope
	dim FItemList
	'//추가
	Dim FRectCateCode

	Public Sub GetItemData(ByVal iid , ByVal RealWaitYN)

		dim strSQL , linkimgurl
		If RealWaitYN = "N" Or RealWaitYN = "" then
			strSQL = "execute [db_academy].[dbo].sp_academy_DIYItemPrdApp @vItemID ='" & CStr(iid) & "'"
		else
			strSQL = "execute [db_academy].[dbo].sp_academy_DIYItemPrdAppReal @vItemID ='" & CStr(iid) & "'"
		End If 

		If RealWaitYN = "N" Then
			linkimgurl = "waitimage"
		Else
			linkimgurl = "webimage"
		End If 

		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType=adOpenStatic
		rsACADEMYget.Locktype=adLockReadOnly
		rsACADEMYget.Open strSQL, dbACADEMYget

		set Prd = new CCategoryPrdItem

		if  not rsACADEMYget.EOF  then

			FResultCount = 1
			rsACADEMYget.Movefirst

				Prd.FItemid					= rsACADEMYget("Itemid")				'상품 코드
				Prd.FcdL					= rsACADEMYget("Cate_large")
				Prd.FcdM					= rsACADEMYget("Cate_mid")
				Prd.FcdS					= rsACADEMYget("Cate_small")

				Prd.FMakerid				= rsACADEMYget("makerid")				'업체 아이디

				Prd.Fitemname				= db2html(rsACADEMYget("itemname"))	'상품명
				Prd.FMakerName				= db2html(rsACADEMYget("makername")) 	'제조사
				'Prd.FOrgprice				= rsACADEMYget("orgprice")				'원가
				Prd.FItemDiv				= rsACADEMYget("itemdiv")				'상품 속성
				Prd.FMileage				= rsACADEMYget("mileage")				'마일리지
				Prd.FSellCash				= rsACADEMYget("sellcash")				'판매가
				Prd.FLimitNo				= rsACADEMYget("limitno")				'한정수량
				Prd.FLimitSold				= rsACADEMYget("LimitSold")			'한정판매수량
				Prd.FKeyWords				= db2html(rsACADEMYget("keyWords"))
				Prd.FDeliverytype			= rsACADEMYget("deliverytype")
				'Prd.FEvalCnt				= rsACADEMYget("evalcnt")
				Prd.FOptionCnt				= rsACADEMYget("optioncnt")
				'Prd.FQnaCnt					= rsACADEMYget("qnaCnt")
				Prd.FItemSource 			= db2html(rsACADEMYget("itemsource"))
				Prd.FSourceArea 			= db2html(rsACADEMYget("sourcearea"))
				Prd.FItemSize 				= db2html(rsACADEMYget("itemsize"))
				Prd.FItemWeight				= db2html(rsACADEMYget("itemWeight"))
				'Prd.FCurrItemCouponIdx 		= rsACADEMYget("curritemcouponidx")

				Prd.FSellYn					= rsACADEMYget("sellyn")
				'Prd.FSaleYn					= rsACADEMYget("saleyn")
				Prd.FLimitYn 				= rsACADEMYget("limityn")
				'Prd.FItemCouponYN			= rsACADEMYget("itemcouponyn")
				'Prd.FItemCouponType			=	rsACADEMYget("itemcoupontype")
				'Prd.FItemCouponValue		= rsACADEMYget("itemcouponvalue")
				'Prd.FUsingHTML				= rsACADEMYget("usinghtml")

				Prd.FItemContent 			= db2html(rsACADEMYget("itemcontent"))
				Prd.FOrderComment			= db2html(Trim(rsACADEMYget("ordercomment")))

				'Prd.FAvailPayType			= rsACADEMYget("AvailPayType")
                
                if (rsACADEMYget("mainimage"))<>"" then
				Prd.FImageMain 				= fingersImgUrl & "/diyitem/"&linkimgurl&"/main/" + GetImageSubFolderByItemid(Prd.FItemid) + "/" + rsACADEMYget("mainimage")
				end if
				Prd.FImageList 				= fingersImgUrl & "/diyitem/"&linkimgurl&"/list/" + GetImageSubFolderByItemid(Prd.FItemid) + "/" + rsACADEMYget("listimage")
				'Prd.FImageList120 			= fingersImgUrl & "/diyitem/webimage/list120/" + GetImageSubFolderByItemid(Prd.FItemid) + "/" + rsACADEMYget("listimage120")
				'Prd.FImageSmall 			= fingersImgUrl & "/diyitem/webimage/small/" + GetImageSubFolderByItemid(Prd.FItemid) + "/" + rsACADEMYget("smallimage")
				Prd.FImageBasic 			= fingersImgUrl & "/diyItem/"&linkimgurl&"/basic/" + GetImageSubFolderByItemid(Prd.FItemid) + "/" + rsACADEMYget("basicimage")
				'Prd.FImageBasicIcon 		= fingersImgUrl & "/diyitem/webimage/basicicon/" + GetImageSubFolderByItemid(Prd.FItemid) + "/C" + rsACADEMYget("basicimage")
				Prd.FRegdate 				= rsACADEMYget("regdate")
				Prd.FBrandName				= db2Html(rsACADEMYget("brandname"))
				Prd.FBrandName_kor			= db2Html(rsACADEMYget("BrandName_Kor"))
				Prd.FBrandUsing				= rsACADEMYget("BrandUsing")			'브랜드 사용 여부
				Prd.FDefaultFreeBeasongLimit = rsACADEMYget("DefaultFreeBeasongLimit")
				Prd.FDefaultDeliverPay		= rsACADEMYget("defaultDeliveryPay")
				Prd.FcateCode				= rsACADEMYget("catecode") ''카테고리
				'Prd.FFavcount				= rsACADEMYget("favcount") ''관심상품 카운트수

				Prd.Fcstodr					= rsACADEMYget("cstodr") ''즉시발송 제작후 발송
				Prd.Frequiremakeday			= rsACADEMYget("requiremakeday") ''제작후 발송기간 ? 일
				Prd.Frequirecontents		= rsACADEMYget("requirecontents") ''특이사항
				Prd.Frefundpolicy			= rsACADEMYget("refundpolicy") ''교환/환불 정책
		else
			FResultCount = 0
		end if

		rsACADEMYget.close

	End Sub


	Public Sub getAddImage(byval itemid , ByVal RealWaitYN)
			dim strSQL,ArrRows,i
			Dim linkimgurl1, linkimgurl2
			
			If RealWaitYN = "N" Then
				linkimgurl1 = "waitcontentsimage"
				linkimgurl2 = "waitimage"
			Else
				linkimgurl1 = "contentsimage"
				linkimgurl2 = "webimage"
			End If 

			If RealWaitYN = "N" Then
				strSQL = "exec [db_academy].[dbo].sp_academy_DIYItemPrd_AddImageWaitApp @vItemid =" & CStr(itemid)
			Else 
				strSQL = "exec [db_academy].[dbo].sp_academy_DIYItemPrd_AddImageWaitAppReal @vItemid =" & CStr(itemid)
			End If 

			rsACADEMYget.CursorLocation = adUseClient
			rsACADEMYget.CursorType=adOpenStatic
			rsACADEMYget.Locktype=adLockReadOnly
			rsACADEMYget.Open strSQL, dbACADEMYget

			If Not rsACADEMYget.EOF Then
				ArrRows 	= rsACADEMYget.GetRows
			End if
			rsACADEMYget.close

			if isArray(ArrRows) then

			FResultCount = Ubound(ArrRows,2) + 1

			redim  FADD(FResultCount)

				For i=0 to FResultCount-1
					Set FADD(i) = new CCategoryPrdItem
					FADD(i).FAddimageGubun	= ArrRows(0,i)
					FADD(i).FAddImageType	= ArrRows(1,i)
					FADD(i).FAddimgText		= ArrRows(3,i)
					IF ArrRows(1,i)="1" Or ArrRows(1,i)="2" Then
						FADD(i).FAddimage 		= fingersImgUrl & "/diyitem/"&linkimgurl1&"/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
					Else
						FADD(i).FAddimage 		= fingersImgUrl & "/diyitem/"&linkimgurl2&"/add" & Cstr(FADD(i).FAddimageGubun) & "/" & GetImageSubFolderByItemid(itemid) & "/" & ArrRows(2,i)
						FADD(i).FAddimageSmall	= fingersImgUrl & "/diyitem/"&linkimgurl2&"/add" & Cstr(FADD(i).FAddimageGubun) & "icon/" & GetImageSubFolderByItemid(itemid) & "/C" & ArrRows(2,i)
					End If
				next
			end if
	End Sub

	public Sub GetOneItemAddImageList()
	    dim sqlstr, i, j
	    dim bufimgadd
	    dim bufimgaddCnt
		Dim FTotalCount
	    
	    sqlStr = "select top 1 imgadd"
		sqlStr = sqlStr & " from db_academy.dbo.tbl_diy_wait_item"
		sqlStr = sqlStr & " where itemid='" & itemid & "'"
		
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
		if  not rsACADEMYget.EOF  then
		    bufimgadd   = rsACADEMYget("imgadd")
		end if
		rsACADEMYget.close
		
		if IsNULL(bufimgadd) then 
		    bufimgaddCnt = 0
		else
		    bufimgadd = split(bufimgadd,",")
		    bufimgaddCnt = UBound(bufimgadd)
		    
		end if
		
        FTotalCount = bufimgaddCnt
        FResultCount = FTotalCount

     
        redim FItemList(FResultCount)
        
        for i=0 to bufimgaddCnt-1
            set FItemList(i) = new CCategoryPrdItem
            FItemList(i).FIDX           = i
            FItemList(i).FITEMID        = itemid
            FItemList(i).FIMGTYPE       = 0
            FItemList(i).FGUBUN         = i+1
            FItemList(i).FADDIMAGE_400  = bufimgadd(i)
            
            FItemList(i).FADDIMAGE_Icon =""
            
            if ((Not IsNULL(FItemList(i).FADDIMAGE_400)) and (FItemList(i).FADDIMAGE_400<>"")) then FItemList(i).FADDIMAGE_400 = fingersImgUrl & "/diyItem/waitimage/add" & CStr(i+1) & "/" & GetImageSubFolderByItemid(FItemList(i).FItemID) + "/"  + FItemList(i).FADDIMAGE_400
        next
        
    end Sub

	'BEST 상품 목록 접수
	PUBLIC SUB getDIYBESTItemList()

		dim strSQL, iLoopCnt

		'// 내용 접수
		strSQL =" EXECUTE db_academy.dbo.sp_academy_DIYItemBest_2016 " &_
					" @disp = '" &FRectCateCode&"'" &_
					" ,@itemid ='" &FRectItemid& "' " &_
					" ,@Makerid ='" &FRectMakerid& "' " &_
					" ,@PgSize='" &FPageSize&"' " &_
					" ,@SellScope = '"&FSellScope&"'"
		rsACADEMYget.CursorLocation = adUseClient
		rsACADEMYget.CursorType = adOpenStatic
		rsACADEMYget.LockType = adLockOptimistic
		rsACADEMYget.PageSize=FPageSize
		rsACADEMYget.Open strSQL, dbACADEMYget
		FResultCount = rsACADEMYget.RecordCount
		
		if (FResultCount<1) then FResultCount=0
		
		redim FItemList(FResultCount)

		iLoopCnt=0
		if  not rsACADEMYget.EOF  then
			do until rsACADEMYget.eof
				set FItemList(iLoopCnt) = new CCategoryPrdItem
				
				FItemList(iLoopCnt).FItemid			= rsACADEMYget("Itemid")
				FItemList(iLoopCnt).FItemName		= db2html(rsACADEMYget("ItemName"))
				FItemList(iLoopCnt).FSellCash		= rsACADEMYget("SellCash")
				FItemList(iLoopCnt).FOrgPrice		= rsACADEMYget("OrgPrice")
				FItemList(iLoopCnt).FMakerId		= rsACADEMYget("MakerId")
				FItemList(iLoopCnt).FBrandName		= db2html(rsACADEMYget("BrandName"))
				FItemList(iLoopCnt).FImageList		= fingersImgUrl & "/diyitem/webimage/list/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" &db2html(rsACADEMYget("ListImage"))
				FItemList(iLoopCnt).FImageList120	= fingersImgUrl & "/diyitem/webimage/list120/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & db2html(rsACADEMYget("ListImage120"))
				FItemList(iLoopCnt).FImageSmall		= fingersImgUrl & "/diyitem/webimage/small/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" &db2html(rsACADEMYget("smallImage"))
				FItemList(iLoopCnt).FImageicon1 	= fingersImgUrl & "/diyitem/webimage/icon1/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & rsACADEMYget("icon1image")
				FItemList(iLoopCnt).FImageicon2 	= fingersImgUrl & "/diyitem/webimage/icon2/" & GetImageSubFolderByItemid(FItemList(iLoopCnt).FItemid) & "/" & rsACADEMYget("icon2image")
					
				FItemList(iLoopCnt).FSellyn			= rsACADEMYget("sellYn")
				FItemList(iLoopCnt).FSaleyn			= rsACADEMYget("SaleYn")
				FItemList(iLoopCnt).FLimityn		= rsACADEMYget("LimitYn")
				FItemList(iLoopCnt).FRegdate		= rsACADEMYget("regdate")
				FItemList(iLoopCnt).FItemcouponyn	= rsACADEMYget("itemcouponYn")
				FItemList(iLoopCnt).FItemcouponvalue= rsACADEMYget("itemCouponValue")
				FItemList(iLoopCnt).FItemcoupontype	= rsACADEMYget("itemCouponType")
				FItemList(iLoopCnt).FEvalcnt		= rsACADEMYget("evalCnt")
				FItemList(iLoopCnt).FItemScore		= rsACADEMYget("itemScore")

				iLoopCnt=iLoopCnt+1
				rsACADEMYget.moveNext
			loop
		end if

		rsACADEMYget.Close

	End Sub
	
	'//핑거스 상품고시 2016 -- 이종화
	Public Sub getItemAddExplain(byval itemid , ByVal RealWaitYN)
			dim strSQL,ArrRows,i
			
			If RealWaitYN = "N" Or RealWaitYN = "" Then
				strSQL = "exec [db_academy].[dbo].[sp_academy_diyitem_AddExplainApp] " & CStr(itemid)
			Else
				strSQL = "exec [db_academy].[dbo].[sp_academy_diyitem_AddExplainAppReal] " & CStr(itemid)
			End If 

			rsACADEMYget.CursorLocation = adUseClient
			rsACADEMYget.CursorType=adOpenStatic
			rsACADEMYget.Locktype=adLockReadOnly
			rsACADEMYget.Open strSQL, dbACADEMYget

			If Not rsACADEMYget.EOF Then
				ArrRows 	= rsACADEMYget.GetRows
			End if
			rsACADEMYget.close

			if isArray(ArrRows) then

			FResultCount = Ubound(ArrRows,2) + 1

			redim  FItemList(FResultCount)

				For i=0 to FResultCount-1
					Set FItemList(i) = new CCategoryPrdItem

					FItemList(i).FInfoname		= ArrRows(0,i)
					FItemList(i).FInfoContent	= ArrRows(1,i)
					FItemList(i).FinfoCode		= ArrRows(2,i)

				next
			end if
	End Sub

	'### 상품상세설명 동영상
	Public Function fnGetItemVideos(byval itemid, ByVal vgubun , ByVal RealWaitYN)
		Dim st_table
		If RealWaitYN = "N" Or RealWaitYN = "" Then
			st_table = "db_academy.dbo.tbl_diy_wait_item_videos"
		Else		
			st_table = "db_academy.dbo.tbl_diy_item_videos"
		End If 
		dim strSQL, vCount
		strSQL = " SELECT TOP 1 videogubun, videotype, videourl, videowidth, videoheight, videofullurl FROM "& st_table &" WHERE videogubun='"&vgubun&"' And itemid = '" & itemid & "'"
		'response.write strSQL
		rsACADEMYget.open strSQL, dbACADEMYget
		set Prd = new CCategoryPrdItem
		if  not rsACADEMYget.EOF  then
			FResultCount = 1
			rsACADEMYget.Movefirst
				Prd.FvideoUrl    	= rsACADEMYget("videourl")
				Prd.FvideoWidth		= rsACADEMYget("videowidth")
				Prd.FvideoHeight	= rsACADEMYget("videoheight")
				Prd.Fvideogubun		= rsACADEMYget("videogubun")
				Prd.FvideoType		= rsACADEMYget("videotype")
				Prd.FvideoFullUrl	= rsACADEMYget("videofullurl")
		Else
			FResultCount = 0
		End IF
		rsACADEMYget.Close

	End Function

	''작가 프로필 사진 / 이름
	Public Function fnDiyItemLecturer(byval makerid)
		dim strSQL
'############		2016-08-12 김진영 하단 주석..프로필에서 보는 이미지와 상이하여 고침 ################
'		strSQL = " SELECT TOP 1 lecturer_name , lecturer_img from db_academy.dbo.tbl_lec_user where lecturer_id = '" & makerid & "'"
'		rsACADEMYget.open strSQL, dbACADEMYget
'		set Prd = new CCategoryPrdItem
'		if  not rsACADEMYget.EOF  then
'			FResultCount = 1
'			rsACADEMYget.Movefirst
'				Prd.Flecturer_name    	= rsACADEMYget("lecturer_name")
'				Prd.Flecturer_img		= webImgUrl & "/image/brandlogo/t1_" & rsACADEMYget("lecturer_img") '//t1 300, t2 200 , t1 150
'		Else
'			FResultCount = 0
'		End IF
'		rsACADEMYget.Close
'#######################################################################################################
		strSQL = " SELECT TOP 1 lecturer_name , isnull(newImage_profile, '') as newImage_profile " &_
				 " ,(select count(*) from db_academy.[dbo].[tbl_academy_Mobile_main] where tabgubun='3' and makerid='" & makerid & "') as best"  &_
				 "	from db_academy.[dbo].[tbl_corner_good] where lecturer_id = '" & makerid & "'"
		rsACADEMYget.open strSQL, dbACADEMYget
		set Prd = new CCategoryPrdItem
		if  not rsACADEMYget.EOF  then
			FResultCount = 1
			rsACADEMYget.Movefirst
				Prd.Flecturer_name    	= rsACADEMYget("lecturer_name")
				Prd.Flecturer_img		= db2html(rsACADEMYget("newImage_profile"))
				Prd.Flecturer_best		= db2html(rsACADEMYget("best"))
		Else
			FResultCount = 0
		End IF
		rsACADEMYget.Close

	End Function
End Class

%>

