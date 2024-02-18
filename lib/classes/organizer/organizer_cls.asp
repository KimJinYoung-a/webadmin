<%
Class organizeItemsCls

	public FImageicon1
	public FImageicon2
	public FImageSmall
	public FImageList120
	public FImageList
	public ForganizerID
	public FCateCode
	public Fitemid
	public FRegDate
	public FisUsing
	public FImg
	public fcontents
	public fuserid
	public fidx
	public foption_value
	public foption_order
	public ftype
	public fkeyword_option_count
	public fcontents_idx_count
	public fcontents_idx
	public FInfoidx
	public FInfoGubun
	public Finfoname 	
	public Finfoimg
	public FinfoPageCnt
	public fyear
	public FDiaryType
	public FBasicImg
	public FListimg
	public FIconImg
	public FgiftYn
	public FonlyYearYn
	public FhitYn
	public FItemCouponType
	public FItemCouponValue
	public fcommentyn
	public FBanneridx	
	public FBannerType
	public FBannerurl
	public FBannerImg
	public FBannerUsing
	public FEvtCode
	public FbannerMapUsing
	public ConIdx
	public ConfImg
	public ConTTxt
 	public fcomment_img
 	public feventgroup_code
 	public fevent_code
 	public fitemname
 	public fevent_start
	public fevent_end 	
	public fweight
 	public fsearch_view
	public fposcode
	public fposname 
	public fimagetype 
	public fimagewidth 
	public fimageheight 
	public fimagepath 
	public flinkpath 
	public fevt_code  	
	public fimagecount 
	public fimage_order 
	public fsearch_order
	public finfo_name
	public fcolor
	public fitemtype
	public forganizer_order
	public fevt_enddate
	public fevt_kind
	public fbrand 
	public fevt_startdate
	public fevt_bannerimg
	public fidx_order
	public fevent_type
	public fevt_name	
	
	public FOW_IDX
	public FOW_TITLE
	public FOW_CONTENTS
	public FOW_MAIN_IMG
	public FOW_SUB_IMG
	public FOW_REGDATE
	public FOW_ISUSING
	public FOW_REALTITLE
	
	Function ImgList
		ImgList ="http://webimage.10x10.co.kr/organizer/2009/list/"& FImg
	End Function
	
	Function ImgBasic
		ImgBasic ="http://webimage.10x10.co.kr/organizer/2009/basic/"& FImg
	End Function

	Function Imgcomment
		Imgcomment ="http://webimage.10x10.co.kr/organizer/2009/comment/"& fcomment_img
	End Function
	
	Function ImgIcon
		ImgIcon ="http://webimage.10x10.co.kr/organizer/2009/icon/"& FImg
	End Function
	
	Public Function getContImgUrl()
		getContImgUrl = "http://webimage.10x10.co.kr/organizer/2009/cont/" & ConfImg
	End Function
	
	public Function getInfoImgUrl()
		getInfoImgUrl = "http://webimage.10x10.co.kr/organizer/2009/info/" & Finfoimg
	End Function
	
end class

Class organizerCls

	Public FItemList()
	Public FItem
	public FResultCount
	public FPageSize
	public FCurrPage
	public Ftotalcount
	public FScrollCount
	public FTotalpage
	public FPageCount
	public FOneItem
	public DiaryPrd
	public frecttype	
	public FRectorganizerID
	public frectidx
	public FYearUse
	public FBannerType
	public FEvtUsing
	public frectitemid
	public FRectIsusing
	public FRectPosCode
	public FRectvaliddate
	public frectcate
	public FRectArrItemid
	public FrectMakerid	
	public frectflagdate			
	public frectevt_code	
	
	public FOW_IDX
	public FOW_TITLE
	public FOW_CONTENTS
	public FOW_MAIN_IMG
	public FOW_SUB_IMG
	public FOW_REGDATE
	public FOW_ISUSING
	public FOW_REALTITLE
		
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	'//admin/organizer/event_edit.asp
	Public Sub geteventone()

		dim strSQL,i

		strSQL =" SELECT top 1 "
		strSQL = strSQL & " idx, evt_code, event_type, isusing, idx_order " + vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_event " + vbcrlf
		strSQL = strSQL & " WHERE idx=" & FRectidx

		rsget.open strSQL,dbget,1

		IF  not rsget.EOF  Then
			set FItem = new organizeItemsCls

			FItem.fidx = rsget("idx")
			FItem.fevt_code = rsget("evt_code")
			FItem.fevent_type = rsget("event_type")
			FItem.fisusing = rsget("isusing")
			FItem.fidx_order = rsget("idx_order")			

		End IF

		rsget.close

	End Sub

	'// 다이어리이벤트관리페이지 /admin/organizer/event.asp
	public Sub geteventList()
		dim strSQL,i

		'갯수 새기
		strSQL =" SELECT count(e.idx) as cnt" + vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_event e" + vbcrlf
		strSQL = strSQL & " join [db_event].[dbo].[tbl_event] AS A " + vbcrlf
		strSQL = strSQL & " on e.evt_code = a.evt_code" + vbcrlf
		strSQL = strSQL & " JOIN  [db_event].[dbo].[tbl_event_display] AS B  " + vbcrlf
		strSQL = strSQL & " ON e.evt_code = B.evt_code " + vbcrlf
		strSQL = strSQL & " WHERE A.evt_using ='Y' and e.event_type='organizer' " + vbcrlf
		
		if frectflagdate = "on" then
			strSQL = strSQL & " and getdate() between A.evt_startdate and A.evt_enddate" + vbcrlf
		end if
		if frectevt_code <> "" then
			strSQL = strSQL & " and e.evt_code = '"& frectevt_code &"'" + vbcrlf
		end if	

		'strSQL = strSQL & " and A.evt_state = 7 and B.evt_bannerimg <> '' and A.evt_kind in (1,16) " + vbcrlf

		'response.write strSQL
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

		'데이터 리스트
		strSQL = ""
		strSQL = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		strSQL = strSQL & " e.idx, e.evt_code, e.event_type, e.isusing, e.idx_order" + vbcrlf
		strSQL = strSQL & " ,a.evt_name,A.evt_code, B.evt_bannerimg" + vbcrlf
		strSQL = strSQL & " , A.evt_startdate, A.evt_enddate, A.evt_kind, B.brand " + vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_event e" + vbcrlf
		strSQL = strSQL & " join [db_event].[dbo].[tbl_event] AS A " + vbcrlf
		strSQL = strSQL & " on e.evt_code = a.evt_code" + vbcrlf
		strSQL = strSQL & " JOIN  [db_event].[dbo].[tbl_event_display] AS B  " + vbcrlf
		strSQL = strSQL & " ON e.evt_code = B.evt_code " + vbcrlf
		strSQL = strSQL & " WHERE A.evt_using ='Y' and e.event_type='organizer' " + vbcrlf
		
		if frectflagdate = "on" then
			strSQL = strSQL & " and getdate() between A.evt_startdate and A.evt_enddate" + vbcrlf
		end if
		if frectevt_code <> "" then
			strSQL = strSQL & " and e.evt_code = '"& frectevt_code &"'" + vbcrlf
		end if	
		
		'strSQL = strSQL & " and A.evt_state = 7 and B.evt_bannerimg <> '' and A.evt_kind in (1,16) " + vbcrlf
		strSQL = strSQL & " ORDER BY e.idx_order  DESC" + vbcrlf

		'response.write strSQL
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new organizeItemsCls

				FItemList(i).fidx = rsget("idx")
				FItemList(i).fevt_name = db2html(rsget("evt_name"))
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fevent_type = rsget("event_type")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx_order = rsget("idx_order")
				FItemList(i).fevt_bannerimg = db2html(rsget("evt_bannerimg"))
				FItemList(i).fevt_startdate = rsget("evt_startdate")
				FItemList(i).fevt_enddate = rsget("evt_enddate")
				FItemList(i).fevt_kind = rsget("evt_kind")
				FItemList(i).fbrand = rsget("brand")				

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub

	''//  기본정보 /admin/organizer/option/pop_organizer_info_reg.asp
	public Function getDiaryInfo(byval idx)
		dim strSQL,i
		
	strSQL = "SELECT idx,info_gubun,info_name,info_img,info_pageCnt,search_View,Info_idx" & vbcrlf
	strSQL = strSQL & " FROM [db_diary2010].[dbo].tbl_organizer_info" & vbcrlf
	strSQL = strSQL & " WHERE idx= "& idx &"" & vbcrlf
	strSQL = strSQL & " ORDER BY info_gubun" & vbcrlf

		rsget.open strSQL,dbget ,1

		FResultCount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(FResultCount)
			i=0
			do until rsget.eof
				set FItemList(i) = new organizeItemsCls
				FItemList(i).FYear = FYearUse
				FItemList(i).FInfoidx = rsget("Info_idx")
				FItemList(i).FInfoGubun = rsget("info_gubun")
				FItemList(i).Finfoname = db2html(rsget("info_name"))
				FItemList(i).Finfoimg = rsget("info_img")
				FItemList(i).FinfoPageCnt = rsget("info_PageCnt")
				FItemList(i).fsearch_view = rsget("search_view")			

				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close
	End Function

'// admin/diary2009/option/pop_diary_info_reg.asp
	public sub fsearch_list()
		dim sqlStr,i

		'데이터 리스트 
		sqlStr = "select " + vbcrlf
		sqlStr = sqlStr & " idx,info_name,search_order" + vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_info_search" + vbcrlf			
		sqlStr = sqlStr & " order by search_order desc" + vbcrlf
		
		'response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1
		
		ftotalcount = rsget.recordcount
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
				set FItemList(i) = new organizeItemsCls

				FItemList(i).fidx = rsget("idx")
				FItemList(i).finfo_name = rsget("info_name")
				FItemList(i).fsearch_order = rsget("search_order")
														
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	''// 다이어리 기본정보 /admin/organizer/option/pop_organizer_info_reg.asp
	Public Sub getDiaryItem(byval idx)

		dim strSQL

		strSQL = "select top 1 m.idx ,m.yearuse, m.diaryType ,m.Itemid ,m.basic_img ,m.List_img" & vbcrlf
		strSQL = strSQL & " ,m.icon_img ,m.isusing ,m.giftYn ,onlyYearYn ,m.hitYn," & vbcrlf
		strSQL = strSQL & " mainimage, listimage,listimage120,smallimage, basicimage," & vbcrlf 																																																																					
		strSQL = strSQL & " cate_large, cate_mid, cate_small, i.regdate ,specialuseritem," & vbcrlf 																																																																					
		strSQL = strSQL & " itemsource, sourcearea, itemsize, sellcash, sellyn, limityn, limitno, limitsold," & vbcrlf																																																																					
		strSQL = strSQL & " itemcontent, (cate_large + cate_mid + cate_small) as itemprecode," & vbcrlf																																																																					
		strSQL = strSQL & " itemname, designercomment, makername,deliverytype,itemdiv,orgprice,sailyn," & vbcrlf																																																																					
		strSQL = strSQL & " usinghtml, mileage, deliverarea," & vbcrlf 																																																																					
		strSQL = strSQL & " ordercomment, reipgodate, brandname, ismobileitem," & vbcrlf 																																																																					
		strSQL = strSQL & " evalcnt,optioncnt," & vbcrlf																																																																					
		strSQL = strSQL & " itemcoupontype, itemcouponvalue" & vbcrlf																																																																					
		strSQL = strSQL & " FROM [db_diary_collection].dbo.tbl_diary_master m" & vbcrlf																																																																					
		strSQL = strSQL & " JOIN [db_item].dbo.tbl_item i on m.itemid= i.itemid" & vbcrlf																																																																					
		strSQL = strSQL & " JOIN [db_item].dbo.tbl_item_contents c on m.itemid = c.itemid" & vbcrlf																																																																					
		strSQL = strSQL & " WHERE m.idx="& idx &"" & vbcrlf																																																																					

		'response.write strSQL
		rsget.open strSQL,dbget,1
		if not rsget.eof then

			set DiaryPrd = new organizeItemsCls

			DiaryPrd.FIdx = rsget("idx")
			DiaryPrd.FYear = rsget("yearuse")
			DiaryPrd.FDiaryType = rsget("diaryType")
			DiaryPrd.FItemid = rsget("Itemid")
			DiaryPrd.FBasicImg = rsget("basic_img")
			DiaryPrd.FListimg = rsget("list_img")
			DiaryPrd.FIconImg = rsget("icon_img")
			DiaryPrd.Fisusing = rsget("isusing")
			DiaryPrd.FgiftYn = rsget("giftYn")
			DiaryPrd.FonlyYearYn = rsget("onlyYearYn")
			DiaryPrd.FhitYn = rsget("hitYn")
			DiaryPrd.FItemCouponType 	=	rsget("itemcoupontype")
			DiaryPrd.FItemCouponValue	= rsget("itemcouponvalue")

		end if
		rsget.close

	End Sub

	'// admin/organizer/option/detail_option.asp
	public Sub fkeyword_option_value()
		dim strSQL,i

		
		'데이터 리스트 		

		strSQL ="select a.idx , a.option_value , b.keyword_option_count"
		strSQL = strSQL & " from db_diary2010.dbo.tbl_organizer_keyword_option a"
		strSQL = strSQL & " left join ("
		strSQL = strSQL & " select keyword_option,count(keyword_option) as keyword_option_count"
		strSQL = strSQL & " from db_diary2010.dbo.tbl_organizer_keyword_master"
		strSQL = strSQL & " where organizerid = "& frectorganizerid &""
		strSQL = strSQL & " group by keyword_option"
		strSQL = strSQL & " ) b"
		strSQL = strSQL & " on a.idx = b.keyword_option"
		strSQL = strSQL & " where type = '"& frecttype &"'" 
		strSQL = strSQL & " order by option_order desc" 
					
		'response.write strSQL
		rsget.pagesize = FPageSize		
		rsget.open strSQL,dbget,1
		
		FTotalCount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		
		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new organizeItemsCls
				
					FItemList(i).fidx = rsget("idx")				
					FItemList(i).foption_value = rsget("option_value")
					FItemList(i).fkeyword_option_count = rsget("keyword_option_count")
					
				i=i+1
				rsget.Movenext

			Loop

		End IF
		
		rsget.close
	End Sub 

	'// admin/organizer/alpha_list.asp
	public Sub falpha_value()
		dim sqlStr,i

		
		'데이터 리스트 		
		sqlStr ="select" & vbcrlf 
		sqlStr = sqlStr & " idx,itemid,isusing" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_alpha" & vbcrlf

		if frectidx <> "" then
		sqlStr = sqlStr & " where idx = "& frectidx &"" & vbcrlf
		end if
		
		sqlStr = sqlStr & " order by idx desc" & vbcrlf
					
		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
		ftotalcount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new organizeItemsCls

				FOneItem.fidx = rsget("idx")
				FOneItem.fitemid = rsget("itemid")
				FOneItem.fisusing = rsget("isusing")
			        			        
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/organizer/index.asp	
	public Sub getorganizerList()
		dim strSQL,i

		'갯수 새기		
		strSQL =" SELECT count(a.organizerid) as cnt" + vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_organizerMaster a" + vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" + vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" + vbcrlf
		strSQL = strSQL & " where 1=1  " + vbcrlf
		
		if frectcate <> "" then
		strSQL = strSQL & " and a.cate= '"& frectcate &"'" + vbcrlf
		end if

		if frectisusing <> "" then
		strSQL = strSQL & " and a.isusing= '"& frectisusing &"'" + vbcrlf
		end if
		
		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " + vbcrlf
		End IF

		IF FRectArrItemid<>"" Then
			strSQL = strSQL & " and a.itemid in ("& FRectArrItemid &") " + vbcrlf
		End IF
		
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
				
		'데이터 리스트 		
		strSQL = ""
		strSQL = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		strSQL = strSQL & " a.color,a.organizerid,a.Cate,a.ItemID,a.RegDate,a.isUsing ,a.BasicImg,a.commentyn" + vbcrlf
		strSQL = strSQL & " ,b.itemname" + vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_organizerMaster a" + vbcrlf
		strSQL = strSQL & " left join db_item.dbo.tbl_item b" + vbcrlf
		strSQL = strSQL & " on a.itemid = b.itemid" + vbcrlf
		strSQL = strSQL & " where 1=1  " + vbcrlf
		
		if frectcate <> "" then
		strSQL = strSQL & " and a.cate= '"& frectcate &"'" + vbcrlf
		end if

		if frectisusing <> "" then
		strSQL = strSQL & " and a.isusing= '"& frectisusing &"'" + vbcrlf
		end if

		IF FrectMakerid<>"" Then
			strSQL = strSQL & " and b.makerid= '"& FrectMakerid &"' " + vbcrlf
		End IF

		IF FRectArrItemid<>"" Then
			strSQL = strSQL & " and a.itemid in ("& FRectArrItemid &") " + vbcrlf
		End IF

		strSQL = strSQL & " order by a.organizerid desc" + vbcrlf
				
		'response.write strSQL
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		rsget.PageSize= FPageSize
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new organizeItemsCls

				FItemList(i).fcolor = rsget("color")				
				FItemList(i).forganizerid = rsget("organizerid")
				FItemList(i).FCateCode = rsget("Cate")
				FItemList(i).Fitemid = rsget("ItemID")
				FItemList(i).FRegDate = rsget("RegDate")
				FItemList(i).FisUsing = rsget("isUsing")
				FItemLIst(i).FImg = rsget("BasicImg")
				FItemLIst(i).fcommentyn = rsget("commentyn")
				FItemLIst(i).fitemname = db2html(rsget("itemname"))
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	End Sub 	

'// admin/organizer/imagemake/imagemake_poscode.asp
    public Sub fposcode_oneitem()		
        dim SqlStr
        SqlStr = "select"
		sqlStr = sqlStr & " poscode,posname,imagetype,imagewidth,imageheight,isusing,imagecount" + vbcrlf        
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode" + vbcrlf
		sqlStr = sqlStr & " where 1=1" + vbcrlf
        SqlStr = SqlStr + " and poscode=" + CStr(FRectPoscode)
         
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new organizeItemsCls
        if Not rsget.Eof then
            
            FOneItem.fposcode = rsget("poscode")
            FOneItem.fposname = db2html(rsget("posname"))
            FOneItem.fimagetype	= rsget("imagetype")
            FOneItem.fimagewidth = rsget("imagewidth")
            FOneItem.fimageheight = rsget("imageheight")
            FOneItem.fisusing = rsget("isusing")
            FOneItem.fimagecount = rsget("imagecount")
                       
        end if
        rsget.close
    end Sub  

'// admin/organizer/imagemake/imagemake_contents.asp   
    public Sub fcontents_oneitem()
        dim sqlStr
        sqlStr = "select top 1" & vbcrlf
		sqlStr = sqlStr & " a.posname,a.imagetype,a.imagewidth,a.imageheight,a.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,b.idx,b.imagepath,b.linkpath,b.evt_code,b.regdate,b.poscode,b.isusing,b.image_order,b.itemtype" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_organizer_poscode_image b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
        sqlStr = sqlStr & " where 1=1" & vbcrlf
        sqlStr = sqlStr & " and b.idx = "& FRectIdx&""

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount
        
        set FOneItem = new organizeItemsCls
        
        if Not rsget.Eof then
    
			FOneItem.fposcode = rsget("poscode")
			FOneItem.fposname = db2html(rsget("posname"))
			FOneItem.fimagetype = rsget("imagetype")
			FOneItem.fimagewidth = rsget("imagewidth")
			FOneItem.fimageheight = rsget("imageheight")
			FOneItem.fisusing = rsget("isusing")
			FOneItem.fidx = rsget("idx")
			FOneItem.fimagepath = db2html(rsget("imagepath"))
			FOneItem.flinkpath = db2html(rsget("linkpath"))
			FOneItem.fevt_code = rsget("evt_code")
			FOneItem.fregdate = rsget("regdate")
			FOneItem.fimagecount = rsget("imagecount") 
			FOneItem.fimage_order = rsget("image_order") 
			FOneItem.fitemtype = rsget("itemtype") 
            
        end if
        rsget.Close
    end Sub

	'// admin/organizer/option/detail_option.asp
	public Sub fkeyword_type()
		dim strSQL,i

		
		'데이터 리스트 		
			
		strSQL ="select type from db_diary2010.dbo.tbl_organizer_keyword_option"
		strSQL = strSQL & " group by type"

		'response.write strSQL & "<br>"
		rsget.pagesize = FPageSize		
		rsget.open strSQL,dbget,1
		
		FTotalCount = rsget.recordcount
		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		
		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).ftype = rsget("type")

				i=i+1
				rsget.Movenext

			Loop

		End IF
		
		rsget.close
	End Sub 

	'// admin/organizer/organizerreg.asp
	Public Sub getDiary()
		
		dim strSQL,i
		
		strSQL =" SELECT top 1 organizerid,Cate,ItemID,RegDate,isUsing ,BasicImg,commentyn , event_code ,eventgroup_code,weight,color,organizer_order "
		strSQL = strSQL & " ,comment_img , event_start ,event_end " + vbcrlf
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_organizerMaster " + vbcrlf
		strSQL = strSQL & " WHERE organizerId=" & FRectorganizerID
		
		rsget.open strSQL,dbget,1
		
		IF  not rsget.EOF  Then
			set FItem = new organizeItemsCls
				
			FItem.forganizerid = rsget("organizerid")
			FItem.FCateCode = rsget("Cate")
			FItem.Fitemid = rsget("ItemID")
			FItem.FRegDate = rsget("RegDate")
			FItem.FisUsing = rsget("isUsing")
			FItem.FImg = rsget("BasicImg")
			FItem.fcommentyn = rsget("commentyn")
			FItem.fevent_code = rsget("event_code")
			FItem.feventgroup_code = rsget("eventgroup_code")
			FItem.fcomment_img = rsget("comment_img")			
			FItem.fevent_code = rsget("event_code")
			FItem.feventgroup_code = rsget("eventgroup_code")	
			FItem.fevent_start = rsget("event_start")
			FItem.fevent_end = rsget("event_end")
			FItem.fweight	= rsget("weight")
			FItem.fcolor	= rsget("color")
			FItem.forganizer_order	= rsget("organizer_order")
								
		End IF	
		rsget.close
	End Sub

	'// admin/organizer/option/keyword_option.asp
	public Sub falpha()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "select" + vbcrlf  
		strSQL = strSQL & " count(idx) as cnt" + vbcrlf 
		strSQL = strSQL & " from db_diary2010.dbo.tbl_organizer_alpha" + vbcrlf 
				
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		
		strSQL ="select top " & Cstr(FPageSize * FCurrPage) + vbcrlf  
		strSQL = strSQL & " idx,itemid,isusing" & vbcrlf
		strSQL = strSQL & " from db_diary2010.dbo.tbl_organizer_alpha" & vbcrlf
		strSQL = strSQL & " order by idx desc" & vbcrlf

		rsget.pagesize = FPageSize		
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		
		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fitemid = rsget("itemid")
				FItemList(i).fisusing = rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End IF
		
		rsget.close
	End Sub 

	'// admin/organizer/option/keyword_option.asp
	public Sub fkeyword_option()
		dim strSQL,i

		'총 갯수 구하기
		strSQL = "select" + vbcrlf  
		strSQL = strSQL & " count(idx) as cnt" + vbcrlf 
		strSQL = strSQL & " from db_diary2010.dbo.tbl_organizer_keyword_option" + vbcrlf 
				
		rsget.Open strSQL,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		
		strSQL ="select top " & Cstr(FPageSize * FCurrPage) + vbcrlf  
		strSQL = strSQL & " idx,option_value,option_order,type,isusing" + vbcrlf  
		strSQL = strSQL & " FROM db_diary2010.dbo.tbl_organizer_keyword_option" + vbcrlf  

		rsget.pagesize = FPageSize		
		rsget.open strSQL,dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1
		
		IF  not rsget.EOF  Then

			i=0

			Do Until rsget.eof
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).foption_value = rsget("option_value")
				FItemList(i).foption_order = rsget("option_order")
				FItemList(i).ftype = rsget("type")
				FItemList(i).fisusing = rsget("isusing")

				i=i+1
				rsget.Movenext

			Loop

		End IF
		
		rsget.close
	End Sub 

	'// admin/organizer/option/keyword_option.asp
	public sub fkeyword_option_edit()
		dim sqlStr,i

		'데이터 리스트 
		sqlStr = "select top 1" + vbcrlf 
		sqlStr = sqlStr & " idx,option_value,option_order,type,isusing" + vbcrlf 
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_keyword_option" + vbcrlf 
		sqlStr = sqlStr & " where idx = '"& frectidx &"'" + vbcrlf 

		'response.write sqlStr &"<br>"

		rsget.Open sqlStr,dbget,1
		ftotalcount = rsget.recordcount

		i=0
		if  not rsget.EOF  then

			do until rsget.EOF
				set FOneItem = new organizeItemsCls

				FOneItem.fidx = rsget("idx")
				FOneItem.foption_value = rsget("option_value")
				FOneItem.foption_order = rsget("option_order")
				FOneItem.ftype = rsget("type")
				FOneItem.fisusing = rsget("isusing")
			        			        
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

'// admin/organizer/imagemake/imagemake_poscode.asp
	public sub fposcode_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select" + vbcrlf
		sqlStr = sqlStr & " count(poscode) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode" + vbcrlf
					
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		

		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " poscode,isusing,posname,imagetype,imagewidth,imageheight,imagecount" + vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode" + vbcrlf			
		sqlStr = sqlStr & " where 1=1" + vbcrlf

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
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fimagecount = rsget("imagecount")
														
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// admin/organizer/imagemake/imagemake_list.asp
	public sub fcontents_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_organizer_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item as i " & vbcrlf
		sqlStr = sqlStr & " on a.evt_code=i.itemid " & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"& FRectIsusing &"'" & vbcrlf		
			end if	

			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if					
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.posname,b.imagetype,b.imagewidth,b.imageheight,b.imagecount" & vbcrlf
		sqlStr = sqlStr & " ,a.idx,a.imagepath,a.linkpath,a.evt_code,a.regdate,a.poscode,a.isusing,a.image_order" & vbcrlf
		sqlStr = sqlStr & " , i.ListImage,i.ListImage120,i.SmallImage,icon1image,i.icon2image" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_poscode_image a" & vbcrlf
		sqlStr = sqlStr & " left join db_diary2010.dbo.tbl_organizer_poscode b" & vbcrlf
		sqlStr = sqlStr & " on a.poscode = b.poscode" & vbcrlf	
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item as i " & vbcrlf
		sqlStr = sqlStr & " on a.evt_code=i.itemid " & vbcrlf
        sqlStr = sqlStr & " where 1=1" & vbcrlf

			if FRectIsusing <> "" then
				sqlStr = sqlStr & " and a.isusing = '"&FRectIsusing&"'" & vbcrlf		
			end if	
			if FRectPosCode <> "" then
				sqlStr = sqlStr & " and a.poscode = "& FRectPosCode &"" & vbcrlf		
			end if	

		sqlStr = sqlStr & " order by a.image_order Desc" + vbcrlf

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
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fposcode = rsget("poscode")
				FItemList(i).fposname = db2html(rsget("posname"))
				FItemList(i).fimagetype = rsget("imagetype")
				FItemList(i).fimagewidth = rsget("imagewidth")
				FItemList(i).fimageheight = rsget("imageheight")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fimagepath = rsget("imagepath")
				FItemList(i).flinkpath = rsget("linkpath")
				FItemList(i).fevt_code = rsget("evt_code")
				FItemList(i).fregdate = rsget("regdate")		
				FItemList(i).fimagecount = rsget("imagecount")
				FItemList(i).fimage_order = rsget("image_order")													
				FItemList(i).FImageList	= "http://webimage.10x10.co.kr/image/list/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("ListImage"))
				FItemList(i).FImageList120	= "http://webimage.10x10.co.kr/image/list120/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & db2html(rsget("ListImage120"))
				FItemList(i).FImageSmall	= "http://webimage.10x10.co.kr/image/small/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" &db2html(rsget("smallImage"))
				FItemList(i).FImageicon1 = "http://webimage.10x10.co.kr/image/icon1/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon1image")
				FItemList(i).FImageicon2 = "http://webimage.10x10.co.kr/image/icon2/" & GetImageSubFolderByItemid(FItemList(i).fevt_code) & "/" & rsget("icon2image")
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// admin/organizer/eval_list.asp
	public sub geteval_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.idx) as cnt" + vbcrlf   
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_eval_list a" & vbcrlf
		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate b" & vbcrlf
		sqlStr = sqlStr & " on a.Eval_idx = b.idx" & vbcrlf
		sqlStr = sqlStr & " where a.isusing = 'Y'" & vbcrlf
						
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " a.idx ,b.regdate ,a.Eval_idx ,a.organizerid , b.contents,b.userid,a.isusing" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizer_eval_list a" & vbcrlf
		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate b" & vbcrlf
		sqlStr = sqlStr & " on a.Eval_idx = b.idx" & vbcrlf
		sqlStr = sqlStr & " where a.isusing = 'Y'" & vbcrlf
		sqlStr = sqlStr & " order by a.idx desc" & vbcrlf
		
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
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fContents = db2html(rsget("Contents"))
				FItemList(i).fUserID = rsget("UserID")				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).forganizerid = rsget("organizerid")
				FItemList(i).fisusing = rsget("isusing")
																												
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'// admin/organizer/eval_list.asp
	public sub feval_list()
		dim sqlStr,i

		'총 갯수 구하기
		sqlStr = "select count(a.itemid) as cnt" + vbcrlf 
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizerMaster a" & vbcrlf
		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate b" & vbcrlf
		sqlStr = sqlStr & " on a.itemid = b.itemid" & vbcrlf
		sqlStr = sqlStr & " where b.IsUsing = 'Y'" & vbcrlf
		
		if frectitemid <> "" then
		sqlStr = sqlStr & " and b.itemid = '" & frectitemid & "'" & vbcrlf			
		end if
		
		rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close
		
		'데이터 리스트 
		sqlStr = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sqlStr = sqlStr & " b.idx ,b.Contents ,b.UserID, b.itemid " & vbcrlf
		sqlStr = sqlStr & " , b.regdate , a.organizerid" & vbcrlf
		sqlStr = sqlStr & " from db_diary2010.dbo.tbl_organizerMaster a" & vbcrlf
		sqlStr = sqlStr & " left join db_board.dbo.tbl_Item_Evaluate b" & vbcrlf
		sqlStr = sqlStr & " on a.itemid = b.itemid" & vbcrlf
		sqlStr = sqlStr & " where b.IsUsing = 'Y'" & vbcrlf

		if frectitemid <> "" then
		sqlStr = sqlStr & " and b.itemid = '" & frectitemid & "'" & vbcrlf			
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
				set FItemList(i) = new organizeItemsCls
				
				FItemList(i).fidx = rsget("idx")
				FItemList(i).fContents = db2html(rsget("Contents"))
				FItemList(i).fUserID = rsget("UserID")
				FItemList(i).fitemid = rsget("itemid")				
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).forganizerid = rsget("organizerid")
																								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub


    public Sub FOneWeekOffice()
        dim sqlStr
        
        If FOW_IDX = "" Then
	        sqlStr = "SELECT TOP 1 " +vbcrlf
			sqlStr = sqlStr & " 	IDX, REALTITLE, TITLE, CONTENTS, MAIN_IMG, SUB_IMG, ISUSING, REGDATE" + vbcrlf
			sqlStr = sqlStr & " From db_diary2010.dbo.tbl_OneWeekOffice " + vbcrlf
	        sqlStr = sqlStr & " ORDER BY IDX DESC " + vbcrlf
	    ELSE
	        sqlStr = "SELECT " +vbcrlf
			sqlStr = sqlStr & " 	IDX, REALTITLE, TITLE, CONTENTS, MAIN_IMG, SUB_IMG, ISUSING, REGDATE" + vbcrlf
			sqlStr = sqlStr & " From db_diary2010.dbo.tbl_OneWeekOffice " + vbcrlf
	        sqlStr = sqlStr & " WHERE IDX = "& FOW_IDX&"" + vbcrlf
		End IF

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new organizeItemsCls

        if Not rsget.Eof then

			FOneItem.FOW_IDX = rsget("IDX")
			FOneItem.FOW_REALTITLE = db2html(rsget("REALTITLE"))
			FOneItem.FOW_TITLE = db2html(rsget("TITLE"))
			FOneItem.FOW_CONTENTS = db2html(rsget("CONTENTS"))
			FOneItem.FOW_MAIN_IMG = rsget("MAIN_IMG")
			FOneItem.FOW_SUB_IMG = rsget("SUB_IMG")
			FOneItem.FOW_REGDATE = rsget("REGDATE")
			FOneItem.FOW_ISUSING = rsget("ISUSING")

        end if
        rsget.Close
    end Sub
    
    
	public sub FOneWeekOfficeList()
		dim strSQL,i

		strSQL = "SELECT " & vbcrlf
		strSQL = strSQL & " 	IDX, REALTITLE, TITLE, CONTENTS, REGDATE" & vbcrlf
		strSQL = strSQL & " From db_diary2010.dbo.tbl_OneWeekOffice " & vbcrlf
		strSQL = strSQL & " ORDER BY IDX DESC" & vbcrlf

		rsget.open strSQL,dbget ,1

		ftotalcount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(ftotalcount)
			i=0
			do until rsget.eof
				set FItemList(i) = new organizeItemsCls

				FItemList(i).FOW_IDX = rsget("IDX")
				FItemList(i).FOW_REALTITLE = db2html(rsget("REALTITLE"))
				FItemList(i).FOW_TITLE = db2html(rsget("TITLE"))
				FItemList(i).FOW_CONTENTS = db2html(rsget("CONTENTS"))
				FItemList(i).FOW_REGDATE = db2html(rsget("REGDATE"))

				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close
	End sub



    public Sub FOrgStory()
        dim sqlStr
        
        If FOW_IDX = "" Then
	        sqlStr = "SELECT TOP 1 " +vbcrlf
			sqlStr = sqlStr & " 	IDX, TITLE, CONTENTS, ISUSING, REGDATE" + vbcrlf
			sqlStr = sqlStr & " From db_diary2010.dbo.tbl_organizer_story " + vbcrlf
	        sqlStr = sqlStr & " ORDER BY IDX DESC " + vbcrlf
	    ELSE
	        sqlStr = "SELECT " +vbcrlf
			sqlStr = sqlStr & " 	IDX, TITLE, CONTENTS, ISUSING, REGDATE" + vbcrlf
			sqlStr = sqlStr & " From db_diary2010.dbo.tbl_organizer_story " + vbcrlf
	        sqlStr = sqlStr & " WHERE IDX = "& FOW_IDX&"" + vbcrlf
		End IF

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        FResultCount = rsget.RecordCount

        set FOneItem = new organizeItemsCls

        if Not rsget.Eof then

			FOneItem.FOW_IDX = rsget("IDX")
			FOneItem.FOW_TITLE = db2html(rsget("TITLE"))
			FOneItem.FOW_CONTENTS = db2html(rsget("CONTENTS"))
			FOneItem.FOW_REGDATE = rsget("REGDATE")
			FOneItem.FOW_ISUSING = rsget("ISUSING")

        end if
        rsget.Close
    end Sub
    
    
	public sub FOrgStoryList()
		dim strSQL,i

		strSQL = "SELECT " & vbcrlf
		strSQL = strSQL & " 	IDX, TITLE, CONTENTS, ISUSING, REGDATE" & vbcrlf
		strSQL = strSQL & " From db_diary2010.dbo.tbl_organizer_story " & vbcrlf
		strSQL = strSQL & " ORDER BY IDX DESC" & vbcrlf

		rsget.open strSQL,dbget ,1

		ftotalcount = rsget.recordcount

		if not rsget.eof then

			redim preserve FItemList(ftotalcount)
			i=0
			do until rsget.eof
				set FItemList(i) = new organizeItemsCls

				FItemList(i).FOW_IDX = rsget("IDX")
				FItemList(i).FOW_TITLE = db2html(rsget("TITLE"))
				FItemList(i).FOW_CONTENTS = db2html(rsget("CONTENTS"))
				FItemList(i).FOW_REGDATE = db2html(rsget("REGDATE"))

				rsget.movenext
				i = i+1
			loop
		end if
		rsget.close
	End sub


	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end class	

'// 다이어리 종류 분류 셀렉트박스
Function SelectList(selName,selVal)
	dim OptArray (7,2),i
	OptArray(0,0)=""
	OptArray(0,1)="전체"
	OptArray(1,0)="10"
	OptArray(1,1)="moleskine"
	OptArray(2,0)="20"
	OptArray(2,1)="frankline"
	OptArray(3,0)="30"
	OptArray(3,1)="innoworks"
	OptArray(4,0)="40"
	OptArray(4,1)="orom"
	OptArray(5,0)="50"
	OptArray(5,1)="matt"
	OptArray(6,0)="60"
	OptArray(6,1)="midori/기타"
	
	response.write "<select name="""& selName &""">" 
	for i=0 To Ubound(OptArray,1)-1
		response.write "<option value='" &OptArray(i,0)&"'"
		IF OptArray(i,0) = selVal THEN 
			response.write " selected"
		End IF
		response.write ">"& OptArray(i,1) &"</option>"
	next
	response.write "</select>"

End Function 

'// 다이어리 종류 분류 일반
Function cateList(selName,selVal)
	dim OptArray (5,2),i
	OptArray(0,0)="10"
	OptArray(0,1)="moleskine"
	OptArray(1,0)="20"
	OptArray(1,1)="frankline"
	OptArray(2,0)="30"
	OptArray(2,1)="innoworks"
	OptArray(3,0)="40"
	OptArray(3,1)="orom"
	OptArray(4,0)="50"
	OptArray(4,1)="matt"
	OptArray(5,0)="60"
	OptArray(5,1)="midori/기타"
	
	for i=0 To 5
		IF OptArray(i,0) = selVal THEN 
			response.write OptArray(i,1)
		End IF
	next
End Function 

function DrawMainPosCodeCombo(selectBoxName,selectedId,changeFlag)
   dim tmp_str,query1
   %>
   <select name="<%=selectBoxName%>" <%= changeFlag %>>
     <option value='' <%if selectedId="" then response.write " selected"%> >전체</option>
   <%
   query1 = " select poscode,posname from db_diary2010.dbo.tbl_organizer_poscode"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("poscode")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("poscode")&"' "&tmp_str&">" + db2html(rsget("posname")) + "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
end function
%>	
	