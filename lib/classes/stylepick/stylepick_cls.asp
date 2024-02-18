<%
'###########################################################
' Description : 스타일픽 클래스
' Hieditor : 2011.04.05 한용민 생성
'###########################################################

Class cstylepick_item
	public fitemidx
	public fitemid
	public fregdate
	public fisusing
	public fcd1
	public fcd2
	public fcd3
	public forderno
	public flastadminid
	public fcatename
	public fcd1name
	public fcd2name
	public fcd3name
	public FMakerid
    public Fitemdiv
    public Fitemgubun
    public Fitemname
    public Fsellcash
    public Fbuycash
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fmileage
    public Flastupdate
    public FsellEndDate
    public Fsellyn
    public Flimityn
    public Fdanjongyn
    public Fsailyn
    public Fisextusing
    public Fmwdiv
    public Fspecialuseritem
    public Fvatinclude
    public Fdeliverytype
    public Fdeliverarea
    public Fdeliverfixday
    public Fismobileitem
    public Fpojangok
    public Flimitno
    public Flimitsold
    public Fevalcnt
    public Foptioncnt
    public Fitemrackcode
    public Fupchemanagecode    
    public Fbrandname
    public Fsmallimage
    public Flistimage
    public Flistimage120
    public Fitemcouponyn
    public Fcurritemcouponidx
    public Fitemcoupontype
    public Fitemcouponvalue    
    public FdefaultFreeBeasongLimit   
    public FdefaultDeliverPay         
    public FdefaultDeliveryType   
    public Fevtidx
    public Ftitle
    public Fsubcopy
    public Fstate
    public Fbanner_img
    public Fstartdate
    public Fenddate
    public Fcomment
    public Fevtitemidx
    public Fstatename
    public fopendate
    public fclosedate
    public fpartMDid
    public fpartwDid
    public fevtitemcnt
    public fpartMDname
    public fpartwDname   
    public fmaincontentscnt
    public fmainidx
    public fmainimage
    public fmainimagelink
    public fcontentsyn
    public fmainctidx
    public fgubun
    public fgubunvalue
    public fcopy
    public flink
    public fcd1pre
    public fcd1next
    public fmainidxmin
    public fmainidxmax
    public fmainidxpre
    public fmainidxnext
    public fRowNum
            
    public function GetLimitEa()
		if FLimitNo-FLimitSold<0 then
			GetLimitEa = 0
		else
			GetLimitEa = FLimitNo-FLimitSold
		end if
	end function
	
    public Function IsSoldOut()
		IsSoldOut = (FSellYn<>"Y") or ((FLimitYn="Y") and (GetLimitEa()<1))
	end function
	
    public Function IsUpcheBeasong()
		if Fdeliverytype="2" or Fdeliverytype="5" or Fdeliverytype="9" or Fdeliverytype="7" then
			IsUpcheBeasong = true
		else
			IsUpcheBeasong = false
		end if
	end function
    				
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class cstylepick
	Public FItemList()
	public foneitem
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	public FPageCount
	Public FCurrPage
	Public FPageSize
	public FTotalPage
	public frectcd1
	public frectcd2
	public frectcd3
	public frectisusing
	public frectcateidx
	public frectcatetype
	public FRectSortDiv
	public frectmakerid
	public FRectItemid
	public FRectItemName
	public FRectSellYN	
	public FRectDanjongyn
	public FRectLimityn
	public FRectMWDiv
	public FRectDeliveryType
	public FRectSailYn
	public FRectCouponYn
	public frectstate
	public frectevtidx
	public frectmainidx
	public frecttitle
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small		
	public frectoverlap
	public frectview
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0	
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'/admin/stylepick/stylepick_event.asp
	public function fnGetmainctList()
        dim sqlStr, sqlsearch, i

		if frectmainidx <> "" then
			sqlsearch = sqlsearch & " and mainidx="&frectmainidx&""
		end if	
	
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and isusing='" + FRectIsUsing + "'"
        end if
					
        '// 본문 내용 접수
		sqlStr = "select"
		sqlStr = sqlStr & " mainctidx ,mainidx ,gubun ,gubunvalue ,isusing ,copy ,link ,lastadminid"
		sqlStr = sqlStr & " from db_giftplus .dbo.tbl_stylepick_main_contents"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by mainctidx asc"
		
		'response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
        
		FResultCount = rsget.RecordCount
		
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new cstylepick_item
                
				FItemList(i).fmainctidx = rsget("mainctidx")
				FItemList(i).fmainidx = rsget("mainidx")
				FItemList(i).fgubun = rsget("gubun")
				FItemList(i).fgubunvalue = rsget("gubunvalue")
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fcopy = db2html(rsget("copy"))
				FItemList(i).flink = db2html(rsget("link"))
				FItemList(i).flastadminid = rsget("lastadminid")
				                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'/admin/stylepick/index_testview.asp
	Public Sub fnGetmain_pageing()
		dim sqlstr,i , sqlsearch
			
		sqlstr = "exec db_giftplus.dbo.ten_stylepick_main_pageing '"&frectmainidx&"','"&frectcd1&"'"
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		
		'response.write sqlstr
		rsget.Open sqlstr, dbget, 1
        ftotalcount = rsget.RecordCount
       
        set FOneItem = new cstylepick_item
        
        if Not rsget.Eof then
		
			foneitem.fRowNum = rsget("RowNum")
			foneitem.fmainidxmin = rsget("mainidxmin")
			foneitem.fmainidxmax = rsget("mainidxmax")
			foneitem.fmainidxpre = rsget("mainidxpre")
			foneitem.fmainidxnext = rsget("mainidxnext")						
			foneitem.fcontentsyn = rsget("contentsyn")
			foneitem.fmainidx = rsget("mainidx")
			foneitem.fcd1 = rsget("cd1")
			foneitem.fmainimage = db2html(rsget("mainimage"))
			foneitem.fmainimagelink = db2html(rsget("mainimagelink"))
			foneitem.fstate = rsget("state")
			foneitem.fstartdate = db2html(rsget("startdate"))
			foneitem.fenddate = db2html(rsget("enddate"))
			foneitem.fisusing = rsget("isusing")
			foneitem.fregdate = rsget("regdate")
			foneitem.flastadminid = rsget("lastadminid")
			foneitem.fopendate = rsget("opendate")
			foneitem.fclosedate = rsget("closedate")
			foneitem.fpartMDid = rsget("partMDid")
			foneitem.fpartWDid = rsget("partWDid")

        end if
        rsget.Close
    end Sub

	'/admin/stylepick/stylepick_main_edit.asp
	Public Sub fnGetmain_item()
		dim sqlstr,i , sqlsearch
		
		if frectmainidx <> "" then
			sqlsearch = sqlsearch & " and m.mainidx = "&frectmainidx&""
		end if
		
		sqlstr = "select top 1"
		sqlStr = sqlStr & " m.mainidx,m.cd1,m.mainimage,m.state,m.startdate,m.enddate,m.isusing,m.regdate"
		sqlStr = sqlStr & " ,m.lastadminid,m.opendate,m.closedate,m.partMDid,m.partWDid,m.mainimagelink"
		sqlStr = sqlStr & " ,m.contentsyn, m.comment,c1.catename"
		sqlStr = sqlStr & " ,(select top 1 cd1"
		sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylepick_cate_cd1"
		sqlStr = sqlStr & " 	where orderno < c1.orderno and isusing='Y'"
		sqlStr = sqlStr & " 	order by orderno desc"
		sqlStr = sqlStr & " 	) as cd1pre"
		sqlStr = sqlStr & " ,(select top 1 cd1"
		sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylepick_cate_cd1"
		sqlStr = sqlStr & " 	where orderno > c1.orderno and isusing='Y'"
		sqlStr = sqlStr & " 	order by orderno asc"
		sqlStr = sqlStr & " 	) as cd1next"		
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_main m"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on m.cd1 = c1.cd1 and c1.isusing='Y'"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cstylepick_item
        
        if Not rsget.Eof then

			foneitem.fcd1pre = rsget("cd1pre")
			foneitem.fcd1next = rsget("cd1next")			
			foneitem.fcomment = db2html(rsget("comment"))
			foneitem.fcontentsyn = rsget("contentsyn")
			foneitem.fmainidx = rsget("mainidx")
			foneitem.fcd1 = rsget("cd1")
			foneitem.fmainimage = db2html(rsget("mainimage"))
			foneitem.fmainimagelink = db2html(rsget("mainimagelink"))
			foneitem.fstate = rsget("state")
			foneitem.fstartdate = db2html(rsget("startdate"))
			foneitem.fenddate = db2html(rsget("enddate"))
			foneitem.fisusing = rsget("isusing")
			foneitem.fregdate = rsget("regdate")
			foneitem.flastadminid = rsget("lastadminid")
			foneitem.fopendate = rsget("opendate")
			foneitem.fclosedate = rsget("closedate")
			foneitem.fpartMDid = rsget("partMDid")
			foneitem.fpartWDid = rsget("partWDid")																																								
			foneitem.fcatename = db2html(rsget("catename"))

        end if
        rsget.Close
    end Sub
    
	'/admin/stylepick/stylepick_event.asp
	public function fnGetmainList()
        dim sqlStr, sqlsearch, i

		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and m.cd1='"&frectcd1&"'"
		end if	
	
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and m.isusing='" + FRectIsUsing + "'"
        end if
		
		If frectstate <> "" THEN
			IF frectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and m.state = 7 and getdate() <= m.startdate"
			ELSEIF frectstate = 7 THEN	'오픈진행중
				sqlsearch  = sqlsearch & " and m.state = 7 and getdate() between m.startdate and m.enddate"			
			elseIF frectstate = 9 THEN	'종료
				sqlsearch  = sqlsearch & " and (m.state = 9 or getdate() >= m.enddate)"
			ELSE
				sqlsearch  = sqlsearch & " and  m.state = "&frectstate & ""
			END IF
		End If
        
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
		sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_main m"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
					
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " m.mainidx,m.cd1,m.mainimage,m.state,m.startdate,m.enddate,m.isusing,m.regdate"
		sqlStr = sqlStr & " ,m.lastadminid,m.opendate,m.closedate,m.partMDid,m.partWDid,c1.catename"
		sqlStr = sqlStr & " ,(case when m.state = 7 and getdate() <= m.startdate then '6'" + vbcrlf
		sqlStr = sqlStr & " 	when m.state = 7 and getdate() between m.startdate and m.enddate then '7'" + vbcrlf
		sqlStr = sqlStr & " 	when m.state = 9 or getdate() >= m.enddate then '9'" + vbcrlf
		sqlStr = sqlStr & " 	else m.state end) as statename" + vbcrlf        
        sqlStr = sqlStr & " ,isnull((select count(mainctidx)"
        sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylepick_main_contents"
        sqlStr = sqlStr & " 	where isusing='Y' and m.mainidx = mainidx),0) as maincontentscnt"
        sqlStr = sqlStr & "	,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = m.partMDid) as partMDname"
        sqlStr = sqlStr & " ,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = m.partWDid) as partWDname"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_main m"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on m.cd1 = c1.cd1 and c1.isusing='Y'"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by m.mainidx desc"
		
		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new cstylepick_item
                
				FItemList(i).fmainidx = rsget("mainidx")
				FItemList(i).fcd1 = rsget("cd1")
				FItemList(i).fmainimage = db2html(rsget("mainimage"))
				FItemList(i).fstate = rsget("state")
				FItemList(i).fstartdate = db2html(rsget("startdate"))
				FItemList(i).fenddate = db2html(rsget("enddate"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).fregdate = rsget("regdate")
				FItemList(i).flastadminid = rsget("lastadminid")
				FItemList(i).fopendate = rsget("opendate")
				FItemList(i).fclosedate = rsget("closedate")
				FItemList(i).fpartMDid = rsget("partMDid")
				FItemList(i).fpartWDid = rsget("partWDid")																																								
				FItemList(i).fmaincontentscnt = rsget("maincontentscnt")
				FItemList(i).fcatename = db2html(rsget("catename"))
				FItemList(i).fstatename = rsget("statename")
				FItemList(i).fpartMDname = db2html(rsget("partMDname"))
				FItemList(i).fpartWDname = db2html(rsget("partWDname"))
				                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'/admin/stylepick/stylepick_event_item.asp
	public function GetevtItemList()
        dim sqlStr, sqlsearch, i

		if frectevtidx <> "" then
			sqlsearch = sqlsearch & " and ei.evtidx='"&frectevtidx&"'"
		end if
		
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            sqlsearch = sqlsearch & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and ei.isusing='" + FRectIsUsing + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectMWDiv="MW" then
            sqlsearch = sqlsearch + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectSailYn<>"" then
            sqlsearch = sqlsearch + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            sqlsearch = sqlsearch + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  sqlsearch = sqlsearch + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event_item ei"
        sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylepick_event e"
        sqlStr = sqlStr & " 	on ei.evtidx = e.evtidx"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on ei.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on i.makerid=c.userid"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'"	
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
		
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)                
        sqlStr = sqlStr & " ei.evtitemidx ,ei.evtidx ,ei.itemid ,ei.regdate ,ei.isusing"
        sqlStr = sqlStr & " ,e.evtidx,e.title,e.subcopy,e.state,e.banner_img,e.startdate,e.enddate"
        sqlStr = sqlStr & " ,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"        
        sqlStr = sqlStr & " ,(c1.catename) as 'cd1name'"
		sqlStr = sqlStr & " ,i.makerid,i.itemdiv,i.itemgubun,i.itemname,i.sellcash,i.buycash,i.orgprice,i.orgsuplycash"	
		sqlStr = sqlStr & " ,i.sailprice,i.sailsuplycash,i.mileage,i.sellEndDate,i.sellyn,i.limityn,i.danjongyn,i.sailyn"	
		sqlStr = sqlStr & " ,i.isextusing,i.mwdiv,i.specialuseritem,i.vatinclude,i.deliverytype,i.deliverarea,i.deliverfixday"	
		sqlStr = sqlStr & " ,i.ismobileitem,i.pojangok,i.limitno,i.limitsold,i.evalcnt,i.optioncnt,i.itemrackcode"	
		sqlStr = sqlStr & " ,i.upchemanagecode,i.brandname,i.smallimage,i.listimage,i.listimage120,i.itemcouponyn"	
		sqlStr = sqlStr & " ,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue,i.itemscore,i.lastupdate"
        sqlStr = sqlStr & " ,IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event_item ei"
        sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylepick_event e"
        sqlStr = sqlStr & " 	on ei.evtidx = e.evtidx"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on ei.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on i.makerid=c.userid"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'"	
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch
		
		IF FRectSortDiv="ne" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="hp" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="lp" Then
			sqlStr = sqlStr & " Order by i.SellCash asc"
		ELSEIF FRectSortDiv="be" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new cstylepick_item

				FItemList(i).fevtitemidx            = rsget("evtitemidx")				               
                FItemList(i).fevtidx            = rsget("evtidx")
                FItemList(i).fcd1            = rsget("cd1")
                FItemList(i).fcd1name = db2html(rsget("cd1name"))
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
                FItemList(i).Fsmallimage        = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
	'/admin/stylepick/stylepick_event_edit.asp		'//admin/stylepick/stylepick_event_item.asp
	Public Sub fnGetEvent_item()
		dim sqlstr,i , sqlsearch
		
		if frectevtidx <> "" then
			sqlsearch = sqlsearch & " and evtidx = "&frectevtidx&""
		end if
		
		sqlstr = "select top 1"
        sqlStr = sqlStr & " e.evtidx,e.title,e.subcopy,e.state,e.banner_img,e.startdate,e.enddate"
        sqlStr = sqlStr & " ,e.isusing,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"
        sqlStr = sqlStr & " ,e.partMDid ,e.partWDid ,e.closedate,c1.catename"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event e"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'" 
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cstylepick_item
        
        if Not rsget.Eof then
			
			foneitem.fcatename            = db2html(rsget("catename"))
            foneitem.fopendate            = db2html(rsget("opendate"))
            foneitem.fclosedate            = db2html(rsget("closedate"))
            foneitem.fevtidx            = rsget("evtidx")
			foneitem.ftitle            = db2html(rsget("title"))
            foneitem.fsubcopy            = db2html(rsget("subcopy"))
			foneitem.fstate            = rsget("state")
            foneitem.fbanner_img            = db2html(rsget("banner_img"))
			foneitem.fstartdate            = db2html(rsget("startdate"))
            foneitem.fenddate            = db2html(rsget("enddate"))
			foneitem.fisusing            = rsget("isusing")
            foneitem.fregdate            = rsget("regdate")
			foneitem.fcomment            = db2html(rsget("comment"))
			foneitem.fpartMDid            = rsget("partMDid")
			foneitem.fpartWDid            = rsget("partWDid")
			foneitem.fcd1            = rsget("cd1")

        end if
        rsget.Close
    end Sub
        

	'/admin/stylepick/stylepick_event.asp	'//admin/stylepick/stylepick_main_search_event.asp
	public function fnGetEventList()
        dim sqlStr, sqlsearch, i

		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and e.cd1='"&frectcd1&"'"
		end if

		if frectevtidx <> "" then
			sqlsearch = sqlsearch & " and e.evtidx="&frectevtidx&""
		end if

		if frecttitle <> "" then
			sqlsearch = sqlsearch & " and e.title like '%"&frecttitle&"%'"
		end if
			
        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and e.isusing='" + FRectIsUsing + "'"
        end if
		
		If frectstate <> "" THEN
			IF frectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and e.state = 7 and getdate() <= e.startdate"
			ELSEIF frectstate = 7 THEN	'오픈진행중
				sqlsearch  = sqlsearch & " and e.state = 7 and getdate() between e.startdate and e.enddate"			
			elseIF frectstate = 9 THEN	'종료
				sqlsearch  = sqlsearch & " and (e.state = 9 or getdate() >= e.enddate)"
			ELSE
				sqlsearch  = sqlsearch & " and  e.state = "&frectstate & ""
			END IF
		End If
        
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event e"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
					
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " e.evtidx,e.title,e.subcopy,e.state,e.banner_img,e.startdate,e.enddate"
        sqlStr = sqlStr & " ,e.isusing,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"
        sqlStr = sqlStr & " ,e.partMDid ,e.partWDid ,e.closedate ,c1.catename"
		sqlStr = sqlStr & " ,(case when e.state = 7 and getdate() <= e.startdate then '6'" + vbcrlf
		sqlStr = sqlStr & " 	when e.state = 7 and getdate() between e.startdate and e.enddate then '7'" + vbcrlf
		sqlStr = sqlStr & " 	when e.state = 9 or getdate() >= e.enddate then '9'" + vbcrlf
		sqlStr = sqlStr & " 	else e.state end) as statename" + vbcrlf        
        sqlStr = sqlStr & " ,isnull((select count(evtitemidx)"
        sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylepick_event_item"
        sqlStr = sqlStr & " 	where isusing='Y' and e.evtidx = evtidx),0) as evtitemcnt"
        sqlStr = sqlStr & "	,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = e.partMDid) as partMDname"
        sqlStr = sqlStr & " ,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = e.partWDid) as partWDname"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylepick_event e"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'"        
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by evtidx desc"
		
		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new cstylepick_item
                
                FItemList(i).fpartMDname            = rsget("partMDname")
                FItemList(i).fpartWDname            = rsget("partWDname")
                FItemList(i).fcatename            = db2html(rsget("catename"))
                FItemList(i).fopendate            = db2html(rsget("opendate"))
                FItemList(i).fclosedate            = db2html(rsget("closedate"))
                FItemList(i).fstatename            = rsget("statename")
                FItemList(i).fevtidx            = rsget("evtidx")
				FItemList(i).ftitle            = db2html(rsget("title"))
                FItemList(i).fsubcopy            = db2html(rsget("subcopy"))
				FItemList(i).fstate            = rsget("state")
                FItemList(i).fbanner_img            = db2html(rsget("banner_img"))
				FItemList(i).fstartdate            = db2html(rsget("startdate"))
                FItemList(i).fenddate            = db2html(rsget("enddate"))
				FItemList(i).fisusing            = rsget("isusing")
                FItemList(i).fregdate            = rsget("regdate")
				FItemList(i).fcomment            = db2html(rsget("comment"))
				FItemList(i).fpartMDid            = rsget("partMDid")
				FItemList(i).fpartWDid            = rsget("partWDid")
				FItemList(i).fcd1            = rsget("cd1")
				FItemList(i).fevtitemcnt            = rsget("evtitemcnt")
				                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function

	'//admin/stylepick/pop_evtitemAddInfo.asp	
	public function GeteventitemList()
        dim sqlStr, sqlsearch, i

        if FRectCate_Large<>"" then
            sqlsearch = sqlsearch + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            sqlsearch = sqlsearch + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            sqlsearch = sqlsearch + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        
		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and si.cd1='"&frectcd1&"'"
		end if

		if frectcd2 <> "" then
			sqlsearch = sqlsearch & " and si.cd2='"&frectcd2&"'"
		end if

		if frectcd3 <> "" then
			sqlsearch = sqlsearch & " and si.cd3='"&frectcd3&"'"
		end if
			
	
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            sqlsearch = sqlsearch & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and si.isusing='" + FRectIsUsing + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectMWDiv="MW" then
            sqlsearch = sqlsearch + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectSailYn<>"" then
            sqlsearch = sqlsearch + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            sqlsearch = sqlsearch + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  sqlsearch = sqlsearch + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if
		
		if frectoverlap = "notoverlap" then
			  sqlsearch = sqlsearch + " and t.itemid is null"
		end if
		
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on si.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on si.cd1 = c1.cd1 and c1.isusing='Y' and isnull(si.cd1,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd2 c2"
		sqlstr = sqlstr & " 	on si.cd2 = c2.cd2 and c2.isusing='Y' and isnull(si.cd2,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd3 c3"
		sqlstr = sqlstr & " 	on si.cd1 = c3.cd3 and c3.isusing='Y' and isnull(si.cd3,'') <> ''"
		
		if frectoverlap = "notoverlap" then
			sqlstr = sqlstr & " left join ("
			sqlstr = sqlstr & " 	select ei.itemid , e.cd1"
			sqlstr = sqlstr & " 	from db_giftplus.dbo.tbl_stylepick_event_item ei"
			sqlstr = sqlstr & " 	join db_giftplus.dbo.tbl_stylepick_event e"
			sqlstr = sqlstr & " 	on ei.evtidx=e.evtidx"
			sqlstr = sqlstr & " 		and ei.isusing='Y' and e.isusing='Y'"
			sqlstr = sqlstr & " 		and e.state < 7 and e.startdate >= getdate()"
			sqlstr = sqlstr & " 	) as t"
			sqlstr = sqlstr & " 	on si.itemid=t.itemid and si.cd1=t.cd1"
		end if
					
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
		
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)                
        sqlStr = sqlStr & " si.itemidx, si.itemid, si.regdate, si.isusing, si.cd1 , si.cd2 ,si.cd3"
        sqlStr = sqlStr & " ,(c1.catename) as 'cd1name',(c2.catename) as 'cd2name',(c3.catename) as 'cd3name'"
		sqlStr = sqlStr & " ,i.makerid,i.itemdiv,i.itemgubun,i.itemname,i.sellcash,i.buycash,i.orgprice,i.orgsuplycash"	
		sqlStr = sqlStr & " ,i.sailprice,i.sailsuplycash,i.mileage,i.sellEndDate,i.sellyn,i.limityn,i.danjongyn,i.sailyn"	
		sqlStr = sqlStr & " ,i.isextusing,i.mwdiv,i.specialuseritem,i.vatinclude,i.deliverytype,i.deliverarea,i.deliverfixday"	
		sqlStr = sqlStr & " ,i.ismobileitem,i.pojangok,i.limitno,i.limitsold,i.evalcnt,i.optioncnt,i.itemrackcode"	
		sqlStr = sqlStr & " ,i.upchemanagecode,i.brandname,i.smallimage,i.listimage,i.listimage120,i.itemcouponyn"	
		sqlStr = sqlStr & " ,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue,i.itemscore,i.lastupdate"
        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on si.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on si.cd1 = c1.cd1 and c1.isusing='Y' and isnull(si.cd1,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd2 c2"
		sqlstr = sqlstr & " 	on si.cd2 = c2.cd2 and c2.isusing='Y' and isnull(si.cd2,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd3 c3"
		sqlstr = sqlstr & " 	on si.cd1 = c3.cd3 and c3.isusing='Y' and isnull(si.cd3,'') <> ''"		

		if frectoverlap = "notoverlap" then
			sqlstr = sqlstr & " left join ("
			sqlstr = sqlstr & " 	select ei.itemid , e.cd1"
			sqlstr = sqlstr & " 	from db_giftplus.dbo.tbl_stylepick_event_item ei"
			sqlstr = sqlstr & " 	join db_giftplus.dbo.tbl_stylepick_event e"
			sqlstr = sqlstr & " 	on ei.evtidx=e.evtidx"
			sqlstr = sqlstr & " 		and ei.isusing='Y' and e.isusing='Y'"
			sqlstr = sqlstr & " 		and e.state < 7 and e.startdate >= getdate()"
			sqlstr = sqlstr & " 	) as t"
			sqlstr = sqlstr & " 	on si.itemid=t.itemid and si.cd1=t.cd1"
		end if

        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch
		
		IF FRectSortDiv="ne" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="hp" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="lp" Then
			sqlStr = sqlStr & " Order by i.SellCash asc"
		ELSEIF FRectSortDiv="be" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new cstylepick_item
                
                FItemList(i).fitemidx            = rsget("itemidx")
                FItemList(i).fcd1            = rsget("cd1")
                FItemList(i).fcd2            = rsget("cd2")
                FItemList(i).fcd3            = rsget("cd3")
                FItemList(i).fcd1name = db2html(rsget("cd1name"))
                FItemList(i).fcd2name = db2html(rsget("cd2name"))
                FItemList(i).fcd3name = db2html(rsget("cd3name"))
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
                FItemList(i).Fsmallimage        = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
	
	'//admin/stylepick/stylepick_item.asp	'//admin/stylepick/stylepick_main_search_item.asp
	public function GetItemList()
        dim sqlStr, sqlsearch, i

        if FRectCate_Large<>"" then
            sqlsearch = sqlsearch + " and i.cate_large='" + FRectCate_Large + "'"
        end if
        
        if FRectCate_Mid<>"" then
            sqlsearch = sqlsearch + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if
        
        if FRectCate_Small<>"" then
            sqlsearch = sqlsearch + " and i.cate_small='" + FRectCate_Small + "'"
        end if
        
		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and si.cd1='"&frectcd1&"'"
		end if

		if frectcd2 <> "" then
			sqlsearch = sqlsearch & " and si.cd2='"&frectcd2&"'"
		end if
		
		if frectcd3 <> "" then
			sqlsearch = sqlsearch & " and si.cd3='"&frectcd3&"'"
		end if
			
		
        if (FRectMakerid <> "") then
            sqlsearch = sqlsearch & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	sqlsearch = sqlsearch & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
            	sqlsearch = sqlsearch & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            sqlsearch = sqlsearch & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if
        
        if (FRectSellYN="YS") then
            sqlsearch = sqlsearch & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            sqlsearch = sqlsearch & " and i.sellyn='" + FRectSellYN + "'"
        end if

        if (FRectIsUsing <> "") then
            sqlsearch = sqlsearch & " and si.isusing='" + FRectIsUsing + "'"
        end if
        
        if FRectDanjongyn="SN" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'Y'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            sqlsearch = sqlsearch + " and i.danjongyn<>'N'"
            sqlsearch = sqlsearch + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            sqlsearch = sqlsearch + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if
        
        if FRectMWDiv="MW" then
            sqlsearch = sqlsearch + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            sqlsearch = sqlsearch + " and i.mwdiv='" + FRectMwDiv + "'"
        end if
		
		if FRectLimityn="Y0" then
            sqlsearch = sqlsearch + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            sqlsearch = sqlsearch + " and i.limityn='" + FRectLimityn + "'"
        end if        
        
        if FRectSailYn<>"" then
            sqlsearch = sqlsearch + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            sqlsearch = sqlsearch + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if
        
        if FRectDeliveryType<>"" then
        	  sqlsearch = sqlsearch + " and i.deliverytype='" + FRectDeliveryType + "'"
        end if

		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on si.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on i.makerid=c.userid"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on si.cd1 = c1.cd1 and c1.isusing='Y' and isnull(si.cd1,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd2 c2"
		sqlstr = sqlstr & " 	on si.cd2 = c2.cd2 and c2.isusing='Y' and isnull(si.cd2,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd3 c3"
		sqlstr = sqlstr & " 	on si.cd1 = c3.cd3 and c3.isusing='Y' and isnull(si.cd3,'') <> ''"			
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
		
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)                
        sqlStr = sqlStr & " si.itemidx, si.itemid, si.regdate, si.isusing, si.cd1 , si.cd2 ,si.cd3"
        sqlStr = sqlStr & " ,(c1.catename) as 'cd1name',(c2.catename) as 'cd2name',(c3.catename) as 'cd3name'"
		sqlStr = sqlStr & " ,i.makerid,i.itemdiv,i.itemgubun,i.itemname,i.sellcash,i.buycash,i.orgprice,i.orgsuplycash"	
		sqlStr = sqlStr & " ,i.sailprice,i.sailsuplycash,i.mileage,i.sellEndDate,i.sellyn,i.limityn,i.danjongyn,i.sailyn"	
		sqlStr = sqlStr & " ,i.isextusing,i.mwdiv,i.specialuseritem,i.vatinclude,i.deliverytype,i.deliverarea,i.deliverfixday"	
		sqlStr = sqlStr & " ,i.ismobileitem,i.pojangok,i.limitno,i.limitsold,i.evalcnt,i.optioncnt,i.itemrackcode"	
		sqlStr = sqlStr & " ,i.upchemanagecode,i.brandname,i.smallimage,i.listimage,i.listimage120,i.itemcouponyn"	
		sqlStr = sqlStr & " ,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue,i.itemscore,i.lastupdate"
        sqlStr = sqlStr & " ,IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
		sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		sqlstr = sqlstr & "		on si.itemid = i.itemid and i.isusing='Y'" + vbcrlf
		sqlStr = sqlStr & " left join [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr & " 	on i.makerid=c.userid"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on si.cd1 = c1.cd1 and c1.isusing='Y' and isnull(si.cd1,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd2 c2"
		sqlstr = sqlstr & " 	on si.cd2 = c2.cd2 and c2.isusing='Y' and isnull(si.cd2,'') <> ''"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd3 c3"
		sqlstr = sqlstr & " 	on si.cd1 = c3.cd3 and c3.isusing='Y' and isnull(si.cd3,'') <> ''"		
        sqlStr = sqlStr & " where i.itemid<>0 " & sqlsearch
		
		IF FRectSortDiv="ne" Then
			sqlStr = sqlStr & " Order by i.itemid desc "
		ELSEIF FRectSortDiv="hp" Then 
			sqlStr = sqlStr & " Order by i.SellCash desc "
		ELSEIF FRectSortDiv="lp" Then
			sqlStr = sqlStr & " Order by i.SellCash asc"
		ELSEIF FRectSortDiv="be" Then
			sqlStr = sqlStr & " Order by i.ItemScore desc "
		ELSE
			sqlStr = sqlStr & " Order by i.itemid desc "
		End IF

		'response.write sqlStr &"<Br>"
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
                set FItemList(i) = new cstylepick_item
                
                FItemList(i).fitemidx            = rsget("itemidx")
                FItemList(i).fcd1            = rsget("cd1")
                FItemList(i).fcd2            = rsget("cd2")
                FItemList(i).fcd3            = rsget("cd3")
                FItemList(i).fcd1name = db2html(rsget("cd1name"))
                FItemList(i).fcd2name = db2html(rsget("cd2name"))
                FItemList(i).fcd3name = db2html(rsget("cd3name"))
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fitemdiv           = rsget("itemdiv")
                FItemList(i).Fitemgubun         = rsget("itemgubun")
                FItemList(i).Fitemname          = db2html(rsget("itemname"))
                FItemList(i).Fsellcash          = rsget("sellcash")
                FItemList(i).Fbuycash           = rsget("buycash")
                FItemList(i).Forgprice          = rsget("orgprice")
                FItemList(i).Forgsuplycash      = rsget("orgsuplycash")
                FItemList(i).Fsailprice         = rsget("sailprice")
                FItemList(i).Fsailsuplycash     = rsget("sailsuplycash")
                FItemList(i).Fmileage           = rsget("mileage")
                FItemList(i).Fregdate           = rsget("regdate")
                FItemList(i).Flastupdate        = rsget("lastupdate")
                FItemList(i).FsellEndDate       = rsget("sellEndDate")
                FItemList(i).Fsellyn            = rsget("sellyn")
                FItemList(i).Flimityn           = rsget("limityn")
                FItemList(i).Fdanjongyn         = rsget("danjongyn")
                FItemList(i).Fsailyn            = rsget("sailyn")
                FItemList(i).Fisusing           = rsget("isusing")
                FItemList(i).Fisextusing        = rsget("isextusing")
                FItemList(i).Fmwdiv             = rsget("mwdiv")
                FItemList(i).Fspecialuseritem   = rsget("specialuseritem")
                FItemList(i).Fvatinclude        = rsget("vatinclude")
                FItemList(i).Fdeliverytype      = rsget("deliverytype")
                FItemList(i).Fdeliverarea       = rsget("deliverarea")
                FItemList(i).Fdeliverfixday     = rsget("deliverfixday")
                FItemList(i).Fismobileitem      = rsget("ismobileitem")
                FItemList(i).Fpojangok          = rsget("pojangok")
                FItemList(i).Flimitno           = rsget("limitno")
                FItemList(i).Flimitsold         = rsget("limitsold")
                FItemList(i).Fevalcnt           = rsget("evalcnt")
                FItemList(i).Foptioncnt         = rsget("optioncnt")
                FItemList(i).Fitemrackcode      = rsget("itemrackcode")
                FItemList(i).Fupchemanagecode   = rsget("upchemanagecode")
                FItemList(i).Fbrandname         = db2html(rsget("brandname"))
                FItemList(i).Fsmallimage        = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                
                rsget.movenext
                i=i+1
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
End Class
	
	
Class cstylepickMenuItem
	public fcd1
	public fcd2
	public fcd3
	public fcatename
	public fisusing
	public forderno
	public flastadminid
	public fitemcount

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class cstylepickMenu
	Public FItemList()
	public foneitem
	Public FResultCount
	Public FTotalCount
	Public FScrollCount
	public FPageCount
	Public FCurrPage
	Public FPageSize
	public FTotalPage
	public frectcd1
	public frectcd2
	public frectcd3
	public frectisusing
	public frectcateidx
	public fitemallcount
		
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		fitemallcount = 0	
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'//stylepick/stylepick_list.asp	'//admin/stylepick/stylepick_collect_testview.asp
	public function fstylepick_cd2_count()
        dim sqlStr, sqlsearch, i
					
        '// 본문 내용 접수
		sqlstr = "exec db_giftplus.dbo.ten_stylepick_cd2_count '"&frectcd1&"'"
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		
		'response.write sqlStr &"<Br>"
        rsget.pagesize = FPageSize
		rsget.Open sqlstr, dbget, 1
		
        FResultCount = rsget.RecordCount
                
        if (FResultCount<1) then FResultCount=0
        
        redim preserve FItemList(FResultCount)

        i=0
        if  not rsget.EOF  then
            rsget.absolutepage = FCurrPage
            do until rsget.EOF
                set FItemList(i) = new cstylepickMenuItem
                
                FItemList(i).fitemcount = rsget("itemcount")
				FItemList(i).fcd2 = rsget("cd2")
				FItemList(i).fcatename = db2html(rsget("catename"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")				
				FItemList(i).flastadminid = db2html(rsget("lastadminid"))
				
				fitemallcount = fitemallcount + rsget("itemcount")
								                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd1()	
		dim sqlStr,i ,sqlsearch
		
		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"&frectisusing&"'"
		end if
		
		'데이터 리스트 
		sqlstr = "select"
		sqlstr = sqlstr & " cd1,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd1"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cstylepickMenuItem
				
				FItemList(i).fcd1 = rsget("cd1")
				FItemList(i).fcatename = db2html(rsget("catename"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")				
				FItemList(i).flastadminid = db2html(rsget("lastadminid"))
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd1_one()
		dim sqlstr,i , sqlsearch
		
		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and cd1 = '"&frectcd1&"'"
		end if
		
		sqlstr = "select"
		sqlstr = sqlstr & " cd1,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd1"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cstylepickMenuItem
        
        if Not rsget.Eof then
			
			foneitem.fcd1 = rsget("cd1")						
			foneitem.fcatename = db2html(rsget("catename"))
			foneitem.fisusing = rsget("isusing")
			foneitem.forderno = rsget("orderno")			
			foneitem.flastadminid = db2html(rsget("lastadminid"))
			           
        end if
        rsget.Close
    end Sub

	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd2()	
		dim sqlStr,i ,sqlsearch

		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"&frectisusing&"'"
		end if

		'데이터 리스트 
		sqlstr = "select"
		sqlstr = sqlstr & " cd2,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd2"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cstylepickMenuItem
				
				FItemList(i).fcd2 = rsget("cd2")
				FItemList(i).fcatename = db2html(rsget("catename"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")				
				FItemList(i).flastadminid = db2html(rsget("lastadminid"))
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd2_one()
		dim sqlstr,i , sqlsearch
		
		if frectcd2 <> "" then
			sqlsearch = sqlsearch & " and cd2 = '"&frectcd2&"'"
		end if
		
		sqlstr = "select"
		sqlstr = sqlstr & " cd2,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd2"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cstylepickMenuItem
        
        if Not rsget.Eof then
			
			foneitem.fcd2 = rsget("cd2")						
			foneitem.fcatename = db2html(rsget("catename"))
			foneitem.fisusing = rsget("isusing")
			foneitem.forderno = rsget("orderno")			
			foneitem.flastadminid = db2html(rsget("lastadminid"))
			           
        end if
        rsget.Close
    end Sub

	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd3()	
		dim sqlStr,i ,sqlsearch

		if frectisusing <> "" then
			sqlsearch = sqlsearch & " and isusing ='"&frectisusing&"'"
		end if
		
		if frectcd3 <> "" then
			sqlsearch = sqlsearch & " and Left(cd3,1) = '"&frectcd3&"'"
		end if

		'데이터 리스트 
		sqlstr = "select"
		sqlstr = sqlstr & " cd3,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd3"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1

		FResultCount = rsget.recordcount

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new cstylepickMenuItem
				
				FItemList(i).fcd3 = rsget("cd3")
				FItemList(i).fcatename = db2html(rsget("catename"))
				FItemList(i).fisusing = rsget("isusing")
				FItemList(i).forderno = rsget("orderno")				
				FItemList(i).flastadminid = db2html(rsget("lastadminid"))
								
				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub
	
	'//admin/stylepick/stylepick_category.asp
	Public Sub getstylepick_cate_cd3_one()
		dim sqlstr,i , sqlsearch
		
		if frectcd3 <> "" then
			sqlsearch = sqlsearch & " and cd3 = '"&frectcd3&"'"
		end if
		
		sqlstr = "select"
		sqlstr = sqlstr & " cd3,catename,isusing,orderno,lastadminid"
		sqlstr = sqlstr & " from db_giftplus.dbo.tbl_stylepick_cate_cd3"
		sqlstr = sqlstr & " where 1=1 " & sqlsearch
		sqlstr = sqlstr & " order by orderno asc"

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cstylepickMenuItem
        
        if Not rsget.Eof then
			
			foneitem.fcd3 = rsget("cd3")						
			foneitem.fcatename = db2html(rsget("catename"))
			foneitem.fisusing = rsget("isusing")
			foneitem.forderno = rsget("orderno")			
			foneitem.flastadminid = db2html(rsget("lastadminid"))
			           
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
End Class

'//카테고리타입 SELECTBOX 타입
function Drawcatetype(selectBoxName,selectedId,changeFlag)
	dim tmp_str,query1
	%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
		<option value='CD1' <%if selectedId="CD1" then response.write " selected"%> >스타일</option>
		<option value='CD2' <%if selectedId="CD2" then response.write " selected"%> >분류</option>
		<!--<option value='CD3' <%if selectedId="CD3" then response.write " selected"%> >카테고리3</option>-->
	</select>
<%
end function

'//카테고리타입
function GETcatetype(v)
	if v = "CD1" then
		GETcatetype = "스타일"
	elseif v = "CD2" then
		GETcatetype = "분류"
	elseif v = "CD3" then
		GETcatetype = "카테고리3"
	else
		GETcatetype = "지정안됨"
	end if						
end function

'//카테고리
function Drawcategory(selectBoxName,selectedId,changeFlag,catetype)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
	<%
	if catetype = "CD1" then
		query1 = " select cd1 ,catename ,isusing ,orderno ,lastadminid"
		query1 = query1 & " from db_giftplus.dbo.tbl_stylepick_cate_cd1 where isusing = 'Y'"
	elseif catetype = "CD2" then
		query1 = " select cd2 ,catename ,isusing ,orderno ,lastadminid"
		query1 = query1 & " from db_giftplus.dbo.tbl_stylepick_cate_cd2"
	elseif catetype = "CD3" then
		query1 = " select cd3 ,catename ,isusing ,orderno ,lastadminid"
		query1 = query1 & " from db_giftplus.dbo.tbl_stylepick_cate_cd3"		
	end if	

	query1 = query1 & " order by orderno asc"
	
	rsget.Open query1,dbget,1
	
	if  not rsget.EOF  then
		do until rsget.EOF

		if catetype = "CD1" then
			if Lcase(selectedId) = Lcase(rsget("cd1")) then
			   tmp_str = " selected"
			end if
		elseif catetype = "CD2" then
			if Lcase(selectedId) = Lcase(rsget("cd2")) then
			   tmp_str = " selected"
			end if
		elseif catetype = "CD3" then
			if Lcase(selectedId) = Lcase(rsget("cd3")) then
			   tmp_str = " selected"
			end if		
		end if

		if catetype = "CD1" then
			response.write("<option value='"&rsget("cd1")&"' "&tmp_str&">" + db2html(rsget("catename")) + "</option>")
		elseif catetype = "CD2" then
			response.write("<option value='"&rsget("cd2")&"' "&tmp_str&">" + db2html(rsget("catename")) + "</option>")
		elseif catetype = "CD3" then
			response.write("<option value='"&rsget("cd3")&"' "&tmp_str&">" + db2html(rsget("catename")) + "</option>")	
		end if
	   			
		tmp_str = ""
		rsget.MoveNext
		loop
	end if
	rsget.close
	response.write("</select>")
end function

'//정렬
function Drawsort(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value="ne" <% if selectedId="ne" then response.write "selected" %>>신상품순</option>
		<option value="be" <% if selectedId="be" then response.write "selected" %>>인기상품순</option>
		<option value="lp" <% if selectedId="lp" then response.write "selected" %>>낮은가격순</option>
		<option value="hp" <% if selectedId="hp" then response.write "selected" %>>높은가격순</option>		
	</select>
<%
end function

'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출안함 , 등록페이지용
function Draweventstate(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>		
		<% if selectedId <= "0" then %><option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option><% end if %>
		<% if selectedId <= "3" then %><option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option><% end if %>
		<% if selectedId <= "5" then %><option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option><% end if %>
		<% if selectedId <= "7" then %><option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option><% end if %>
		<% if selectedId <= "9" then %><option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option><% end if %>
	</select>
<%
end function

'//메인페이지 , 이벤트 공통함수		'/오픈예정 노출함 , 검색페이지용
function Draweventstate2(selectBoxName,selectedId,changeFlag)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
		<option value="" <%if selectedId="" then response.write " selected"%>>선택</option>
		<option value="0" <% if selectedId="0" then response.write "selected" %>>등록대기</option>
		<option value="3" <% if selectedId="3" then response.write "selected" %>>이미지등록요청</option>
		<option value="5" <% if selectedId="5" then response.write "selected" %>>오픈요청</option>
		<option value="6" <% if selectedId="6" then response.write "selected" %>>오픈예정</option>
		<option value="7" <% if selectedId="7" then response.write "selected" %>>오픈</option>
		<option value="9" <% if selectedId="9" then response.write "selected" %>>종료</option>
	</select>
<%
end function

'//메인페이지 , 이벤트 모두 공통
function geteventstate(v)
	if v = "0" then
		geteventstate = "등록대기"
	elseif v = "3" then
		geteventstate = "이미지등록요청"
	elseif v = "5" then
		geteventstate = "오픈요청"
	elseif v = "6" then
		geteventstate = "오픈예정"			
	elseif v = "7" then
		geteventstate = "오픈"
	elseif v = "9" then
		geteventstate = "종료"				
	end if						
end function

'/담당MD 리스트가져오기 (팀장 미만,직원 이상)
Sub sbGetpartid(ByVal selName, ByVal sIDValue, ByVal sScript,part_sn)
	Dim strSql, arrList, intLoop
	
	if part_sn = "" then exit sub
	
	strSql = " SELECT userid, username"
	strSql = strSql & " FROM db_partner.dbo.tbl_user_tenbyten "   
	strSql = strSql & " WHERE part_sn ='"&part_sn&"' and  posit_sn>='4' and  posit_sn<='8' and   isUsing=1" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	strSql = strSql & "	and (statediv ='Y' or (statediv ='N' and datediff(dd,retireday,getdate())<=0))" & vbcrlf
	strSql = strSql & " and userid <> ''" & vbcrlf 
	strSql = strSql & " order by posit_sn, empno" & vbcrlf
	
	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not rsget.eof THEN
	arrList = rsget.getRows()	
	End IF
	rsget.close
%>
	<select name="<%=selName%>" <%=sScript%>>
	<option value="">선택</option>
	<%   
	If isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
	%>
	<option value="<%=arrList(0,intLoop)%>" <%if Cstr(arrList(0,intLoop)) = Cstr(sIDValue) then %>selected<%end if%>><%=arrList(1,intLoop)%></option>
	<%   		
		Next
	End IF
	%>
	</select>
<%	
End Sub
%>
