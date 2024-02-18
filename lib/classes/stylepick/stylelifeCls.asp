<%
Class cStyleLife_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	public fidx
    public fpartMDname
    public fpartwDname
    public fcatename
    public fopendate
    public fclosedate
    public fstatename
    public frectidx
    public ftitle
    public Fsubcopy
    public Fstate
    public Fbanner_img
    public Fstartdate
    public Fenddate
    public Fcomment
    public fisusing
    public fregdate
    public fpartMDid
    public fpartwDid
	public fcd1
	public fcd2
	public fcd3
	public fitemcnt
	public ftitle_img
	public flastadminid
	public fitemidx
	public fcd1name
	public fcd2name
	public fcd3name
	public FItemID
	public FMakerid
	public Fitemdiv
	public Fitemgubun
	public FItemName
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
    public fsortno
	
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
	
	
End Class

Class ClsStyleLife

	public FItemList()
	public FOneItem
	public FIdx
	
	public FGubun
	public FCate1
	public FCate2
	public FCate3
	public FOnly1
	public FOnly2
	public FSort1
	public FSort2
	public FItemID
	public FItemName
	public FMakerID
	
	public FLCate
	public FMCate
	public FSCate
	
	public frectcd1
	public frectcd2
	public frectcd3
	public frectisusing
	public frectstate
	public frectidx
	public frectmainidx
	public frecttitle
	
	public FRectSortDiv
	public FRectMakerid
	public FRectItemid
	public FRectItemName
	public FRectSellYN
	public FRectDanjongyn
	public FRectLimityn
	public FRectMWDiv
	public FRectDeliveryType
	public FRectSailYn
	public FRectCouponYn
	
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public frectoverlap
	
	public ftotalcount
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalPage
	public FScrollCount

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	
	
	public Function FStyleLifeItemList
		Dim sqlStr, i, vSubQuery, iDelCnt, vSorting
		
		vSubQuery = " AND I.itemid NOT IN(select itemid from [db_giftplus].[dbo].[tbl_stylelife_notuse_item]) "
		
		If FGubun <> "x" Then
			If FCate1 <> "" Then
				vSubQuery = vSubQuery & " AND I.cate_large = '" & Trim(FCate1) & "' "
			End If
			
			If FCate2 <> "" Then
				vSubQuery = vSubQuery & " AND I.cate_mid = '" & Trim(FCate2) & "' "
			End If
			
			If FCate3 <> "" Then
				vSubQuery = vSubQuery & " AND I.cate_small = '" & Trim(FCate3) & "' "
			End If
		End IF
		
		If FItemID <> "" Then
			vSubQuery = vSubQuery & " AND I.itemid IN(" & Trim(FItemID) & ")"
		End If
		
		If FItemName <> "" Then
			vSubQuery = vSubQuery & " AND I.itemname like '%" & Trim(FItemName) & "%' "
		End If
		
		If FMakerID <> "" Then
			vSubQuery = vSubQuery & " AND I.makerid = '" & Trim(FMakerID) & "' "
		End If
		
		If FOnly1 <> "" Then
			vSubQuery = vSubQuery & " AND I.sellyn = 'Y' "
		End If
		
		If FOnly2 <> "" Then
			vSubQuery = vSubQuery & " AND SI.itemid is Null "
		End If
		
		IF FSort1 <> "" Then
			vSubQuery = vSubQuery & " AND SI.cd1 = '" & FSort1 & "' "
		End If
		
		vSorting = " ORDER BY "

		IF FSort2 = "ne" Then
			vSorting = vSorting & " I.itemid DESC "
		ElseIf FSort2 = "hp" Then
			vSorting = vSorting & " I.sellcash DESC, I.itemid DESC "
		ElseIf FSort2 = "lp" Then
			vSorting = vSorting & " I.sellcash ASC, I.itemid DESC "
		ElseIf FSort2 = "be" Then
			vSorting = vSorting & " I.itemscore DESC, I.itemid DESC "
		ElseIf FSort2 = "lp" Then
			vSorting = vSorting & " I.sellcash ASC, I.itemid DESC "
		ElseIf FSort2 = "mk" Then
			vSorting = vSorting & " I.makerid ASC, I.itemid DESC "
		End If

		sqlStr = "SELECT COUNT(DISTINCT I.itemid) " & _
				 "			FROM [db_item].[dbo].[tbl_item] AS I " & _
				 "	LEFT JOIN [db_giftplus].[dbo].[tbl_stylepick_item] AS SI ON I.itemid = SI.itemid " & _
				 "	WHERE I.IsUsing = 'Y' " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget(0)
		rsget.Close
		
		IF ftotalcount > 0 THEN
			iDelCnt =  ((FCurrPage - 1) * FPageSize )
			sqlStr = "SELECT DISTINCT Top " & (FPageSize) & " I.itemid, I.itemname, I.smallimage, " & _
					 "			IsNull(" & _
					 "				STUFF( " & _
					 "						( " & _
					 "							SELECT ',' + CAST(cd1 AS VARCHAR(3)) " & _
					 "							FROM [db_giftplus].[dbo].[tbl_stylepick_item] " & _
					 "							WHERE itemid = SI.itemid " & _
					 "							FOR XML PATH('')  " & _
					 "						), 1, 1, '' " & _
					 "				)" & _
					 "				,'') AS cd1 " & _
					 "				, I.sellcash, I.itemscore, I.makerid, (I.cate_large + I.cate_mid) AS itemcate " & _
					 "			FROM [db_item].[dbo].[tbl_item] AS I " & _
					 "	LEFT JOIN [db_giftplus].[dbo].[tbl_stylepick_item] AS SI ON I.itemid = SI.itemid " & _
					 "	WHERE I.IsUsing = 'Y' " & _
					 "	" & vSubQuery & " " & _
					 "		AND I.itemid NOT IN " & _
					 "		( " & _
					 "			SELECT III.itemid FROM " & _
					 "			( " & _
					 "				SELECT DISTINCT TOP "&iDelCnt&" II.itemid, II.sellcash, II.itemscore, II.makerid FROM [db_item].[dbo].[tbl_item] AS II " & _
					 "				LEFT JOIN [db_giftplus].[dbo].[tbl_stylepick_item] AS SII ON II.itemid = SII.itemid " & _
					 "				WHERE II.IsUsing = 'Y' " & Replace(Replace(vSubQuery,"I.","II."),"SI.","SII.") & " " & _
					 "				" & Replace(Replace(vSorting,"I.","II."),"SI.","SII.") & " "&_
					 "			) AS III " & _
					 "		) " & _
					 "	" & vSorting & " "
			'response.write sqlStr
			rsget.Open sqlStr, dbget, 1

			IF not rsget.EOF THEN
				FStyleLifeItemList = rsget.getRows() 
			END IF
			rsget.Close
		END IF
	End Function
	
	
	public Function FStyleLifeItemMidCateList
		Dim sqlStr, i, vSubQuery, iDelCnt, vSorting
		
		If FLCate <> "" Then
			vSubQuery = vSubQuery & " AND I.cate_large = '" & Trim(FLCate) & "' "
		End If
		
		If FMCate <> "" Then
			vSubQuery = vSubQuery & " AND I.cate_mid = '" & Trim(FMCate) & "' "
		End If
		
		If FSCate <> "" Then
			vSubQuery = vSubQuery & " AND I.cate_small = '" & Trim(FSCate) & "' "
		End If
		
		If FCate1 <> "" Then
			vSubQuery = vSubQuery & " AND SI.cd1 = '" & Trim(FCate1) & "' "
		End If
		
		If FCate2 <> "" Then
			vSubQuery = vSubQuery & " AND SI.cd2 = '" & Trim(FCate2) & "' "
		End If
		
		If FCate3 <> "" Then
			vSubQuery = vSubQuery & " AND SI.cd3 = '" & Trim(FCate3) & "' "
		End If
		
		sqlStr = "SELECT COUNT(DISTINCT SI.itemid) " & _
				 "			FROM [db_giftplus].[dbo].[tbl_stylepick_item] AS SI " & _
				 "		INNER JOIN [db_item].[dbo].[tbl_item] AS I ON I.itemid = SI.itemid " & _
				 "	WHERE 1=1 " & _
				 "	" & vSubQuery & " "
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget(0)
		rsget.Close
		
		IF ftotalcount > 0 THEN
			iDelCnt =  ((FCurrPage - 1) * FPageSize )
			sqlStr = "SELECT DISTINCT Top " & (FPageSize) & " I.itemid, I.itemname, I.smallimage, " & _
					 "			 I.sellcash, I.itemscore, I.makerid " & _
					 "			 , (SELECT catename FROM [db_giftplus].[dbo].[tbl_stylepick_cate_cd3] WHERE cd3 = SI.cd3) AS catename " & _
					 "			FROM [db_giftplus].[dbo].[tbl_stylepick_item] AS SI " & _
					 "	INNER JOIN [db_item].[dbo].[tbl_item] AS I ON I.itemid = SI.itemid " & _
					 "	WHERE 1=1 " & _
					 "	" & vSubQuery & " " & _
					 "		AND SI.itemid NOT IN " & _
					 "		( " & _
					 "			SELECT III.itemid FROM " & _
					 "			( " & _
					 "				SELECT DISTINCT TOP "&iDelCnt&" II.itemid, II.sellcash, II.itemscore, II.makerid FROM [db_giftplus].[dbo].[tbl_stylepick_item] AS SII " & _
					 "				INNER JOIN [db_item].[dbo].[tbl_item] AS II ON II.itemid = SII.itemid " & _
					 "				WHERE 1=1 " & Replace(Replace(vSubQuery,"I.","II."),"SI.","SII.") & " " & _
					 "				" & Replace(Replace(vSorting,"I.","II."),"SI.","SII.") & " "&_
					 "			) AS III " & _
					 "		) " & _
					 "	" & vSorting & " "
			'response.write sqlStr
			rsget.Open sqlStr, dbget, 1

			IF not rsget.EOF THEN
				FStyleLifeItemMidCateList = rsget.getRows() 
			END IF
			rsget.Close
		END IF
	End Function
	
	
	public function fnGetThemeList()
        dim sqlStr, sqlsearch, i

		if frectcd1 <> "" then
			sqlsearch = sqlsearch & " and e.cd1='"&frectcd1&"'"
		end if

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and e.idx="&frectidx&""
		end if

		if frecttitle <> "" then
			sqlsearch = sqlsearch & " and e.title like '%"&frecttitle&"%'"
		end if
		
		If frectstate <> "" THEN
			IF frectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and e.state = 7 and getdate() <= e.startdate"
			ELSE
				sqlsearch  = sqlsearch & " and  e.state = "&frectstate & ""
			END IF
		End If
        
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_theme e"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
					
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " e.idx,e.title,e.subcopy,e.state,e.banner_img,e.title_img,e.startdate,e.enddate,e.sortno"
        sqlStr = sqlStr & " ,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"
        sqlStr = sqlStr & " ,e.partMDid ,e.partWDid ,e.closedate ,c1.catename"
		sqlStr = sqlStr & " ,(case when e.state = 7 and getdate() <= e.startdate then '6'" + vbcrlf
		sqlStr = sqlStr & " 	else e.state end) as statename" + vbcrlf        
        sqlStr = sqlStr & " ,isnull((select count(itemidx)"
        sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylelife_theme_item"
        sqlStr = sqlStr & " 	where e.idx = idx),0) as itemcnt"
        sqlStr = sqlStr & "	,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = e.partMDid) as partMDname"
        sqlStr = sqlStr & " ,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = e.partWDid) as partWDname"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_theme e"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'"        
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by sortno desc, idx desc, startdate DESC"
		
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
                set FItemList(i) = new cStyleLife_oneitem
                
                FItemList(i).fpartMDname            = rsget("partMDname")
                FItemList(i).fpartWDname            = rsget("partWDname")
                FItemList(i).fcatename            = db2html(rsget("catename"))
                FItemList(i).fopendate            = db2html(rsget("opendate"))
                FItemList(i).fclosedate            = db2html(rsget("closedate"))
                FItemList(i).fstatename            = rsget("statename")
                FItemList(i).fidx            = rsget("idx")
				FItemList(i).ftitle            = db2html(rsget("title"))
                FItemList(i).fsubcopy            = db2html(rsget("subcopy"))
				FItemList(i).fstate            = rsget("state")
                FItemList(i).fbanner_img            = db2html(rsget("banner_img"))
                FItemList(i).ftitle_img            = db2html(rsget("title_img"))
				FItemList(i).fstartdate            = db2html(rsget("startdate"))
                FItemList(i).fenddate            = db2html(rsget("enddate"))
                FItemList(i).fregdate            = rsget("regdate")
				FItemList(i).fcomment            = db2html(rsget("comment"))
				FItemList(i).fpartMDid            = rsget("partMDid")
				FItemList(i).fpartWDid            = rsget("partWDid")
				FItemList(i).fcd1            = rsget("cd1")
				FItemList(i).fitemcnt            = rsget("itemcnt")
				FItemList(i).fsortno            = rsget("sortno")
				                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
    
	Public Sub fnGetTheme_item()
		dim sqlstr,i , sqlsearch
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if
		
		sqlstr = "select top 1"
        sqlStr = sqlStr & " e.idx,e.title,e.subcopy,e.state,e.banner_img,e.title_img,e.startdate,e.enddate"
        sqlStr = sqlStr & " ,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"
        sqlStr = sqlStr & " ,e.partMDid ,e.partWDid ,e.closedate,c1.catename,e.sortno"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_theme e"
		sqlstr = sqlstr & " left join db_giftplus.dbo.tbl_stylepick_cate_cd1 c1"
		sqlstr = sqlstr & " 	on e.cd1 = c1.cd1 and c1.isusing='Y'" 
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cStyleLife_oneitem
        
        if Not rsget.Eof then
			
			foneitem.fcatename            = db2html(rsget("catename"))
            foneitem.fopendate            = db2html(rsget("opendate"))
            foneitem.fclosedate            = db2html(rsget("closedate"))
            foneitem.fidx            = rsget("idx")
			foneitem.ftitle            = db2html(rsget("title"))
            foneitem.fsubcopy            = db2html(rsget("subcopy"))
			foneitem.fstate            = rsget("state")
            foneitem.fbanner_img            = db2html(rsget("banner_img"))
            foneitem.ftitle_img            = db2html(rsget("title_img"))
			foneitem.fstartdate            = db2html(rsget("startdate"))
            foneitem.fenddate            = db2html(rsget("enddate"))
            foneitem.fregdate            = rsget("regdate")
			foneitem.fcomment            = db2html(rsget("comment"))
			foneitem.fpartMDid            = rsget("partMDid")
			foneitem.fpartWDid            = rsget("partWDid")
			foneitem.fcd1            = rsget("cd1")
			foneitem.fsortno            = rsget("sortno")

        end if
        rsget.Close
    end Sub
    
    
	public function GetItemList()
        dim sqlStr, sqlsearch, i

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and ei.idx='"&frectidx&"'"
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
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_theme_item ei"
        sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylelife_theme e"
        sqlStr = sqlStr & " 	on ei.idx = e.idx"
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
        sqlStr = sqlStr & " ei.itemidx ,ei.itemid ,ei.regdate"
        sqlStr = sqlStr & " ,e.idx,e.title,e.subcopy,e.state,e.banner_img,e.title_img,e.startdate,e.enddate"
        sqlStr = sqlStr & " ,e.regdate,e.comment,e.lastadminid,e.cd1,e.opendate"        
        sqlStr = sqlStr & " ,(c1.catename) as 'cd1name'"
		sqlStr = sqlStr & " ,i.makerid,i.itemdiv,i.itemgubun,i.itemname,i.sellcash,i.buycash,i.orgprice,i.orgsuplycash"	
		sqlStr = sqlStr & " ,i.sailprice,i.sailsuplycash,i.mileage,i.sellEndDate,i.sellyn,i.limityn,i.danjongyn,i.sailyn,i.isusing"	
		sqlStr = sqlStr & " ,i.isextusing,i.mwdiv,i.specialuseritem,i.vatinclude,i.deliverytype,i.deliverarea,i.deliverfixday"	
		sqlStr = sqlStr & " ,i.ismobileitem,i.pojangok,i.limitno,i.limitsold,i.evalcnt,i.optioncnt,i.itemrackcode"	
		sqlStr = sqlStr & " ,i.upchemanagecode,i.brandname,i.smallimage,i.listimage,i.listimage120,i.itemcouponyn"	
		sqlStr = sqlStr & " ,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue,i.itemscore,i.lastupdate"
        sqlStr = sqlStr & " ,IsNULL(c.defaultFreeBeasongLimit,0) as defaultFreeBeasongLimit"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliverPay,0) as defaultDeliverPay"
        sqlStr = sqlStr & " ,IsNULL(c.defaultDeliveryType,'') as defaultDeliveryType"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_theme_item ei"
        sqlStr = sqlStr & " join db_giftplus.dbo.tbl_stylelife_theme e"
        sqlStr = sqlStr & " 	on ei.idx = e.idx"
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
                set FItemList(i) = new cStyleLife_oneitem

				FItemList(i).fitemidx            = rsget("itemidx")				               
                FItemList(i).fidx            = rsget("idx")
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
    
    
	public function GetTmemeitemList()
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
        
        If frectcd1 <> "0P0" Then
			if frectcd1 <> "" then
				sqlsearch = sqlsearch & " and si.cd1='"&frectcd1&"'"
			end if
		End If

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
            sqlsearch = sqlsearch & " and i.isusing='" + FRectIsUsing + "'"
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
		If frectcd1 = "0P0" Then
	        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
			sqlstr = sqlstr & "	left join [db_giftplus].dbo.tbl_stylepick_item si" + vbcrlf
		Else
	        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
			sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		End IF
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
			sqlstr = sqlstr & " 	from db_giftplus.dbo.tbl_stylelife_theme_item ei"
			sqlstr = sqlstr & " 	join db_giftplus.dbo.tbl_stylelife_theme e"
			sqlstr = sqlstr & " 	on ei.idx=e.idx"
			sqlstr = sqlstr & " 		and ei.isusing='Y'"
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
        sqlStr = sqlStr & " si.itemidx, i.itemid, si.regdate, si.isusing, si.cd1 , si.cd2 ,si.cd3"
        sqlStr = sqlStr & " ,(c1.catename) as 'cd1name',(c2.catename) as 'cd2name',(c3.catename) as 'cd3name'"
		sqlStr = sqlStr & " ,i.makerid,i.itemdiv,i.itemgubun,i.itemname,i.sellcash,i.buycash,i.orgprice,i.orgsuplycash"	
		sqlStr = sqlStr & " ,i.sailprice,i.sailsuplycash,i.mileage,i.sellEndDate,i.sellyn,i.limityn,i.danjongyn,i.sailyn"	
		sqlStr = sqlStr & " ,i.isextusing,i.mwdiv,i.specialuseritem,i.vatinclude,i.deliverytype,i.deliverarea,i.deliverfixday"	
		sqlStr = sqlStr & " ,i.ismobileitem,i.pojangok,i.limitno,i.limitsold,i.evalcnt,i.optioncnt,i.itemrackcode"	
		sqlStr = sqlStr & " ,i.upchemanagecode,i.brandname,i.smallimage,i.listimage,i.listimage120,i.itemcouponyn"	
		sqlStr = sqlStr & " ,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue,i.itemscore,i.lastupdate"
		If frectcd1 = "0P0" Then
	        sqlStr = sqlStr & " from [db_item].[dbo].tbl_item i"
			sqlstr = sqlstr & "	left join [db_giftplus].dbo.tbl_stylepick_item si" + vbcrlf
		Else
	        sqlStr = sqlStr & " from [db_giftplus].dbo.tbl_stylepick_item si"
			sqlstr = sqlstr & "	join [db_item].[dbo].tbl_item i" + vbcrlf
		End IF
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
			sqlstr = sqlstr & " 	from db_giftplus.dbo.tbl_stylelife_theme_item ei"
			sqlstr = sqlstr & " 	join db_giftplus.dbo.tbl_stylelife_theme e"
			sqlstr = sqlstr & " 	on ei.idx=e.idx"
			sqlstr = sqlstr & " 		and ei.isusing='Y'"
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
                set FItemList(i) = new cStyleLife_oneitem
                
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
    
    
	public function fnGetWeeklyList()
        dim sqlStr, sqlsearch, i

		if frectidx <> "" then
			sqlsearch = sqlsearch & " and w.idx = '"&frectidx&"'"
		end if

		if frecttitle <> "" then
			sqlsearch = sqlsearch & " and w.title like '%"&frecttitle&"%'"
		end if
		
		If frectstate <> "" THEN
			IF frectstate = 6 THEN	'오픈예정
				sqlsearch  = sqlsearch & " and w.state = 7 and getdate() <= w.startdate"
			ELSE
				sqlsearch  = sqlsearch & " and  w.state = "&frectstate & ""
			END IF
		End If
        
		'// 결과수 카운트
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_weekly w"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

		'response.write sqlStr &"<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
		
		if FTotalCount < 1 then exit function
					
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " w.idx, w.title, w.state, w.banner_img, w.title_img, w.startdate, "
        sqlStr = sqlStr & " w.regdate, w.comment, w.lastadminid, "
        sqlStr = sqlStr & " w.partMDid ,w.partWDid "
		sqlStr = sqlStr & " ,(case when w.state = 7 and getdate() <= w.startdate then '6'" + vbcrlf
		sqlStr = sqlStr & " 	else w.state end) as statename" + vbcrlf        
        sqlStr = sqlStr & " ,isnull((select count(itemidx)"
        sqlStr = sqlStr & " 	from db_giftplus.dbo.tbl_stylelife_weekly_item"
        sqlStr = sqlStr & " 	where w.idx = idx),0) as itemcnt"
        sqlStr = sqlStr & "	,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = w.partMDid) as partMDname"
        sqlStr = sqlStr & " ,(SELECT username from db_partner.dbo.tbl_user_tenbyten WHERE userid = w.partWDid) as partWDname"
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_weekly w"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch
        sqlStr = sqlStr & " order by w.startdate DESC, w.idx desc"
		
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
                set FItemList(i) = new cStyleLife_oneitem
                
                FItemList(i).fpartMDname            = rsget("partMDname")
                FItemList(i).fpartWDname            = rsget("partWDname")
                FItemList(i).fstatename            = rsget("statename")
                FItemList(i).fidx            = rsget("idx")
				FItemList(i).ftitle            = db2html(rsget("title"))
				FItemList(i).fstate            = rsget("state")
                FItemList(i).fbanner_img            = db2html(rsget("banner_img"))
                FItemList(i).ftitle_img            = db2html(rsget("title_img"))
				FItemList(i).fstartdate            = db2html(rsget("startdate"))
                FItemList(i).fregdate            = rsget("regdate")
				FItemList(i).fcomment            = db2html(rsget("comment"))
				FItemList(i).fpartMDid            = rsget("partMDid")
				FItemList(i).fpartWDid            = rsget("partWDid")
				FItemList(i).fitemcnt            = rsget("itemcnt")
				                
                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
    
    
	Public Sub fnGetWeekly_item()
		dim sqlstr,i , sqlsearch
		
		if frectidx <> "" then
			sqlsearch = sqlsearch & " and idx = "&frectidx&""
		end if
		
		sqlstr = "select top 1"
        sqlStr = sqlStr & " w.idx, w.title, w.cd1, w.state,w.banner_img,w.title_img,w.startdate "
        sqlStr = sqlStr & " ,w.regdate,w.comment,w.lastadminid"
        sqlStr = sqlStr & " ,w.partMDid ,w.partWDid "
        sqlStr = sqlStr & " from db_giftplus.dbo.tbl_stylelife_weekly w"
        sqlStr = sqlStr & " where 1=1 " & sqlsearch

        'response.write sqlStr&"<br>"
        rsget.Open SqlStr, dbget, 1
        ftotalcount = rsget.RecordCount
        
        set FOneItem = new cStyleLife_oneitem
        
        if Not rsget.Eof then
			
            foneitem.fidx            = rsget("idx")
			foneitem.ftitle            = db2html(rsget("title"))
			foneitem.fcd1            	= rsget("cd1")
			foneitem.fstate            = rsget("state")
            foneitem.fbanner_img            = db2html(rsget("banner_img"))
            foneitem.ftitle_img            = db2html(rsget("title_img"))
			foneitem.fstartdate            = db2html(rsget("startdate"))
            foneitem.fregdate            = rsget("regdate")
			foneitem.fcomment            = db2html(rsget("comment"))
			foneitem.fpartMDid            = rsget("partMDid")
			foneitem.fpartWDid            = rsget("partWDid")

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


Function StyleLifeItemComma(value)
	Dim vBody, i
	For i = LBound(Split(value,",")) To UBound(Split(value,","))
		vBody = vBody & StyleName(Split(value,",")(i)) & ","
	Next
	vBody = Left(vBody,Len(vBody)-1)
	StyleLifeItemComma = vBody
End Function

Function StyleNameSelectBox(itemid,itemcate)
	Dim vBody
	vBody = "<select name='stylecate' class='select' onChange='nowChangeCate(this.value)'>"
	vBody = vBody & "<option value=''>- 스타일미지정 -</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|010'>클래식</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|020'>큐트</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|040'>모던</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|050'>네추럴</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|060'>오리엔탈</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|070'>팝</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|080'>로맨틱</option>"
	vBody = vBody & "<option value='"&itemid&"|"&itemcate&"|090'>빈티지</option>"
	vBody = vBody & "</select>"
	StyleNameSelectBox = vBody
End Function

Function StyleName(code)
	SELECT Case code
		Case "010"
			StyleName = "클래식"
		Case "020"
			StyleName = "큐트"
		Case "030"
			StyleName = "댄디"
		Case "040"
			StyleName = "모던"
		Case "050"
			StyleName = "네추럴"
		Case "060"
			StyleName = "오리엔탈"
		Case "070"
			StyleName = "팝"
		Case "080"
			StyleName = "로맨틱"
		Case "090"
			StyleName = "빈티지"
		Case Else
			StyleName = ""
	End SELECT
End Function

Function fnStyleCate2(cate)
	SELECT Case Left(cate,3)
		Case "010", "020"
			fnStyleCate2 = "010"
		Case "030", "035"
			fnStyleCate2 = "040"
		Case "040", "050", "045", "055", "060"
			fnStyleCate2 = "020"
		Case "080", "090", "070", "075"
			fnStyleCate2 = "030"
		Case "100"
			fnStyleCate2 = "050"
		Case "110"
			fnStyleCate2 = "060"
		Case "025"
			fnStyleCate2 = "060"
		Case Else
			fnStyleCate2 = ""
	End SELECT
End Function

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
	strSql = strSql & " WHERE part_sn IN("&part_sn&") and  posit_sn>='4' and  posit_sn<='12' and   isUsing=1" & vbcrlf

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

function Drawcategory(selectBoxName,selectedId,changeFlag,catetype)
dim tmp_str,query1
%>
	<select name="<%=selectBoxName%>" <%= changeFlag %>>
	<option value='' <%if selectedId="" then response.write " selected"%> >선택</option>
	<%
	if catetype = "CD1" then
		query1 = " select cd1 ,catename ,isusing ,orderno ,lastadminid"
		query1 = query1 & " from db_giftplus.dbo.tbl_stylepick_cate_cd1"
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
%>