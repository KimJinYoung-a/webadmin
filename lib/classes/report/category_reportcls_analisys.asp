<%

Class CCategoryReportItem
	public Fselltotal
	public Fbuytotal
	public Fsellcnt
	public FCLarge
	public FCLName

	public FCmid
	public FCMName

	public FCsmall
	public FCSName

	public maxt
	public maxc


	public FItemNo
	public FItemID
	public FItemCost
	public FItemName
	public FItemOptionStr
	public FBuycash
	public FMakerid
	public FSellcash

	public FImageSmall

    public Fitemcouponyn     
    public Fcurritemcouponidx
    public Fitemcoupontype   
    public Fitemcouponvalue  
    public Fcouponbuyprice   
    public Forgprice
    public Forgsuplycash
    public Fsailprice
    public Fsailsuplycash
    public Fsailyn 
    
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

 
	
		'// 상품 쿠폰 여부
	public Function IsCouponItem() '!
			IsCouponItem = (FItemCouponYN="Y")
	end Function

	'// 세일포함 실제가격
	public Function getRealPrice() '! 
		getRealPrice = FSellCash 
	end Function
	
		'// 쿠폰 할인가
	public Function GetCouponDiscountPrice() '?
		Select case Fitemcoupontype
			case "1" ''% 쿠폰
				GetCouponDiscountPrice = CLng(Fitemcouponvalue*getRealPrice/100)
			case "2" ''원 쿠폰
				GetCouponDiscountPrice = Fitemcouponvalue
			case "3" ''무료배송 쿠폰
			    GetCouponDiscountPrice = 0
			case else
				GetCouponDiscountPrice = 0
		end Select

    end Function
    
    public Function GetCouponAssignPrice() '!
		if (IsCouponItem) then
			GetCouponAssignPrice = getRealPrice - GetCouponDiscountPrice
		else
			GetCouponAssignPrice = getRealPrice
		end if
	end Function
end Class

class CCategoryReport
	public maxt
	public maxc

	public FItemList()

	public FPageSize
	public FCurrPage
	public FTotalPage
  public FPageCount
	public FTotalCount
	public FResultCount
  public FScrollCount
	public FRectFromDate
	public FRectToDate

	public FRectCD1
	public FRectCD2
	public FRectCD3

	public FRectDispY
	public FRectSellY
	public FRectRdsite
	public FRectOldJumun
	public FRectSitename
	public FRectDepth1
	public FRectGpRdsite

	public FRectOrdertype
	public FRectChannelDiv
  public FRectPurchasetype
  public FRectInc3pl
	public FRectDispCate
	public FRectOptExists
	
	public FRectSellChannelDiv
	
	public sub SearchCategorySellrePort()
		Dim sql, i

   		maxt = -1
    	maxc = -1



			sql = "select top 100 count(d.itemno) as sellcnt, sum(d.itemcost*d.itemno) as sumtotal,"
			sql = sql + " l.code_large, l.code_nm"

			if FRectOldJumun="on" then
				sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m"
				sql = sql + " 	Join [db_log].[dbo].tbl_old_order_detail_2003 d"
				sql = sql + " 	on m.orderserial=d.orderserial"
			else
				sql = sql + " from [db_order].[dbo].tbl_order_master m "
				sql = sql + " 	Join [db_order].[dbo].tbl_order_detail d"
				sql = sql + " 	on m.orderserial=d.orderserial"
			end if

			sql = sql + " Join [db_item].[dbo].tbl_item i"
			sql = sql + " on d.itemid=i.itemid"
			sql = sql + " Join [db_item].[dbo].tbl_cate_large l"
			sql = sql + " on i.cate_large=l.code_large"
			sql = sql + " where m.regdate>='" & FRectFromDate & "'"
			sql = sql + " and m.regdate<'" & FRectToDate & "'"
			sql = sql + " and m.ipkumdiv>3"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and m.jumundiv<>9"
			sql = sql + " and d.itemid<>0"
			sql = sql + " and d.cancelyn<>'Y'"
			sql = sql + " group by  l.code_large,  l.code_nm"
			sql = sql + " order by l.code_large"


        if FRectOldJumun="on" then
            db3_rsget.CursorLocation = adUseClient
		    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
			''rsget.Open sql,dbget,1

			FResultCount = db3_rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until db3_rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).Fselltotal = db3_rsget("sumtotal")
				FItemList(i).Fsellcnt = db3_rsget("sellcnt")
				FItemList(i).FCLarge = db3_rsget("code_large")
				FItemList(i).FCLName = db3_rsget("code_nm")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if
				db3_rsget.MoveNext
				i = i + 1
			loop
			db3_rsget.close
        else
            rsget.CursorLocation = adUseClient
		    rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
			''rsget.Open sql,dbget,1

			FResultCount = rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				FItemList(i).FCLarge = rsget("code_large")
				FItemList(i).FCLName = rsget("code_nm")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if
				rsget.MoveNext
				i = i + 1
			loop
			rsget.close
        end if
	end sub


	public sub SearchCategorySellrePortDetail()
		Dim sql, i

    	maxt = -1
    	maxc = -1


		'#################################################'
		'데이타.'
		'#################################################'

			sql = "select top 100 count(d.itemno) as sellcnt, sum(d.itemcost*d.itemno) as sumtotal,"
			sql = sql + " l.code_mid,l.code_nm"

			if FRectOldJumun="on" then
				sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
				sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
			else
				sql = sql + " from [db_order].[dbo].tbl_order_master m, "
				sql = sql + " [db_order].[dbo].tbl_order_detail d,"
			end if

			sql = sql + "[db_item].[dbo].tbl_item i,"
			sql = sql + " [db_item].[dbo].tbl_cate_mid l"
			sql = sql + " where m.orderserial=d.orderserial"
			sql = sql + " and m.regdate>='" & FRectFromDate & "'"
			sql = sql + " and m.regdate<'" & FRectToDate & "'"
			sql = sql + " and m.ipkumdiv>3"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and m.jumundiv<>9"
			sql = sql + " and d.itemid<>0"
			sql = sql + " and d.cancelyn<>'Y'"
			sql = sql + " and d.itemid=i.itemid"
			sql = sql + " and i.cate_large='" + FRectCD1 + "'"
			sql = sql + " and l.code_large='" + FRectCD1 + "'"
			sql = sql + " and i.cate_mid=l.code_mid"
			sql = sql + " group by l.code_mid,code_nm"
			sql = sql + " order by l.code_mid"

			'response.write sql
	    if FRectOldJumun="on" then
			db3_rsget.CursorLocation = adUseClient
		    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = db3_rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until db3_rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FCmid  = db3_rsget("code_mid")
				FItemList(i).FCMName = db3_rsget("code_nm")
				FItemList(i).Fselltotal = db3_rsget("sumtotal")
				FItemList(i).Fsellcnt = db3_rsget("sellcnt")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
			loop
			db3_rsget.close
        else
            rsget.Open sql,dbget,1
			FResultCount = rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FCmid  = rsget("code_mid")
				FItemList(i).FCMName = rsget("code_nm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
			loop
			rsget.close
        end if
	end sub

	public sub SearchCategorySellrePortSubDetail()
		Dim sql, i

    	maxt = -1
    	maxc = -1


		''#################################################
		''데이타.
		''#################################################
        ''[db_item].[dbo].tbl_item -> [db_item].[dbo].tbl_item_Category 로 수정 .. 특정쿼리 느림..?

			sql = "select top 100 count(d.itemno) as sellcnt, sum(d.itemcost*d.itemno) as sumtotal,"
			sql = sql + " i.code_small,l.code_nm"
			if FRectOldJumun="on" then
				sql = sql + " from [db_log].[dbo].tbl_old_order_master_2003 m,"
				sql = sql + " [db_log].[dbo].tbl_old_order_detail_2003 d,"
			else
				sql = sql + " from [db_order].[dbo].tbl_order_master m, "
				sql = sql + " [db_order].[dbo].tbl_order_detail d,"
			end if

			sql = sql + " [db_item].[dbo].tbl_item_Category i,"
			sql = sql + " [db_item].[dbo].tbl_cate_small l"
			sql = sql + " where m.orderserial=d.orderserial"
			sql = sql + " and m.regdate>='" & FRectFromDate & "'"
			sql = sql + " and m.regdate<'" & FRectToDate & "'"
			sql = sql + " and m.ipkumdiv>'3'"
			sql = sql + " and m.cancelyn='N'"
			sql = sql + " and m.jumundiv<>'9'"
			sql = sql + " and d.itemid<>0"
			sql = sql + " and d.cancelyn<>'Y'"
			sql = sql + " and d.itemid=i.itemid"
			sql = sql + " and i.code_large='" + FRectCD1 + "'"
			sql = sql + " and i.code_mid='" + FRectCD2 + "'"
			sql = sql + " and i.code_div='D'"
			sql = sql + " and i.code_large=l.code_large"
			sql = sql + " and i.code_mid=l.code_mid"
			sql = sql + " and i.code_small=l.code_small"
			sql = sql + " group by i.code_small,l.code_nm"
			sql = sql + " order by i.code_small"

'response.write sql

        if FRectOldJumun="on" then
			db3_rsget.CursorLocation = adUseClient
		    db3_rsget.Open sql,db3_dbget,adOpenForwardOnly, adLockReadOnly
			FResultCount = db3_rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until db3_rsget.eof

				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FCsmall  = db3_rsget("code_small")
				FItemList(i).FCSName = db3_rsget("code_nm")
				FItemList(i).Fselltotal = db3_rsget("sumtotal")
				FItemList(i).Fsellcnt = db3_rsget("sellcnt")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				db3_rsget.MoveNext
				i = i + 1
			loop

			db3_rsget.close
        else
            rsget.Open sql,dbget,1
			FResultCount = rsget.RecordCount
		    redim preserve FItemList(FResultCount)
			do until rsget.eof

				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FCsmall  = rsget("code_small")
				FItemList(i).FCSName = rsget("code_nm")
				FItemList(i).Fselltotal = rsget("sumtotal")
				FItemList(i).Fsellcnt = rsget("sellcnt")
				if Not IsNull(FItemList(i).Fselltotal) then
					maxt = MaxVal(maxt,FItemList(i).Fselltotal)
					maxc = MaxVal(maxc,FItemList(i).Fsellcnt)
				end if

				rsget.MoveNext
				i = i + 1
			loop

			rsget.close
        end if
	end sub

	public Sub SearchCategoryBestseller()
		dim sqlStr
		dim i

		sqlStr = "select top " + CStr(FPageSize)
		If FRectOptExists = "on" Then
			sqlStr = sqlStr + " sum(d.itemno) as itemno ,sum(d.itemno*d.itemcost) as sellsum, sum(d.itemno*d.buycash)as buysum, d.itemid, d.itemcost ,"
			sqlStr = sqlStr + " d.itemname, d.makerid, d.itemoptionname, i.smallimage"
		Else
			sqlStr = sqlStr + " sum(d.itemno) as itemno ,sum(d.itemno*d.itemcost) as sellsum, sum(d.itemno*d.buycash)as buysum, d.itemid, 0 as itemcost, "
			sqlStr = sqlStr + " d.itemname, d.makerid, '' as itemoptionname, i.smallimage"
		End If
		if FRectOldJumun="on" then
			sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
			sqlStr = sqlStr + " Join [db_log].[dbo].tbl_old_order_detail_2003 d"
			sqlStr = sqlStr + " on m.orderserial=d.orderserial"
		else
			sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m "
			sqlStr = sqlStr + " Join [db_order].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " on m.orderserial=d.orderserial"
		end if
		sqlStr = sqlStr + " Join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr + " on d.itemid=i.itemid"
	 IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가 
				sqlStr = sqlStr + "  JOIN db_item.dbo.tbl_display_cate_item as dc "
				sqlStr = sqlStr + " on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
	END IF
		If FRectPurchasetype <> "" Then
    		sqlStr = sqlStr & " LEFT JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
    	End IF

    	sqlStr = sqlStr & "       left join db_partner.dbo.tbl_partner p2"
	    sqlStr = sqlStr & "       on m.sitename=p2.id "

		sqlStr = sqlStr + " where m.ipkumdiv>'3'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		if (FRectFromDate<>"") then
			sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectFromDate) + "'"
		end if

		if (FRectToDate<>"") then
			sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectToDate) + "'"
		end if

		if (FRectCD1<>"") then
			sqlStr = sqlStr + " and i.cate_large=" + CStr(FRectCD1) + ""  ''인덱스 안타게..?
		end if

		if (FRectCD2<>"") then
			sqlStr = sqlStr + " and i.cate_mid=" + CStr(FRectCD2) + ""
		end if

		if (FRectCD3<>"") then
			sqlStr = sqlStr + " and i.cate_small=" + CStr(FRectCD3) + ""
		end if

		if FRectDispY="on" then
			sqlStr = sqlStr + " and i.dispyn='Y'"
		end if

		if FRectSellY="on" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if

		if (FRectChannelDiv<>"") then
			if FRectChannelDiv="web" then
				sqlStr = sqlStr & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
			elseif FRectChannelDiv="jaehu" then
				sqlStr = sqlStr & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			elseif FRectChannelDiv="mjaehu" then
				sqlStr = sqlStr & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			elseif FRectChannelDiv="mobile" then
				sqlStr = sqlStr & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") and m.accountdiv<>'50' "
			elseif FRectChannelDiv="ipjum" then
				sqlStr = sqlStr & " and isNULL(m.rdsite,m.sitename) in(" & fnChannelDiv(FRectChannelDiv) & ") "
			end if
		end if

		if FRectRdsite="on" then
			sqlStr = sqlStr + " and m.rdsite in ('mobile','mobile_kakaotalk','mobile_nate','mobile_kakaotms') "
		elseif FRectRdsite<>"" then
			sqlStr = sqlStr + " and (m.rdsite='" & FRectRdsite & "' or m.sitename='" & FRectRdsite & "')"
		end if
        If FRectPurchasetype <> "" Then
		    sqlStr = sqlStr & " and p.purchasetype = '" & FRectPurchasetype &"'"
	    End IF

	    ''2014/01/27추가
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
                sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')<>''"
            end if
        else
            sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')=''"
        end if
		If FRectOptExists = "on" Then
			sqlStr = sqlStr + " group by d.itemid, d.itemcost, d.itemname, d.makerid, d.itemoptionname ,i.smallimage"
		Else
			sqlStr = sqlStr + " group by d.itemid, d.itemname, d.makerid, i.smallimage"
		End If

		'정렬방법
		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr + " order by sellsum Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr + " order by sum(d.itemno*(d.itemcost-d.buycash)) Desc"
			Case "unitCost"
				'객단가순
				If FRectOptExists = "on" Then
					sqlStr = sqlStr + " order by d.itemcost Desc"
				else
					sqlStr = sqlStr + " order by sum(d.itemcost)/sum(Case When d.itemno>0 then d.itemno else 1 end) Desc"
				end if
			Case Else
				'수량순
				sqlStr = sqlStr + " order by itemno Desc, sellsum desc"
		end Select

'response.write sqlStr
'response.end
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		''rsget.Open sqlStr,dbget,1
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).Fselltotal       = rsget("sellsum")
				FItemList(i).Fbuytotal       = rsget("buysum")

				FItemList(i).FItemNo       = rsget("itemno")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemCost       = rsget("itemcost")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerid		= rsget("makerid")

				FItemList(i).FImageSmall	= rsget("smallimage")

				if IsNULL(FItemList(i).FImageSmall) then

				else
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
				end if
				rsget.movenext
				i=i+1
			loop
		rsget.Close
	end Sub


    '' anal 서버로 변경
    public Sub ONSearchCategoryBestseller()
		dim sqlStr
		dim i

		sqlStr = "select top " + CStr(FPageSize)
		If FRectOptExists = "on" Then
			sqlStr = sqlStr + " sum(d.itemno) as itemno ,sum(d.itemno*d.itemcost) as sellsum, sum(d.itemno*d.buycash)as buysum, d.itemid, d.itemcost ,"
			sqlStr = sqlStr + " i.itemname, d.makerid, d.itemoptionname, i.smallimage"
		Else
			sqlStr = sqlStr + " sum(d.itemno) as itemno ,sum(d.itemno*d.itemcost) as sellsum, sum(d.itemno*d.buycash)as buysum, d.itemid, 0 as itemcost, "
			sqlStr = sqlStr + " i.itemname, d.makerid, '' as itemoptionname, i.smallimage"
		End If
		
		    sqlStr = sqlStr + " , i.sellcash, i.buycash ,i.orgprice, i.orgsuplycash, i.sailprice,i.sailsuplycash,i.sailyn,i.itemcouponyn,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue "		
		
		'' 쓰이는곳이 없음.. 주석처리..
		''    sqlStr = sqlStr + " , Case itemCouponyn When 'Y' then (Select top 1 couponbuyprice From [db_item].[dbo].tbl_item_coupon_detail Where itemcouponidx=i.curritemcouponidx and itemid=i.itemid) end as couponbuyprice "		 
		
		'if FRectOldJumun="on" then
		'	sqlStr = sqlStr + " from [db_log].[dbo].tbl_old_order_master_2003 m"
		'	sqlStr = sqlStr + " Join [db_log].[dbo].tbl_old_order_detail_2003 d"
		'	sqlStr = sqlStr + " on m.orderserial=d.orderserial"
		'else
			sqlStr = sqlStr + " from [db_analyze_data_raw].[dbo].tbl_order_master m "
			sqlStr = sqlStr + " Join [db_analyze_data_raw].[dbo].tbl_order_detail d"
			sqlStr = sqlStr + " on m.orderserial=d.orderserial"
		'end if
		
		sqlStr = sqlStr + " Join [db_analyze_data_raw].[dbo].tbl_item i"
		sqlStr = sqlStr + " on d.itemid=i.itemid"
	 IF FRectDispCate<>"" THEN	'2014-02-27 정윤정 전시카테고리 검색 추가 
				sqlStr = sqlStr + "  JOIN [db_analyze_data_raw].dbo.tbl_display_cate_item as dc "
				sqlStr = sqlStr + " on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
	 END IF
		If FRectPurchasetype <> "" Then
    		sqlStr = sqlStr & " LEFT JOIN [db_analyze_data_raw].[dbo].[tbl_partner] as p on d.makerid = p.id "
    	End IF

    	sqlStr = sqlStr & "       left join [db_analyze_data_raw].dbo.tbl_partner p2"
	    sqlStr = sqlStr & "       on m.sitename=p2.id "

		sqlStr = sqlStr + " where m.ipkumdiv>'3'"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"

		if (FRectFromDate<>"") then
			sqlStr = sqlStr + " and m.regdate >='" + CStr(FRectFromDate) + "'"
		end if

		if (FRectToDate<>"") then
			sqlStr = sqlStr + " and m.regdate <'" + CStr(FRectToDate) + "'"
		end if

		if (FRectCD1<>"") then
			sqlStr = sqlStr + " and i.cate_large=" + CStr(FRectCD1) + ""  ''인덱스 안타게..?
		end if

		if (FRectCD2<>"") then
			sqlStr = sqlStr + " and i.cate_mid=" + CStr(FRectCD2) + ""
		end if

		if (FRectCD3<>"") then
			sqlStr = sqlStr + " and i.cate_small=" + CStr(FRectCD3) + ""
		end if

		if FRectDispY="on" then
			sqlStr = sqlStr + " and i.dispyn='Y'"
		end if

		if FRectSellY="on" then
			sqlStr = sqlStr + " and i.sellyn='Y'"
		end if
 

        if (FRectSellChannelDiv<>"") then   '변경 2015.04.17 정윤정
       		sqlStr = sqlStr & " and m.beadaldiv in ("&getChannelvalue2ArrIDxGroup(FRectSellChannelDiv)&")"
    	end if
 
		
	    if FRectRdsite="on" then '모바일+APP 사이트만
	        sqlStr = sqlStr & " and m.beadaldiv in ('4','5','7','8')"
	    end if
    	
        If FRectPurchasetype <> "" Then
		    sqlStr = sqlStr & " and p.purchasetype = '" & FRectPurchasetype &"'"
	    End IF

	    ''2014/01/27추가
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
                sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')<>''"
            end if
        else
            sqlStr = sqlStr & " and isNULL(p2.tplcompanyid,'')=''"
        end if
		If FRectOptExists = "on" Then
			sqlStr = sqlStr + " group by d.itemid, d.itemcost, i.itemname, d.makerid, d.itemoptionname ,i.smallimage,i.sellcash, i.buycash ,i.orgprice, i.orgsuplycash, i.sailprice,i.sailsuplycash,i.sailyn,i.itemcouponyn,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue, i.itemid "
		Else
			sqlStr = sqlStr + " group by d.itemid, i.itemname, d.makerid, i.smallimage,i.sellcash, i.buycash ,i.orgprice, i.orgsuplycash, i.sailprice,i.sailsuplycash,i.sailyn,i.itemcouponyn,i.curritemcouponidx,i.itemcoupontype,i.itemcouponvalue, i.itemid "
		End If

		'정렬방법
		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr + " order by sellsum Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr + " order by sum(d.itemno*(d.itemcost-d.buycash)) Desc"
			Case "unitCost"
				'객단가순
				If FRectOptExists = "on" Then
					sqlStr = sqlStr + " order by d.itemcost Desc"
				else
					sqlStr = sqlStr + " order by sum(d.itemcost)/sum(Case When d.itemno>0 then d.itemno else 1 end) Desc"
				end if
			Case Else
				'수량순
				sqlStr = sqlStr + " order by itemno Desc, sellsum desc"
		end Select

''tbl_item_coupon_detail 없음..
response.write sqlStr
'response.end
        rsAnalget.CursorLocation = adUseClient
		rsAnalget.Open sqlStr,dbAnalget,adOpenForwardOnly, adLockReadOnly
		''rsAnalget.Open sqlStr,dbAnalget,1
		FResultCount = rsAnalget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsAnalget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).Fselltotal       = rsAnalget("sellsum")
				FItemList(i).Fbuytotal       = rsAnalget("buysum")

				FItemList(i).FItemNo       = rsAnalget("itemno")
				FItemList(i).FItemID       = rsAnalget("itemid")
				FItemList(i).FItemCost       = rsAnalget("itemcost")
				FItemList(i).FItemName     = db2html(rsAnalget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsAnalget("itemoptionname"))
				FItemList(i).FMakerid		= rsAnalget("makerid")

				FItemList(i).FImageSmall	= rsAnalget("smallimage")

				if IsNULL(FItemList(i).FImageSmall) then

				else
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
				end if
				
				FItemList(i).Fsellcash      = rsAnalget("sellcash")
                FItemList(i).Fbuycash       = rsAnalget("buycash")
                FItemList(i).Forgprice          = rsAnalget("orgprice")
                FItemList(i).Forgsuplycash      = rsAnalget("orgsuplycash")
                FItemList(i).Fsailprice         = rsAnalget("sailprice")
                FItemList(i).Fsailsuplycash     = rsAnalget("sailsuplycash")
                FItemList(i).Fsailyn            = rsAnalget("sailyn") 
                FItemList(i).Fitemcouponyn      = rsAnalget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsAnalget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsAnalget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsAnalget("itemcouponvalue")
                '''FItemList(i).Fcouponbuyprice    = rsAnalget("couponbuyprice")	'쿠폰적용 매입가 //쓰이는곳이 없음..
				
				rsAnalget.movenext
				i=i+1
			loop
		rsAnalget.Close
	end Sub
	
	'2013/12/11 김진영..네이버관련해서 추가
	Public Sub OutmallSearchCategoryBestseller()
		Dim sqlStr, i
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize)
		sqlStr = sqlStr & " sum(d.itemno) as itemno ,sum(d.itemno*d.itemcost) as sellsum, sum(d.itemno*d.buycash)as buysum, d.itemid, d.itemcost, "
		sqlStr = sqlStr & " d.itemname, d.makerid, d.itemoptionname, i.smallimage "
		If FRectOldJumun = "on" Then
			sqlStr = sqlStr & " FROM [db_log].[dbo].tbl_old_order_master_2003 m "
			sqlStr = sqlStr & " Join [db_log].[dbo].tbl_old_order_detail_2003 d on m.orderserial=d.orderserial "
		Else
			sqlStr = sqlStr & " FROM [db_order].[dbo].tbl_order_master m "
			sqlStr = sqlStr & " Join [db_order].[dbo].tbl_order_detail d on m.orderserial=d.orderserial "
		End If
		sqlStr = sqlStr & " Join [db_item].[dbo].tbl_item i on d.itemid = i.itemid"
		sqlStr = sqlStr & " Join db_item.dbo.tbl_Outmall_RdsiteGubun as R on R.gubun='"&FRectSitename&"' and m.rdsite = R.rdsite "
		sqlStr = sqlStr & " where m.ipkumdiv > '1'"
		sqlStr = sqlStr & " and m.cancelyn = 'N'"
		sqlStr = sqlStr & " and d.cancelyn <> 'Y'"
		sqlStr = sqlStr & " and d.itemid <> 0"

		If (FRectFromDate <> "") Then
			sqlStr = sqlStr & " and m.regdate >='" & CStr(FRectFromDate) & "'"
		End If

		If (FRectToDate<>"") Then
			sqlStr = sqlStr & " and m.regdate <'" & CStr(FRectToDate) & "'"
		End If

		If (FRectDepth1 <> "") Then
			sqlStr = sqlStr & " and i.dispcate1 = '" & CStr(FRectDepth1) & "'"
		End If

		If (FRectGpRdsite <> "") Then
			sqlStr = sqlStr & " and R.rdsite = '" & CStr(FRectGpRdsite) & "'"
		End If
		sqlStr = sqlStr & " group by d.itemid, d.itemcost, d.itemname, d.makerid, d.itemoptionname ,i.smallimage"

		'정렬방법
		Select Case FRectOrdertype
			Case "totalprice"
				'매출순
				sqlStr = sqlStr + " order by sellsum Desc"
	    	Case "gain"
	    		'수익순
	            sqlStr = sqlStr + " order by sum(d.itemno*(d.itemcost-d.buycash)) Desc"
			Case "unitCost"
				'객단가순
				sqlStr = sqlStr + " order by d.itemcost Desc"
			Case Else
				'수량순
				sqlStr = sqlStr + " order by itemno Desc, sellsum desc"
		end Select
        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		redim preserve FItemList(FResultCount)
		do until rsget.eof
			set FItemList(i) = new CCategoryReportItem
				FItemList(i).Fselltotal       = rsget("sellsum")
				FItemList(i).Fbuytotal       = rsget("buysum")
				FItemList(i).FItemNo       = rsget("itemno")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemCost       = rsget("itemcost")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FMakerid		= rsget("makerid")
				FItemList(i).FImageSmall	= rsget("smallimage")
				if IsNULL(FItemList(i).FImageSmall) then

				else
					FItemList(i).FImageSmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).FImageSmall
				end if
				rsget.movenext
			i=i+1
		loop
		rsget.Close
	End Sub

    '' 6개월 이전 자료 DataMart로 변경 2013/10/17
	public Sub CategorySearchBestsellerList()
		dim sqlStr
		dim i

		''#################################################
		''데이타.
		''#################################################
		sqlStr = "select top " & CStr(FPageSize)
		sqlStr = sqlStr & "	sum(d.itemno) as sm, d.itemid, d.buycash, d.itemcost, d.itemname, d.makerid, d.itemoptionname "

		if FRectOldJumun="on" then
			'6개월 이전 자료
			sqlStr = sqlStr & " from [db_log].[dbo].tbl_old_order_master_2003 m "
			sqlStr = sqlStr & " Join [db_log].[dbo].tbl_old_order_detail_2003 d "
			sqlStr = sqlStr & "		on m.orderserial=d.orderserial "
		else
			'최근 자료
			sqlStr = sqlStr & " from [db_order].[dbo].tbl_order_master m "
			sqlStr = sqlStr & " Join [db_order].[dbo].tbl_order_detail d "
			sqlStr = sqlStr & "		on m.orderserial=d.orderserial "
		end if

        sqlStr = sqlStr & "	 Join  [db_item].[dbo].tbl_item i "   
        sqlStr = sqlStr & "		on d.itemid=i.itemid"
		sqlStr = sqlStr & "		and cate_large='" + FRectCD1 + "' "
		if (FRectCD2<>"") then
		    sqlStr = sqlStr & "		and cate_mid='" + FRectCD2 + "' "
	    end if
	    if (FRectCD3<>"") then
		    ''sqlStr = sqlStr & "		and cate_small='" + FRectCD3 + "'"
		    sqlStr = sqlStr & "		and cate_large+cate_mid+cate_small='" + FRectCD1+FRectCD2+FRectCD3 + "'"
        end if

		sqlStr = sqlStr & " where m.ipkumdiv>3 "
		sqlStr = sqlStr & "	and m.jumundiv<>9 "
		sqlStr = sqlStr & "	and m.cancelyn='N' "
		sqlStr = sqlStr & "	and d.itemid<>0 "
		sqlStr = sqlStr & "	and d.cancelyn<>'Y' "

		'결과 기간설정
		if (FRectFromDate<>"") then
			sqlStr = sqlStr & " and m.regdate >='" & CStr(FRectFromDate) & "' "
		end if
		if (FRectToDate<>"") then
			sqlStr = sqlStr & " and m.regdate <'" & CStr(FRectToDate) & "' "
		end if


		sqlStr = sqlStr & " group by d.itemid, d.buycash, d.itemcost, d.itemname, d.makerid, d.itemoptionname "
		sqlStr = sqlStr & " order by sm Desc "
'rw sqlStr
    'response.write "점검중-서동석" '' too slow => 데이터 마트로 돌림
	'response.end
	    
    if (TRUE or FRectOldJumun="on") then
        if (FRectOldJumun<>"on") then
            response.write "1시간 지연 데이터"
        end if
        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		''rsget.Open sqlStr,dbget,1
		FResultCount = db3_rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until db3_rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FItemNo       = db3_rsget("sm")
				FItemList(i).FItemID       = db3_rsget("itemid")
				FItemList(i).FItemCost       = db3_rsget("itemcost")
				FItemList(i).FItemName     = db2html(db3_rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(db3_rsget("itemoptionname"))
				FItemList(i).FBuycash		= db3_rsget("buycash")
				FItemList(i).FMakerid		= db3_rsget("makerid")
				db3_rsget.movenext
				i=i+1
			loop
		db3_rsget.Close
	else
	    response.write "점검중" '' too slow => 데이터 마트로 돌림
	    response.end
	
	    rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.recordCount
		''올림.
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		do until rsget.eof
				set FItemList(i) = new CCategoryReportItem
				FItemList(i).FItemNo       = rsget("sm")
				FItemList(i).FItemID       = rsget("itemid")
				FItemList(i).FItemCost       = rsget("itemcost")
				FItemList(i).FItemName     = db2html(rsget("itemname"))
				FItemList(i).FItemOptionStr= db2html(rsget("itemoptionname"))
				FItemList(i).FBuycash		= rsget("buycash")
				FItemList(i).FMakerid		= rsget("makerid")
				rsget.movenext
				i=i+1
			loop
		rsget.Close
    end if
	end Sub

	Private Sub Class_Initialize()

		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	function MaxVal(a,b)
		if (CDbl(a)> CDbl(b)) then
			MaxVal=a
		else
			MaxVal=b
		end if
	end function
end class

Sub Drawsitename(selectboxname, sitename, onchgYN)
	dim strSQL, tem_str, chgscript
	strSQL = " SELECT gubun FROM "
	strSQL = strSQL & " db_item.dbo.tbl_Outmall_RdsiteGubun "
	strSQL = strSQL & " GROUP BY gubun"
	rsget.Open strSQL, dbget, 1
	If onchgYN = "Y" Then
		chgscript = "onchange=jsGrouplist(this.value)"
	End If
	response.write "<select name='" & selectboxname & "' class='select' "&chgscript&" >"
	response.write "<option value=''>--선택--</option>"
	If not rsget.EOF then
		Do until rsget.EOF
			If Lcase(sitename) = Lcase(rsget("gubun")) Then
				tem_str = " selected"
			End If
			response.write "<option value='" & rsget("gubun") & "' " & tem_str & ">" & rsget("gubun") & "</option>"
			tem_str = ""
			rsget.movenext
		Loop
	End If
	rsget.close
	response.write "</select>"
End Sub

Sub RdsiteGubunList(gubun, selectGpRdsitename, GpRdsite)
	dim strSQL, tem_str
	strSQL = " SELECT rdsite, explain FROM "
	strSQL = strSQL & " db_item.dbo.tbl_Outmall_RdsiteGubun "
	strSQL = strSQL & " WHERE gubun='"&gubun&"' "
	rsget.Open strSQL, dbget, 1
	response.write "<select name='" & selectGpRdsitename & "' class='select'>"
	response.write "<option value=''>--선택--</option>"
	If not rsget.EOF then
		Do until rsget.EOF
			If Lcase(GpRdsite) = Lcase(rsget("rdsite")) Then
				tem_str = " selected"
			End If
			response.write "<option value='" & rsget("rdsite") & "' " & tem_str & ">" & rsget("explain") & "</option>"
			tem_str = ""
			rsget.movenext
		Loop
	End If
	rsget.close
	response.write "</select>"
End Sub
%>