<%
'###############################################
' Discription : 모바일 mdpick 클래스
' History : 2013.12.17 한용민
'###############################################

Class cmdpick_oneitem
	public fitemisusing
	public fidx
	public fitemid
	public fisusing
	public forderno
	public fregdate
	public flastdate
	public fregadminid
	public flastadminid
	public Fmakerid
	public Fcate_large
	public Fcate_mid
	public Fcate_small
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
	public fbasicimage
	public Fitemcouponyn
	public Fcurritemcouponidx
	public Fitemcoupontype
	public Fitemcouponvalue
	public Fstartdate
	public Fenddate

    public function IsEndDateExpired()
        IsEndDateExpired = now()>Cdate(Fenddate)
    end function	
End Class

Class cmdpick
	Public FItemList()
	Public FOneItem
	Public FTotalCount
	Public FPageSize
	Public FCurrPage
	Public FResultCount
	Public FTotalPage
	Public FPageCount
	Public FScrollCount
	
	public FRectMakerid
	public FRectItemDiv
	public FRectItemid
	public FRectItemName
	public FRectSellYN
	public FRectitemIsUsing
	public FRectDanjongyn
	public FRectMWDiv
	public FRectLimityn
	public FRectCate_Large
	public FRectCate_Mid
	public FRectCate_Small
	public FRectDispCate
	public FRectSailYn
	public FRectCouponYn
	public FRectVatYn
	public FRectIsOversea
	public FRectisUsing
	public frectidx
	
	'//admin/mobile/mdpick/mdpick_edit.asp
	Public Sub getmdpick_one
		Dim sqlStr, i, sqlsearch
		
		if frectisusing<>"" then
			sqlsearch = sqlsearch & " and isusing='"& frectisusing &"'"
		end if
		if FRectIdx<>"" then
			sqlsearch = sqlsearch & " and idx="& FRectIdx &""
		end if
		
		sqlStr = "SELECT TOP 1"
		sqlStr = sqlStr & " idx, itemid, isusing, orderno, regdate, lastdate, regadminid, lastadminid, startdate, enddate"
		sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_mobile_main_mdpick"
		sqlStr = sqlStr & " WHERE 1=1 " & sqlsearch
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
		ftotalcount = rsget.recordcount
		
        SET FOneItem = new cmdpick_oneitem
	        If Not rsget.Eof then

                FOneItem.fstartdate            = rsget("startdate")
                FOneItem.fenddate            = rsget("enddate")                
                FOneItem.Fidx            = rsget("idx")
                FOneItem.Fitemid         = rsget("itemid")
                FOneItem.Fisusing        = rsget("isusing")
                FOneItem.Forderno        = rsget("orderno")
                FOneItem.Fregdate        = rsget("regdate")
                FOneItem.Flastdate       = rsget("lastdate")
                FOneItem.Fregadminid     = rsget("regadminid")
                FOneItem.Flastadminid    = rsget("lastadminid")
					
        	End If
        rsget.Close
	End Sub
	
	'//admin/mobile/mdpick/mdpick_list.asp
	public function getmdpick_list()
        dim sqlStr, addSql, i

        '// 추가 쿼리
        if (FRectMakerid <> "") then
            addSql = addSql & " and i.makerid='" + FRectMakerid + "'"
        end if

        if (FRectItemDiv<> "") then
            addSql = addSql & " and i.itemdiv='" + FRectItemDiv + "'"
        end if

        if (FRectItemid <> "") then
            if right(trim(FRectItemid),1)="," then
            	FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + Left(FRectItemid,Len(FRectItemid)-1) + ")"
            else
				FRectItemid = Replace(FRectItemid,",,",",")
            	addSql = addSql & " and i.itemid in (" + FRectItemid + ")"
            end if
        end if

        if (FRectItemName <> "") then
            addSql = addSql & " and i.itemname like '%" + html2db(FRectItemName) + "%'"
        end if

        if (FRectSellYN="YS") then
            addSql = addSql & " and i.sellyn<>'N'"
        elseif (FRectSellYN <> "") then
            addSql = addSql & " and i.sellyn='" + FRectSellYN + "'"
        end if
        if (FRectIsUsing <> "") then
            addSql = addSql & " and k.isusing='" + FRectIsUsing + "'"
        end if
        if (FRectitemIsUsing <> "") then
            addSql = addSql & " and i.isusing='" + FRectitemIsUsing + "'"
        end if

        if FRectDanjongyn="SN" then
            addSql = addSql + " and i.danjongyn<>'Y'"
            addSql = addSql + " and i.danjongyn<>'M'"
        elseif FRectDanjongyn="YM" then
            addSql = addSql + " and i.danjongyn<>'N'"
            addSql = addSql + " and i.danjongyn<>'S'"
        elseif FRectDanjongyn<>"" then
            addSql = addSql + " and i.danjongyn='" + FRectDanjongyn + "'"
        end if

        if FRectMWDiv="MW" then
            addSql = addSql + " and (i.mwdiv='M' or i.mwdiv='W')"
        elseif FRectMWDiv<>"" then
            addSql = addSql + " and i.mwdiv='" + FRectMwDiv + "'"
        end if

		if FRectLimityn="Y0" then
            addSql = addSql + " and i.limityn='Y' and (i.limitno-i.limitsold<1)"
        elseif FRectLimityn<>"" then
            addSql = addSql + " and i.limityn='" + FRectLimityn + "'"
        end if

        if FRectCate_Large<>"" then
            addSql = addSql + " and i.cate_large='" + FRectCate_Large + "'"
        end if

        if FRectCate_Mid<>"" then
            addSql = addSql + " and i.cate_mid='" + FRectCate_Mid + "'"
        end if

        if FRectCate_Small<>"" then
            addSql = addSql + " and i.cate_small='" + FRectCate_Small + "'"
        end if

		if FRectDispCate<>"" then
			addSql = addSql + " and i.itemid in (select itemid from db_item.dbo.tbl_display_cate_item where catecode like '" + FRectDispCate + "%' and isDefault='y') "
		end if

        if FRectSailYn<>"" then
            addSql = addSql + " and i.sailyn='" + FRectSailYn + "'"
        end if

        if FRectCouponYn<>"" then
            addSql = addSql + " and i.itemCouponyn='" + FRectCouponYn + "'"
        end if

        if FRectVatYn<>"" then
            addSql = addSql + " and i.vatinclude='" + FRectVatYn + "'"
        end if

        if FRectIsOversea<>"" then
			addSql = addSql + " and i.deliverOverseas='" + FRectIsOversea + "'"
			if FRectIsOversea="Y" then
				addSql = addSql + " and i.itemWeight>0 "
			else
				addSql = addSql + " and i.itemWeight<=0 "
			end if
        end if
        
		sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_main_mdpick k"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on k.itemid=i.itemid"
        sqlStr = sqlStr & " where 1=1 " & addSql

		'response.write  sqlStr & "<Br>"
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close
                
        '// 본문 내용 접수
        sqlStr = "select top " + Cstr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " k.idx, k.itemid, k.isusing, k.orderno, k.regdate, k.lastdate, k.regadminid, k.lastadminid, k.startdate, k.enddate"
        sqlStr = sqlStr & " ,i.listimage120, i.listimage, i.basicimage, i.smallimage, i.itemcouponvalue, i.itemcoupontype, i.curritemcouponidx"
        sqlStr = sqlStr & " ,i.itemcouponyn, i.brandname, i.upchemanagecode, i.itemrackcode, i.optioncnt, i.evalcnt"
        sqlStr = sqlStr & " ,i.limitsold, i.limitno, i.pojangok, i.ismobileitem, i.deliverfixday, i.deliverarea, i.deliverytype"
        sqlStr = sqlStr & " ,i.vatinclude, i.specialuseritem, i.mwdiv, i.isextusing, i.isusing as itemisusing, i.sailyn, i.danjongyn"
        sqlStr = sqlStr & " ,i.limityn, i.sellyn, i.sellEndDate, i.lastupdate, i.regdate, i.mileage, i.sailsuplycash"
        sqlStr = sqlStr & " ,i.sailprice, i.orgsuplycash, i.orgprice, i.buycash, i.sellcash, i.itemname, i.itemgubun"
        sqlStr = sqlStr & " ,i.itemdiv, i.cate_small, i.cate_mid, i.cate_large, i.makerid"
        sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_main_mdpick k"
        sqlStr = sqlStr & " join [db_item].[dbo].tbl_item i"
		sqlStr = sqlStr & " 	on k.itemid=i.itemid"
        sqlStr = sqlStr & " where 1=1 " & addSql		
		sqlStr = sqlStr & " Order by k.orderno asc, k.idx desc"

		'response.write  sqlStr & "<Br>"
        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1

        FtotalPage =  Clng(FTotalCount\FPageSize)
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
                set FItemList(i) = new cmdpick_oneitem
				
				FItemList(i).fstartdate            = rsget("startdate")
				FItemList(i).fenddate            = rsget("enddate")
                FItemList(i).Fidx            = rsget("idx")
                FItemList(i).Fitemid         = rsget("itemid")
                FItemList(i).Fisusing        = rsget("isusing")
                FItemList(i).Forderno        = rsget("orderno")
                FItemList(i).Fregdate        = rsget("regdate")
                FItemList(i).Flastdate       = rsget("lastdate")
                FItemList(i).Fregadminid     = rsget("regadminid")
                FItemList(i).Flastadminid    = rsget("lastadminid")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fcate_large        = rsget("cate_large")
                FItemList(i).Fcate_mid          = rsget("cate_mid")
                FItemList(i).Fcate_small        = rsget("cate_small")
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
                FItemList(i).fitemisusing           = rsget("itemisusing")
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
                FItemList(i).Fitemcouponyn      = rsget("itemcouponyn")
                FItemList(i).Fcurritemcouponidx = rsget("curritemcouponidx")
                FItemList(i).Fitemcoupontype    = rsget("itemcoupontype")
                FItemList(i).Fitemcouponvalue   = rsget("itemcouponvalue")
                FItemList(i).Fsmallimage        = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
                FItemList(i).Flistimage         = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage")
                FItemList(i).Flistimage120      = "http://webimage.10x10.co.kr/image/list120/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("listimage120")
				FItemList(i).fbasicimage 		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("basicimage")

                rsget.movenext
                i=i+1
            loop
        end if
        rsget.Close
    end function
	
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function
	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function
	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function	
End Class
%>