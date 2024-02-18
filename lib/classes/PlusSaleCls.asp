<%

'// select top 100 *
'// from
'// db_sitemaster.dbo.tbl_buy_benefit
'// -- buy_benefit_idx, benefit_type, benefit_title, benefit_subtitle, benefit_start_dt, benefit_end_dt, whole_target_yn, use_yn, channel_www_yn, channel_mob_yn, channel_app_yn, mob_info_contents, www_info_contents, show_rank, reg_dt, reg_admin_id, last_update_dt, last_update_admin_id
'// -- info_contents_mobile, info_contents_www

'// select top 100 *
'// from
'// db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group
'// -- benefit_group_no, buy_benefit_idx, group_type, group_name, sort_no, use_yn, condition_amount, delivery_type, catecode, makerid, evtcode, evt_buy_condition

'// select top 100 *
'// from
'// db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item
'// -- plus_sale_item_idx, benefit_group_no, itemid, plus_sale_price, plus_sale_pct, plus_sale_buyprice, sale_burden_type, limit_yn, limit_cnt, max_buy_cnt, badge_contents, notice, sort_no, sell_cnt, use_yn, opt_cnt

'// BuyBenefitItem
'// BuyBenefitPlusSaleGroupItem
'// BuyBenefitPlusSaleGroupItemItem

class CBuyBenefitItem
    public Fbuy_benefit_idx
    public Fbenefit_type
    public Fbenefit_type_name
    public Fbenefit_title
    public Fbenefit_subtitle
    public Fbenefit_start_dt
    public Fbenefit_end_dt
    public Fwhole_target_yn
    public Fuse_yn
    public Fchannel_www_yn
    public Fchannel_mob_yn
    public Fchannel_app_yn
    public Fmob_info_contents
    public Fwww_info_contents
    public Finfo_contents_mobile
    public Finfo_contents_www
    public Fshow_rank
    public Freg_dt
    public Freg_admin_id
    public Flast_update_dt
    public Flast_update_admin_id

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CBuyBenefitStatItem
    public Fbuy_benefit_idx
    public Fbenefit_title
    public Fbenefit_group_no
    public Fgroup_name
    public FtargetOrderCount
    public ForderCnt
    public FItemCnt

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CBuyBenefitPlusSaleGroupItem
    public Fbenefit_group_no
    public Fbuy_benefit_idx
    public Fgroup_type
    public Fgroup_type_name
    public Fgroup_name
    public Fsort_no
    public Fuse_yn
    public Fcondition_amount
    public Fdelivery_type
    public Fdelivery_type_name
    public Fcatecode
    public Fmakerid
    public Fevtcode
    public Fevt_buy_condition
    public Fevt_buy_condition_name

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CBuyBenefitPlusSaleGroupItemItem
    public Fplus_sale_item_idx
    public Fbenefit_group_no
    public Fitemid
    public Fplus_sale_price
    public Fplus_sale_pct
    public Fplus_sale_buyprice
    public Fsale_burden_type
    public Fsale_burden_type_name
    public Flimit_yn
    public Flimit_cnt
    public Fmax_buy_cnt
    public Fbadge_contents
    public Fnotice
    public Fsort_no
    public Fsell_cnt
    public Fuse_yn
    public Fopt_cnt
    public Fsellcash
    public Fitemname
    public Fmwdiv

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

class CBuyBenefitPlusSaleSoldItem
    public Fitemid
    public Fitemoption
    public Fitemname
    public Fitemoptionname
    public Fitemno
    public Fmeachul

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end class

Class CBuyBenefit
	public FItemList()
	public FOneItem
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount

    public FRectBuyBenefitIdx
    public FRectBenefitGroupNo
    public FRectPlusSaleItemIdx
    public FRectUseYN
    public FRectIdx
    public FRectKeyword
    public FRectViewDate
    public FRectExistYN

	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	public Sub GetBuyBenefitList()
        dim i, sqlStr, addSql

        addSql = ""
        if (FRectUseYN <> "") then
            addSql = addSql & " and m.use_yn = '" & FRectUseYN & "' "
        end if

        if (FRectIdx <> "") then
            addSql = addSql & " and m.buy_benefit_idx='" & FRectIdx& "' "
        end if

        if (FRectKeyword <> "") then
            addSql = addSql & " and m.benefit_title like '%" & FRectKeyword & "%' "
        end if

        if (FRectViewDate <> "") then
            addSql = addSql & " and '" & FRectViewDate & "' between m.benefit_start_dt and m.benefit_end_dt "
        end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit m "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = " select top " & FPageSize*FCurrPage & " m.buy_benefit_idx, m.benefit_type, m.benefit_title, m.benefit_subtitle, m.benefit_start_dt, m.benefit_end_dt, m.whole_target_yn, m.use_yn, m.channel_www_yn, m.channel_mob_yn, m.channel_app_yn, m.mob_info_contents, m.www_info_contents, m.show_rank, m.reg_dt, m.reg_admin_id, IsNull(last_update_dt, m.reg_dt) as last_update_dt, IsNull(m.last_update_admin_id, m.reg_admin_id) as last_update_admin_id, c1.pcomm_name as benefit_type_name, m.info_contents_mobile, m.info_contents_www "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit m "
        sqlStr = sqlStr & " 	left join [db_partner].[dbo].tbl_partner_comm_code c1 "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and c1.pcomm_isusing = 'Y' "
        sqlStr = sqlStr & " 		and c1.pcomm_group = 'PSBenefitType' "
        sqlStr = sqlStr & " 		and m.benefit_type = c1.pcomm_cd "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by m.buy_benefit_idx desc "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBuyBenefitItem

                FItemList(i).Fbuy_benefit_idx           = rsget("buy_benefit_idx")
                FItemList(i).Fbenefit_type            	= rsget("benefit_type")
                FItemList(i).Fbenefit_type_name         = rsget("benefit_type_name")
                FItemList(i).Fbenefit_title            	= rsget("benefit_title")
                FItemList(i).Fbenefit_subtitle          = rsget("benefit_subtitle")
                FItemList(i).Fbenefit_start_dt          = rsget("benefit_start_dt")
                FItemList(i).Fbenefit_end_dt            = rsget("benefit_end_dt")
                FItemList(i).Fwhole_target_yn           = rsget("whole_target_yn")
                FItemList(i).Fuse_yn            		= rsget("use_yn")
                FItemList(i).Fchannel_www_yn            = rsget("channel_www_yn")
                FItemList(i).Fchannel_mob_yn            = rsget("channel_mob_yn")
                FItemList(i).Fchannel_app_yn            = rsget("channel_app_yn")
                FItemList(i).Fmob_info_contents         = rsget("mob_info_contents")
                FItemList(i).Fwww_info_contents         = rsget("www_info_contents")
                FItemList(i).Fshow_rank            		= rsget("show_rank")
                FItemList(i).Freg_dt            		= rsget("reg_dt")
                FItemList(i).Freg_admin_id            	= rsget("reg_admin_id")
                FItemList(i).Flast_update_dt            = rsget("last_update_dt")
                FItemList(i).Flast_update_admin_id      = rsget("last_update_admin_id")

                FItemList(i).Finfo_contents_mobile      = db2html(rsget("info_contents_mobile"))
                FItemList(i).Finfo_contents_www      	= db2html(rsget("info_contents_www"))

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
    end sub

    public sub GetCBuyBenefitMasterOne
		dim sqlStr, addSql

        if (FRectBuyBenefitIdx <> "") then
            addSql = " and m.buy_benefit_idx = " & FRectBuyBenefitIdx
        else
            addSql = " and 1 <> 1 "
        end if

        sqlStr = " select top 1 m.buy_benefit_idx, m.benefit_type, m.benefit_title, m.benefit_subtitle, m.benefit_start_dt, m.benefit_end_dt, m.whole_target_yn, m.use_yn, m.channel_www_yn, m.channel_mob_yn, m.channel_app_yn, m.mob_info_contents, m.www_info_contents, m.show_rank, m.reg_dt, m.reg_admin_id, IsNull(last_update_dt, m.reg_dt) as last_update_dt, IsNull(m.last_update_admin_id, m.reg_admin_id) as last_update_admin_id, c1.pcomm_name as benefit_type_name, m.info_contents_mobile, m.info_contents_www "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit m "
        sqlStr = sqlStr & " 	left join [db_partner].[dbo].tbl_partner_comm_code c1 "
        sqlStr = sqlStr & " 	on "
        sqlStr = sqlStr & " 		1 = 1 "
        sqlStr = sqlStr & " 		and c1.pcomm_isusing = 'Y' "
        sqlStr = sqlStr & " 		and c1.pcomm_group = 'PSBenefitType' "
        sqlStr = sqlStr & " 		and m.benefit_type = c1.pcomm_cd "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

        set FOneItem = new CBuyBenefitItem

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then

			FOneItem.Fbuy_benefit_idx           = rsget("buy_benefit_idx")
			FOneItem.Fbenefit_type            	= rsget("benefit_type")
			FOneItem.Fbenefit_type_name         = rsget("benefit_type_name")
			FOneItem.Fbenefit_title            	= rsget("benefit_title")
			FOneItem.Fbenefit_subtitle          = rsget("benefit_subtitle")
			FOneItem.Fbenefit_start_dt          = rsget("benefit_start_dt")
			FOneItem.Fbenefit_end_dt            = rsget("benefit_end_dt")
			FOneItem.Fwhole_target_yn           = rsget("whole_target_yn")
			FOneItem.Fuse_yn            		= rsget("use_yn")
			FOneItem.Fchannel_www_yn            = rsget("channel_www_yn")
			FOneItem.Fchannel_mob_yn            = rsget("channel_mob_yn")
			FOneItem.Fchannel_app_yn            = rsget("channel_app_yn")
			FOneItem.Fmob_info_contents         = rsget("mob_info_contents")
			FOneItem.Fwww_info_contents         = rsget("www_info_contents")
			FOneItem.Fshow_rank            		= rsget("show_rank")
			FOneItem.Freg_dt            		= rsget("reg_dt")
			FOneItem.Freg_admin_id            	= rsget("reg_admin_id")
			FOneItem.Flast_update_dt            = rsget("last_update_dt")
			FOneItem.Flast_update_admin_id      = rsget("last_update_admin_id")

			FOneItem.Finfo_contents_mobile      = db2html(rsget("info_contents_mobile"))
			FOneItem.Finfo_contents_www      	= db2html(rsget("info_contents_www"))
		end if
		rsget.Close

    end sub

    Public Sub GetBuyBenefitSoldItemList()
        dim i, sqlStr, addSql

        sqlStr = " exec [db_datamart].[dbo].[usp_TEN_Buy_Benifit_Stat_ItemList] " & FRectBuyBenefitIdx
        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly

		FResultCount = db3_rsget.RecordCount
        FTotalCount = db3_rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if

		redim preserve FItemList(FResultCount)

		if  not db3_rsget.EOF  then
			i = 0
			do until db3_rsget.eof
				set FItemList(i) = new CBuyBenefitPlusSaleSoldItem

                FItemList(i).Fitemid  = db3_rsget("itemid")
                FItemList(i).Fitemoption  = db3_rsget("itemoption")
                FItemList(i).Fitemname  = db3_rsget("itemname")
                FItemList(i).Fitemoptionname  = db3_rsget("itemoptionname")
                FItemList(i).Fitemno  = db3_rsget("itemno")
                FItemList(i).Fmeachul  = db3_rsget("¸ÅÃâ")

				db3_rsget.MoveNext
				i = i + 1
			loop
		end if
		db3_rsget.close
    end sub

    Public Sub GetBuyBenefitGroupList()
        dim i, sqlStr, addSql

        addSql = ""
        if (FRectBuyBenefitIdx <> "") then
            addSql = " and g.buy_benefit_idx = " & FRectBuyBenefitIdx
        else
            addSql = " and 1 <> 1 "
        end if

        if (FRectUseYN <> "") then
            addSql = addSql & " and g.use_yn = '" & FRectUseYN & "' "
        end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group g "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = " select top " & FPageSize*FCurrPage & " g.benefit_group_no, g.buy_benefit_idx, g.group_type, g.group_name, g.sort_no, g.use_yn, g.condition_amount, g.delivery_type, g.catecode, g.makerid, g.evtcode, g.evt_buy_condition, c1.pcomm_name as group_type_name, c2.pcomm_name as delivery_type_name, c3.pcomm_name as evt_buy_condition_name "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group g "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c1 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c1.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c1.pcomm_group = 'PSGroupType' "
		sqlStr = sqlStr & "			and g.group_type = c1.pcomm_cd "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c2 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c2.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c2.pcomm_group = 'PSDeliveryType' "
		sqlStr = sqlStr & "			and g.delivery_type = c2.pcomm_cd "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c3 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c3.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c3.pcomm_group = 'PSBuyCondition' "
		sqlStr = sqlStr & "			and g.evt_buy_condition = c3.pcomm_cd "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by g.sort_no "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBuyBenefitPlusSaleGroupItem

                FItemList(i).Fbenefit_group_no  = rsget("benefit_group_no")
                FItemList(i).Fbuy_benefit_idx   = rsget("buy_benefit_idx")
                FItemList(i).Fgroup_type        = rsget("group_type")
                FItemList(i).Fgroup_type_name   = rsget("group_type_name")
                FItemList(i).Fgroup_name        = rsget("group_name")
                FItemList(i).Fsort_no           = rsget("sort_no")
                FItemList(i).Fuse_yn           	= rsget("use_yn")
                FItemList(i).Fcondition_amount  = rsget("condition_amount")
                FItemList(i).Fdelivery_type     = rsget("delivery_type")
                FItemList(i).Fdelivery_type_name     = rsget("delivery_type_name")
                FItemList(i).Fcatecode          = rsget("catecode")
                FItemList(i).Fmakerid           = rsget("makerid")
                FItemList(i).Fevtcode           = rsget("evtcode")
                FItemList(i).Fevt_buy_condition = rsget("evt_buy_condition")
                FItemList(i).Fevt_buy_condition_name = rsget("evt_buy_condition_name")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
    End Sub

    Public Sub GetBuyBenefitGroupItemList()
        dim i, sqlStr, addSql

        addSql = ""
        if (FRectBenefitGroupNo <> "") then
            addSql = " and gi.benefit_group_no = " & FRectBenefitGroupNo
        else
            addSql = " and 1 <> 1 "
        end if

        if (FRectUseYN <> "") then
            addSql = addSql & " and gi.use_yn = '" & FRectUseYN & "' "
        end if

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item gi "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

		'response.write sqlStr &"<br>"
		rsget.Open sqlStr, dbget, 1
			FTotalCount = rsget("cnt")
		rsget.close


        sqlStr = " select top " & FPageSize*FCurrPage & " gi.plus_sale_item_idx, gi.benefit_group_no, gi.itemid, gi.plus_sale_price, gi.plus_sale_pct, gi.plus_sale_buyprice, gi.sale_burden_type, gi.limit_yn, gi.limit_cnt, gi.max_buy_cnt, gi.badge_contents, gi.notice, gi.sort_no, gi.sell_cnt, gi.use_yn, gi.opt_cnt, c1.pcomm_name as sale_burden_type_name, i.orgprice as sellcash "
        sqlStr = sqlStr & " , i.itemname, i.mwdiv "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item gi "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c1 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c1.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c1.pcomm_group = 'PSBurdenType' "
		sqlStr = sqlStr & "			and gi.sale_burden_type = c1.pcomm_cd "
		sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item] i on gi.itemid = i.itemid "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql
        sqlStr = sqlStr & " order by gi.sort_no "

		''response.write sqlStr &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open sqlStr, dbget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
        if (FtotalPage < 1) then
            FtotalPage = 1
        end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)

		if  not rsget.EOF  then
			i = 0
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CBuyBenefitPlusSaleGroupItemItem

                FItemList(i).Fplus_sale_item_idx  	= rsget("plus_sale_item_idx")
                FItemList(i).Fbenefit_group_no  	= rsget("benefit_group_no")
                FItemList(i).Fitemid  				= rsget("itemid")
                FItemList(i).Fplus_sale_price  		= rsget("plus_sale_price")
                FItemList(i).Fplus_sale_pct  		= rsget("plus_sale_pct")
                FItemList(i).Fplus_sale_buyprice  	= rsget("plus_sale_buyprice")
                FItemList(i).Fsale_burden_type  	= rsget("sale_burden_type")
                FItemList(i).Fsale_burden_type_name = rsget("sale_burden_type_name")
                FItemList(i).Flimit_yn  			= rsget("limit_yn")
                FItemList(i).Flimit_cnt  			= rsget("limit_cnt")
                FItemList(i).Fmax_buy_cnt  			= rsget("max_buy_cnt")
                FItemList(i).Fbadge_contents  		= rsget("badge_contents")
                FItemList(i).Fnotice  				= rsget("notice")
                FItemList(i).Fsort_no  				= rsget("sort_no")
                FItemList(i).Fsell_cnt  			= rsget("sell_cnt")
                FItemList(i).Fuse_yn  				= rsget("use_yn")
			    FItemList(i).Fopt_cnt				= rsget("opt_cnt")
                FItemList(i).Fsellcash				= rsget("sellcash")

                FItemList(i).Fitemname				= db2html(rsget("itemname"))
                FItemList(i).Fmwdiv					= rsget("mwdiv")

				rsget.MoveNext
				i = i + 1
			loop
		end if
		rsget.close
    End Sub

    public sub GetCBuyBenefitGroupOne
		dim sqlStr, addSql

        if (FRectBenefitGroupNo <> "") then
            addSql = " and g.benefit_group_no = " & FRectBenefitGroupNo
        else
            addSql = " and 1 <> 1 "
        end if

        sqlStr = " select top 1 g.benefit_group_no, g.buy_benefit_idx, g.group_type, g.group_name, g.sort_no, g.use_yn, g.condition_amount, g.delivery_type, g.catecode, g.makerid, g.evtcode, g.evt_buy_condition, c1.pcomm_name as group_type_name, c2.pcomm_name as delivery_type_name, c3.pcomm_name as evt_buy_condition_name "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group g "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c1 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c1.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c1.pcomm_group = 'PSGroupType' "
		sqlStr = sqlStr & "			and g.group_type = c1.pcomm_cd "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c2 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c2.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c2.pcomm_group = 'PSDeliveryType' "
		sqlStr = sqlStr & "			and g.delivery_type = c2.pcomm_cd "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c3 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c3.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c3.pcomm_group = 'PSBuyCondition' "
		sqlStr = sqlStr & "			and g.evt_buy_condition = c3.pcomm_cd "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

        set FOneItem = new CBuyBenefitPlusSaleGroupItem

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then
			FOneItem.Fbenefit_group_no  = rsget("benefit_group_no")
			FOneItem.Fbuy_benefit_idx   = rsget("buy_benefit_idx")
			FOneItem.Fgroup_type        = rsget("group_type")
			FOneItem.Fgroup_type_name   = rsget("group_type_name")
			FOneItem.Fgroup_name        = rsget("group_name")
			FOneItem.Fsort_no           = rsget("sort_no")
			FOneItem.Fuse_yn           	= rsget("use_yn")
			FOneItem.Fcondition_amount  = rsget("condition_amount")
			FOneItem.Fdelivery_type     = rsget("delivery_type")
			FOneItem.Fdelivery_type_name     = rsget("delivery_type_name")
			FOneItem.Fcatecode          = rsget("catecode")
			FOneItem.Fmakerid           = rsget("makerid")
			FOneItem.Fevtcode           = rsget("evtcode")
			FOneItem.Fevt_buy_condition = rsget("evt_buy_condition")
			FOneItem.Fevt_buy_condition_name = rsget("evt_buy_condition_name")
        end if
		rsget.Close
    end sub

    Public Sub GetBuyBenefitGroupItemOne()
		dim sqlStr, addSql

        if (FRectPlusSaleItemIdx <> "") then
            addSql = " and gi.plus_sale_item_idx = " & FRectPlusSaleItemIdx
        else
            addSql = " and 1 <> 1 "
        end if

        sqlStr = " select top 1 gi.plus_sale_item_idx, gi.benefit_group_no, gi.itemid, gi.plus_sale_price, gi.plus_sale_pct, gi.plus_sale_buyprice, gi.sale_burden_type, gi.limit_yn, gi.limit_cnt, gi.max_buy_cnt, gi.badge_contents, gi.notice, gi.sort_no, gi.sell_cnt, gi.use_yn, gi.opt_cnt, c1.pcomm_name as sale_burden_type_name, i.orgprice as sellcash "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " db_sitemaster.dbo.tbl_buy_benefit_plus_sale_group_item gi "
        sqlStr = sqlStr & "		left join [db_partner].[dbo].tbl_partner_comm_code c1 "
		sqlStr = sqlStr & "		on "
		sqlStr = sqlStr & "			1 = 1 "
		sqlStr = sqlStr & "			and c1.pcomm_isusing = 'Y' "
		sqlStr = sqlStr & "			and c1.pcomm_group = 'PSBurdenType' "
		sqlStr = sqlStr & "			and gi.sale_burden_type = c1.pcomm_cd "
        sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item] i on gi.itemid = i.itemid "
        sqlStr = sqlStr & " where 1 = 1 "
        sqlStr = sqlStr & addSql

        set FOneItem = new CBuyBenefitPlusSaleGroupItemItem

		'response.write sqlStr & "<br>"
		rsget.Open sqlStr, dbget, 1
		FResultCount = rsget.RecordCount
		FtotalCount = rsget.RecordCount
		if Not rsget.Eof then
			FOneItem.Fplus_sale_item_idx  	= rsget("plus_sale_item_idx")
			FOneItem.Fbenefit_group_no  	= rsget("benefit_group_no")
			FOneItem.Fitemid  				= rsget("itemid")
			FOneItem.Fplus_sale_price  		= rsget("plus_sale_price")
			FOneItem.Fplus_sale_pct  		= rsget("plus_sale_pct")
			FOneItem.Fplus_sale_buyprice  	= rsget("plus_sale_buyprice")
			FOneItem.Fsale_burden_type  	= rsget("sale_burden_type")
			FOneItem.Fsale_burden_type_name = rsget("sale_burden_type_name")
			FOneItem.Flimit_yn  			= rsget("limit_yn")
			FOneItem.Flimit_cnt  			= rsget("limit_cnt")
			FOneItem.Fmax_buy_cnt  			= rsget("max_buy_cnt")
			FOneItem.Fbadge_contents  		= rsget("badge_contents")
			FOneItem.Fnotice  				= rsget("notice")
			FOneItem.Fsort_no  				= rsget("sort_no")
			FOneItem.Fsell_cnt  			= rsget("sell_cnt")
			FOneItem.Fuse_yn  				= rsget("use_yn")
			FOneItem.Fopt_cnt				= rsget("opt_cnt")
			FOneItem.Fsellcash				= rsget("sellcash")
        end if
		rsget.Close
    End Sub

    Public Sub GetBuyBenefitStat()
		dim sqlStr, addSql, i

        addSql = ""
        if (FRectUseYN <> "") then
            addSql = addSql & " and b.use_yn = '" & FRectUseYN & "' "
        end if

        if (FRectIdx <> "") then
            addSql = addSql & " and b.buy_benefit_idx='" & FRectIdx& "' "
        end if

        if (FRectKeyword <> "") then
            addSql = addSql & " and b.benefit_title like '%" & FRectKeyword & "%' "
        end if

        if (FRectExistYN <> "") then
            addSql = addSql & " and IsNull(s.target_order_count, 0)>0 "
        end if

        sqlStr = " select count(distinct cast(b.buy_benefit_idx as varchar(10)) +'/'+ cast(isNull(g.benefit_group_no,0) as varchar(10))) AS cnt, CEILING(CAST(Count(distinct cast(b.buy_benefit_idx as varchar(10)) +'/'+ cast(isNull(g.benefit_group_no,0) as varchar(10))) AS FLOAT)/" & FPageSize & ") AS totPg  "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_sitemaster].[dbo].[tbl_buy_benefit] b with(nolock) "
        sqlStr = sqlStr & " 	left join [db_sitemaster].[dbo].[tbl_buy_benefit_plus_sale_group] g with(nolock) on b.buy_benefit_idx = g.buy_benefit_idx "
        sqlStr = sqlStr & " 	left join [db_sitemaster].[dbo].[tbl_buy_benefit_plus_sale_group_item] i with(nolock) on g.benefit_group_no = i.benefit_group_no "
        sqlStr = sqlStr & " 	        and g.use_yn = 'Y' "
        sqlStr = sqlStr & " 	left join [db_datamart].[dbo].[tbl_buy_benefit_plus_sale_group_stat] s on g.benefit_group_no = s.benefit_group_no "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 " & addSql
        sqlStr = sqlStr & " 	and b.benefit_type = 'P' "
		db3_rsget.Open sqlStr, db3_dbget, 1
            FtotalCount = db3_rsget("cnt")
            FtotalPage = db3_rsget("totPg")
        db3_rsget.Close

        sqlStr = " select "
        sqlStr = sqlStr & " 	b.buy_benefit_idx, isNull(g.benefit_group_no,0) as benefit_group_no, b.benefit_title, isNull(g.group_name,'') as group_name, count(distinct d.orderserial) as orderCnt, IsNull(sum(d.itemno),0) as ItemCnt, IsNull(s.target_order_count, -1) as targetOrderCount "
        sqlStr = sqlStr & " from "
        sqlStr = sqlStr & " 	[db_sitemaster].[dbo].[tbl_buy_benefit] b with(nolock) "
        sqlStr = sqlStr & " 	left join [db_sitemaster].[dbo].[tbl_buy_benefit_plus_sale_group] g with(nolock) on b.buy_benefit_idx = g.buy_benefit_idx "
        sqlStr = sqlStr & " 	left join [db_sitemaster].[dbo].[tbl_buy_benefit_plus_sale_group_item] i with(nolock) on g.benefit_group_no = i.benefit_group_no "
        sqlStr = sqlStr & " 	        and g.use_yn = 'Y' "
        sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_detail] d with(nolock) on i.plus_sale_item_idx = d.plus_sale_item_idx "
        sqlStr = sqlStr & " 	        and d.cancelyn <> 'Y' "
        sqlStr = sqlStr & " 	join [db_order].[dbo].[tbl_order_master] m with(nolock) on m.orderserial = d.orderserial "
        sqlStr = sqlStr & " 	        and m.cancelyn = 'N' "
        sqlStr = sqlStr & " 	        and m.jumundiv not in (6, 9) "
        sqlStr = sqlStr & " 	left join [db_datamart].[dbo].[tbl_buy_benefit_plus_sale_group_stat] s on g.benefit_group_no = s.benefit_group_no "
        sqlStr = sqlStr & " where "
        sqlStr = sqlStr & " 	1 = 1 " & addSql
        sqlStr = sqlStr & " 	and b.benefit_type = 'P' "
        sqlStr = sqlStr & " group by "
        sqlStr = sqlStr & " 	b.buy_benefit_idx, isNull(g.benefit_group_no,0), b.benefit_title, isNull(g.group_name,''), IsNull(s.target_order_count, -1) "
        sqlStr = sqlStr & " order by "
        sqlStr = sqlStr & " 	b.buy_benefit_idx desc "
        sqlStr = sqlStr & " OFFSET " & (FCurrPage-1)*FPageSize & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY "

		'response.write sqlStr &"<br>"
        db3_rsget.Open sqlStr, db3_dbget, 1
		FResultCount = db3_rsget.RecordCount

		redim preserve FItemList(FResultCount)

		if  not db3_rsget.EOF  then
			i = 0
			do until db3_rsget.eof
				set FItemList(i) = new CBuyBenefitStatItem

                FItemList(i).Fbuy_benefit_idx  		= db3_rsget("buy_benefit_idx")
                FItemList(i).Fbenefit_group_no  	= db3_rsget("benefit_group_no")
                FItemList(i).Fbenefit_title  		= db3_rsget("benefit_title")
                FItemList(i).Fgroup_name  			= db3_rsget("group_name")
                FItemList(i).ForderCnt  			= db3_rsget("orderCnt")
                FItemList(i).FItemCnt  				= db3_rsget("ItemCnt")
                FItemList(i).FtargetOrderCount  	= db3_rsget("targetOrderCount")

				db3_rsget.MoveNext
				i = i + 1
			loop
		end if
		db3_rsget.close
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

end Class

%>
