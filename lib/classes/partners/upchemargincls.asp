<%
Class CUpcheMarginItem
    public FMakerid
    public FBrandName
    public FPartnerUsingYn
    public FBrandUsingYn

    public FDefaultOnlineMwDiv
    public FDefaultOnlineMargin


    public FOnlineMCount
    public FOnlineWCount
    public FOnlineUCount

    public FOnlineMAvgMargin
    public FOnlineWAvgMargin
    public FOnlineUAvgMargin

    public Fgroupid
    public Fcompany_no
    public Fcompany_name

    public FdefaultFreeBeasongLimit
    public FdefaultDeliveryType
    public FdefaultDeliverPay

    public FS000comm_cd
    public FS000defaultmargin
    public FS000defaultsuplymargin

    public FS800comm_cd
    public FS800defaultmargin
    public FS800defaultsuplymargin

    public FS870comm_cd
    public FS870defaultmargin
    public FS870defaultsuplymargin

    public FS700comm_cd
    public FS700defaultmargin
    public FS700defaultsuplymargin

    public FT000comm_cd
    public FT000defaultmargin
    public FT000defaultsuplymargin

    public FY000comm_cd
    public FY000defaultmargin
    public FY000defaultsuplymargin

    function getOnlinedefaultDlvTypeName()
        dim buf
        if (FdefaultDeliveryType="9") then
            buf="업체<font color='red'>조건</font>"
        elseif (FdefaultDeliveryType="7") then
            buf="업체<font color='blue'>착불</font>"
        end if

        if (FdefaultFreeBeasongLimit<>0) then
            buf = buf & CHKIIF(buf="","","<br>")&ForMatNumber(FdefaultFreeBeasongLimit,0) & "미만"
        end if

        if (FdefaultDeliverPay<>0) then
            buf = buf & CHKIIF(buf="","(반품시) ","<br>")&ForMatNumber(FdefaultDeliverPay,0) & "원"
        end if

        getOnlinedefaultDlvTypeName = buf
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Class CUpcheMargin

	public FItemList()
	public FOneItem

	public FTotalCount
	public FResultCount

	public FCurrPage
	public FTotalPage
	public FPageSize
	public FScrollCount

	public FRectMwDiv
	public FRectMakerid
	public FRectbrandUsingYn
	public FRectCateCode

	public Sub GetUpcheOnlineMarginList()
	    dim sqlStr, i

	    sqlStr = " select count(c.userid) as cnt "
	    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr + " 	join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " 	on c.userid=p.id"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and p.userdiv='9999'" ''2013/11/27 추가
        sqlStr = sqlStr + " and c.userdiv='02'"   ''2013/11/27 추가
        if (FRectbrandUsingYn<>"") then
            sqlStr = sqlStr + " and c.isusing='" & FRectbrandUsingYn & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and c.userid='" & FRectMakerid & "'"
        end if

        if (FRectCateCode<>"") then
            sqlStr = sqlStr + " and c.catecode='" & FRectCateCode & "'"
        end if

        if (FRectMwDiv<>"") then
            sqlStr = sqlStr + " and c.maeipdiv='" & FRectMwDiv & "'"
        end if

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close


	    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " c.userid, c.socname_kor, c.maeipdiv, c.defaultMargine, c.isusing, p.isusing,"
        sqlStr = sqlStr + " IsNULL(T.cntM,0) as cntM, IsNULL(T.cntW,0) as cntW, IsNULL(T.cntU,0) as cntU,"
        sqlStr = sqlStr + " IsNULL(T.mgnM,0) as mgnSumM,IsNULL(T.mgnW,0) as mgnSumW,IsNULL(T.mgnU,0) as mgnSumU,"
        sqlStr = sqlStr + " c.isusing as BrandUsingYn, p.isusing as PartnerUsingYn"
        sqlStr = sqlStr + " ,g.groupid, g.company_no, g.company_name, c.defaultFreeBeasongLimit, c.defaultDeliveryType,c.defaultDeliverPay"
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr + " 	join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " 	on c.userid=p.id"
        sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner_group g"
        sqlStr = sqlStr + " 	on p.groupid=g.groupid"
        sqlStr = sqlStr + " 	left join ("
        sqlStr = sqlStr + " 		select makerid, sum(case when mwdiv='M' then 1 else 0 end)  as cntM"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='W' then 1 else 0 end)  as cntW"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='U' then 1 else 0 end)  as cntU"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='M' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnM"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='W' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnW"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='U' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnU"
        sqlStr = sqlStr + " 		from [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " 		where isusing='Y'"
        sqlStr = sqlStr + " 		and orgprice<>0"
        sqlStr = sqlStr + " 		group by makerid"
        sqlStr = sqlStr + " 	) T on c.userid=T.makerid"

        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and p.userdiv='9999'"
        sqlStr = sqlStr + " and c.userdiv='02'"
        if (FRectbrandUsingYn<>"") then
            sqlStr = sqlStr + " and c.isusing='" & FRectbrandUsingYn & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and c.userid='" & FRectMakerid & "'"
        end if

        if (FRectMwDiv<>"") then
            sqlStr = sqlStr + " and c.maeipdiv='" & FRectMwDiv & "'"
        end if

        if (FRectCateCode<>"") then
            sqlStr = sqlStr + " and c.catecode='" & FRectCateCode & "'"
        end if

        sqlStr = sqlStr + " order by c.userid"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.recordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
			    set FItemList(i) = new CUpcheMarginItem

			    FItemList(i).FMakerid             = rsget("userid")
                FItemList(i).FBrandName           = db2html(rsget("socname_kor"))
                FItemList(i).FPartnerUsingYn      = rsget("PartnerUsingYn")
                FItemList(i).FBrandUsingYn        = rsget("BrandUsingYn")

                FItemList(i).FDefaultOnlineMwDiv  = rsget("maeipdiv")
                FItemList(i).FDefaultOnlineMargin = rsget("defaultMargine")

                FItemList(i).FOnlineMCount        = rsget("cntM")
                FItemList(i).FOnlineWCount        = rsget("cntW")
                FItemList(i).FOnlineUCount        = rsget("cntU")

                if (FItemList(i).FOnlineMCount<>0) then
                    FItemList(i).FOnlineMAvgMargin    = CLng(rsget("mgnSumM")/FItemList(i).FOnlineMCount*10)/10
                end if

                if (FItemList(i).FOnlineWCount<>0) then
                    FItemList(i).FOnlineWAvgMargin    = CLng(rsget("mgnSumW")/FItemList(i).FOnlineWCount*10)/10
                end if

                if (FItemList(i).FOnlineUCount<>0) then
                    FItemList(i).FOnlineUAvgMargin    = CLng(rsget("mgnSumU")/FItemList(i).FOnlineUCount*10)/10
                end if

                FItemList(i).Fgroupid           = rsget("groupid")
                FItemList(i).Fcompany_no        = rsget("company_no")
                FItemList(i).Fcompany_name      = rsget("company_name")

                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")

			    i=i+1
				rsget.movenext
			loop
		end if

        rsget.Close
    end Sub


	public Sub GetUpcheTotalMarginList()
	    dim sqlStr, i

	    sqlStr = " select count(c.userid) as cnt "
	    sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr + " 	 join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " 	on c.userid=p.id"
        sqlStr = sqlStr + " where 1=1" ''p.isusing='Y'
        sqlStr = sqlStr + " and p.userdiv='9999'" ''2013/11/27 추가
        sqlStr = sqlStr + " and c.userdiv='02'"   ''2013/11/27 추가

        if (FRectbrandUsingYn<>"") then
            sqlStr = sqlStr + " and c.isusing='" & FRectbrandUsingYn & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and c.userid='" & FRectMakerid & "'"
        end if

        if (FRectCateCode<>"") then
            sqlStr = sqlStr + " and c.catecode='" & FRectCateCode & "'"
        end if

        if (FRectMwDiv<>"") then
            sqlStr = sqlStr + " and c.maeipdiv='" & FRectMwDiv & "'"
        end if

        rsget.Open sqlStr,dbget,1
			FTotalCount = rsget("cnt")
		rsget.Close

        ''S000,S800,S870,S700,T000,Y000
        ''streetshop000,streetshop800,streetshop870,streetshop700,ithinksoop000,ygentshop1000
        ''직영,가맹,도매,해외,아이띵소,대행
	    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " c.userid, c.socname_kor, c.maeipdiv, c.defaultMargine, c.isusing, p.isusing,"
        sqlStr = sqlStr + " IsNULL(T.cntM,0) as cntM, IsNULL(T.cntW,0) as cntW, IsNULL(T.cntU,0) as cntU,"
        sqlStr = sqlStr + " IsNULL(T.mgnM,0) as mgnSumM,IsNULL(T.mgnW,0) as mgnSumW,IsNULL(T.mgnU,0) as mgnSumU,"
        sqlStr = sqlStr + " c.isusing as BrandUsingYn, p.isusing as PartnerUsingYn"
        sqlStr = sqlStr + " ,S000.comm_cd as S000comm_cd, S000.defaultmargin as S000defaultmargin, S000.defaultsuplymargin as S000defaultsuplymargin"
        sqlStr = sqlStr + " ,S800.comm_cd as S800comm_cd, S800.defaultmargin as S800defaultmargin, S800.defaultsuplymargin as S800defaultsuplymargin"
        sqlStr = sqlStr + " ,S870.comm_cd as S870comm_cd, S870.defaultmargin as S870defaultmargin, S870.defaultsuplymargin as S870defaultsuplymargin"
        sqlStr = sqlStr + " ,S700.comm_cd as S700comm_cd, S700.defaultmargin as S700defaultmargin, S700.defaultsuplymargin as S700defaultsuplymargin"
        sqlStr = sqlStr + " ,T000.comm_cd as T000comm_cd, T000.defaultmargin as T000defaultmargin, T000.defaultsuplymargin as T000defaultsuplymargin"
        sqlStr = sqlStr + " ,Y000.comm_cd as Y000comm_cd, Y000.defaultmargin as Y000defaultmargin, Y000.defaultsuplymargin as Y000defaultsuplymargin"
        sqlStr = sqlStr + " ,g.groupid, g.company_no, g.company_name, c.defaultFreeBeasongLimit, c.defaultDeliveryType,c.defaultDeliverPay"
        sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
        sqlStr = sqlStr + " 	join [db_partner].[dbo].tbl_partner p"
        sqlStr = sqlStr + " 	on c.userid=p.id"
        sqlStr = sqlStr + " 	left join [db_partner].[dbo].tbl_partner_group g"
        sqlStr = sqlStr + " 	on p.groupid=g.groupid"
        sqlStr = sqlStr + " 	left join ("
        sqlStr = sqlStr + " 		select makerid, sum(case when mwdiv='M' then 1 else 0 end)  as cntM"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='W' then 1 else 0 end)  as cntW"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='U' then 1 else 0 end)  as cntU"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='M' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnM"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='W' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnW"
        sqlStr = sqlStr + " 				, sum(case when mwdiv='U' then (100-orgsuplycash/orgprice*100) else 0 end)  as mgnU"
        sqlStr = sqlStr + " 		from [db_item].[dbo].tbl_item "
        sqlStr = sqlStr + " 		where isusing='Y'"
        sqlStr = sqlStr + " 		and orgprice<>0"
        sqlStr = sqlStr + " 		group by makerid"
        sqlStr = sqlStr + " 	) T on c.userid=T.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer S000"
        sqlStr = sqlStr + " 		on S000.shopid='streetshop000'"
        sqlStr = sqlStr + " 		and c.userid=S000.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer S800"
        sqlStr = sqlStr + " 		on S800.shopid='streetshop800'"
        sqlStr = sqlStr + " 		and c.userid=S800.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer S870"
        sqlStr = sqlStr + " 		on S870.shopid='streetshop870'"
        sqlStr = sqlStr + " 		and c.userid=S870.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer S700"
        sqlStr = sqlStr + " 		on S700.shopid='streetshop700'"
        sqlStr = sqlStr + " 		and c.userid=S700.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer T000"
        sqlStr = sqlStr + " 		on T000.shopid='ithinksoop000'"
        sqlStr = sqlStr + " 		and c.userid=T000.makerid"
        sqlStr = sqlStr + " 	left join [db_shop].[dbo].tbl_shop_designer Y000"
        sqlStr = sqlStr + " 		on Y000.shopid='ygentshop1000'"
        sqlStr = sqlStr + " 		and c.userid=Y000.makerid"



        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and p.userdiv='9999'" ''2013/11/27 추가
        sqlStr = sqlStr + " and c.userdiv='02'"   ''2013/11/27 추가
        if (FRectbrandUsingYn<>"") then
            sqlStr = sqlStr + " and c.isusing='" & FRectbrandUsingYn & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and c.userid='" & FRectMakerid & "'"
        end if

        if (FRectCateCode<>"") then
            sqlStr = sqlStr + " and c.catecode='" & FRectCateCode & "'"
        end if

        if (FRectMwDiv<>"") then
            sqlStr = sqlStr + " and c.maeipdiv='" & FRectMwDiv & "'"
        end if
        sqlStr = sqlStr + " order by c.userid"

        rsget.pagesize = FPageSize
        rsget.Open sqlStr,dbget,1
        FResultCount = rsget.recordCount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

        redim preserve FItemList(FResultCount)
		i=0

		if Not rsget.Eof then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
			    set FItemList(i) = new CUpcheMarginItem

			    FItemList(i).FMakerid             = rsget("userid")
                FItemList(i).FBrandName           = db2html(rsget("socname_kor"))
                FItemList(i).FPartnerUsingYn      = rsget("PartnerUsingYn")
                FItemList(i).FBrandUsingYn        = rsget("BrandUsingYn")

                FItemList(i).FDefaultOnlineMwDiv  = rsget("maeipdiv")
                FItemList(i).FDefaultOnlineMargin = rsget("defaultMargine")

                FItemList(i).FOnlineMCount        = rsget("cntM")
                FItemList(i).FOnlineWCount        = rsget("cntW")
                FItemList(i).FOnlineUCount        = rsget("cntU")

                if (FItemList(i).FOnlineMCount<>0) then
                    FItemList(i).FOnlineMAvgMargin    = CLng(rsget("mgnSumM")/FItemList(i).FOnlineMCount*10)/10
                end if

                if (FItemList(i).FOnlineWCount<>0) then
                    FItemList(i).FOnlineWAvgMargin    = CLng(rsget("mgnSumW")/FItemList(i).FOnlineWCount*10)/10
                end if

                if (FItemList(i).FOnlineUCount<>0) then
                    FItemList(i).FOnlineUAvgMargin    = CLng(rsget("mgnSumU")/FItemList(i).FOnlineUCount*10)/10
                end if

                FItemList(i).Fgroupid           = rsget("groupid")
                FItemList(i).Fcompany_no        = rsget("company_no")
                FItemList(i).Fcompany_name      = rsget("company_name")

                FItemList(i).FdefaultFreeBeasongLimit   = rsget("defaultFreeBeasongLimit")
                FItemList(i).FdefaultDeliveryType       = rsget("defaultDeliveryType")
                FItemList(i).FdefaultDeliverPay         = rsget("defaultDeliverPay")

                ''S000,S800,S870,S700,T000,Y000
                FItemList(i).FS000comm_cd              = rsget("S000comm_cd")
                FItemList(i).FS000defaultmargin        = rsget("S000defaultmargin")
                FItemList(i).FS000defaultsuplymargin   = rsget("S000defaultsuplymargin")

                FItemList(i).FS800comm_cd              = rsget("S800comm_cd")
                FItemList(i).FS800defaultmargin        = rsget("S800defaultmargin")
                FItemList(i).FS800defaultsuplymargin   = rsget("S800defaultsuplymargin")

                FItemList(i).FS870comm_cd              = rsget("S870comm_cd")
                FItemList(i).FS870defaultmargin        = rsget("S870defaultmargin")
                FItemList(i).FS870defaultsuplymargin   = rsget("S870defaultsuplymargin")

                FItemList(i).FS700comm_cd              = rsget("S700comm_cd")
                FItemList(i).FS700defaultmargin        = rsget("S700defaultmargin")
                FItemList(i).FS700defaultsuplymargin   = rsget("S700defaultsuplymargin")

                FItemList(i).FT000comm_cd              = rsget("T000comm_cd")
                FItemList(i).FT000defaultmargin        = rsget("T000defaultmargin")
                FItemList(i).FT000defaultsuplymargin   = rsget("T000defaultsuplymargin")

                FItemList(i).FY000comm_cd              = rsget("Y000comm_cd")
                FItemList(i).FY000defaultmargin        = rsget("Y000defaultmargin")
                FItemList(i).FY000defaultsuplymargin   = rsget("Y000defaultsuplymargin")


			    i=i+1
				rsget.movenext
			loop
		end if

        rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 12
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
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