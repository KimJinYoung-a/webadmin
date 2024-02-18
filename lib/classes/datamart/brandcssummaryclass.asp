<%
Class CBrandCSItem
    public FID
    public Fdivcd
    public Fgubun01
    public Fgubun02

    public Fdivcd_Name
    public Fgubun01_Name
    public Fgubun02_Name

    public Forderserial
    public Fitemid
    public Fitemoption
    public Fconfirmitemno
    public Fisupchebeasong

    public Fitemname
    public Fitemoptionname

    public FRegdate
    public Ffinishdate

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CBrandCSSummaryItem
    public Fyyyymm
    public Fmakerid
    public Fisupchebeasong
    public Fdivcd
    public Fgubun01
    public Fgubun02
    public Fdivname
    public Fgubun01name
    public Fgubun02name
    public Fcnt

    public FCNT_1
    public FCNT_2
    public FCNT_3
    public FCNT_4
    public FCNT_5
    public FCNT_6
    public FCNT_7

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CBrandCSSummary
    public FItemList()
	public FOneItem

    public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage

    public FRectYYYYMM
    public FRectStartDate
    public FRectEndDate
    public FRectMakerid
    public FRectDivCd
    public FRectCDL
    public FRectGubun02Arr
    public FRectIsupchebeasong

    public FRectNotIncludeETC  'CD99, CF99, CH99, CE99

    public Sub getBrandCsList()
        dim sqlStr, i

        sqlStr = " select count(*) as cnt "
        sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list a"
        sqlStr = sqlStr + " 	Join db_cs.dbo.tbl_new_as_detail d"
        sqlStr = sqlStr + " 		on a.id=d.masterid"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and d.itemid<>0"

        if (FRectStartDate<>"") then
            sqlStr = sqlStr + " and a.finishdate>='" + FRectStartDate + "'"
        end if

        if (FRectEndDate<>"") then
            sqlStr = sqlStr + " and a.finishdate<'" + FRectEndDate + "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr + " and d.isupchebeasong='" + FRectIsupchebeasong + "'"
        end if

        if (FRectDivCd<>"") then
            if (FRectDivCd="T012") then
                sqlStr = sqlStr + " and a.divcd in ('A000','A001','A002')"
            else
                sqlStr = sqlStr + " and a.divcd='" + FRectDivCd + "'"
            end if
        end if

        if (FRectGubun02Arr<>"") then
            sqlStr = sqlStr + " and a.gubun02 in (" + FRectGubun02Arr + ")"
        end if



        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget("cnt")
        rsget.Close

        '' too slow;;
        sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " a.id, c1.comm_name as divcd_Name,c3.comm_name as gubun02_Name,a.orderserial, a.finishdate,"
        sqlStr = sqlStr + " d.itemid,d.itemoption,d.makerid, d.confirmitemno,d.isupchebeasong, d.itemname, d.itemoptionname,"
        sqlStr = sqlStr + " d.confirmitemno, d.isupchebeasong"
        sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list a"
        sqlStr = sqlStr + " 	Join db_cs.dbo.tbl_new_as_detail d"
        sqlStr = sqlStr + " 		on a.id=d.masterid"
        sqlStr = sqlStr + " 	left join db_cs.dbo.tbl_cs_comm_code c1"
        sqlStr = sqlStr + " 		on a.divcd=c1.comm_cd"
        sqlStr = sqlStr + " 	left join db_cs.dbo.tbl_cs_comm_code c3"
        sqlStr = sqlStr + " 		on a.gubun02=c3.comm_cd"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and d.itemid<>0"

        if (FRectStartDate<>"") then
            sqlStr = sqlStr + " and a.finishdate>='" + FRectStartDate + "'"
        end if

        if (FRectEndDate<>"") then
            sqlStr = sqlStr + " and a.finishdate<'" + FRectEndDate + "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr + " and d.makerid='" + FRectMakerid + "'"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr + " and d.isupchebeasong='" + FRectIsupchebeasong + "'"
        end if

        if (FRectDivCd<>"") then
            if (FRectDivCd="T012") then
                sqlStr = sqlStr + " and a.divcd in ('A000','A001','A002')"
            else
                sqlStr = sqlStr + " and a.divcd='" + FRectDivCd + "'"
            end if
        end if

        if (FRectGubun02Arr<>"") then
            sqlStr = sqlStr + " and a.gubun02 in (" + FRectGubun02Arr + ")"
        end if
        sqlStr = sqlStr + " and deleteyn='N'"
		''속도향상
		''and a.id >= 1500000
        sqlStr = sqlStr + " order by a.id"
		''response.write sqlStr


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
			do until rsget.eof

			    set FItemList(i) = new CBrandCSItem
			    FItemList(i).FID            = rsget("id")
			    FItemList(i).Fdivcd_Name    = rsget("divcd_Name")
			    FItemList(i).Fgubun02_Name  = rsget("gubun02_Name")

			    FItemList(i).Forderserial  = rsget("orderserial")
			    FItemList(i).Fitemid        = rsget("itemid")
			    FItemList(i).Fitemoption    = rsget("itemoption")
			    FItemList(i).Fitemname      = rsget("itemname")
			    FItemList(i).Fitemoptionname = rsget("itemoptionname")

			    FItemList(i).Fconfirmitemno    = rsget("confirmitemno")
			    FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")
			    FItemList(i).Ffinishdate    = rsget("finishdate")
			    i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end Sub

    public Sub getBrandCsSUMList()
        dim sqlStr, i

		sqlStr = " select top 100 "
		sqlStr = sqlStr + " 	c1.comm_name as divcd_name "
		sqlStr = sqlStr + " 	,c3.comm_name as gubun02_name "
		sqlStr = sqlStr + " 	,d.itemid "
		sqlStr = sqlStr + " 	,d.makerid "
		sqlStr = sqlStr + " 	,sum(d.confirmitemno) as totconfirmitemno "
		sqlStr = sqlStr + " 	,d.isupchebeasong "
		sqlStr = sqlStr + " 	,d.itemname "
		sqlStr = sqlStr + " from db_cs.dbo.tbl_new_as_list a "
		sqlStr = sqlStr + " inner join db_cs.dbo.tbl_new_as_detail d on a.id = d.masterid "
		sqlStr = sqlStr + " left join db_cs.dbo.tbl_cs_comm_code c1 on a.divcd = c1.comm_cd "
		sqlStr = sqlStr + " left join db_cs.dbo.tbl_cs_comm_code c3 on a.gubun02 = c3.comm_cd "
		sqlStr = sqlStr + " where 1 = 1 "
		sqlStr = sqlStr + " 	and d.itemid <> 0 "
		sqlStr = sqlStr + " 	and a.finishdate >= '" + CStr(FRectStartDate) + "' "
		sqlStr = sqlStr + " 	and a.finishdate < '" + CStr(FRectEndDate) + "' "
		sqlStr = sqlStr + " 	and d.makerid = '" + CStr(FRectMakerid) + "' "

		if Not application("Svr_Info")="Dev" then
			'// 속도향상
			sqlStr = sqlStr + " 	and a.id >= 1500000 "
		end if

        if (FRectDivCd<>"") then
            if (FRectDivCd="T012") then
                sqlStr = sqlStr + " and a.divcd in ('A000','A001','A002')"
            else
                sqlStr = sqlStr + " and a.divcd='" + FRectDivCd + "'"
            end if
        end if

        if (FRectGubun02Arr<>"") then
            sqlStr = sqlStr + " and a.gubun02 in (" + FRectGubun02Arr + ")"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr + " and d.isupchebeasong='" + FRectIsupchebeasong + "'"
        end if

		sqlStr = sqlStr + " 	and deleteyn = 'N' "

		sqlStr = sqlStr + " group by "
		sqlStr = sqlStr + " 	c1.comm_name "
		sqlStr = sqlStr + " 	,c3.comm_name "
		sqlStr = sqlStr + " 	,d.itemid "
		sqlStr = sqlStr + " 	,d.makerid "
		sqlStr = sqlStr + " 	,d.isupchebeasong "
		sqlStr = sqlStr + " 	,d.itemname "
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	sum(d.confirmitemno) desc "
		sqlStr = sqlStr + " 	,c1.comm_name "
		sqlStr = sqlStr + " 	,c3.comm_name "
        rsget.Open sqlStr,dbget,1
            FTotalCount = rsget.RecordCount
			FResultCount = FTotalCount

			redim preserve FItemList(FResultCount)
			i=0
			if  not rsget.EOF  then
				rsget.absolutepage = FCurrPage
				do until rsget.eof

					set FItemList(i) = new CBrandCSItem
					''FItemList(i).FID            = rsget("id")
					FItemList(i).Fdivcd_Name    = rsget("divcd_Name")
					FItemList(i).Fgubun02_Name  = rsget("gubun02_Name")

					''FItemList(i).Forderserial  = rsget("orderserial")
					FItemList(i).Fitemid        = rsget("itemid")
					''FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).Fitemname      = rsget("itemname")
					''FItemList(i).Fitemoptionname = rsget("itemoptionname")

					FItemList(i).Fconfirmitemno    = rsget("totconfirmitemno")
					FItemList(i).Fisupchebeasong    = rsget("isupchebeasong")
					''FItemList(i).Ffinishdate    = rsget("finishdate")
					i=i+1
					rsget.moveNext
				loop
			end if
        rsget.Close

    end Sub

    public Sub getBrandCsSummary_GubunGroup()
        dim sqlStr
        sqlStr = "select "
        sqlStr = sqlStr + " s.yyyymm, s.makerid, s.isupchebeasong "
        if (FRectDivCd="T012") then
            sqlStr = sqlStr + " ,s.divcd , '' as divname"
        else
            sqlStr = sqlStr + " ,s.divcd , s.divname"
        end if
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CD05','CF05') then cnt else 0 end) as CNT_1"  ''품절
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF06','CG01') then cnt else 0 end) as CNT_2"  ''출고지연
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CE01','CE02') then cnt else 0 end) as CNT_3"  '' 상품불량,불만족
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF03','CF04','CF01') then cnt else 0 end) as CNT_4"  '' 상품누락,사은품누락,오발송
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CE04','CE03') then cnt else 0 end) as CNT_5"  ''상품등록오류
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF02','CG02','CG03') then cnt else 0 end) as CNT_6"  '' 상품파손,택배사파손,분실
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CD01','CB04') then cnt else 0 end) as CNT_7"  '' 단순변심,고객변심
        sqlStr = sqlStr + " ,sum(s.cnt) as sumCnt"
        sqlStr = sqlStr + " from db_datamart.dbo.tbl_cs_monthly_BrandAS_summary s"
        sqlStr = sqlStr + " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and s.yyyymm='" & FRectYYYYMM & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and s.makerid='" & FRectMakerid & "'"
        end if

        if (FRectDivCd="T012") then
            sqlStr = sqlStr & " and s.divcd in ('A000','A001','A002')"
        elseif (FRectDivCd<>"") then
            sqlStr = sqlStr & " and s.divcd='" & FRectDivCd & "'"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr & " and s.isupchebeasong='" & FRectIsupchebeasong & "'"
        end if

        if (FRectNotIncludeETC="on") then
            sqlStr = sqlStr & " and s.gubun02 not in ('CD99','CF99','CE99','CH99')"
        end if

        if (FRectCDL<>"") then
            sqlStr = sqlStr & " and s.makerid in (Select userid From TENDB.db_user.dbo.tbl_user_c where cateCode='" & FRectCDL & "')"
        end if

        sqlStr = sqlStr + " group by s.yyyymm, s.makerid, s.isupchebeasong"
        if (FRectDivCd="T012") then
            sqlStr = sqlStr + " , s.divcd"
        else
            sqlStr = sqlStr + " , s.divcd, s.divname"
        end if
        sqlStr = sqlStr + " order by yyyymm desc, sumCnt desc"

        db3_rsget.Open sqlStr,db3_dbget,1

        FTotalCount  = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
		FResultCount = FTotalCount

        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBrandCSSummaryItem
				FItemList(i).Fyyyymm        = db3_rsget("yyyymm")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fisupchebeasong= db3_rsget("isupchebeasong")
                FItemList(i).Fdivcd         = db3_rsget("divcd")
                FItemList(i).Fdivname       = db2Html(db3_rsget("divname"))



                FItemList(i).FCNT_1       = db3_rsget("CNT_1")
                FItemList(i).FCNT_2       = db3_rsget("CNT_2")
                FItemList(i).FCNT_3       = db3_rsget("CNT_3")
                FItemList(i).FCNT_4       = db3_rsget("CNT_4")
                FItemList(i).FCNT_5       = db3_rsget("CNT_5")
                FItemList(i).FCNT_6       = db3_rsget("CNT_6")
                FItemList(i).FCNT_7       = db3_rsget("CNT_7")

                FItemList(i).Fcnt           = db3_rsget("sumCnt")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end Sub

    public Sub getBrandCsSummary_GubunGroupNew()
        dim sqlStr, addSqlStr

		addSqlStr = ""

        if (FRectYYYYMM<>"") then
            addSqlStr = addSqlStr & " and s.yyyymm='" & FRectYYYYMM & "'"
        end if

        if (FRectMakerid<>"") then
            addSqlStr = addSqlStr & " and s.makerid='" & FRectMakerid & "'"
        end if

        if (FRectDivCd="T012") then
            addSqlStr = addSqlStr & " and s.divcd in ('A000','A001','A002')"
        elseif (FRectDivCd<>"") then
            addSqlStr = addSqlStr & " and s.divcd='" & FRectDivCd & "'"
        end if

        if (FRectIsupchebeasong<>"") then
            addSqlStr = addSqlStr & " and s.isupchebeasong='" & FRectIsupchebeasong & "'"
        end if

        if (FRectNotIncludeETC="on") then
            addSqlStr = addSqlStr & " and s.gubun02 not in ('CD99','CF99','CE99','CH99')"
        end if

        if (FRectCDL<>"") then
            addSqlStr = addSqlStr & " and s.makerid in (Select userid From TENDB.db_user.dbo.tbl_user_c where cateCode='" & FRectCDL & "')"
        end if

		'// ====================================================================
        sqlStr = "select s.yyyymm "
        sqlStr = sqlStr + " from db_datamart.dbo.tbl_cs_monthly_BrandAS_summary s"
        sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSqlStr

        sqlStr = sqlStr + " group by s.yyyymm, s.makerid, s.isupchebeasong"
        if (FRectDivCd="T012") then
            sqlStr = sqlStr + " , s.divcd"
        else
            sqlStr = sqlStr + " , s.divcd, s.divname"
        end if

		sqlStr = "select count(*) as cnt from (" + sqlStr + ") T "
		''response.write sqlStr

        db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close


		'// ====================================================================
        sqlStr = "select "
        sqlStr = sqlStr + " top " & (FPagesize*FCurrPage) & " s.yyyymm, s.makerid, s.isupchebeasong "
        if (FRectDivCd="T012") then
            sqlStr = sqlStr + " ,s.divcd , '' as divname"
        else
            sqlStr = sqlStr + " ,s.divcd , s.divname"
        end if
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CD05','CF05') then cnt else 0 end) as CNT_1"  ''품절
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 = 'CE01' then cnt else 0 end) as CNT_2"  			'' 상품불량
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 = 'CF01' then cnt else 0 end) as CNT_3"  			'' 오발송
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 = 'CF02' then cnt else 0 end) as CNT_4"  			'' 상품파손
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF03','CF04') then cnt else 0 end) as CNT_5"  '' 상품누락,사은품누락
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF06','CG01') then cnt else 0 end) as CNT_6"  ''출고지연
        sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CD01','CB04') then cnt else 0 end) as CNT_7"  '' 단순변심,고객변심
        'CE02,CE04,CE03,CG02,CG03

        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF06','CG01') then cnt else 0 end) as CNT_2"  ''출고지연
        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CE01','CE02') then cnt else 0 end) as CNT_3"  '' 상품불량,불만족
        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF03','CF04','CF01') then cnt else 0 end) as CNT_4"  '' 상품누락,사은품누락,오발송
        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CE04','CE03') then cnt else 0 end) as CNT_5"  ''상품등록오류
        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CF02','CG02','CG03') then cnt else 0 end) as CNT_6"  '' 상품파손,택배사파손,분실
        'sqlStr = sqlStr + " ,sum(Case when s.gubun02 in ('CD01','CB04') then cnt else 0 end) as CNT_7"  '' 단순변심,고객변심
        sqlStr = sqlStr + " ,sum(s.cnt) as sumCnt"
        sqlStr = sqlStr + " from db_datamart.dbo.tbl_cs_monthly_BrandAS_summary s"
        sqlStr = sqlStr + " where 1=1"

		sqlStr = sqlStr + addSqlStr

        sqlStr = sqlStr + " group by s.yyyymm, s.makerid, s.isupchebeasong"
        if (FRectDivCd="T012") then
            sqlStr = sqlStr + " , s.divcd"
        else
            sqlStr = sqlStr + " , s.divcd, s.divname"
        end if
        sqlStr = sqlStr + " order by yyyymm desc, sumCnt desc"

		'' response.write sqlStr
		'' response.end
        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBrandCSSummaryItem
				FItemList(i).Fyyyymm        = db3_rsget("yyyymm")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fisupchebeasong= db3_rsget("isupchebeasong")
                FItemList(i).Fdivcd         = db3_rsget("divcd")
                FItemList(i).Fdivname       = db2Html(db3_rsget("divname"))



                FItemList(i).FCNT_1       = db3_rsget("CNT_1")
                FItemList(i).FCNT_2       = db3_rsget("CNT_2")
                FItemList(i).FCNT_3       = db3_rsget("CNT_3")
                FItemList(i).FCNT_4       = db3_rsget("CNT_4")
                FItemList(i).FCNT_5       = db3_rsget("CNT_5")
                FItemList(i).FCNT_6       = db3_rsget("CNT_6")
                FItemList(i).FCNT_7       = db3_rsget("CNT_7")

                FItemList(i).Fcnt           = db3_rsget("sumCnt")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end Sub

    public Sub getBrandCssummary()
        dim sqlStr
        sqlStr = "select count(*) as cnt"
        sqlStr = sqlStr + " from db_datamart.dbo.tbl_cs_monthly_BrandAS_summary s"
        sqlStr = sqlStr + " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and s.yyyymm='" & FRectYYYYMM & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and s.makerid='" & FRectMakerid & "'"
        end if

        if (FRectDivCd="T012") then
            sqlStr = sqlStr & " and s.divcd in ('A000','A001','A002')"
        elseif (FRectDivCd<>"") then
            sqlStr = sqlStr & " and s.divcd='" & FRectDivCd & "'"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr & " and s.isupchebeasong='" & FRectIsupchebeasong & "'"
        end if

        if (FRectNotIncludeETC="on") then
            sqlStr = sqlStr & " and s.gubun02 not in ('CD99','CF99','CE99','CH99')"
        end if

        if (FRectCDL<>"") then
            sqlStr = sqlStr & " and s.makerid in (Select userid From TENDB.db_user.dbo.tbl_user_c where cateCode='" & FRectCDL & "')"
        end if

        db3_rsget.Open sqlStr,db3_dbget,1
            FTotalCount = db3_rsget("cnt")
        db3_rsget.Close

        sqlStr = "select top " & (FPagesize*FCurrPage) & " s.* "
        'sqlStr = sqlStr + " , p.groupid"
        sqlStr = sqlStr + " from db_datamart.dbo.tbl_cs_monthly_BrandAS_summary s"
        'sqlStr = sqlStr + "     left join db_partner.dbo.tbl_partner p"
        'sqlStr = sqlStr + "     on s.makerid=p.id"
        sqlStr = sqlStr + " where 1=1"
        if (FRectYYYYMM<>"") then
            sqlStr = sqlStr & " and s.yyyymm='" & FRectYYYYMM & "'"
        end if

        if (FRectMakerid<>"") then
            sqlStr = sqlStr & " and s.makerid='" & FRectMakerid & "'"
        end if

        if (FRectDivCd="T012") then
            sqlStr = sqlStr & " and s.divcd in ('A000','A001','A002')"
        elseif (FRectDivCd<>"") then
            sqlStr = sqlStr & " and s.divcd='" & FRectDivCd & "'"
        end if

        if (FRectIsupchebeasong<>"") then
            sqlStr = sqlStr & " and s.isupchebeasong='" & FRectIsupchebeasong & "'"
        end if

        if (FRectNotIncludeETC="on") then
            sqlStr = sqlStr & " and s.gubun02 not in ('CD99','CF99','CE99','CH99')"
        end if

        if (FRectCDL<>"") then
            sqlStr = sqlStr & " and s.makerid in (Select userid From TENDB.db_user.dbo.tbl_user_c where cateCode='" & FRectCDL & "')"
        end if

        sqlStr = sqlStr + " order by s.yyyymm desc, s.cnt desc"

		'' response.write sqlStr
		'' response.end
        db3_rsget.pagesize = FPageSize
        db3_rsget.Open sqlStr,db3_dbget,1

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = db3_rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if (FResultCount<1) then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
		    db3_rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CBrandCSSummaryItem
				FItemList(i).Fyyyymm        = db3_rsget("yyyymm")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fisupchebeasong= db3_rsget("isupchebeasong")
                FItemList(i).Fdivcd         = db3_rsget("divcd")
                FItemList(i).Fgubun01       = db3_rsget("gubun01")
                FItemList(i).Fgubun02       = db3_rsget("gubun02")
                FItemList(i).Fdivname       = db2Html(db3_rsget("divname"))
                FItemList(i).Fgubun01name   = db2Html(db3_rsget("gubun01name"))
                FItemList(i).Fgubun02name   = db2Html(db3_rsget("gubun02name"))
                FItemList(i).Fcnt           = db3_rsget("cnt")


				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end Sub

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 100
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
