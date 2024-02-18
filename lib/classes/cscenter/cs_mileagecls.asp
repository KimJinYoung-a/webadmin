<%
'###########################################################
' Description : 고객센터 마일리지관리 클래스
' History : 이상구 생성
'###########################################################

Class CCSCenterMileageSummaryItem

        public Ftotalbuymileage         '6개월이내 구매마일리지
        public Ftotalbonusmileage
        public Ftotalspendmileage
        public Ftotaloldbuymileage      '6개월이전 구매마일리지
        public Facademymileage          '아카데미(핑거스) 구매마일리지

        public FrealExpiredMileage      '만료된 마일리지

        public function getTotalBuymileage()
            getTotalBuymileage = CLng(Ftotalbuymileage) + CLng(Ftotaloldbuymileage) + CLng(Facademymileage)
        end function

        public function getCurrentMileage()
            getCurrentMileage = getTotalBuymileage + CLng(Ftotalbonusmileage) - CLng(Ftotalspendmileage) - CLng(FrealExpiredMileage)
        end function

        Private Sub Class_Initialize()
            FrealExpiredMileage = 0
        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class


Class CUserCurrentMileageItem
    public Fuserid
    public Fjumunmileage                '6개월이내 구매마일리지
    public Fbonusmileage
    public Fspendmileage
    public Flastupdate
    public Fflowerjumunmileage          '6개월이전 구매마일리지
    public Facademymileage              '아카데미(핑거스) 구매마일리지

    public FrealExpiredMileage          '만료된 마일리지

    public function getTotalBuymileage()
        getTotalBuymileage = Fjumunmileage + Fflowerjumunmileage + Facademymileage
    end function

    public function getCurrentMileage()
        getCurrentMileage = getTotalBuymileage + Fbonusmileage - Fspendmileage - FrealExpiredMileage
    end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

end Class

Class CSMileageItem
	Public Forderserial
	Public Fcustomername
	Public Fuserid
	Public Ftitle
	Public Fwriteuser
	Public Fusername
	Public Ffinishuser
	Public Frefundresult
	Public Fregdate
	Public Ffinishdate
    Public Fcontents_jupsu

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

end Class

Class CExpireYearItem
    public FregYear
    public Fuserid
    public Fexpiredate
    public Fbonusgainmileage
    public Fordergainmileage
    public Forderminusmileage
    public FpreYearAssignedSpendmileage
    public FrealExpiredMileage

    public FspendMileage
    public FaccumulateGainSum
    public FaccumulateOrderMinusMileage
    public FaccumulateExpiredSum

    public function getKorExpireDateStr()
        getKorExpireDateStr = Left(Fexpiredate,4) &"년 " & Mid(Fexpiredate,6,2) & "월 " & Mid(Fexpiredate,9,2) & "일"
    end function

    public function getMayExpireTotal()
        getMayExpireTotal = getGainMileage-Fspendmileage-FrealExpiredMileage

        if (getMayExpireTotal<1) then getMayExpireTotal=0
    end function

    public function getGainMileage()
        getGainMileage = Fbonusgainmileage + Fordergainmileage + Forderminusmileage
    end function

    public function getYearMaySpendMileage_OLD()
        dim acctremain
        acctremain = (FaccumulateGainSum + FaccumulateOrderMinusMileage) - Fspendmileage
        if (acctremain=<0) then
            getYearMaySpendMileage_OLD = getGainMileage
        elseif (acctremain>=getGainMileage) then
            getYearMaySpendMileage_OLD = 0
        else
            getYearMaySpendMileage_OLD = getGainMileage-acctremain
        end if
    end function

    public function getYearMaySpendMileage()
        dim acctremain

        ' FaccumulateGainSum  ' 누적적립마일리지
        ' FaccumulateOrderMinusMileage  ' 누적주문마일너스마일리지
        ' FaccumulateExpiredSum ' 누적소멸마일리지
        acctremain = (FaccumulateGainSum + FaccumulateOrderMinusMileage - (FaccumulateExpiredSum - FrealExpiredMileage)) - Fspendmileage

        if (acctremain=<0) or (acctremain<=FrealExpiredMileage) then
            getYearMaySpendMileage = getGainMileage - FrealExpiredMileage
        elseif (acctremain>=getGainMileage) then
            getYearMaySpendMileage = 0
        elseif (getGainMileage<=FrealExpiredMileage) then
            getYearMaySpendMileage = 0
        else
            getYearMaySpendMileage = getGainMileage - FrealExpiredMileage -acctremain
        end if
    end function

    public function getYearMayRemainMileage()
        getYearMayRemainMileage = getGainMileage - getYearMaySpendMileage - FrealExpiredMileage
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExpireMonthItem
    public Fregmonth
    public Fuserid
    public Fexpiredate
    public Fbonusgainmileage
    public Fordergainmileage
    public Forderminusmileage
    public FrealExpiredMileage
    public FspendMileage
    public FaccumulateGainSum
    public FaccumulateOrderMinusMileage
    public FaccumulateExpiredSum

    public function getKorExpireDateStr()
        getKorExpireDateStr = Left(Fexpiredate,4) &"년 " & Mid(Fexpiredate,6,2) & "월 " & Mid(Fexpiredate,9,2) & "일"
    end function

    public function getMayExpireTotal()
        getMayExpireTotal = getGainMileage-Fspendmileage-FrealExpiredMileage

        if (getMayExpireTotal<1) then getMayExpireTotal=0
    end function

    public function getGainMileage()
        getGainMileage = Fbonusgainmileage + Fordergainmileage + Forderminusmileage
    end function

    public function getMonthlyMaySpendMileage()
        dim acctremain

        ' FaccumulateGainSum  ' 누적적립마일리지
        ' FaccumulateOrderMinusMileage  ' 누적주문마일너스마일리지
        ' FaccumulateExpiredSum ' 누적소멸마일리지
        acctremain = (FaccumulateGainSum + FaccumulateOrderMinusMileage - (FaccumulateExpiredSum - FrealExpiredMileage)) - Fspendmileage

        if (acctremain=<0) or (acctremain<=FrealExpiredMileage) then
            getMonthlyMaySpendMileage = getGainMileage - FrealExpiredMileage
        elseif (acctremain>=getGainMileage) then
            getMonthlyMaySpendMileage = 0
        elseif (getGainMileage<=FrealExpiredMileage) then
            getMonthlyMaySpendMileage = 0
        else
            getMonthlyMaySpendMileage = getGainMileage - FrealExpiredMileage -acctremain
        end if
    end function

    public function getMonthlyMayRemainMileage()
        getMonthlyMayRemainMileage = getGainMileage - getMonthlyMaySpendMileage - FrealExpiredMileage
    end function

    Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class CCSCenterMileageItem
        public Fid
        public Fuserid
        public Fmileage
        public Fjukyocd
        public Fjukyo
        public Fregdate
        public Forderserial
        public Fdeleteyn

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CCSCenterMileage
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

        public FRectUserID
        public FRectDeleteYn
        public FRectExpireDate
		public FRectStartDate
		public FRectEndDate

		Public FRectGrpBy
		Public FRectWriteUser

        ' /cscenter/mileage/popAdminExpireMileSummary.asp    ' /cscenter/mileage/popAdminExpireMileMonthlySummary.asp
        public Sub getUserCurrentMileage()
            dim sqlStr

            if FRectUserID="" or isnull(FRectUserID) then exit Sub

            sqlStr = "select m.userid, m.jumunmileage, m.bonusmileage, m.spendmileage, m.flowerjumunmileage, m.academymileage "
    		sqlStr = sqlStr + " , IsNULL(m.ExpiredMile,0) as realExpiredMileage, m.lastupdate "
    		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_current_mileage m with (nolock)"
    		sqlStr = sqlStr + " where m.userid='" + FRectUserID + "'"

            'response.write sqlStr & "<br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

            FResultCount = rsget.RecordCount
            ftotalcount = rsget.RecordCount

            redim preserve FItemList(FResultCount)
            if  not rsget.EOF  then
                set FOneItem = new CUserCurrentMileageItem

                FOneItem.Fuserid              = rsget("userid")
                FOneItem.Fjumunmileage        = rsget("jumunmileage")
                FOneItem.Fbonusmileage        = rsget("bonusmileage")
                FOneItem.Fspendmileage        = rsget("spendmileage")
                FOneItem.Flastupdate          = rsget("lastupdate")
                FOneItem.Fflowerjumunmileage  = rsget("flowerjumunmileage")
                FOneItem.Facademymileage      = rsget("academymileage")
                FOneItem.FrealExpiredMileage  = rsget("realExpiredMileage")
            end if
            rsget.close
        end sub

        '' Old function
        public Sub GetCSCenterMileageSummary()
                dim i,sqlStr

'                sqlStr = " select top 1 IsNull(T.tmile,0) as totalbuymileage, IsNull(c.bonusmileage,0) as totalbonusmileage, IsNull(c.spendmileage,0) as totalspendmileage, "
'                sqlStr = sqlStr + " IsNull(c.jumunmileage,0) as notused1,IsNull(c.flowerjumunmileage,0) as totaloldbuymileage, IsNull(c.academymileage,0) as academymileage "
'                sqlStr = sqlStr + " from "
'                sqlStr = sqlStr + " ( "
'                sqlStr = sqlStr + "     select sum(m.totalmileage) as tmile "
'                sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m "
'                sqlStr = sqlStr + "     where m.userid='" + CStr(FRectUserID) + "' "
'                sqlStr = sqlStr + "     and m.userid<>'' "
'                sqlStr = sqlStr + "     and m.cancelyn='N' "
'                sqlStr = sqlStr + "     and m.ipkumdiv >=4 "
'                sqlStr = sqlStr + "     and m.sitename='10x10' "
'                sqlStr = sqlStr + " ) as T "
'                sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_current_mileage c on c.userid='" + CStr(FRectUserID) + "' and c.userid <> '' "

'                sqlStr = " select top 1 c.* from [db_user].[dbo].tbl_user_current_mileage c "
'                sqlStr = sqlStr + " where c.userid='" + CStr(FRectUserID) + "'"

                sqlStr = "select m.jumunmileage, m.bonusmileage, m.spendmileage, m.flowerjumunmileage, m.academymileage "
        		sqlStr = sqlStr + " , IsNULL(e.realExpiredMileage,0) as realExpiredMileage "
        		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_current_mileage m"
        		sqlStr = sqlStr + " left join ("
        		sqlStr = sqlStr + "     select userid, sum(realExpiredMileage) as realExpiredMileage"
                sqlStr = sqlStr + "     from db_user.dbo.tbl_mileage_Year_Expire"
                sqlStr = sqlStr + "     where userid='" + FRectUserID + "'"
                sqlStr = sqlStr + "     group by userid"
                sqlStr = sqlStr + " ) e on m.userid=e.userid"
        		sqlStr = sqlStr + " where m.userid='" + FRectUserID + "'"

                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new CCSCenterMileageSummaryItem

                                FItemList(i).Ftotalbuymileage           = rsget("jumunmileage")
                                FItemList(i).Ftotalbonusmileage         = rsget("bonusmileage")
                                FItemList(i).Ftotalspendmileage         = rsget("spendmileage")
                                FItemList(i).Ftotaloldbuymileage        = rsget("flowerjumunmileage")
                                FItemList(i).Facademymileage            = rsget("academymileage")

                                FItemList(i).FrealExpiredMileage        = rsget("realExpiredMileage")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        public Sub GetCSCenterMileageList()
                dim i,sqlStr

                sqlStr = " select top 500 m.id,m.userid,m.mileage,m.jukyocd,m.jukyo,m.regdate,m.orderserial,m.deleteyn "
                sqlStr = sqlStr + " from [db_user].[dbo].tbl_mileagelog m "
                sqlStr = sqlStr + " where m.userid='" + CStr(FRectUserID) + "' "
                if (FRectDeleteYn<>"") then
                    sqlStr = sqlStr + " and m.deleteyn='" + CStr(FRectDeleteYn) + "' "
                end if
                sqlStr = sqlStr + " order by m.regdate desc "

                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new CCSCenterMileageItem

                                FItemList(i).Fid                = rsget("id")
                                FItemList(i).Fuserid            = rsget("userid")
                                FItemList(i).Fmileage           = rsget("mileage")
                                FItemList(i).Fjukyocd           = rsget("jukyocd")
                                FItemList(i).Fjukyo             = rsget("jukyo")
                                FItemList(i).Fregdate           = rsget("regdate")
                                FItemList(i).Forderserial       = rsget("orderserial")
                                FItemList(i).Fdeleteyn          = rsget("deleteyn")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        ' 소멸될 마일리지 월별 합계.    ' 2023.07.21 한용민 생성
        ' /cscenter/mileage/cs_mileage.asp	' /cscenter/mileage/popAdminExpireMileMonthlySummary.asp
        public Sub getNextExpireMileageMonthlySum()
            dim sqlStr

            if FRectUserid="" or isnull(FRectUserid) then exit Sub
            if FRectExpireDate="" or isnull(FRectExpireDate) then exit Sub

            sqlStr = " select"
            sqlStr = sqlStr & " c.userid, c.spendmileage, '"&FRectExpireDate&"' as expiredate"
            sqlStr = sqlStr & " , IsNULL(T.bonusgainmileage,0) as bonusgainmileage, IsNULL(T.ordergainmileage,0) as ordergainmileage"
            sqlStr = sqlStr & " , IsNULL(T.orderminusmileage,0) as orderminusmileage"
            sqlStr = sqlStr & " , IsNULL(T.realExpiredMileage,0) as realExpiredMileage"
            sqlStr = sqlStr & " from db_user.[dbo].tbl_user_current_mileage c with (nolock)"
            sqlStr = sqlStr & " left join ("
            sqlStr = sqlStr & "     select"
            sqlStr = sqlStr & "     e.userid, sum(e.bonusgainmileage) as bonusgainmileage, sum(e.ordergainmileage) as ordergainmileage"
            sqlStr = sqlStr & "     , sum(e.orderminusmileage) as orderminusmileage"
            sqlStr = sqlStr & "     , sum(e.realExpiredMileage) as realExpiredMileage"
            sqlStr = sqlStr & "     from db_user.dbo.tbl_mileageMonthlyExpire e with (nolock)"
            sqlStr = sqlStr & "     where e.userid='" & FRectUserid & "'"
            sqlStr = sqlStr & "     and e.expiredate<='" & FRectExpireDate & "'"
            sqlStr = sqlStr & "     group by e.userid"
            sqlStr = sqlStr & " ) T"
            sqlStr = sqlStr & "     on c.userid=T.userid"
            sqlStr = sqlStr & " where c.userid='" & FRectUserid & "'"

            'response.write sqlStr &"<br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

            FTotalCount = rsget.RecordCount
            FResultCount = rsget.RecordCount

            if  not rsget.EOF  then
    			set FOneItem = new CExpireMonthItem
                FOneItem.Fuserid                      = rsget("userid")
                FOneItem.Fexpiredate                  = FRectExpireDate
                FOneItem.Fbonusgainmileage            = rsget("bonusgainmileage")
                FOneItem.Fordergainmileage            = rsget("ordergainmileage")
                FOneItem.Forderminusmileage           = rsget("orderminusmileage")
                FOneItem.FrealExpiredMileage          = rsget("realExpiredMileage")
                FOneItem.FspendMileage                = rsget("spendMileage")
            else
                '' 만료 예정내역이 없을 경우.
                set FOneItem = new CExpireMonthItem
                FOneItem.Fuserid = FRectUserid
                FOneItem.Fexpiredate = FRectExpireDate
                FOneItem.Fbonusgainmileage  = 0
                FOneItem.Fordergainmileage  = 0
                FOneItem.FrealExpiredMileage = 0
                FOneItem.FspendMileage = 0
    		end if

    		rsget.Close
        end sub

        ' 다음년초에 Expire될 마일리지 합계.    ' 서동석 생성
        ' /cscenter/mileage/cs_mileage.asp  ' /cscenter/mileage/popAdminExpireMileMonthlySummary.asp
        public Sub getNextExpireMileageSum()
            dim sqlStr

            if FRectUserid="" or isnull(FRectUserid) then exit Sub
            if FRectExpireDate="" or isnull(FRectExpireDate) then exit Sub

            sqlStr = " select c.userid, c.spendmileage,'"&FRectExpireDate&"' as expiredate,"
            sqlStr = sqlStr & " IsNULL(T.bonusgainmileage,0) as bonusgainmileage, IsNULL(T.ordergainmileage,0) as ordergainmileage,"
            sqlStr = sqlStr & " IsNULL(T.orderminusmileage,0) as orderminusmileage, IsNULL(T.preYearAssignedSpendMileage,0) as preYearAssignedSpendMileage,"
            sqlStr = sqlStr & " IsNULL(T.realExpiredMileage,0) as realExpiredMileage"
            sqlStr = sqlStr & " from db_user.[dbo].tbl_user_current_mileage c with (nolock)"
            sqlStr = sqlStr & " left join ("
            sqlStr = sqlStr & "     select e.userid, sum(e.bonusgainmileage) as bonusgainmileage, sum(e.ordergainmileage) as ordergainmileage,"
            sqlStr = sqlStr & "     sum(e.orderminusmileage) as orderminusmileage, sum(e.preYearAssignedSpendMileage) as preYearAssignedSpendMileage,"
            sqlStr = sqlStr & "     sum(e.realExpiredMileage) as realExpiredMileage"
            sqlStr = sqlStr & "     from db_user.dbo.tbl_mileage_Year_Expire e with (nolock)"
            sqlStr = sqlStr & "     where e.userid='" & FRectUserid & "'"
            sqlStr = sqlStr & "     and e.expiredate<='" & FRectExpireDate & "'"
            sqlStr = sqlStr & "     group by e.userid"
            sqlStr = sqlStr & " ) T"
            sqlStr = sqlStr & "     on c.userid=T.userid"
            sqlStr = sqlStr & " where c.userid='" & FRectUserid & "'"

            'response.write sqlStr &"<br>"
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

            FTotalCount = rsget.RecordCount
            FResultCount = FTotalCount

            if  not rsget.EOF  then
    			set FOneItem = new CExpireYearItem
                FOneItem.Fuserid                      = rsget("userid")
                FOneItem.Fexpiredate                  = FRectExpireDate
                FOneItem.Fbonusgainmileage            = rsget("bonusgainmileage")
                FOneItem.Fordergainmileage            = rsget("ordergainmileage")
                FOneItem.Forderminusmileage           = rsget("orderminusmileage")
                FOneItem.FpreYearAssignedSpendmileage = rsget("preYearAssignedSpendmileage")
                FOneItem.FrealExpiredMileage          = rsget("realExpiredMileage")
                FOneItem.FspendMileage                = rsget("spendMileage")
            else
                '' 만료 예정내역이 없을 경우.
                set FOneItem = new CExpireYearItem
                FOneItem.Fuserid = FRectUserid
                FOneItem.Fexpiredate = FRectExpireDate
                FOneItem.Fbonusgainmileage  = 0
                FOneItem.Fordergainmileage  = 0
                FOneItem.FpreYearAssignedSpendmileage = 0
                FOneItem.FrealExpiredMileage = 0
                FOneItem.FspendMileage = 0
    		end if

    		rsget.Close
        end sub

        ' 소멸될 마일리지 월별 합계.    ' 2023.07.21 한용민 생성
        ' /cscenter/mileage/popAdminExpireMileSummary.asp    ' /cscenter/mileage/popAdminExpireMileMonthlySummary.asp
        public Sub getNextExpireMileageMonthlyList()
            dim sqlStr, i
            dim t_accumulateGainSum, t_accumulateOrderMinusMileage, t_accumulateExpiredSum

            if FRectUserid="" or isnull(FRectUserid) then exit Sub

            sqlStr = " select"
            sqlStr = sqlStr & " e.regmonth, e.userid, e.expiredate, e.bonusgainmileage, e.ordergainmileage, e.orderminusmileage"
            sqlStr = sqlStr & " , e.realExpiredMileage"
            sqlStr = sqlStr & " , IsNULL(c.spendmileage,0) as spendmileage"
            sqlStr = sqlStr & " from db_user.dbo.tbl_mileageMonthlyExpire e with (nolock)"
            sqlStr = sqlStr & " left join db_user.[dbo].tbl_user_current_mileage c with (nolock)"
            sqlStr = sqlStr & "     on e.userid=c.userid"
            sqlStr = sqlStr & " where e.userid='" & FRectUserid & "'"

            if (FRectExpireDate<>"") then
                sqlStr = sqlStr & " and e.expiredate='" & FRectExpireDate & "'"
            end if

            sqlStr = sqlStr & " order by e.regmonth asc"

            'response.write sqlStr &"<br>"
            rsget.pagesize = FPageSize
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if
    		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

            if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		t_accumulateGainSum =0
    		t_accumulateOrderMinusMileage =0
    		t_accumulateExpiredSum =0
    		if  not rsget.EOF  then
    			rsget.absolutepage = FCurrPage
    			do until rsget.eof
        			set FItemList(i) = new CExpireMonthItem
        			FItemList(i).Fregmonth                       = rsget("regmonth")
                    FItemList(i).Fuserid                        = rsget("userid")
                    FItemList(i).Fexpiredate                    = rsget("expiredate")
                    FItemList(i).Fbonusgainmileage              = rsget("bonusgainmileage")
                    FItemList(i).Fordergainmileage              = rsget("ordergainmileage")
                    FItemList(i).Forderminusmileage             = rsget("orderminusmileage")
                    FItemList(i).FrealExpiredMileage            = rsget("realExpiredMileage")

                    FItemList(i).FspendMileage                  = rsget("spendMileage")

                    ''누적 적립마일리지
                    t_accumulateGainSum                         = t_accumulateGainSum + FItemList(i).Fbonusgainmileage + FItemList(i).Fordergainmileage
                    t_accumulateOrderMinusMileage               = t_accumulateOrderMinusMileage + FItemList(i).Forderminusmileage
                    t_accumulateExpiredSum                      = t_accumulateExpiredSum + FItemList(i).FrealExpiredMileage
                    FItemList(i).FaccumulateGainSum             = t_accumulateGainSum
                    FItemList(i).FaccumulateOrderMinusMileage   = t_accumulateOrderMinusMileage
                    FItemList(i).FaccumulateExpiredSum          = t_accumulateExpiredSum
                    i=i+1
    				rsget.moveNext
    			loop
    		end if

    		rsget.Close
        end Sub

        ''다음년초에 Expire될 마일리지 년도별 합계. ' 서동석 생성
        ' /cscenter/mileage/popAdminExpireMileSummary.asp    ' /cscenter/mileage/popAdminExpireMileMonthlySummary.asp
        public Sub getNextExpireMileageYearList()
            dim sqlStr,i
            dim t_accumulateGainSum, t_accumulateOrderMinusMileage, t_accumulateExpiredSum

		    if FRectUserid="" or isnull(FRectUserid) then exit Sub

            sqlStr = " select e.regYear, e.userid, e.expiredate, e.bonusgainmileage, e.ordergainmileage,"
            sqlStr = sqlStr & " e.orderminusmileage, e.preYearAssignedSpendMileage,"
            sqlStr = sqlStr & " e.realExpiredMileage,"
            sqlStr = sqlStr & " IsNULL(c.spendmileage,0) as spendmileage"
            sqlStr = sqlStr & " from db_user.dbo.tbl_mileage_Year_Expire e with (nolock)"
            sqlStr = sqlStr & " left join db_user.[dbo].tbl_user_current_mileage c with (nolock)"
            sqlStr = sqlStr & " on e.userid=c.userid"
            sqlStr = sqlStr & " where e.userid='" & FRectUserid & "'"
            if (FRectExpireDate<>"") then
                sqlStr = sqlStr & " and e.expiredate='" & FRectExpireDate & "'"
            end if
            sqlStr = sqlStr & " order by e.regYear"

            'response.write sqlStr & "<Br>"
            rsget.pagesize = FPageSize
            rsget.CursorLocation = adUseClient
            rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if
    		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

            if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		t_accumulateGainSum =0
    		t_accumulateOrderMinusMileage =0
    		t_accumulateExpiredSum =0
    		if  not rsget.EOF  then
    			rsget.absolutepage = FCurrPage
    			do until rsget.eof
        			set FItemList(i) = new CExpireYearItem
        			FItemList(i).FregYear                       = rsget("regYear")
                    FItemList(i).Fuserid                        = rsget("userid")
                    FItemList(i).Fexpiredate                    = rsget("expiredate")
                    FItemList(i).Fbonusgainmileage              = rsget("bonusgainmileage")
                    FItemList(i).Fordergainmileage              = rsget("ordergainmileage")
                    FItemList(i).Forderminusmileage             = rsget("orderminusmileage")
                    FItemList(i).FpreYearAssignedSpendmileage   = rsget("preYearAssignedSpendmileage")
                    FItemList(i).FrealExpiredMileage            = rsget("realExpiredMileage")

                    FItemList(i).FspendMileage                  = rsget("spendMileage")

                    ''누적 적립마일리지
                    t_accumulateGainSum                         = t_accumulateGainSum + FItemList(i).Fbonusgainmileage + FItemList(i).Fordergainmileage
                    t_accumulateOrderMinusMileage               = t_accumulateOrderMinusMileage + FItemList(i).Forderminusmileage
                    t_accumulateExpiredSum                      = t_accumulateExpiredSum + FItemList(i).FrealExpiredMileage
                    FItemList(i).FaccumulateGainSum             = t_accumulateGainSum
                    FItemList(i).FaccumulateOrderMinusMileage   = t_accumulateOrderMinusMileage
                    FItemList(i).FaccumulateExpiredSum          = t_accumulateExpiredSum
                    i=i+1
    				rsget.moveNext
    			loop
    		end if

    		rsget.Close
        end Sub

		Public Sub getCSMileage()
            dim sqlStr,i

			Select Case FRectGrpBy
				Case "writeuser"
					sqlStr = " select top " & FPageSize*FCurrPage & " '' as orderserial, '' as customername, '' as userid, '' as title, a.writeuser, t.username, '' as finishuser, sum(r.refundresult) as refundresult, '' as regdate, convert(varchar(7),a.finishdate,121) as finishdate, '' as contents_jupsu "
				Case "title"
					sqlStr = " select top " & FPageSize*FCurrPage & " '' as orderserial, '' as customername, '' as userid, a.title, '' as writeuser, '' as username, '' as finishuser, sum(r.refundresult) as refundresult, '' as regdate, convert(varchar(7),a.finishdate,121) as finishdate, '' as contents_jupsu "
				Case Else
					sqlStr = " select top " & FPageSize*FCurrPage & " a.orderserial, a.customername, a.userid, a.title, a.writeuser, t.username, a.finishuser, r.refundresult, a.regdate, a.finishdate, a.contents_jupsu "
			End Select

			sqlStr = sqlStr & " from "
			sqlStr = sqlStr & " 	[db_cs].[dbo].tbl_new_as_list a "
			sqlStr = sqlStr & " 	join [db_cs].[dbo].tbl_as_refund_info r "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		a.id = r.asid "
			sqlStr = sqlStr & " 	left join [db_partner].[dbo].[tbl_user_tenbyten] t "
			sqlStr = sqlStr & " 	on "
			sqlStr = sqlStr & " 		t.userid = a.writeuser "
			sqlStr = sqlStr & " where "
			sqlStr = sqlStr & " 	1 = 1 "
			sqlStr = sqlStr & " 	and a.currstate = 'B007' "
			sqlStr = sqlStr & " 	and a.deleteyn = 'N' "
			sqlStr = sqlStr & " 	and r.isCSServiceRefund = 'Y' "

			If (FRectWriteUser <> "") Then
				sqlStr = sqlStr & " 	and a.writeuser = '" & FRectWriteUser & "' "
			End If

			sqlStr = sqlStr & " 	and a.finishdate >= '" & FRectStartDate & "' "
			sqlStr = sqlStr & " 	and a.finishdate <= '" & FRectEndDate & " 23:59:59' "

			Select Case FRectGrpBy
				Case "writeuser"
					sqlStr = sqlStr & " group by "
					sqlStr = sqlStr & " 	a.writeuser, t.username, convert(varchar(7),a.finishdate,121) "
					sqlStr = sqlStr & " order by "
					sqlStr = sqlStr & " 	convert(varchar(7),a.finishdate,121) desc, t.username, a.writeuser "
				Case "title"
					sqlStr = sqlStr & " group by "
					sqlStr = sqlStr & " 	a.title, convert(varchar(7),a.finishdate,121) "
					sqlStr = sqlStr & " order by "
					sqlStr = sqlStr & " 	convert(varchar(7),a.finishdate,121) desc, a.title "
				Case Else
					sqlStr = sqlStr & " order by "
					sqlStr = sqlStr & " 	a.finishdate desc "
			End Select

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
        			set FItemList(i) = new CSMileageItem

        			FItemList(i).Forderserial         	= rsget("orderserial")
					FItemList(i).Fcustomername          = rsget("customername")
					FItemList(i).Fuserid                = rsget("userid")
					FItemList(i).Ftitle                 = rsget("title")
					FItemList(i).Fwriteuser             = rsget("writeuser")
					FItemList(i).Fusername             	= rsget("username")
					FItemList(i).Ffinishuser            = rsget("finishuser")
					FItemList(i).Frefundresult          = rsget("refundresult")
					FItemList(i).Fregdate               = rsget("regdate")
					FItemList(i).Ffinishdate            = rsget("finishdate")
                    FItemList(i).Fcontents_jupsu        = rsget("contents_jupsu")

                    i=i+1
    				rsget.moveNext
    			loop
    		end if

    		rsget.Close
		end Sub


        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
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
