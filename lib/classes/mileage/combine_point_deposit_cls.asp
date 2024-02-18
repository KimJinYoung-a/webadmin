<%
'###########################################################
' Description :  온라인 오프라인 마일리지 & 예치금 통합관리
' History : 2013.11.12 한용민 생성
'###########################################################
function drawIFRS15_MonthData(byval grp1, byval iselstr, byVal iboxname, byVal iaddstr)
	Dim sqlStr, i
	Dim grpGubunSsnName : grpGubunSsnName = "IFRS_grp11_"
	Dim ArrRows : ArrRows = session(grpGubunSsnName&grp1)
	Dim ret
	IF isArray(ArrRows) then
		'rw "IFRS_grp1_"&grp1
	else
		sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_IFRS15_Get_grp1List] '"&grp1&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		i = 0
		If not db3_rsget.EOF Then
			ArrRows = db3_rsget.getRows()
			session(grpGubunSsnName&grp1) = ArrRows
		End If
		db3_rsget.Close
	End If

	ret = "<select name='"&iboxname&"' id='"&iboxname&"' "&iaddstr&">"
    ret = ret&"<option value='' "&CHKIIF(iselstr="","selected","")&">선택</option>"
    if isArray(ArrRows) then
        For i=0 To UBound(ArrRows,2)
            if (CStr(iselstr)=CStr(ArrRows(0,i))) then
                ret = ret&"<option value='"&ArrRows(0,i)&"' selected >"&fnChanageName(ArrRows(0,i))&"</option>"
            else
                ret = ret&"<option value='"&ArrRows(0,i)&"'>"&fnChanageName(ArrRows(0,i))&"</option>"
            end if
		Next
    end if
    ret = ret&"</select>"

	drawIFRS15_MonthData = ret
end function

class cpointDepositDtl_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FtargetGbn
	public Forderserial
	public FsubOrderserial
	public FDtlDesc
	public Fyyyymmdd
	public Fuserid
	public FiPoint

	public Fcancelyn
	public Fipkumdiv
	public Fbuyname
	public Fcanceldate
    public Fidx
	public FregUserid

    Public function GetYYYYMMDD()
        dim yyyy, mm, dd

        if IsNull(Forderserial) then
            GetYYYYMMDD = Left(Fyyyymmdd, 10)
            Exit Function
        end if

        yyyy = "20" & Left(Forderserial, 2)
        mm = Right(Left(Forderserial, 4), 2)
        dd = Right(Left(Forderserial, 6), 2)
        GetYYYYMMDD = yyyy & "-" & mm & "-" & dd
    end function

end class

class ccombine_point_deposit_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fyyyymm
	public fsrcGbn
	public ftargetGbn
	public fGbnCd
	public faccpointsum
	public fpointsum
	public fsrcGbnname
	public ftargetGbnname
	public fGbnCdname
	public fORD
	public fRTD
	public fSFT
	public fSPO
	public fGNI
	public fGNC
	public fSPE
	public fGNE	'이벤트 마일리지 적립(구분없음)
	public fGOE	'이벤트 마일리지 구매 적립
	public fGPE	'이벤트 마일리지 프로모션 적립
	public fETC
    public fXPR
end class

class ccombine_point_deposit
	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FTotalSum
	public FScrollCount
	public FTotalPage

	public FRectsrcGbn
	public FRecttargetGbn
	public frectGbnCd
	public FRectStartdate
	public FRectEndDate
    public FRectYYYYMM

	public FRecttargetSub

	public function getIFRS15_MonthData(byref colRows)
		Dim sqlStr, i
        sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_IFRS15_getMonthly] '"&FRectStartdate&"','"&FRectEndDate&"','"&FRecttargetGbn&"','"&FRecttargetSub&"'"

		db3_rsget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		FResultCount = db3_rsget.RecordCount
		Redim preserve FItemList(FResultCount)
		i = 0
		If not db3_rsget.EOF Then
			colRows = Array()
			For Each fld In db3_rsget.Fields
				reDim Preserve colRows(UBound(colRows) + 1)
				colRows(UBound(colRows))=fld.Name
			Next

			getIFRS15_MonthData = db3_rsget.getRows()

		End If
		db3_rsget.Close
	end function

	'/admin/maechul/managementsupport/combine_point_deposit_month.asp
	public function fcombine_point_deposit_month
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectsrcGbn<>"" then
			sqlsearch = sqlsearch + " and s.srcGbn='" & FRectsrcGbn & "'"
		end if
		if FRecttargetGbn<>"" then
			if FRecttargetGbn="ONAC" then
				sqlsearch = sqlsearch + " and s.targetGbn in ('ON','AC')"
			else
				sqlsearch = sqlsearch + " and s.targetGbn='"&FRecttargetGbn&"'"
			end if
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " A.*"
        sql = sql & " ,(select sum(pointsum)"
        sql = sql & "  from db_summary.dbo.tbl_monthly_PointDeposit_summary "
        sql = sql & " where Gbncd not in ('GOE','GPE')"		'GNE 분리데이터 제외 합계(중복됨)
        sql = sql & " and yyyymm<=A.yyyymm"
        if FRectsrcGbn<>"" then
			sql = sql + " and srcGbn='" & FRectsrcGbn & "'"
		end if
		if FRecttargetGbn<>"" then
			if FRecttargetGbn="ONAC" then
				sql = sql + " and targetGbn in ('ON','AC')"
			else
				sql = sql + " and targetGbn='"&FRecttargetGbn&"'"
			end if
		end if
        sql = sql & " ) as accpointsum"
        sql = sql & " from ("
		sql = sql & " select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " s.yyyymm"
		sql = sql & " ,sum(CASE WHEN s.Gbncd not in ('GOE','GPE') THEN s.pointsum else 0 END) as pointsum"		'GNE 분리데이터 제외 합계(중복됨)
		sql = sql & " ,sum(CASE WHEN s.Gbncd='ORD' THEN s.pointsum else 0 END) as ORD"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='RTD' THEN s.pointsum else 0 END) as RTD"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='SFT' THEN s.pointsum else 0 END) as SFT"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='SPO' THEN s.pointsum else 0 END) as SPO"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='GNI' THEN s.pointsum else 0 END) as GNI"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='GNC' THEN s.pointsum else 0 END) as GNC"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='SPE' THEN s.pointsum else 0 END) as SPE"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='GNE' THEN s.pointsum else 0 END) as GNE"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='GOE' THEN s.pointsum else 0 END) as GOE"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='GPE' THEN s.pointsum else 0 END) as GPE"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='ETC' THEN s.pointsum else 0 END) as ETC"
		sql = sql & " ,sum(CASE WHEN s.Gbncd='XPR' THEN s.pointsum else 0 END) as XPR"
		sql = sql & " from db_summary.dbo.tbl_monthly_PointDeposit_summary s"
		sql = sql & " where 1=1 " & sqlsearch
		sql = sql & " group by s.yyyymm"
		sql = sql & " ) A"
		sql = sql & " order by A.yyyymm desc"

		''response.write sql & "<Br>"
		''response.end

		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new ccombine_point_deposit_oneitem

				FItemList(i).fyyyymm			= rsget("yyyymm")
				FItemList(i).fORD			= rsget("ORD")
				FItemList(i).fRTD			= rsget("RTD")
				FItemList(i).fSFT			= rsget("SFT")
				FItemList(i).fSPO			= rsget("SPO")
				FItemList(i).fGNI			= rsget("GNI")
				FItemList(i).fGNC			= rsget("GNC")
				FItemList(i).fSPE			= rsget("SPE")
				FItemList(i).fGNE			= rsget("GNE")
				FItemList(i).fGOE			= rsget("GOE")
				FItemList(i).fGPE			= rsget("GPE")
				FItemList(i).fETC			= rsget("ETC")
                FItemList(i).fXPR			= rsget("XPR")
				FItemList(i).fpointsum      = rsget("pointsum")

                FItemList(i).faccpointsum   = rsget("accpointsum")
				'LogMiletotalprice
				'LogTotalmileage
				rsget.movenext
				i = i + 1
			Loop
		End If

		rsget.close
	end function

    public function fcombine_point_deposit_Detail_list()
        dim i , sqlstr
        sqlStr ="db_summary.[dbo].[sp_Ten_monthly_PointDeposit_summary_GIFT_DetailCNT_New]('"&FRectsrcGbn&"','"&FRectYYYYMM&"','"&FRecttargetGbn&"','"&frectGbnCd&"')"
      '' rw sqlStr
       'response.end

		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		''sqlStr ="[db_jungsan].[dbo].[sp_Ten_getMonthJungsanAdmCnt]('"&FRectYYYYMM&"','"&FRectMakerid&"','"&FRectJGubun&"','"&FRecttargetGbn&"','"&FRectgroupid&"','"&FRectTaxType&"',"&FRectFinishFlag&")"
		''rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotalCount = rsget("CNT")
			FTotalSum   = rsget("iPoint")
		END IF
		rsget.close

		IF FTotalCount > 0 THEN


            sqlstr = "db_summary.[dbo].[sp_Ten_monthly_PointDeposit_summary_GIFT_DetailList_New]('"&FRectsrcGbn&"','"&FRectYYYYMM&"','"&FRecttargetGbn&"','"&frectGbnCd&"',"&FPageSize&","&FCurrPage&")"
            ''rw sqlstr
            ''response.end
            rsget.CursorLocation = adUseClient
    	    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

    		FResultCount = rsget.RecordCount
    		FtotalPage =  CInt(FTotalCount\FPageSize)
    		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
    			FtotalPage = FtotalPage +1
    		end if

    		if (FResultCount<1) then FResultCount=0

    		redim preserve FItemList(FResultCount)
    		i=0
    		if  not rsget.EOF  then
    			do until rsget.EOF
    				set FItemList(i) = new cpointDepositDtl_oneitem
    				FItemList(i).FtargetGbn         = rsget("targetGbn")
                    FItemList(i).Forderserial       = rsget("orderserial")
                    FItemList(i).FsubOrderserial    = rsget("subOrderserial")
                    FItemList(i).FDtlDesc           = rsget("DtlDesc")
                    FItemList(i).Fyyyymmdd          = rsget("yyyymmdd")
                    FItemList(i).Fuserid            = rsget("userid")
                    FItemList(i).FiPoint            = rsget("iPoint")

                    FItemList(i).Fcancelyn          = rsget("cancelyn")
                    FItemList(i).Fipkumdiv          = rsget("ipkumdiv")
                    FItemList(i).Fbuyname           = rsget("buyname")
                    FItemList(i).Fcanceldate        = rsget("canceldate")
                    FItemList(i).Fidx        		= rsget("idx")
					if FRectsrcGbn="G" and frectGbnCd="GNE" then
						FItemList(i).FregUserid        	= rsget("regUserid")
					end if

                    if IsNull(FItemList(i).Fyyyymmdd) then
                        FItemList(i).Fyyyymmdd = ""
                    end if

    				rsget.movenext
    				i=i+1
    			loop
    		end if
    		rsget.Close
    	END IF
    end function

	'/admin/maechul/managementsupport/combine_point_deposit_list.asp
	public function fcombine_point_deposit_list
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and s.yyyymm <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectsrcGbn<>"" then
			sqlsearch = sqlsearch + " and s.srcGbn='" & FRectsrcGbn & "'"
		end if
		if FRecttargetGbn<>"" then
			if FRecttargetGbn="ONAC" then
				sqlsearch = sqlsearch + " and s.targetGbn in ('ON','AC')"
			else
				sqlsearch = sqlsearch + " and s.targetGbn='"&FRecttargetGbn&"'"
			end if
		end if
		if frectGbnCd<>"" then
			sqlsearch = sqlsearch + " and s.GbnCd='" & frectGbnCd & "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " yyyymm, srcGbn, targetGbn, GbnCd, isnull(pointsum,0) as pointsum"
		sql = sql & " ,(select top 1 codename"
		sql = sql & " 		from db_shop.dbo.tbl_offshop_commoncode"
		sql = sql & " 		where useyn='Y' and codekind='srcGbn' and codegroup='MAIN' and codeid=s.srcGbn) as srcGbnname"
		sql = sql & " ,(select top 1 codename"
		sql = sql & " 		from db_shop.dbo.tbl_offshop_commoncode"
		sql = sql & " 		where useyn='Y' and codekind='targetGbn' and codegroup='MAIN' and codeid=s.targetGbn) as targetGbnname"
		sql = sql & " ,(select top 1 codename"
		sql = sql & " 		from db_shop.dbo.tbl_offshop_commoncode"
		sql = sql & " 		where useyn='Y' and codekind='GbnCd' and codegroup='MAIN' and codeid=s.GbnCd) as GbnCdname"
		sql = sql & " from db_summary.dbo.tbl_monthly_PointDeposit_summary s"
		sql = sql & " where 1=1 " & sqlsearch
		sql = sql & " order by yyyymm desc, srcGbn asc, targetGbn asc, GbnCd asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new ccombine_point_deposit_oneitem

				FItemList(i).fyyyymm			= rsget("yyyymm")
				FItemList(i).fsrcGbn			= rsget("srcGbn")
				FItemList(i).ftargetGbn			= rsget("targetGbn")
				FItemList(i).fGbnCd			= rsget("GbnCd")
				FItemList(i).fpointsum			= rsget("pointsum")
				FItemList(i).fsrcGbnname			= db2html(rsget("srcGbnname"))
				FItemList(i).ftargetGbnname			= db2html(rsget("targetGbnname"))
				FItemList(i).fGbnCdname			= db2html(rsget("GbnCdname"))

				rsget.movenext
				i = i + 1
			Loop
		End If

		rsget.close
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

'값치환 함수
function fnChanageName(byVal val)
	if not(isNull(val) or val="") then
		fnChanageName = replace(val,"사용마일(출고)","사용마일(정산확정)")
	end if
end function
%>
