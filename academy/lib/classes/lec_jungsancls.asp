<%
class CJungsanMasterItem
	public Fid
	public Fdesignerid
	public Fgroupid
	public Fyyyymm
	public Ftitle
	public Fub_cnt
	public Fub_totalsellcash
	public Fub_totalsuplycash
	public Fub_comment
	public Fme_cnt
	public Fme_totalsellcash
	public Fme_totalsuplycash
	public Fme_comment
	public Fwi_cnt
	public Fwi_totalsellcash
	public Fwi_totalsuplycash
	public Fwi_comment
	public Fet_cnt
	public Fet_totalsellcash
	public Fet_totalsuplycash
	public Fet_comment
	public Fsh_cnt
	public Fsh_totalsellcash
	public Fsh_totalsuplycash
	public Fsh_comment

	public Fregdate
	public Fcancelyn
	public Ffinishflag
	public Fipkumdate
	public Ftaxregdate
	public Fbigo

	public FDesignerEmail

	public Fjungsan_bank
	public Fjungsan_date
	public Fjungsan_acctno
	public Fjungsan_acctname
	public Fcompany_name
	public Fjungsan_gubun

	public Ftaxinputdate
	public Fcompany_no
	public Fceoname
	public Fcompany_address
	public Fcompany_address2

	public Fdifferencekey
	public Ftaxtype
	public FTaxLinkidx
	public Fneotaxno

	public FFixsegumil
	public Fbankingupflag

	public function getDbDate()
		dim sqlstr
		sqlstr = " select convert(varchar(10),getdate(),21) as nowdate "
		rsget.Open sqlStr,dbget,1
		getDbDate = CDate(rsget("nowdate"))
		rsget.Close
	end function

	public function GetNormalTaxDate()
		if Not(IsNULL(FFixsegumil)) and (FFixsegumil<>"") then
			GetNormalTaxDate = FFixsegumil
		else
			''if Fjungsan_date="말일" then
				GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-1)
			''else
			''	GetNormalTaxDate = dateserial(Left(Fyyyymm,4),Right(Fyyyymm,2)+1,1-2)
			''end if
		end if
	end function

	public function GetPreFixSegumil()
		dim thisdate, maytaxdate
		dim ithis1day , ithis21day, premonth1day, premonth21day

		thisdate = getDbDate()
		maytaxdate = GetNormalTaxDate()
        
        '' 12일까지 마감할 경우 13으로 세팅
		premonth1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"01")
		premonth21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2)-1,"13")
		ithis1day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"01")
		ithis21day = dateserial(Left(thisdate,4),Mid(thisdate,6,2),"13")

		if (thisdate>=ithis21day) then
			GetPreFixSegumil = ithis1day
		elseif (maytaxdate<premonth21day)  then
			GetPreFixSegumil = premonth1day
		else
			GetPreFixSegumil = maytaxdate
		end if
	end function




	public function IsElecTaxExists()
		IsElecTaxExists = Not(IsNULL(FTaxLinkidx) or (FTaxLinkidx="")) and (Ffinishflag>=3)
	end function


	''//세금계산서
	public function IsElecTaxCase()
		IsElecTaxCase = (Ftaxtype="01") and (Fjungsan_gubun="일반과세") and (Ffinishflag<3)
	end function


	''//계산서
	public function IsElecFreeTaxCase()
		IsElecFreeTaxCase = (Ftaxtype="02") 'and (Fjungsan_gubun="면세")
	end function


	''//간이, 원천, 기타
	public function IsElecSimpleBillCase()
		IsElecSimpleBillCase = (Ftaxtype="03") and (Ffinishflag<3)
	end function

	public function GetSimpleTaxtypeName()
		if Ftaxtype="01" then
			GetSimpleTaxtypeName = "과세"
		elseif Ftaxtype="02" then
			GetSimpleTaxtypeName = "면세"
		elseif Ftaxtype="03" then
			GetSimpleTaxtypeName = "간이"
		end if
	end function

	public function GetTaxtypeNameColor()
		if Ftaxtype="01" then
			GetTaxtypeNameColor = "#000000"
		elseif Ftaxtype="02" then
			GetTaxtypeNameColor = "#FF3333"
		elseif Ftaxtype="03" then
			GetTaxtypeNameColor = "#3333FF"
		end if
	end function

	public function GetTotalSellcash()
		GetTotalSellcash = Fub_totalsellcash + Fme_totalsellcash + Fwi_totalsellcash + Fet_totalsellcash + Fsh_totalsellcash
	end function

	public function GetTotalSuplycash()
		GetTotalSuplycash = Fub_totalsuplycash + Fme_totalsuplycash + Fwi_totalsuplycash + Fet_totalsuplycash + Fsh_totalsuplycash
	end function
	
	''원천징수대상자 정산금액
    public function GetTotalWithHoldingJungSanSum()
        dim ototalsum
        dim TreePercentTax
        ototalsum = GetTotalSuplycash
        TreePercentTax = Fix(Fix(ototalsum*0.03)/10)*10
        
        GetTotalWithHoldingJungSanSum = ototalsum
        
        ''3%세금이 1000원 이하이면 세금없음.
        if TreePercentTax<=1000 then Exit function
        
        GetTotalWithHoldingJungSanSum = ototalsum - TreePercentTax - Fix(Fix(TreePercentTax*0.1)/10)*10
        
	end function

	public function GetTotalTaxSuply()
		if Ftaxtype="01" then
			GetTotalTaxSuply = CLng(GetTotalSuplycash / 1.1)
		else
			GetTotalTaxSuply = GetTotalSuplycash
		end if
	end function

	public function GetTotalTaxVat()
		GetTotalTaxVat = GetTotalSuplycash - GetTotalTaxSuply
	end function

	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
			GetStateName = "업체확인대기"
		elseif Ffinishflag="2" then
			GetStateName = "업체확인완료"
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		else

		end if
	end function

	public function GetStateColor()
		if Ffinishflag="0" then
			GetStateColor = "#000000"
		elseif Ffinishflag="1" then
			GetStateColor = "#448888"
		elseif Ffinishflag="2" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="3" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="7" then
			GetStateColor = "#FF0000"
		else

		end if
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class

class CJungsanDetailItem
	public Fid
	public Fmasteridx
	public Fgubuncd
	public Fdetailidx
	public Fmastercode
	public Fbuyname
	public Freqname
	public Fitemid
	public Fitemoption
	public Fitemname
	public Fitemoptionname
	public Fitemno
	public Fsellcash
	public Fsuplycash

	public FOrgSellCash
	public FOrgSuplyCash

	public FExecDate
	public Fcomment

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CLecJungsanItem
	public FIdx
	public FOrderSerial
	public FItemId
	public FItemOption
	public FItemName
	public FItemOptionName
	public FItemNo
	public FBuyCash
	public FSellCash

	public FCurrState
	public FBeasongDate
	public FUpcheSongjangNo
	public FRegDate
	public FIpkumDate
	public FBuyName
	public FJumunDiv

	public FIpkumDiv
	public FMWDiv

	public Flec_date

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub

end Class


Class CJungsanSummaryByTaxDateItem
    public Ftaxregdate        
    public Fjungsansum_susi   
    public Fjungsansum_31date  
    public Fjungsansum_15date   
    public Fjungsansum_etcdate
    public Fewol_jungsansum   
    public Fnext_jungsansum 
                       
    public Ffixedsum          
    public Fipkumsum          
                       
    public Ftot_jungsanprice  

    Private Sub Class_Initialize()
        Ftaxregdate        = 0
        Fjungsansum_susi   = 0
        Fjungsansum_31date = 0
        Fjungsansum_15date = 0
        Fjungsansum_etcdate= 0
        Fewol_jungsansum   = 0
        Fnext_jungsansum   = 0
                    
        Ffixedsum          = 0
        Fipkumsum          = 0
                    
        Ftot_jungsanprice  = 0
	End Sub

	Private Sub Class_Terminate()

    End Sub
end Class

class CJungsanSumaryItem
	public Fyyyymm
	public Ftot

	''매입가
	public Fuptot
	public Fmetot
	public Fwitot
	public Fshtot
	public Fettot

	''판매가
	public Fupselltot
	public Fmeselltot
	public Fwiselltot
	public Fshselltot
	public Fetselltot

	public Ffinishflag
	public Fjungsan_date

	public Ftotflag_notconfirmsum
	public Ftotflag_confirmsum
	public Ftotflag_ipkumsum
    
    public Ffixedthissum
    public Ffixednextsum
    
	public function GetStateName()
		if Ffinishflag="0" then
			GetStateName = "수정중"
		elseif Ffinishflag="1" then
			GetStateName = "업체확인대기"
		elseif Ffinishflag="2" then
			GetStateName = "업체확인완료"
		elseif Ffinishflag="3" then
			GetStateName = "정산확정"
		elseif Ffinishflag="7" then
			GetStateName = "입금완료"
		else

		end if
	end function

	public function GetStateColor()
		if Ffinishflag="0" then
			GetStateColor = "#000000"
		elseif Ffinishflag="1" then
			GetStateColor = "#448888"
		elseif Ffinishflag="2" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="3" then
			GetStateColor = "#0000FF"
		elseif Ffinishflag="7" then
			GetStateColor = "#FF0000"
		else

		end if
	end function

	public function getTotSum()
		getTotSum = Fuptot + Fmetot + Fwitot + Fshtot + Fettot
	end function

	public function getTotSellcashSum()
		getTotSellcashSum = Fupselltot + Fmeselltot + Fwiselltot + Fshselltot + Fetselltot
	end function

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CLecJungsan
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
    
    public FTotalSum
    
	public FRectid
	public FRectGubun
	public FRectDesigner
	public FRectMastercodes
	public FRectOrder
	public FRectYYYYMM
	public FRectPreYYYYMM
	public FRectStartDay
	public FRectEndDay

	public FWitakInsserted
	public FRectDesignerViewOnly

	public FCurrmastercode
	public FPremastercode
	public FRectState
	public FRectIpkumilNot

	public FRectNotIncludeWonChon
	public FRectOnlyIncludeWonChon
	public FRectOnlyIncludeNoTax
	public FRectOnlyIncludeSimpleTax

	public FRectDifferencekey

	public FRectOnlyElecTax
	public FRectOnlyNotElecTax

	public FRectBankingupflag
	public FRectNotYYYYMM
    
    public FRectStartYYYYMM
    public FRectEndYYYYMM
    
    public FRectFixStateExiste
    public FRectfinishflag
    public FRectTaxRegDate
    public FRectJungsanDate
    
    public FRectTaxType
    public FRectTaxDate
	public FRectLectureID
    
	public function LecJungsanDetailListSum()
		dim sqlStr,i
		sqlStr = "select T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.itemno, T.sellcash, T.suplycash,"
		sqlStr = sqlStr + " i.sellcash as orgsellcash, i.buycash as orgsuplycash"
		sqlStr = sqlStr + " from ("
			sqlStr = sqlStr + "select d.itemid,d.itemoption,d.itemname,d.itemoptionname,sum(d.itemno) as itemno,d.sellcash,"
			sqlStr = sqlStr + " d.suplycash"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail d "
			sqlStr = sqlStr + " where d.masteridx=" + CStr(FRectid)

			if FRectgubun<>"" then
				sqlStr = sqlStr + " and gubuncd='" + FRectgubun + "'"
			end if
			sqlStr = sqlStr + " group by d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.sellcash,d.suplycash"
		sqlStr = sqlStr + " ) as T"
		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on T.itemid=i.itemid"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = rsget("itemname")
				FItemList(i).Fitemoptionname= rsget("itemoptionname")
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FOrgsellcash      = rsget("orgsellcash")
				FItemList(i).FOrgsuplycash  	= rsget("orgsuplycash")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function


	public function LecJungsanDetailList()
		dim sqlStr,i
		if (FRectgubun="upche") or (FRectgubun="witaksell") then
			sqlStr = "select distinct j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash"
			sqlStr = sqlStr + ", convert(varchar(10),d.beasongdate,21) as execdate"
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d on d.idx=j.detailidx"
		elseif FRectgubun="maeip" then
			sqlStr = "select distinct j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash, convert(varchar(10),d.executedt,21) as execdate "
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master d on d.code=j.mastercode"
		else
			sqlStr = "select distinct j.id,j.masteridx,j.gubuncd,j.detailidx,j.mastercode,j.buyname,j.reqname,"
			sqlStr = sqlStr + "j.itemid,j.itemoption,j.itemname,j.itemoptionname,j.itemno,j.sellcash,"
			sqlStr = sqlStr + "j.suplycash, convert(varchar(10),d.executedt,21) as execdate "
			sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_detail j"

			sqlStr = sqlStr + " left join "
			sqlStr = sqlStr + " [db_storage].[dbo].tbl_acount_storage_master d on d.code=j.mastercode"
		end if
		sqlStr = sqlStr + " where j.masteridx=" + CStr(FRectid)

		if FRectgubun<>"" then
			sqlStr = sqlStr + " and j.gubuncd='" + FRectgubun + "'"
		end if

		if FRectOrder="itemid" then
			sqlStr = sqlStr + " order by j.itemid, j.itemoption, j.mastercode desc"
		else
			sqlStr = sqlStr + " order by j.mastercode"
		end if

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fid            = rsget("id")
				FItemList(i).Fmasteridx     = rsget("masteridx")
				FItemList(i).Fgubuncd       = rsget("gubuncd")
				FItemList(i).Fdetailidx     = rsget("detailidx")
				FItemList(i).Fmastercode    = rsget("mastercode")
				FItemList(i).Fbuyname       = db2html(rsget("buyname"))
				FItemList(i).Freqname       = db2html(rsget("reqname"))
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = db2html(rsget("itemname"))
				FItemList(i).Fitemoptionname= db2html(rsget("itemoptionname"))
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				FItemList(i).FExecDate      = rsget("execdate")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

	public function LecJungsanDetailListByYYYYMM()
		dim sqlStr,i

		sqlStr = "select d.id,d.masteridx,d.gubuncd,d.detailidx,d.mastercode,d.buyname,d.reqname,"
		sqlStr = sqlStr + "d.itemid,d.itemoption,d.itemname,d.itemoptionname,d.itemno,d.sellcash,"
		sqlStr = sqlStr + "d.suplycash"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
		sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " where m.designerid='" + CStr(FRectdesigner) + "'"
		sqlStr = sqlStr + " and m.yyyymm='" + CStr(FRectYYYYMM) + "'"
		sqlStr = sqlStr + " and m.id=d.masteridx"
		if FRectdifferencekey<>"" then
			sqlStr = sqlStr + " and m.differencekey=" + CStr(FRectdifferencekey)
		end if
		sqlStr = sqlStr + " and d.gubuncd='" + FRectgubun + "'"
		sqlStr = sqlStr + " order by d.mastercode desc, d.detailidx"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanDetailItem

				FItemList(i).Fid            = rsget("id")
				FItemList(i).Fmasteridx     = rsget("masteridx")
				FItemList(i).Fgubuncd       = rsget("gubuncd")
				FItemList(i).Fdetailidx     = rsget("detailidx")
				FItemList(i).Fmastercode    = rsget("mastercode")
				FItemList(i).Fbuyname       = rsget("buyname")
				FItemList(i).Freqname       = rsget("reqname")
				FItemList(i).Fitemid        = rsget("itemid")
				FItemList(i).Fitemoption    = rsget("itemoption")
				FItemList(i).Fitemname      = rsget("itemname")
				FItemList(i).Fitemoptionname= rsget("itemoptionname")
				FItemList(i).Fitemno        = rsget("itemno")
				FItemList(i).Fsellcash      = rsget("sellcash")
				FItemList(i).Fsuplycash  	= rsget("suplycash")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function

    public function LecJungsanSummaryBySegumDate()
        dim sqlStr,i
        ''taxregdate IsNULL = 원천징수 등.
        
        sqlStr = " select m.taxregdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='수시') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as jungsansum_susi," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='말일') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as jungsansum_31date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and (g.jungsan_date='15일') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as jungsansum_15date," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm=convert(varchar(7),m.taxregdate,21)) and ((g.jungsan_date is NULL) or (g.jungsan_date not in('수시','말일','15일'))) then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as jungsansum_etcdate," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.yyyymm<>convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as ewol_jungsansum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as fixedsum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='7') then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as ipkumsum," + VbCrlf
        sqlStr = sqlStr + " sum(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) as tot_jungsanprice" + VbCrlf
        sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m" + VbCrlf
        sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group g " + VbCrlf
        sqlStr = sqlStr + "     on m.groupid=g.groupid" + VbCrlf
        sqlStr = sqlStr + " where m.finishflag >=3" + VbCrlf

        if (FRectStartDay<>"") then
            sqlStr = sqlStr + " and m.taxregdate>='" + FRectStartDay + "'" + VbCrlf
        end if
        
        if (FRectEndDay<>"") then
            sqlStr = sqlStr + " and m.taxregdate<'" + FRectEndDay + "'" + VbCrlf
        end if
        
        sqlStr = sqlStr + " group by m.taxregdate" + VbCrlf
        sqlStr = sqlStr + " order by m.taxregdate desc " + VbCrlf

        
        rsget.Open sqlStr, dbget, 1
        
        FResultCount = rsget.RecordCount
        FTotalCount = FResultCount
        
        if FResultCount<1 then FResultCount=0
        
        redim preserve FItemList(FResultCount)
        
		if  not rsget.EOF  then
		    i = 0
		    rsget.absolutepage = FCurrPage
		    do until rsget.eof
		    
			set FItemList(i) = new CJungsanSummaryByTaxDateItem
			

            FItemList(i).Ftaxregdate         = rsget("taxregdate")
            FItemList(i).Fjungsansum_susi    = rsget("jungsansum_susi")
            FItemList(i).Fjungsansum_31date  = rsget("jungsansum_31date")
            FItemList(i).Fjungsansum_15date  = rsget("jungsansum_15date")
            FItemList(i).Fjungsansum_etcdate = rsget("jungsansum_etcdate")
            FItemList(i).Fewol_jungsansum    = rsget("ewol_jungsansum")
            
            FItemList(i).Ffixedsum          = rsget("fixedsum")
            FItemList(i).Fipkumsum          = rsget("ipkumsum")
            
            FItemList(i).Ftot_jungsanprice  = rsget("tot_jungsanprice")
            
            
			rsget.MoveNext
			i = i + 1
		loop
		
	    end if
	
        rsget.Close
        
    end function
    
	public function LecJungsanSummary0()
		dim sqlStr,i
		sqlStr = "select m.yyyymm,"
		sqlStr = sqlStr + " IsNull(p.jungsan_date,'') as jungsan_date, "
		sqlStr = sqlStr + " Sum(m.ub_totalsuplycash) as uptot, "
		sqlStr = sqlStr + " Sum(m.me_totalsuplycash) as metot, "
		sqlStr = sqlStr + " Sum(m.wi_totalsuplycash) as witot, "
		sqlStr = sqlStr + " Sum(m.sh_totalsuplycash) as shtot, "
		sqlStr = sqlStr + " Sum(m.et_totalsuplycash) as ettot, "
        
        sqlStr = sqlStr + " Sum(m.ub_totalsellcash) as upselltot, "
		sqlStr = sqlStr + " Sum(m.me_totalsellcash) as meselltot, "
		sqlStr = sqlStr + " Sum(m.wi_totalsellcash) as wiselltot, "
		sqlStr = sqlStr + " Sum(m.sh_totalsellcash) as shselltot, "
		sqlStr = sqlStr + " Sum(m.et_totalsellcash) as etselltot, "
		
		sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag='7') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_ipkumsum,"
      	sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag='3') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_confirmsum,"
      	sqlStr = sqlStr + " Sum(CASE "
        sqlStr = sqlStr + "     WHEN (m.finishflag <'3') THEN (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash)"
        sqlStr = sqlStr + "     ELSE 0"
      	sqlStr = sqlStr + "     END ) as totflag_notconfirmsum,"
      	
      	''정산일 기준으로 입금예정금액 산출.
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm=convert(varchar(7),taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as fixedthissum," + VbCrlf
        ''sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (m.yyyymm<>convert(varchar(7),taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as fixednextsum " + VbCrlf
        
        ''금월 기준으로 입금예정금액 산출. taxregdate IsNULL = 원천징수 등.
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as fixedthissum," + VbCrlf
        sqlStr = sqlStr + " sum(case when (m.finishflag='3') and (convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21))  then (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.sh_totalsuplycash + m.et_totalsuplycash) else 0 end) as fixednextsum" + VbCrlf
        
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where 1=1"
		
		if (FRectStartYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm>='" + FRectStartYYYYMM + "'" + VbCrlf
        end if
        
        if (FRectEndYYYYMM<>"") then
            sqlStr = sqlStr + " and m.yyyymm<='" + FRectEndYYYYMM + "'" + VbCrlf
        end if
        
		sqlStr = sqlStr + " group by m.yyyymm, p.jungsan_date"
		if FRectFixStateExiste<>"" then
		    ''미처리 내역이 있는것..
            sqlStr = sqlStr + " having sum(case when (m.finishflag<=3) then (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash) else 0 end)<>0"
		end if
		sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanSumaryItem

				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Fuptot       	   = rsget("uptot")
				FItemList(i).Fmetot            = rsget("metot")
				FItemList(i).Fwitot            = rsget("witot")
				FItemList(i).Fshtot            = rsget("shtot")
				FItemList(i).Fettot            = rsget("ettot")
				
				FItemList(i).Fupselltot       	   = rsget("upselltot")
				FItemList(i).Fmeselltot            = rsget("meselltot")
				FItemList(i).Fwiselltot            = rsget("wiselltot")
				FItemList(i).Fshselltot            = rsget("shselltot")
				FItemList(i).Fetselltot            = rsget("etselltot")

				FItemList(i).Ftotflag_notconfirmsum  = rsget("totflag_notconfirmsum")
				FItemList(i).Ftotflag_confirmsum     = rsget("totflag_confirmsum")
				FItemList(i).Ftotflag_ipkumsum       = rsget("totflag_ipkumsum")

                FItemList(i).Ffixedthissum      = rsget("fixedthissum")
                FItemList(i).Ffixednextsum      = rsget("fixednextsum")
                
				FItemList(i).Fjungsan_date     = rsget("jungsan_date")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function

	public function LecJungsanSummary()
		dim sqlStr,i
		sqlStr = "select m.yyyymm, sum(IsNull(m.ub_totalsuplycash,0)) as uptot,"
		sqlStr = sqlStr + " sum(IsNull(m.me_totalsuplycash,0)) as metot,"
		sqlStr = sqlStr + " sum(IsNull(m.wi_totalsuplycash,0)) as witot, "
		sqlStr = sqlStr + " sum(IsNull(m.sh_totalsuplycash,0)) as shtot, "
		sqlStr = sqlStr + " sum(IsNull(m.et_totalsuplycash,0)) as ettot, m.finishflag,"
		sqlStr = sqlStr + " IsNull(p.jungsan_date,'') as jungsan_date"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where m.cancelyn='N'"
		sqlStr = sqlStr + " and m.finishflag<7"
		sqlStr = sqlStr + " group by m.yyyymm, m.finishflag, p.jungsan_date"
		sqlStr = sqlStr + " order by m.yyyymm desc, p.jungsan_date, m.finishflag"

		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanSumaryItem

				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Fuptot       	   = rsget("uptot")
				FItemList(i).Fmetot            = rsget("metot")
				FItemList(i).Fwitot            = rsget("witot")
				FItemList(i).Fshtot            = rsget("shtot")
				FItemList(i).Fettot            = rsget("ettot")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fjungsan_date     = rsget("jungsan_date")
				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end function
    
    public function LecJungsanFixedList()
		dim sqlStr,i
		sqlStr = "select m.id, m.designerid, m.groupid, m.yyyymm, m.title, m.ub_cnt,"
		sqlStr = sqlStr + " m.ub_totalsellcash,"
		sqlStr = sqlStr + " m.ub_totalsuplycash,"
		sqlStr = sqlStr + " m.ub_comment,"
		sqlStr = sqlStr + " m.me_cnt, m.me_totalsellcash,"
		sqlStr = sqlStr + " m.me_totalsuplycash, m.me_comment,"
		sqlStr = sqlStr + " m.wi_cnt, m.wi_totalsellcash,"
		sqlStr = sqlStr + " m.wi_totalsuplycash, m.wi_comment,"
		sqlStr = sqlStr + " m.et_cnt, m.et_totalsellcash,"
		sqlStr = sqlStr + " m.et_totalsuplycash, m.et_comment,"
		sqlStr = sqlStr + " m.sh_cnt, m.sh_totalsellcash,"
		sqlStr = sqlStr + " m.sh_totalsuplycash, m.sh_comment,"

		sqlStr = sqlStr + " m.regdate,m.cancelyn,m.finishflag,convert(varchar(10),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + " convert(varchar(10),m.taxregdate,20) as taxregdate, m.bigo, "
		sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no,p.ceoname,p.company_address,p.company_address2,"
		sqlStr = sqlStr + " m.taxinputdate, m.differencekey, m.taxtype, m.taxlinkidx, m.neotaxno, m.bankingupflag"
		sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + "     left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		
		if FRectfinishflag="ALL" then
		    sqlStr = sqlStr + " where m.finishflag>=3"
		elseif FRectfinishflag<>"" then
		    sqlStr = sqlStr + " where m.finishflag='" + FRectfinishflag + "'"
		else
		    sqlStr = sqlStr + " where m.finishflag='3'"
        end if
        
        if (FRectTaxRegDate<>"") then
            sqlStr = sqlStr + " and m.taxregdate='" + FRectTaxRegDate + "'"
        end if
        
        '' AA 전월 정산내역 중 발행일이 전월 & 정산일 수시/15일
        '' BB 전월 정산내역 중 발행일이 전월 & 정산일 말일    
        '' CC 전전월 이하 정산내역 중 발행일이 전월           
        '' DD 발행일이 현재월 이상       
        '' EE 정상발행 전체   
        '' FF 이월발행 전체 (비정상발행)                            
        '' ZZ 발행일이 빈값이거나, 그 외 날짜               
        if FRectGubun="ZZ" then
            sqlStr = sqlStr + " and m.taxregdate is NULL"
        elseif FRectGubun="AA" then
            sqlStr = sqlStr + " and (IsNULL(p.jungsan_date,'')='' or p.jungsan_date<>'말일')"
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="BB" then
            sqlStr = sqlStr + " and p.jungsan_date='말일'"
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="CC" then
            sqlStr = sqlStr + " and m.yyyymm<convert(varchar(7),m.taxregdate,21)"
            sqlStr = sqlStr + " and convert(varchar(7),getdate(),21)>convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="DD" then
            sqlStr = sqlStr + " and convert(varchar(7),getdate(),21)<=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="EE" then
            sqlStr = sqlStr + " and m.yyyymm=convert(varchar(7),m.taxregdate,21)"
        elseif FRectGubun="FF" then
            sqlStr = sqlStr + " and m.yyyymm<>convert(varchar(7),m.taxregdate,21)"
        end if
        
        if FRectJungsanDate="NULL" then
            sqlStr = sqlStr + " and IsNULL(p.jungsan_date,'')=''"
        elseif FRectJungsanDate<>"" then
            sqlStr = sqlStr + " and p.jungsan_date='" + FRectJungsanDate + "'"
        end if
        
        if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if
		
		if FRectbankingupflag<>"" then
		    sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if 
		
		if FRectYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if
		
		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if
		
		
		
        sqlStr = sqlStr + " order by m.neotaxno, m.taxinputdate"
        
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
        
        if FResultCount<1 then FResultCount=0
        
		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanMasterItem

				FItemList(i).Fid               = rsget("id")
				FItemList(i).Fdesignerid       = rsget("designerid")
				FItemList(i).Fgroupid		   = rsget("groupid")
				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Ftitle            = rsget("title")
				FItemList(i).Fub_cnt           = rsget("ub_cnt")
				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
				FItemList(i).Fme_cnt           = rsget("me_cnt")
				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))

				FItemList(i).Fet_cnt           = rsget("et_cnt")
				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))


				FItemList(i).Fregdate          = rsget("regdate")
				FItemList(i).Fcancelyn         = rsget("cancelyn")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fipkumdate        = rsget("ipkumdate")
				FItemList(i).Ftaxregdate       = rsget("taxregdate")
				FItemList(i).Fbigo			   = db2html(rsget("bigo"))
				FItemList(i).FDesignerEmail		= rsget("jungsan_email")

				FItemList(i).Fjungsan_bank		= rsget("jungsan_bank")
				FItemList(i).Fjungsan_date		= rsget("jungsan_date")
				FItemList(i).Fjungsan_acctno		= rsget("jungsan_acctno")
				FItemList(i).Fjungsan_acctname		= rsget("jungsan_acctname")
				FItemList(i).Fcompany_name		= db2html(rsget("company_name"))

				FItemList(i).Fjungsan_gubun		= db2html(rsget("jungsan_gubun"))
				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
				FItemList(i).Fceoname	= db2html(rsget("ceoname"))
				FItemList(i).Fcompany_address	= db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2	= db2html(rsget("company_address2"))

				FItemList(i).Fdifferencekey = rsget("differencekey")
				FItemList(i).Ftaxtype = rsget("taxtype")
				FItemList(i).FTaxLinkidx = rsget("taxlinkidx")
				FItemList(i).Fneotaxno = rsget("neotaxno")

				FItemList(i).Fbankingupflag = rsget("bankingupflag")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

    end function


	public function LecJungsanMasterList()
		dim sqlStr,i
		
		sqlStr = "select count(m.id) as cnt,"
		sqlStr = sqlStr + " sum(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash) as ttlsuplycash"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c,"
		sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where m.id<>0"
        sqlStr = sqlStr + " and m.designerid=c.userid "
        sqlStr = sqlStr + " and c.userdiv='14'"
        
        ''sqlStr = sqlStr + " and ((c.userdiv='14') or (c.catecode in ('94','95')))"
        
        if (FRectDesigner="") And (FRectLectureID="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end If

		if FRectLectureID<>"" then
			sqlStr = sqlStr + " and c.userid='" + FRectLectureID + "'"
		end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		if FRectOnlyIncludeSimpleTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='간이과세'"
		end if

		if FRectOnlyElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is Not NULL"
		end if

		if FRectOnlyNotElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is NULL"
		end if

		if FRectBankingupflag<>"" then
			sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if
		
		if (FRectTaxType<>"") then
		    sqlStr = sqlStr + " and m.taxtype='" + FRectTaxType + "'"
		end if
		
		if (FRectTaxDate<>"") then
		    sqlStr = sqlStr + " and m.taxregdate='" + FRectTaxDate + "'"
		end if
		
		rsget.Open sqlStr,dbget,1
		    FTotalCount = rsget("cnt")
		    FTotalSum   = rsget("ttlsuplycash")
		rsget.Close
		
		IF IsNULL(FTotalSum) then FTotalSum=0
		
		sqlStr = "select m.id, m.designerid, m.groupid, m.yyyymm, m.title, m.ub_cnt,"
		sqlStr = sqlStr + " m.ub_totalsellcash,"
		sqlStr = sqlStr + " m.ub_totalsuplycash,"
		sqlStr = sqlStr + " m.ub_comment,"
		sqlStr = sqlStr + " m.me_cnt, m.me_totalsellcash,"
		sqlStr = sqlStr + " m.me_totalsuplycash, m.me_comment,"
		sqlStr = sqlStr + " m.wi_cnt, m.wi_totalsellcash,"
		sqlStr = sqlStr + " m.wi_totalsuplycash, m.wi_comment,"
		sqlStr = sqlStr + " m.et_cnt, m.et_totalsellcash,"
		sqlStr = sqlStr + " m.et_totalsuplycash, m.et_comment,"
		sqlStr = sqlStr + " m.sh_cnt, m.sh_totalsellcash,"
		sqlStr = sqlStr + " m.sh_totalsuplycash, m.sh_comment,"

		sqlStr = sqlStr + " m.regdate,m.cancelyn,m.finishflag,convert(varchar(10),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + " convert(varchar(10),m.taxregdate,20) as taxregdate, m.bigo, "
		sqlStr = sqlStr + " p.jungsan_email,p.jungsan_bank,p.jungsan_date,p.jungsan_acctno,"
		sqlStr = sqlStr + " p.jungsan_acctname,p.company_name, p.jungsan_gubun,p.company_no,p.ceoname,p.company_address,p.company_address2,"
		sqlStr = sqlStr + " m.taxinputdate, m.differencekey, m.taxtype, m.taxlinkidx, m.neotaxno, m.bankingupflag"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c,"
		sqlStr = sqlStr + " [db_jungsan].[dbo].tbl_designer_jungsan_master m"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner_group p on m.groupid=p.groupid"
		sqlStr = sqlStr + " where m.id<>0"
        sqlStr = sqlStr + " and m.designerid=c.userid "
        sqlStr = sqlStr + " and c.userdiv='14'"
        ''sqlStr = sqlStr + " and ((c.userdiv='14') or (c.catecode in ('94','95')))"

		if (FRectDesigner="") And (FRectLectureID="") and (FRectYYYYMM<>"") then
			sqlStr = sqlStr + " and m.yyyymm='" + FRectYYYYMM + "'"
		end if

		if FRectDesignerViewOnly=true then
			sqlStr = sqlStr + " and m.finishflag>0"
		end if

		if FRectDesigner<>"" then
			sqlStr = sqlStr + " and m.designerid='" + FRectDesigner + "'"
		end if

		if FRectID<>"" then
			sqlStr = sqlStr + " and m.id=" + CStr(FRectID)
		end If
		
		if FRectLectureID<>"" then
			sqlStr = sqlStr + " and c.userid='" + FRectLectureID + "'"
		end if

		if FRectState<>"" then
			sqlStr = sqlStr + " and m.finishflag='" + FRectState + "'"
		end if

		if FRectIpkumilNot<>"" then
			sqlStr = sqlStr + " and p.jungsan_date<>'" + FRectIpkumilNot + "'"
		end if

		if FRectNotIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun<>'원천징수'"
			sqlStr = sqlStr + " and p.jungsan_gubun<>'간이과세'"
		end if

		if FRectOnlyIncludeWonChon<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='원천징수'"
		end if

		if FRectOnlyIncludeNoTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='면세'"
		end if

		if FRectOnlyIncludeSimpleTax<>"" then
			sqlStr = sqlStr + " and p.jungsan_gubun='간이과세'"
		end if

		if FRectOnlyElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is Not NULL"
		end if

		if FRectOnlyNotElecTax<>"" then
			sqlStr = sqlStr + " and m.neotaxno is NULL"
		end if

		if FRectBankingupflag<>"" then
			sqlStr = sqlStr + " and m.bankingupflag='" + FRectBankingupflag + "'"
		end if

		if FRectNotYYYYMM<>"" then
			sqlStr = sqlStr + " and m.yyyymm<>'" + FRectNotYYYYMM + "'"
		end if
        
        if (FRectTaxType<>"") then
		    sqlStr = sqlStr + " and m.taxtype='" + FRectTaxType + "'"
		end if
		
		if (FRectTaxDate<>"") then
		    sqlStr = sqlStr + " and m.taxregdate='" + FRectTaxDate + "'"
		end If
		
		If FrectOrder="marginD" Or FrectOrder="marginA" Then
			sqlStr = sqlStr + " and (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash)<>0"
		end If

		if FrectOrder="totalsellD" then
			sqlStr = sqlStr + " order by (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash) desc"
		elseif FrectOrder="totalsellA" then
			sqlStr = sqlStr + " order by (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash) asc"
		elseif FrectOrder="marginD" then
			sqlStr = sqlStr + " order by (((m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash)-(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash))/(m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash)) desc"
		elseif FrectOrder="marginA" then
			sqlStr = sqlStr + " order by (((m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash)-(m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash))/(m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash)) asc"
		elseif FrectOrder="jungsanD" then
			sqlStr = sqlStr + " order by (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash) desc"
		elseif FrectOrder="jungsanA" then
			sqlStr = sqlStr + " order by (m.ub_totalsuplycash + m.me_totalsuplycash + m.wi_totalsuplycash + m.et_totalsuplycash + m.sh_totalsuplycash) asc"
		else
			sqlStr = sqlStr + " order by (m.ub_totalsellcash + m.me_totalsellcash + m.wi_totalsellcash + m.et_totalsellcash + m.sh_totalsellcash) desc"
		end If

		
		
''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CJungsanMasterItem

				FItemList(i).Fid               = rsget("id")
				FItemList(i).Fdesignerid       = rsget("designerid")
				FItemList(i).Fgroupid		   = rsget("groupid")
				FItemList(i).Fyyyymm           = rsget("yyyymm")
				FItemList(i).Ftitle            = rsget("title")
				FItemList(i).Fub_cnt           = rsget("ub_cnt")
				FItemList(i).Fub_totalsellcash = rsget("ub_totalsellcash")
				FItemList(i).Fub_totalsuplycash= rsget("ub_totalsuplycash")
				FItemList(i).Fub_comment       = db2html(rsget("ub_comment"))
				FItemList(i).Fme_cnt           = rsget("me_cnt")
				FItemList(i).Fme_totalsellcash = rsget("me_totalsellcash")
				FItemList(i).Fme_totalsuplycash= rsget("me_totalsuplycash")
				FItemList(i).Fme_comment       = db2html(rsget("me_comment"))
				FItemList(i).Fwi_cnt           = rsget("wi_cnt")
				FItemList(i).Fwi_totalsellcash = rsget("wi_totalsellcash")
				FItemList(i).Fwi_totalsuplycash= rsget("wi_totalsuplycash")
				FItemList(i).Fwi_comment       = db2html(rsget("wi_comment"))

				FItemList(i).Fet_cnt           = rsget("et_cnt")
				FItemList(i).Fet_totalsellcash = rsget("et_totalsellcash")
				FItemList(i).Fet_totalsuplycash= rsget("et_totalsuplycash")
				FItemList(i).Fet_comment       = db2html(rsget("et_comment"))
				FItemList(i).Fsh_cnt           = rsget("sh_cnt")
				FItemList(i).Fsh_totalsellcash = rsget("sh_totalsellcash")
				FItemList(i).Fsh_totalsuplycash= rsget("sh_totalsuplycash")
				FItemList(i).Fsh_comment       = db2html(rsget("sh_comment"))


				FItemList(i).Fregdate          = rsget("regdate")
				FItemList(i).Fcancelyn         = rsget("cancelyn")
				FItemList(i).Ffinishflag       = rsget("finishflag")
				FItemList(i).Fipkumdate        = rsget("ipkumdate")
				FItemList(i).Ftaxregdate       = rsget("taxregdate")
				FItemList(i).Fbigo			   = db2html(rsget("bigo"))
				FItemList(i).FDesignerEmail		= rsget("jungsan_email")

				FItemList(i).Fjungsan_bank		= rsget("jungsan_bank")
				FItemList(i).Fjungsan_date		= rsget("jungsan_date")
				FItemList(i).Fjungsan_acctno		= rsget("jungsan_acctno")
				FItemList(i).Fjungsan_acctname		= rsget("jungsan_acctname")
				FItemList(i).Fcompany_name		= db2html(rsget("company_name"))

				FItemList(i).Fjungsan_gubun		= db2html(rsget("jungsan_gubun"))
				FItemList(i).Ftaxinputdate	= rsget("taxinputdate")
				FItemList(i).Fcompany_no	= db2html(rsget("company_no"))
				FItemList(i).Fceoname	= db2html(rsget("ceoname"))
				FItemList(i).Fcompany_address	= db2html(rsget("company_address"))
				FItemList(i).Fcompany_address2	= db2html(rsget("company_address2"))

				FItemList(i).Fdifferencekey = rsget("differencekey")
				FItemList(i).Ftaxtype = rsget("taxtype")
				FItemList(i).FTaxLinkidx = rsget("taxlinkidx")
				FItemList(i).Fneotaxno = rsget("neotaxno")

				FItemList(i).Fbankingupflag = rsget("bankingupflag")


				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close

	end function



	public sub SearchLectureJungsanList()
		dim sqlStr,i

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " d.detailidx, d.orderserial, d.itemid, d.itemname, d.itemoption,"
		sqlStr = sqlStr + "  d.itemno, d.itemoptionname, d.itemcost, d.buycash,"
		sqlStr = sqlStr + "  d.currstate, convert(varchar(19),d.beasongdate,20) as beasongdate,"
		sqlStr = sqlStr + "  d.songjangno, m.buyname, convert(varchar(19),m.regdate,20) as regdate, convert(varchar(19),m.ipkumdate,20) as ipkumdate,"
		sqlStr = sqlStr + "  m.jumundiv, i.lec_date"
		sqlStr = sqlStr + " from [110.93.128.83].[db_academy].[dbo].tbl_academy_order_master m,"
		sqlStr = sqlStr + " [110.93.128.83].[db_academy].[dbo].tbl_academy_order_detail d,"
		sqlStr = sqlStr + " [110.93.128.83].[db_academy].[dbo].tbl_lec_item i"

		sqlStr = sqlStr + " where m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.ipkumdiv>3"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.itemid=i.idx"
		sqlStr = sqlStr + " and d.makerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " and i.lec_date='" + FRectYYYYMM + "'"
		sqlStr = sqlStr + " and m.orderserial not in ("
		sqlStr = sqlStr + " 	select d.mastercode from "
		sqlStr = sqlStr + " 	[db_jungsan].[dbo].tbl_designer_jungsan_master m,"
		sqlStr = sqlStr + " 	[db_jungsan].[dbo].tbl_designer_jungsan_detail d"
		sqlStr = sqlStr + " 	where m.id=d.masteridx"
		sqlStr = sqlStr + " 	and m.designerid='" + FRectDesigner + "'"
		sqlStr = sqlStr + " 	and d.gubuncd='upche'"
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " order by d.orderserial desc"

		rsget.pagesize = FPageSize
		rsget.Open sqlStr,dbget,1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CLecJungsanItem
				FItemList(i).FIdx           = rsget("detailidx")
				FItemList(i).FOrderSerial   = rsget("orderserial")
				FItemList(i).FItemId        = rsget("itemid")
				FItemList(i).FItemOption    = rsget("itemoption")
				FItemList(i).FItemName      = db2html(rsget("itemname"))
				FItemList(i).FItemOptionName= db2html(rsget("itemoptionname"))
				FItemList(i).FItemNo        = rsget("itemno")
				FItemList(i).FBuyCash       = rsget("buycash")
				FItemList(i).FSellCash      = rsget("itemcost")
				FItemList(i).FCurrState      = rsget("currstate")
				FItemList(i).FBeasongdate      = rsget("beasongdate")
				FItemList(i).FRegDate		 = rsget("regdate")
				FItemList(i).FIpkumDate		 = rsget("ipkumdate")
				FItemList(i).FBuyName		= rsget("buyname")
				FItemList(i).FJumunDiv		= rsget("jumundiv")
				FItemList(i).FUpcheSongjangNo		= rsget("songjangno")

				FItemList(i).Flec_date		= rsget("lec_date")

				i=i+1
				rsget.moveNext
			loop
		end if

		rsget.Close
	end sub


	Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage = 1
		FPageSize = 300
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function
end Class
%>