<%
function I_getLogCheckTypeName(ichktype)
    if isNULL(ichktype) then Exit function

    SELECT CASE ichktype
        CASE 0 : I_getLogCheckTypeName="일반"
        CASE 7 : I_getLogCheckTypeName="주문원장"
        CASE 9 : I_getLogCheckTypeName="쿠폰(취소)"
        CASE 8 : I_getLogCheckTypeName="출고후취소"
        CASE 99 : I_getLogCheckTypeName="마스터디테일"
        CASE ELSE I_getLogCheckTypeName = ichktype
    END SELECT
end function

Class COrderLogCheckItem
    public FQueIdx
    public Forderserial
    public Fchktype
    public FQueregdt
    public Fsitename
    public Fjumundiv
    public Fcancelyn
    public FIpkumdiv

    public Fitemid
    public Fitemoption
    public Fitemno
    public Fmakerid
    public Fitemcost
    public Freducedprice
    public Fbuycash
    public Fdcancelyn

    public Fbeasongdate
    public Fdlvfinishdt
    public Fjungsanfixdate

    ' public Ftotalsum
    ' public Fsubtotalprice
    ' public FsubtotalpriceCouponNotApplied
    ' public Ftencardspend	  
    ' public Fmiletotalprice

    ' public Flgorgitemcost
    ' public FlgitemCNTsum
    ' public FlgitemcostCouponNotAppliedSum
    ' public FlgitemcostSum
    ' public FlgreducedPriceSum
    ' public FlgbuycashSum
    ' public FlgbuycashcouponNotAppliedSum

    public function getIpkumdivname()
        if isNULL(FIpkumdiv) then Exit function

        if (FIpkumdiv="2") then
            getIpkumdivname = "입금 대기"
        elseif (FIpkumdiv="4") then
            getIpkumdivname = "결제 완료"
        elseif (FIpkumdiv="5") then
            getIpkumdivname = "상품 준비"
        elseif (FIpkumdiv="6") then
            getIpkumdivname = "출고 준비"
        elseif (FIpkumdiv="7") then
            getIpkumdivname = "일부 출고"
        elseif (FIpkumdiv="8") then
            getIpkumdivname = "출고 완료"
        end if
    end function

    public function getCancelynName()
        getCancelynName = ""
        if isNULL(Fcancelyn) then Exit function
        
        if (Fcancelyn<>"N") then getCancelynName= ""
         
    end function

    public function getJumundivName()
        if isNULL(Fjumundiv) then Exit function

        SELECT CASE Fjumundiv
            CASE "9" : getJumundivName="반품"
            CASE "6" : getJumundivName="교환"
            
            CASE ELSE getJumundivName = Fjumundiv
        END SELECT
    end function

    public function getLogCheckTypeName()
        Dim ichkname : ichkname=I_getLogCheckTypeName(Fchktype)
        
        getLogCheckTypeName = replace(ichkname,"일반","")
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CJungsanCheckOrderItem
	public Forderserial
    public Fsitename
    public Fjumundiv
    public Fcancelyn

    public Fmakerid
    public Fitemgubun
    public Fitemid
    public Fitemoption
    public Fitemno
    public Fdcancelyn

    public FitemcostCouponNotApplied
    public Fitemcost
    public Freducedprice
    public Fbuycash
    public FbuycashcouponNotApplied

    public Fbeasongdate
    public Fdlvfinishdt
    public Fjungsanfixdate
    public Fomwdiv
    public Fvatinclude
    public FodlvType
    public Fmileage

    public Fsuborderserial
    public Flgitemno
    public Flgmakerid
    public Flgorgitemcost
    public FlgitemcostCouponNotApplied
    public Flgitemcost
    public FlgreducedPrice
    public FanbunPriceDetailSUM
    public FanbunCouponPriceDetailSUM
    public FanbunAppliedPriceDetailSUM
    public Flgbuycash
    public FlgupcheJungsanCash
    public Flgmileage
    public Flgvatinclude
    public Flgomwdiv
    public FlgodlvType
    public FDTLactDate
    public FDTLipkumdate
    public FDTLsitename
    public FDTLtargetGbn
    public FDTLbeadaldiv
    public FDTLjFixedDt
    public Flgbeasongdate

    public FchkType
    ' public Fbuyname
	' public Freqname
	' public FreqZipAddr

    public function getLogCheckTypeName()
        Dim ichkname : ichkname=I_getLogCheckTypeName(Fchktype)
        
        getLogCheckTypeName = ichkname
    end function

    public function getDCancelynName()
        if isNULL(Fdcancelyn) then Exit function

        if (Fdcancelyn<>"N") then getDCancelynName=Fdcancelyn
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Class CExtOrderJungsanCheckItem
	public Fsitename
	public FordCnt
	public FChgOrdCNT
	public FretOrdCNT
	public ForgOrderserial
	public Fauthcode
	public Fitemid
	public Fitemoption
	public Fitemno
	public FitemcostSum
	public FreducedpriceSum
	public FbeasongMonth
	public Forgsongjangdiv
	public Forgsongjangno
	public Forgdlvfinishdt
	public Forgjungsanfixdate

	public FMinus_itemno			
	public FMinus_itemcostSum		
	public FMinus_reducedpriceSum	
	public FMinus_beasongmonth

	public FextItemNoSum	
	public FextitemcostSum
	public FextReducedPriceSum
	public FextMeachulMonth

	public Fcomment
	public Fdiffitemno
	public FdiffSum
	public Fjorgorderserial


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class



Class CJungsanCommentItem
	public Frowidx
	public Forderserial
	public Fitemid
	public Fitemoption
	public Freguserid
	public Fcomment
	public Fregdate
	public Fdeldate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CJungsanCheck
    public FItemList()
	public FOneItem

	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage

	' public FRectSellSite
	' public FRectJungsanType
	' public FRectGroupGubun

	public FRectStartDate
	public FRectEndDate
	public FRectYYYYMM
	public FRectDiffType

	public FRectSearchField
	public FRectSearchText

	public FRowNo
	public FSumItemNo
	public FSumitemcost
	public FSumMeachulPrice
	public FSumReducedPrice
	public FSumOwnCouponPrice
	public FSumTenCouponPrice
	public FSumCommPrice
	public FSumJungsanPrice
	public FMiMapTTLCnt
	public FSumJungsanPrice_ETC

	public FRectMiMap
	public FRectVatYn
	public FRectReturnOnly
	public FRectErrexists
	public FRectExceptItemCostZero
	public FRectDlvMonth
	public FRectMiMapMinus

	public FRectMakerid
	public FRectItemid
	public FRectReturnExcept
	public FRectMinusGainOnly
	public FRectDiffType2

	public FRectOrderserial
	public FRectItemOption

	public FRectCheckBySum
	public FonlyErrNoExists
	public FRectErrorType
	public FRectAccerrtype

	public FdiffnoSum
	public FdiffsumSum
	public FErrAsignSum

    public function getLogDiffByOrderserial()
        Dim sqlStr
		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiff_ViewByOrderserial] '"&FRectOrderserial&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CJungsanCheckOrderItem
                ''orderserial	sitename	jumundiv	cancelyn	makerid	itemgubun	itemid	itemoption	itemno	dcancelyn	itemcostCouponNotApplied	itemcost	reducedprice	buycash	buycashcouponNotApplied	
                ''beasongdate	dlvfinishdt	jungsanfixdate	omwdiv	vatinclude	odlvType	mileage
                ''suborderserial	lgitemno	lgmakerid	lgorgitemcost	lgitemcostCouponNotApplied	lgitemcost	lgreducedPrice	anbunPriceDetailSUM	anbunCouponPriceDetailSUM	anbunAppliedPriceDetailSUM	lgbuycash
                ''upcheJungsanCash	lgmileage	lgvatinclude	lgomwdiv	lgodlvType	DTLactDate	DTLipkumdate	DTLsitename	DTLtargetGbn	DTLbeadaldiv	DTLjFixedDt

				FItemList(i).Forderserial	= db3_rsget("orderserial")
                FItemList(i).Fsitename	    = db3_rsget("sitename")
                FItemList(i).Fjumundiv		= db3_rsget("jumundiv")
                FItemList(i).Fcancelyn		= db3_rsget("cancelyn")

				FItemList(i).Fmakerid		= db3_rsget("makerid")
				FItemList(i).Fitemgubun		= db3_rsget("itemgubun")
                FItemList(i).Fitemid		= db3_rsget("itemid")
                FItemList(i).Fitemoption	= db3_rsget("itemoption")
                FItemList(i).Fitemno		= db3_rsget("itemno")
                FItemList(i).Fdcancelyn		= db3_rsget("dcancelyn")
                FItemList(i).FitemcostCouponNotApplied		= db3_rsget("itemcostCouponNotApplied")
                FItemList(i).Fitemcost		= db3_rsget("itemcost")
                FItemList(i).Freducedprice		= db3_rsget("reducedprice")
                FItemList(i).Fbuycash		= db3_rsget("buycash")
                FItemList(i).FbuycashcouponNotApplied		= db3_rsget("buycashcouponNotApplied")
                FItemList(i).Fbeasongdate	= db3_rsget("beasongdate")
                FItemList(i).Fdlvfinishdt	= db3_rsget("dlvfinishdt")
                FItemList(i).Fjungsanfixdate	= db3_rsget("jungsanfixdate")
                FItemList(i).Fomwdiv		= db3_rsget("omwdiv")
                FItemList(i).Fvatinclude		= db3_rsget("vatinclude")
                FItemList(i).FodlvType		= db3_rsget("odlvType")
                FItemList(i).Fmileage		= db3_rsget("mileage")

                FItemList(i).Fsuborderserial		= db3_rsget("suborderserial")
                FItemList(i).Flgitemno		= db3_rsget("lgitemno")
                FItemList(i).Flgmakerid		= db3_rsget("lgmakerid")
                FItemList(i).Flgorgitemcost		= db3_rsget("lgorgitemcost")
                FItemList(i).FlgitemcostCouponNotApplied		= db3_rsget("lgitemcostCouponNotApplied")
                FItemList(i).Flgitemcost		= db3_rsget("lgitemcost")
                FItemList(i).FlgreducedPrice		= db3_rsget("lgreducedPrice")
                FItemList(i).FanbunPriceDetailSUM		= db3_rsget("anbunPriceDetailSUM")
                FItemList(i).FanbunCouponPriceDetailSUM		= db3_rsget("anbunCouponPriceDetailSUM")
                FItemList(i).FanbunAppliedPriceDetailSUM		= db3_rsget("anbunAppliedPriceDetailSUM")
                FItemList(i).Flgbuycash		    = db3_rsget("lgbuycash")
                FItemList(i).FlgupcheJungsanCash		= db3_rsget("upcheJungsanCash")
                FItemList(i).Flgmileage		    = db3_rsget("lgmileage")
                FItemList(i).Flgvatinclude		= db3_rsget("lgvatinclude")
                FItemList(i).Flgomwdiv		    = db3_rsget("lgomwdiv")
                FItemList(i).FlgodlvType		= db3_rsget("lgodlvType")
                FItemList(i).FDTLactDate		= db3_rsget("DTLactDate")
                FItemList(i).FDTLipkumdate		= db3_rsget("DTLipkumdate")
                FItemList(i).FDTLsitename		= db3_rsget("DTLsitename")
                FItemList(i).FDTLtargetGbn		= db3_rsget("DTLtargetGbn")
                FItemList(i).FDTLbeadaldiv		= db3_rsget("DTLbeadaldiv")
                FItemList(i).FDTLjFixedDt		= db3_rsget("DTLjFixedDt")
                FItemList(i).Flgbeasongdate     = db3_rsget("lgbeasongdate")

                FItemList(i).FchkType       = db3_rsget("chkType")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

    end function
    
    public function getLogDiffList()
		Dim sqlStr

        IF (FRectDiffType="") or (FRectDiffType<"100") then 
		    sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiff_GetByQueue] "&FPageSize&","&CHKIIF(FRectDiffType="","NULL",FRectDiffType)
        ELSEIF (FRectDiffType>="100") and (FRectDiffType<"200") then 
            sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiff_Get_ON] "&FPageSize&",'"&FRectYYYYMM&"',"&FRectDiffType
        ELSEIF (FRectDiffType>="200") and (FRectDiffType<"300") then 
            sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiff_Get_OF] "&FPageSize&",'"&FRectYYYYMM&"',"&FRectDiffType
        ELSEIF (FRectDiffType>="300") and (FRectDiffType<"400") then 
            sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiffJungsan_Get_ON] "&FPageSize&",'"&FRectYYYYMM&"',"&FRectDiffType
        ELSEIF (FRectDiffType>="400") and (FRectDiffType<"500") then 
            sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiffJungsan_Get_OF] "&FPageSize&",'"&FRectYYYYMM&"',"&FRectDiffType
        ELSEIF (FRectDiffType>="900") and (FRectDiffType<="999") then 
            sqlStr = " exec [db_datamart].[dbo].[usp_Ten_OrderLogDiffFixdateCheck] "&FPageSize&",'"&FRectYYYYMM&"',"&FRectDiffType
        END IF

        if (sqlStr="") then Exit function

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new COrderLogCheckItem
                ''idx	orderserial	chktype	regdt	sitename	jumundiv	cancelyn	totalsum	subtotalprice	subtotalpriceCouponNotApplied	tencardspend	miletotalprice
                ''lgorgitemcost	lgitemCNTsum	lgitemcostCouponNotAppliedSum	lgitemcostSum	reducedPriceSum	buycashSum	buycashcouponNotAppliedSum

                FItemList(i).FQueIdx	    = db3_rsget("idx")
				FItemList(i).Forderserial	= db3_rsget("orderserial")
				FItemList(i).Fchktype		= db3_rsget("chktype")
				FItemList(i).FQueregdt		= db3_rsget("regdt")
				FItemList(i).Fsitename	    = db3_rsget("sitename")
                FItemList(i).Fjumundiv		= db3_rsget("jumundiv")
                FItemList(i).Fcancelyn		= db3_rsget("cancelyn")
				FItemList(i).Fipkumdiv		= db3_rsget("ipkumdiv")

                FItemList(i).Fitemid        = db3_rsget("itemid")
                FItemList(i).Fitemoption    = db3_rsget("itemoption")
                FItemList(i).Fitemno        = db3_rsget("itemno")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fitemcost      = db3_rsget("itemcost")
                FItemList(i).Freducedprice  = db3_rsget("reducedprice")
                FItemList(i).Fbuycash       = db3_rsget("buycash")
                FItemList(i).Fdcancelyn		= db3_rsget("dcancelyn")

                FItemList(i).Fbeasongdate   = db3_rsget("beasongdate")
                FItemList(i).Fdlvfinishdt   = db3_rsget("dlvfinishdt")
                FItemList(i).Fjungsanfixdate= db3_rsget("jungsanfixdate")

                ' FItemList(i).Ftotalsum		= db3_rsget("totalsum")
				' FItemList(i).Fsubtotalprice	= db3_rsget("subtotalprice")
				' FItemList(i).FsubtotalpriceCouponNotApplied		= db3_rsget("subtotalpriceCouponNotApplied")
				' FItemList(i).Ftencardspend	  = db3_rsget("tencardspend")
				' FItemList(i).Fmiletotalprice  = db3_rsget("miletotalprice")

                ' FItemList(i).Flgorgitemcost   = db3_rsget("lgorgitemcostSum")
				' FItemList(i).FlgitemCNTsum	  = db3_rsget("lgitemCNTsum")
                ' FItemList(i).FlgitemcostCouponNotAppliedSum	  = db3_rsget("lgitemcostCouponNotAppliedSum")
                ' FItemList(i).FlgitemcostSum	  = db3_rsget("lgitemcostSum")
                ' FItemList(i).FlgreducedPriceSum = db3_rsget("lgreducedPriceSum")
                ' FItemList(i).FlgbuycashSum	  = db3_rsget("lgbuycashSum")
                ' FItemList(i).FlgbuycashcouponNotAppliedSum	  = db3_rsget("lgbuycashcouponNotAppliedSum")

				' FItemList(i).Fitemoption	= db3_rsget("itemoption")
				' FItemList(i).Fitemname		= db3_rsget("itemname")
				' FItemList(i).Fitemoptionname	= db3_rsget("itemoptionname")
				' FItemList(i).Fmakerid			= db3_rsget("makerid")
				' FItemList(i).Fupcheconfirmdate	= db3_rsget("upcheconfirmdate")
				' FItemList(i).FitemcostcouponnotApplied = db3_rsget("itemcostcouponnotApplied")
				' FItemList(i).Fitemcost		= db3_rsget("itemcost")
				' FItemList(i).Freducedprice	= db3_rsget("reducedprice")
				' FItemList(i).Fitemno		= db3_rsget("itemno")
				' FItemList(i).Fodlvfixday	= db3_rsget("odlvfixday")
				' FItemList(i).Fsongjangdiv	= db3_rsget("songjangdiv")
				' FItemList(i).Fsongjangno	= db3_rsget("songjangno")
				' FItemList(i).Fbeasongdate	= db3_rsget("beasongdate")
				' FItemList(i).Fdlvfinishdt	= db3_rsget("dlvfinishdt")
				' FItemList(i).Fjungsanfixdate	= db3_rsget("jungsanfixdate")
				' FItemList(i).Fbuycash		= db3_rsget("buycash")
				' FItemList(i).Fomwdiv		= db3_rsget("omwdiv")

				' FItemList(i).Fcomment		= db3_rsget("comment")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end Function



'''--------------------------------


	public function getExtjungsanCommentList()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_Comment_List] '"&FRectOrderserial&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanCommentItem
				
				FItemList(i).Frowidx		= db3_rsget("rowidx")
				FItemList(i).Forderserial	= db3_rsget("orderserial")
				FItemList(i).Fitemid		= db3_rsget("itemid")
				FItemList(i).Fitemoption	= db3_rsget("itemoption")
				FItemList(i).Freguserid		= db3_rsget("reguserid")
				FItemList(i).Fcomment		= db3_rsget("comment")
				FItemList(i).Fregdate		= db3_rsget("regdate")
				FItemList(i).Fdeldate		= db3_rsget("deldate")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
	end function

	public function getOutJungsanCheckCSInfo()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTMALL_Jungsan_CheckCSInfo] '"&FRectOrderserial&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanCheckCSItem
				FItemList(i).Fcsid          = rsget("id")
				FItemList(i).Fdivcd         = rsget("divcd")
				FItemList(i).FdivName       = rsget("divName")
				FItemList(i).Fwriteuser     = rsget("writeuser")
				FItemList(i).Ffinishuser    = rsget("finishuser")
				FItemList(i).Ftitle         = rsget("title")
				FItemList(i).Fcurrstate     = rsget("currstate")
				FItemList(i).Fregdate       = rsget("regdate")
				FItemList(i).Ffinishdate    = rsget("finishdate")
				FItemList(i).Fconfirmdate   = rsget("confirmdate")
				FItemList(i).Fdeletedate    = rsget("deletedate")
				FItemList(i).Fdeleteyn      = rsget("deleteyn")
				FItemList(i).Frequireupche  = rsget("requireupche")
				FItemList(i).Fmakerid  		= rsget("makerid")
				FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
				FItemList(i).Fsongjangno    = rsget("songjangno")
				FItemList(i).Fextsitename   = rsget("extsitename")

				FItemList(i).Frefasid				= rsget("refasid")
				FItemList(i).Frefminusorderserial	= rsget("refminusorderserial")
				FItemList(i).Frefchangeorderserial	= rsget("refchangeorderserial")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Function

	public function getOutJungsanCheckOrderInfo()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTMALL_Jungsan_CheckOrderInfo] '"&FRectOrderserial&"'"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanCheckOrderItem
				FItemList(i).Forderserial	= rsget("orderserial")
				FItemList(i).Fbuyname		= rsget("buyname")
				FItemList(i).Freqname		= rsget("reqname")
				FItemList(i).FreqZipAddr	= rsget("reqZipAddr")
				FItemList(i).Fipkumdiv		= rsget("ipkumdiv")
				FItemList(i).Fcancelyn		= rsget("cancelyn")
				FItemList(i).Fdcancelyn		= rsget("dcancelyn")
				FItemList(i).Fregdate		= rsget("regdate")
				FItemList(i).Fipkumdate		= rsget("ipkumdate")
				FItemList(i).Fbaljudate		= rsget("baljudate")
				FItemList(i).Fbeadaldiv		= rsget("beadaldiv")
				FItemList(i).Fsitename		= rsget("sitename")
				FItemList(i).Fjumundiv		= rsget("jumundiv")
				FItemList(i).Fidx			= rsget("idx")
				FItemList(i).Fitemid		= rsget("itemid")
				FItemList(i).Fitemoption	= rsget("itemoption")
				FItemList(i).Fitemname		= rsget("itemname")
				FItemList(i).Fitemoptionname	= rsget("itemoptionname")
				FItemList(i).Fmakerid			= rsget("makerid")
				FItemList(i).Fupcheconfirmdate	= rsget("upcheconfirmdate")
				FItemList(i).FitemcostcouponnotApplied = rsget("itemcostcouponnotApplied")
				FItemList(i).Fitemcost		= rsget("itemcost")
				FItemList(i).Freducedprice	= rsget("reducedprice")
				FItemList(i).Fitemno		= rsget("itemno")
				FItemList(i).Fodlvfixday	= rsget("odlvfixday")
				FItemList(i).Fsongjangdiv	= rsget("songjangdiv")
				FItemList(i).Fsongjangno	= rsget("songjangno")
				FItemList(i).Fbeasongdate	= rsget("beasongdate")
				FItemList(i).Fdlvfinishdt	= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= rsget("jungsanfixdate")
				FItemList(i).Fbuycash		= rsget("buycash")
				FItemList(i).Fomwdiv		= rsget("omwdiv")

				FItemList(i).Fcomment		= rsget("comment")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end Function

	public function getExtJungsanOrderDiffList()
		Dim sqlStr
		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_OrderDiffCheck] '" & FRectSellSite & "','" & FRectStartDate & "','" & FRectEndDate & "', '" & FRectDiffType & "', '" & FPageSize & "','"&FRectDlvMonth&"',"&CHKIIF(FRectCheckBySum<>"",1,"NULL")&""

        rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				'FItemList(i).FextJungsanType		= rsget("extJungsanType")
				'FItemList(i).FextCommPrice			= rsget("extCommPrice")
				'FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				'FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				'FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				'FItemList(i).FsiteNo				= rsget("siteNo")

				FItemList(i).Forderitemcost			= rsget("itemcost")
				FItemList(i).Forderreducedprice		= rsget("reducedprice")
				FItemList(i).Forderitemno			= rsget("itemno")
				FItemList(i).Forderbeasongdate		= rsget("beasongdate")
				if NOT isNULL(FItemList(i).Forderbeasongdate) THEN
					FItemList(i).Forderbeasongdate = LEFT(FItemList(i).Forderbeasongdate,10)
				end if

				FItemList(i).Fdtlcancelyn		= rsget("dtlcancelyn")
				FItemList(i).Fmastercancelyn	= rsget("mastercancelyn")

				FItemList(i).Fdlvfinishdt		= rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= rsget("jungsanfixdate")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public function getExtJungsanOrderDiffList_replica()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_OUTAMLL_Jungsan_OrderDiffCheck_replica] '" & FRectSellSite & "','" & FRectStartDate & "','" & FRectEndDate & "', '" & FRectDiffType & "', '" & FPageSize & "','"&FRectDlvMonth&"',"&CHKIIF(FRectCheckBySum<>"",1,"NULL")&""

        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= db3_rsget("sellsite")
				FItemList(i).FextOrderserial		= db3_rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= db3_rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= db3_rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= db3_rsget("extItemNo")
				FItemList(i).FextItemCost			= db3_rsget("extItemCost")
				FItemList(i).FextReducedPrice		= db3_rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= db3_rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= db3_rsget("extTenCouponPrice")
				'FItemList(i).FextJungsanType		= db3_rsget("extJungsanType")
				'FItemList(i).FextCommPrice			= db3_rsget("extCommPrice")
				'FItemList(i).FextTenMeachulPrice	= db3_rsget("extTenMeachulPrice")
				'FItemList(i).FextTenJungsanPrice	= db3_rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= db3_rsget("extMeachulDate")
				'FItemList(i).FextJungsanDate		= db3_rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= db3_rsget("OrgOrderserial")
				FItemList(i).Fitemid				= db3_rsget("itemid")
				FItemList(i).Fitemoption			= db3_rsget("itemoption")
				'FItemList(i).FsiteNo				= db3_rsget("siteNo")

				FItemList(i).Forderitemcost			= db3_rsget("itemcost")
				FItemList(i).Forderreducedprice		= db3_rsget("reducedprice")
				FItemList(i).Forderitemno			= db3_rsget("itemno")
				FItemList(i).Forderbeasongdate		= db3_rsget("beasongdate")
				if NOT isNULL(FItemList(i).Forderbeasongdate) THEN
					FItemList(i).Forderbeasongdate = LEFT(FItemList(i).Forderbeasongdate,10)
				end if

				FItemList(i).Fdtlcancelyn		= db3_rsget("dtlcancelyn")
				FItemList(i).Fmastercancelyn	= db3_rsget("mastercancelyn")

				FItemList(i).Fdlvfinishdt		= db3_rsget("dlvfinishdt")
				FItemList(i).Fjungsanfixdate	= db3_rsget("jungsanfixdate")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function getExtOrderJungsanDiffList()
		Dim sqlStr
		sqlStr = " exec [db_dataSummary].[dbo].[usp_Ten_Check_Outmall_OrderVsOutJungsan] '"&FRectDlvMonth&"','" & FRectSellSite & "',"&CHKIIF(FRectDiffType="","NULL",FRectDiffType)&","&CHKIIF(FRectDiffType2="","NULL",FRectDiffType2)&""

        db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				
				set FItemList(i) = new CExtOrderJungsanCheckItem

				FItemList(i).Fsitename 			= db3_rsget("sitename")
				FItemList(i).FordCnt			= db3_rsget("ordCnt")
				FItemList(i).FChgOrdCNT			= db3_rsget("ChgOrdCNT")
				FItemList(i).FretOrdCNT			= db3_rsget("retOrdCNT")
				FItemList(i).ForgOrderserial	= db3_rsget("orgOrderserial")
				FItemList(i).Fauthcode			= db3_rsget("authcode")
				FItemList(i).Fitemid			= db3_rsget("itemid")
				FItemList(i).Fitemoption		= db3_rsget("itemoption")
				FItemList(i).Fitemno			= db3_rsget("itemno")
				FItemList(i).FitemcostSum		= db3_rsget("itemcostSum")
				FItemList(i).FreducedpriceSum	= db3_rsget("reducedpriceSum")
				FItemList(i).FbeasongMonth		= db3_rsget("beasongMonth")
				FItemList(i).Forgsongjangdiv	= db3_rsget("orgsongjangdiv")
				FItemList(i).Forgsongjangno		= db3_rsget("orgsongjangno")
				FItemList(i).Forgdlvfinishdt	= db3_rsget("orgdlvfinishdt")
				FItemList(i).Forgjungsanfixdate	= db3_rsget("orgjungsanfixdate")

				FItemList(i).FMinus_itemno			= db3_rsget("Minus_itemno")
				FItemList(i).FMinus_itemcostSum		= db3_rsget("Minus_itemcostSum")
				FItemList(i).FMinus_reducedpriceSum	= db3_rsget("Minus_reducedpriceSum")
				FItemList(i).FMinus_beasongmonth		= db3_rsget("Minus_beasongmonth")

				FItemList(i).FextItemNoSum			= db3_rsget("extItemNoSum")
				FItemList(i).FextitemcostSum		= db3_rsget("extitemcostSum")
				FItemList(i).FextReducedPriceSum	= db3_rsget("extReducedPriceSum")
				FItemList(i).FextMeachulMonth		= db3_rsget("extMeachulMonth")

				FItemList(i).Fcomment				= db3_rsget("comment")
				FItemList(i).Fdiffitemno			= db3_rsget("diffitemno")
				FItemList(i).FdiffSum				= db3_rsget("diffSum")

				FItemList(i).Fjorgorderserial		= db3_rsget("jorgorderserial")  ''정산내역이 있는지 여부판단.

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanByItemDW()
		Dim sqlStr, i
        '' @styyyymmdd ,@edyyyymmdd ,@sellsite varchar(32) = NULL,@makerid varchar(32) = NULL
        '',@itemid int = NULL,@returnExcept int = 0 -- 반품정산건 제외	,@minuscOnly int = 0 -- 마이너스 상품이 있는경우만.
		sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSumDtlByItemCNT] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "'," & FRectItemid & ", " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ""
''rw sqlStr
        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly
		    FTotalCount = rsSTSget("CNT")
        rsSTSget.Close

        if (FTotalCount<1) then
            FResultCount = 0
            Exit function
        end if

        FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if

        sqlStr = " exec [db_statistics_order].[dbo].[usp_TEN_XSite_GainSumDtlByItem] '"&FRectStartdate&"','"&FRectEndDate&"','" & FRectSellSite & "','" & FRectMakerid & "'," & FRectItemid & ", " & CHKIIF(FRectReturnExcept="","NULL",FRectReturnExcept) & ", " & CHKIIF(FRectMinusGainOnly="","NULL",FRectMinusGainOnly) & ","&FPageSize&","&FCurrPage
        rsSTSget.CursorLocation = adUseClient
		rsSTSget.Open sqlStr,dbSTSget,adOpenForwardOnly,adLockReadOnly

		FResultCount = rsSTSget.RecordCount
		if FResultCount<0 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsSTSget.EOF  then
			do until rsSTSget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsSTSget("sellsite")
				FItemList(i).FextOrderserial		= rsSTSget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsSTSget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsSTSget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsSTSget("extItemNo")

				FItemList(i).FextItemCost			= rsSTSget("extItemCost")
				FItemList(i).FextReducedPrice		= rsSTSget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsSTSget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsSTSget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsSTSget("extJungsanType")
				FItemList(i).FextCommPrice			= rsSTSget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsSTSget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsSTSget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsSTSget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsSTSget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsSTSget("OrgOrderserial")
				FItemList(i).Fitemid				= rsSTSget("itemid")
				FItemList(i).Fitemoption			= rsSTSget("itemoption")
				'FItemList(i).FsiteNo				= rsSTSget("siteNo")
				'FItemList(i).FMinusOrderserial		= rsSTSget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if

				FItemList(i).Fmakerid			= rsSTSget("makerid")
				FItemList(i).Fmwdiv				= rsSTSget("omwdiv")
				FItemList(i).Ftenbuycash		= rsSTSget("tenbuycash")
				FItemList(i).Fjungsangain		= rsSTSget("jungsangain")

				i=i+1
				rsSTSget.moveNext
			loop
		end if
		rsSTSget.Close
	end Function

	public function GetExtJungsanCheckTargetList()
		dim i, sqlStr

        sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_CheckRequireList] '" & FRectSellSite & "','"&FRectStartdate&"','"&FRectEndDate&"','"&FRectDiffType&"'"
        rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsanMapCheckListTmpOrder()
		dim i, sqlStr
		dim iextOrderserial, iOrgOrderserial, iextitemid

		if (FRectSearchField="extOrderserial") then
			iextOrderserial = FRectSearchText
		elseif (FRectSearchField="OrgOrderserial") then
			iOrgOrderserial = FRectSearchText
		elseif (FRectSearchField="extitemid") then
			iextitemid = FRectSearchText
		end if

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MiMapCheckList_TmpOrder] '" & FRectSellSite & "', '" & iextOrderserial & "', '" & iOrgOrderserial & "','"&iextitemid&"' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtOrderTmpItem
				FItemList(i).FOutMallOrderSeq	= rsget("OutMallOrderSeq")
				FItemList(i).FOrderSerial		= rsget("OrderSerial")
				FItemList(i).FOrgDetailKey		= rsget("OrgDetailKey")
				FItemList(i).FSellSite			= rsget("SellSite")
				FItemList(i).FOutMallOrderSerial	= rsget("OutMallOrderSerial")
				FItemList(i).FSellDate			= rsget("SellDate")
				FItemList(i).FPayDate			= rsget("PayDate")
				FItemList(i).FmatchItemID		= rsget("matchItemID")
				FItemList(i).Fmatchitemoption	= rsget("matchitemoption")
				FItemList(i).Fsellprice			= rsget("sellprice")
				FItemList(i).Frealsellprice		= rsget("realsellprice")
				FItemList(i).FItemOrderCount	= rsget("ItemOrderCount")
				FItemList(i).ForderDlvPay		= rsget("orderDlvPay")
				FItemList(i).Fsendstate			= rsget("sendstate")
				FItemList(i).FoutMallGoodsNo	= rsget("outMallGoodsNo")
				FItemList(i).Fref_outmallorderserial	= rsget("ref_outmallorderserial")
				FItemList(i).FbeasongNum11st	= rsget("beasongNum11st")
				
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsanMapCheckList()
		dim i, sqlStr
		dim iextOrderserial, iOrgOrderserial, iextitemid

		if (FRectSearchField="extOrderserial") then
			iextOrderserial = FRectSearchText
		elseif (FRectSearchField="OrgOrderserial") then
			iOrgOrderserial = FRectSearchText
		elseif (FRectSearchField="extitemid") then
			iextitemid = FRectSearchText
		end if

		sqlStr = " exec [db_jungsan].[dbo].[usp_Ten_OUTAMLL_Jungsan_MiMapCheckList] '" & FRectSellSite & "', '" & FRectJungsanType & "', '" & iextOrderserial & "', '" & iOrgOrderserial & "','"&iextitemid&"' "

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")
				FItemList(i).Fref_Slice_extOrderserSeq = rsget("ref_Slice_extOrderserSeq")
				
				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if

				FItemList(i).FExtitemid = rsget("Extitemid")
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
	end function

	public function GetExtJungsan()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		 if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		        addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		 elseif (FRectSellSite<>"") then
			    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		 end if

		if (FRectJungsanType <> "") then
			addSqlStr = addSqlStr + " and j.extJungsanType = '" + CStr(FRectJungsanType) + "' "
		end if

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField="extOrderserial") then
				addSqlStr = addSqlStr + " and (LEFT(j.extOrderserial,"&LEN(FRectSearchText)&")='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " 	or j.extOrgOrderserial='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " )"
			elseif (FRectSearchField="matchitemid") then
				addSqlStr = addSqlStr + " and j.itemid='"&FRectSearchText&"'"&vbCRLF
			else
				addSqlStr = addSqlStr + " and j." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectMiMap<>"") then
			addSqlStr = addSqlStr + " and isNULL(j.OrgOrderserial,'')=''"
		end if

		if (FRectMiMapMinus<>"") then
			addSqlStr = addSqlStr + " and j.extitemno<0 and isNULL(j.MinusOrderserial,'')=''"
		end if

		if (FRectVatYn<>"") then
			addSqlStr = addSqlStr + " and extvatyn='"&FRectVatYn&"'"
		end if

		if (FRectReturnOnly<>"") then
			addSqlStr = addSqlStr + " and extitemno<1"
		end if

		if (FRectErrexists<>"") then
			''FSumitemcost-FSumMeachulPrice-(FSumOwnCouponPrice+FSumTenCouponPrice)
			''FSumJungsanPrice-(FSumMeachulPrice-FSumCommPrice+FSumOwnCouponPrice)

			addSqlStr = addSqlStr + " and ("
			addSqlStr = addSqlStr + "	(extItemCost-extReducedPrice-extOwnCouponPrice-extTenCouponPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extTenJungsanPrice-extReducedPrice+extCommPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extReducedPrice<>extTenMeachulPrice)"
			addSqlStr = addSqlStr + " )"

		end if

		if (FRectExceptItemCostZero<>"") then
			addSqlStr = addSqlStr + " and extItemCost<>0"
		end if

		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
	    sqlStr = sqlStr + " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		' response.write sqlstr & "<Br>"
    	rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if CLng(FCurrPage)>CLng(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		dim sqlSum
		sqlSum = " select  count(*) cnt, sum(j.extitemNO) as itemno , sum(j.extItemCost*j.extItemno) as itemcost "
		sqlSum = sqlSum & " , sum(j.extTenMeachulPrice*j.extItemNo) as MeachulPrice "
		sqlSum = sqlSum & ", sum(j.extReducedPrice*j.extItemNo) as  ReducedPrice "
		sqlSum = sqlSum & ", sum(j.extOwnCouponPrice*j.extItemNo) as  OwnCouponPrice "
		sqlSum = sqlSum & ", sum(j.extTenCouponPrice*j.extItemNo) as TenCouponPrice "
		sqlSum = sqlSum & ", sum(j.extCommPrice*j.extItemNo) as CommPrice "
		sqlSum = sqlSum & ", sum(j.extTenJungsanPrice*j.extItemNo) as JungsanPrice "
		sqlSum = sqlSum & ", sum((CASE WHEN j.extJungsanType not in ('C','D') then j.extTenJungsanPrice*j.extItemNo else 0 END)) as JungsanPrice_ETC "
		sqlSum = sqlSum & ", sum(CASE WHEN isNULL(j.OrgOrderserial,'')='' THEN 1 ELSE 0 END) as MiMapTTLCnt"
		sqlSum = sqlSum & " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
		sqlSum = sqlSum & " where 1=1 "
		sqlSum = sqlSum & addSqlStr

		rsget.CursorLocation = adUseClient
		rsget.Open sqlSum,dbget,adOpenForwardOnly,adLockReadOnly
		if  not rsget.EOF  then
			FRowNo 			= rsget("cnt")
			FSumItemNo 			= rsget("itemno")
			FSumitemcost 		= rsget("itemcost")
			FSumMeachulPrice 	= rsget("MeachulPrice")
			FSumReducedPrice	= rsget("ReducedPrice")
			FSumOwnCouponPrice 	= rsget("OwnCouponPrice")
			FSumTenCouponPrice 	= rsget("TenCouponPrice")
			FSumCommPrice 		= rsget("CommPrice")
			FSumJungsanPrice 	= rsget("JungsanPrice")
			FMiMapTTLCnt		= rsget("MiMapTTLCnt")

			FSumJungsanPrice_ETC = rsget("JungsanPrice_ETC")
		end if
		rsget.close

		'// ====================================================================
		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " j.* "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by j.extMeachulDate desc, j.sellsite, j.extOrderserial, j.extOrderserSeq "

		' response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanItem

				FItemList(i).Fsellsite				= rsget("sellsite")
				FItemList(i).FextOrderserial		= rsget("extOrderserial")
				FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
				FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
				FItemList(i).FextItemNo				= rsget("extItemNo")
				FItemList(i).FextItemCost			= rsget("extItemCost")
				FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
				FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
				FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
				FItemList(i).FextJungsanType		= rsget("extJungsanType")
				FItemList(i).FextCommPrice			= rsget("extCommPrice")
				FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
				FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
				FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
				FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
				FItemList(i).Fitemid				= rsget("itemid")
				FItemList(i).Fitemoption			= rsget("itemoption")
				FItemList(i).FsiteNo				= rsget("siteNo")
				FItemList(i).FMinusOrderserial		= rsget("MinusOrderserial")

				FItemList(i).FextTenJungsanPrice_ETC = 0
				if (FItemList(i).FextJungsanType<>"C") and (FItemList(i).FextJungsanType<>"D") then
					FItemList(i).FextTenJungsanPrice_ETC = FItemList(i).FextTenJungsanPrice
				end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

	public function GetExtJungsanExcelDown()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		      addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		elseif (FRectSellSite<>"") then
		    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		end if

		if (FRectJungsanType <> "") then
			addSqlStr = addSqlStr + " and j.extJungsanType = '" + CStr(FRectJungsanType) + "' "
		end if

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if (FRectSearchField <> "") and (FRectSearchText <> "") then
			if (FRectSearchField="extOrderserial") then
				addSqlStr = addSqlStr + " and (LEFT(j.extOrderserial,"&LEN(FRectSearchText)&")='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " 	or j.extOrgOrderserial='"&FRectSearchText&"'"
				addSqlStr = addSqlStr + " )"
			else
				addSqlStr = addSqlStr + " and j." + CStr(FRectSearchField) + " = '" + CStr(FRectSearchText) + "' "
			end if
		end if

		if (FRectMiMap<>"") then
			addSqlStr = addSqlStr + " and isNULL(j.OrgOrderserial,'')=''"
		end if

		if (FRectVatYn<>"") then
			addSqlStr = addSqlStr + " and extvatyn='"&FRectVatYn&"'"
		end if

		if (FRectReturnOnly<>"") then
			addSqlStr = addSqlStr + " and extitemno<1"
		end if

		if (FRectErrexists<>"") then
			''FSumitemcost-FSumMeachulPrice-(FSumOwnCouponPrice+FSumTenCouponPrice)
			''FSumJungsanPrice-(FSumMeachulPrice-FSumCommPrice+FSumOwnCouponPrice)

			addSqlStr = addSqlStr + " and ("
			addSqlStr = addSqlStr + "	(extItemCost-extReducedPrice-extOwnCouponPrice-extTenCouponPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extTenJungsanPrice-extReducedPrice+extCommPrice<>0)"
			addSqlStr = addSqlStr + " 	or (extReducedPrice<>extTenMeachulPrice)"
			addSqlStr = addSqlStr + " )"

		end if

		if (FRectExceptItemCostZero<>"") then
			addSqlStr = addSqlStr + " and extItemCost<>0"
		end if

		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""
		''addSqlStr = addSqlStr + ""

		'// ====================================================================
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(FPageSize) + ") as totPg"
	    sqlStr = sqlStr + " from db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

		''response.write sqlstr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
			FResultCount = FTotalCount
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		'// ====================================================================
		'' 최대 2만건만 하자.
		sqlStr = "select top 20000 j.sellsite,extMeachulDate,extOrderserial,extOrderserSeq,extOrgOrderserial "
		sqlStr = sqlStr + " ,extItemNo,extItemCost,extOwnCouponPrice"
		sqlStr = sqlStr + " ,extTenCouponPrice,extReducedPrice,extTenMeachulPrice,extCommPrice"
		sqlStr = sqlStr + " ,extTenJungsanPrice,OrgOrderserial"
		sqlStr = sqlStr + " ,itemid,itemoption,siteNo,extJungsanDate,extJungsanType"
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
	    sqlStr = sqlStr + " where 1=1"
		sqlStr = sqlStr + addSqlStr

    	sqlStr = sqlStr + " order by j.extMeachulDate desc, j.sellsite, j.extOrderserial, j.extOrderserSeq "

		''response.write sqlStr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

		if  not rsget.EOF  then
			GetExtJungsanExcelDown = rsget.getRows
		end if

		rsget.close()

		' redim preserve FItemList(FResultCount)
		' i=0
		' if  not rsget.EOF  then
		' 	rsget.absolutepage = 1
		' 	do until rsget.eof
		' 		set FItemList(i) = new CExtJungsanItem

		' 		FItemList(i).Fsellsite				= rsget("sellsite")
		' 		FItemList(i).FextOrderserial		= rsget("extOrderserial")
		' 		FItemList(i).FextOrderserSeq		= rsget("extOrderserSeq")
		' 		FItemList(i).FextOrgOrderserial		= rsget("extOrgOrderserial")
		' 		FItemList(i).FextItemNo				= rsget("extItemNo")
		' 		FItemList(i).FextItemCost			= rsget("extItemCost")
		' 		FItemList(i).FextReducedPrice		= rsget("extReducedPrice")
		' 		FItemList(i).FextOwnCouponPrice		= rsget("extOwnCouponPrice")
		' 		FItemList(i).FextTenCouponPrice		= rsget("extTenCouponPrice")
		' 		FItemList(i).FextJungsanType		= rsget("extJungsanType")
		' 		FItemList(i).FextCommPrice			= rsget("extCommPrice")
		' 		FItemList(i).FextTenMeachulPrice	= rsget("extTenMeachulPrice")
		' 		FItemList(i).FextTenJungsanPrice	= rsget("extTenJungsanPrice")
		' 		FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
		' 		FItemList(i).FextJungsanDate		= rsget("extJungsanDate")
		' 		FItemList(i).FOrgOrderserial		= rsget("OrgOrderserial")
		' 		FItemList(i).Fitemid				= rsget("itemid")
		' 		FItemList(i).Fitemoption			= rsget("itemoption")
		' 		FItemList(i).FsiteNo				= rsget("siteNo")

		' 		i=i+1
		' 		rsget.moveNext
		' 	loop
		' end if
    end function

	public function GetExtJungsanFixedErrDetailListByOrder()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETLIST_ByOrder] '"&FRectOrderserial&"',"&CHKIIF(FRectItemID="","NULL",FRectItemID)&",'"&FRectItemOption&"'"
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanFixedErrItem

				FItemList(i).Fyyyymm					= db3_rsget("yyyymm")
				FItemList(i).Fsellsite					= db3_rsget("sellsite")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).Fitemid					= db3_rsget("itemid")
				FItemList(i).Fitemoption				= db3_rsget("itemoption")
				FItemList(i).Fitemnosum					= db3_rsget("itemnosum")
				FItemList(i).Freducedsum				= db3_rsget("reducedsum")
				FItemList(i).Fbuycashsum				= db3_rsget("buycashsum")
				FItemList(i).FextItemNoSum				= db3_rsget("extItemNoSum")
				FItemList(i).FextreducedpriceSum		= db3_rsget("extreducedpriceSum")
				FItemList(i).FextTenJungsanPriceSum		= db3_rsget("extTenJungsanPriceSum")
				FItemList(i).Fupddt				= db3_rsget("upddt")
				FItemList(i).FErrAsignMonth		= db3_rsget("ErrAsignMonth")
				FItemList(i).FErrAsignSum		= db3_rsget("ErrAsignSum")
				FItemList(i).Fdiffthis			= db3_rsget("diffthis")

				' FItemList(i).FacctErrsum		= db3_rsget("acctErrsum")
				' FItemList(i).FaccAsgnErrSum		= db3_rsget("accAsgnErrSum")
				' FItemList(i).FaccTTLErrSum		= db3_rsget("accTTLErrSum")
				
				FItemList(i).FOutMallOrderSerial = db3_rsget("OutMallOrderSerial")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanFixedErrDetailList()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETCNT] '"&FRectSellSite&"','"&FRectYYYYMM&"','"&FRectJungsanType&"',"&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&","&CHKIIF(FRectAccerrtype<>"",FRectAccerrtype,"NULL")

		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
		if NOT db3_rsget.Eof then
			FTotalCount = db3_rsget("cnt")
			FdiffnoSum  = db3_rsget("diffnoSum")
			FdiffsumSum = db3_rsget("diffsumSum")
			FErrAsignSum = db3_rsget("ErrAsignSum")
		end if
		db3_rsget.Close

		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_FIXED_DIFF_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectSellSite&"','"&FRectYYYYMM&"','"&FRectJungsanType&"',"&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&","&CHKIIF(FRectAccerrtype<>"",FRectAccerrtype,"NULL")
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanFixedErrItem

				FItemList(i).Fyyyymm					= db3_rsget("yyyymm")
				FItemList(i).Fsellsite					= db3_rsget("sellsite")
				FItemList(i).Forderserial				= db3_rsget("orderserial")
				FItemList(i).Fitemid					= db3_rsget("itemid")
				FItemList(i).Fitemoption				= db3_rsget("itemoption")
				FItemList(i).Fitemnosum					= db3_rsget("itemnosum")
				FItemList(i).Freducedsum				= db3_rsget("reducedsum")
				FItemList(i).Fbuycashsum				= db3_rsget("buycashsum")
				FItemList(i).FextItemNoSum				= db3_rsget("extItemNoSum")
				FItemList(i).FextreducedpriceSum		= db3_rsget("extreducedpriceSum")
				FItemList(i).FextTenJungsanPriceSum		= db3_rsget("extTenJungsanPriceSum")
				FItemList(i).Fupddt				= db3_rsget("upddt")
				FItemList(i).FErrAsignMonth		= db3_rsget("ErrAsignMonth")
				FItemList(i).FErrAsignSum		= db3_rsget("ErrAsignSum")
				FItemList(i).Fdiffthis			= db3_rsget("diffthis")

				FItemList(i).FaccErrNoSum		= db3_rsget("accErrNoSum")
				FItemList(i).FacctErrsum		= db3_rsget("acctErrsum")
				FItemList(i).FaccAsgnErrSum		= db3_rsget("accAsgnErrSum")
				FItemList(i).FaccTTLErrSum		= db3_rsget("accTTLErrSum")
				
				FItemList(i).FOutMallOrderSerial= db3_rsget("OutMallOrderSerial")
				FItemList(i).Fcomment			= db3_rsget("comment")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanErrDetailList()
		dim i, sqlStr

		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_GETCNT] '"&FRectSellSite&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectJungsanType&"',"&CHKIIF(FonlyErrNoExists<>"",1,0)&","&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&""
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
		if NOT db3_rsget.Eof then
			FTotalCount = db3_rsget("cnt")
			FdiffnoSum  = db3_rsget("diffnoSum")
			FdiffsumSum = db3_rsget("diffsumSum")
		end if
		db3_rsget.Close

		if FTotalCount < 1 then
			FResultCount = 0
			exit function
		end if


		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectSellSite&"','"&FRectStartDate&"','"&FRectEndDate&"','"&FRectJungsanType&"',"&CHKIIF(FonlyErrNoExists<>"",1,0)&","&CHKIIF(FRectErrorType<>"",FRectErrorType,"NULL")&""
		db3_rsget.CursorLocation = adUseClient
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanErrItem

				FItemList(i).Fsitename		= db3_rsget("sitename")
				FItemList(i).Fyyyymmdd		= db3_rsget("yyyymmdd")
				FItemList(i).Fauthcode		= db3_rsget("authcode")
				FItemList(i).Foorderserial	= db3_rsget("oorderserial")
				FItemList(i).Fitemid		= db3_rsget("itemid")
				FItemList(i).Fitemoption	= db3_rsget("itemoption")
				FItemList(i).Fdiffno		= db3_rsget("diffno")
				FItemList(i).Fdiffsum		= db3_rsget("diffsum")
				
				FItemList(i).Fjumundiv			= db3_rsget("jumundiv")
				FItemList(i).Flinkorderserial	= db3_rsget("linkorderserial")
				FItemList(i).FregDt				= db3_rsget("regDt")
				FItemList(i).FupdDt				= db3_rsget("updDt")

				FItemList(i).Fcomment			= db3_rsget("comment")
				FItemList(i).Ferrortype			= db3_rsget("errortype")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close

	end function

	public function GetExtJungsanDiff()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		sqlStr = " exec [db_datasummary].[dbo].[usp_Ten_OUT_Jungsan_DIFF_MonthList] '" & FRectSellSite & "', '" & FRectDiffType & "' "

        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.Open  sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = db3_rsget.RecordCount
		FTotalCount  = FResultCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanDiffItem
				FItemList(i).Fsitename       = db3_rsget("sitename")
				FItemList(i).Fyyyymm         = db3_rsget("yyyymm")
				FItemList(i).FTMeachulItem   = db3_rsget("TMeachulItem")
				FItemList(i).FTMeachulDLV    = db3_rsget("TMeachulDLV")
				FItemList(i).FTbuycashItem   = db3_rsget("TbuycashItem")
				FItemList(i).FTbuycashDLV    = db3_rsget("TbuycashDLV")
								
				FItemList(i).FXMeachulItem   = db3_rsget("XMeachulItem")
				FItemList(i).FXMeachulDLV    = db3_rsget("XMeachulDLV")
				FItemList(i).FXJungsanItem   = db3_rsget("XJungsanItem")
				FItemList(i).FXJungsanDLV    = db3_rsget("XJungsanDLV")
				FItemList(i).FregDt          = db3_rsget("regDt")
				FItemList(i).FupdDt      	 = db3_rsget("updDt")

				FItemList(i).FmonthItemDiff  = db3_rsget("monthItemDiff")
				FItemList(i).FmonthdlvDiff   = db3_rsget("monthdlvDiff")
				FItemList(i).FdiffITEMsum    = db3_rsget("diffITEMsum")
				FItemList(i).FdiffDlvsum     = db3_rsget("diffDlvsum")

				FItemList(i).FmonthItemDiffMapErr	= db3_rsget("monthItemDiffMapErr")
				FItemList(i).FmonthdlvDiffTMapErr	= db3_rsget("monthdlvDiffTMapErr")

				FItemList(i).FmonthItemDiffNoExists	= db3_rsget("monthItemDiffNoExists")
				FItemList(i).FmonthdlvDiffTNoExists	= db3_rsget("monthdlvDiffTNoExists")

				FItemList(i).FMonthDiffSum = db3_rsget("MonthDiffSum")
				FItemList(i).FMonthErrAsignSum  = db3_rsget("MonthErrAsignSum")
				FItemList(i).FMonthnotAssignErr = db3_rsget("MonthnotAssignErr")
				FItemList(i).FMonthErrAsignItemSum = db3_rsget("MonthErrAsignItemSum")

				FItemList(i).FMonthErrAsignItemSumReqCheck = db3_rsget("MonthErrAsignItemSumReqCheck")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function


	public function GetExtJungsanDiff_OLD()
	    dim i, sqlStr, addSqlStr

		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_Count] '" & FRectYYYYMM & "', '" & FRectSellSite & "', '" & FRectDiffType & "' "

		''response.write sqlstr & "<Br>"
    	db3_rsget.Open sqlStr,db3_dbget,1
			FTotalCount = db3_rsget("cnt")
		db3_rsget.Close

		if FTotalCount<1 then exit function

		'// ====================================================================
		sqlStr = " exec [db_datamart].[dbo].[usp_Ten_GetExtSiteMeachulDiff_List] '" & FRectYYYYMM & "', '" & FRectSellSite & "', '" & FRectDiffType & "', '" & FPageSize & "', '" & FCurrPage & "' "

        db3_rsget.CursorLocation = adUseClient
    	db3_rsget.CursorType = adOpenStatic
    	db3_rsget.LockType = adLockOptimistic

		db3_rsget.Open sqlStr,db3_dbget,1

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if
		'
		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1
		'
		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
			''rsget.absolutepage = FCurrPage
			do until db3_rsget.eof
				set FItemList(i) = new CExtJungsanDiffItem_OLD

				FItemList(i).Fyyyymm				= db3_rsget("yyyymm")
				FItemList(i).Fsellsite				= db3_rsget("sitename")

				FItemList(i).Forderserial			= db3_rsget("orderserial")
				FItemList(i).FMeachulPriceSUM		= db3_rsget("MeachulPriceSUM")
				FItemList(i).FextMeachulPriceSUM	= db3_rsget("extMeachulPriceSUM")
				FItemList(i).FMeachulPriceSUM1		= db3_rsget("MeachulPriceSUM1")
				FItemList(i).FextMeachulPriceSUM1	= db3_rsget("extMeachulPriceSUM1")
				FItemList(i).FMeachulPriceSUM2		= db3_rsget("MeachulPriceSUM2")
				FItemList(i).FextMeachulPriceSUM2	= db3_rsget("extMeachulPriceSUM2")
				FItemList(i).FMeachulPriceSUM3		= db3_rsget("MeachulPriceSUM3")
				FItemList(i).FextMeachulPriceSUM3	= db3_rsget("extMeachulPriceSUM3")

				i=i+1
				db3_rsget.moveNext
			loop
		end if
		db3_rsget.Close
    end function

	public function GetExtJungsanStatistic()
	    dim i, sqlStr, addSqlStr

		'// ====================================================================
		addSqlStr = ""

		if (FRectStartDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate >= '" + CStr(FRectStartDate) + "' "
		end if

		if (FRectEndDate <> "") then
			addSqlStr = addSqlStr + " and j.extMeachulDate < '" + CStr(FRectEndDate) + "' "
		end if

		if(FRectSellSite <> "") then
		    if (FRectSellSite="ssg6006") or (FRectSellSite="ssg6007") then
		        addSqlStr = addSqlStr + " and j.sellsite+isNull(j.siteno,'') = '" + CStr(FRectSellSite) + "' "
		    else
			    addSqlStr = addSqlStr + " and j.sellsite = '" + CStr(FRectSellSite) + "' "
		    end if
		end if

		'// ====================================================================
		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "

		if (FRectGroupGubun = "sellsite") then
			sqlStr = sqlStr + " 	j.sellsite "
		else
			sqlStr = sqlStr + " 	j.extMeachulDate "
		end if

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extitemCost*j.extItemNo else 0 end) as totExtitemCostProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extReducedPrice*j.extItemNo else 0 end) as totExtReducedPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extOwnCouponPrice*j.extItemNo else 0 end) as totExtOwnCouponPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenCouponPrice*j.extItemNo else 0 end) as totExtTenCouponPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceProduct "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceDeliver "

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extitemCost*j.extItemNo else 0 end) as totExtitemCostDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extReducedPrice*j.extItemNo else 0 end) as totExtReducedPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extOwnCouponPrice*j.extItemNo else 0 end) as totExtOwnCouponPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenCouponPrice*j.extItemNo else 0 end) as totExtTenCouponPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceDeliver "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'D' then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceDeliver "

		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extTenMeachulPrice*j.extItemNo else 0 end) as totExtTenMeachulPriceEtc "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extCommPrice*j.extItemNo else 0 end) as totExtCommPriceEtc "
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') not in ('C', 'D') then j.extTenJungsanPrice*j.extItemNo else 0 end) as totExtTenJungsanPriceEtc "
		sqlStr = sqlStr + " 	, sum(j.extTenMeachulPrice*j.extItemNo) as totExtTenMeachulPrice "
		sqlStr = sqlStr + " 	, sum(j.extCommPrice*j.extItemNo) as totExtCommPrice "
		sqlStr = sqlStr + " 	, sum(j.extTenJungsanPrice*j.extItemNo) as totExtTenJungsanPrice "

		sqlStr = sqlStr + " 	, sum(j.extitemCost*j.extItemNo) as totExtitemCost "
		sqlStr = sqlStr + " 	, sum(j.extReducedPrice*j.extItemNo) as totExtReducedPrice "
		sqlStr = sqlStr + " 	, sum(j.extOwnCouponPrice*j.extItemNo) as totExtOwnCouponPrice "
		sqlStr = sqlStr + " 	, sum(j.extTenCouponPrice*j.extItemNo) as totExtTenCouponPrice "

		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(extJungsanDate,'')<>'' THEN 0 ELSE 1 END) as MiMapp, count(*) as RowCnt "
		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(extJungsanDate,'')='' and IsNull(j.extJungsanType, 'C') = 'C'  THEN 1 ELSE 0 END) as MiMapp_C"
		sqlStr = sqlStr + " 	, sum(case when IsNull(j.extJungsanType, 'C') = 'C' then 1 else 0 END) as RowCnt_C "

		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(OrgOrderserial,'')='' THEN 1 ELSE 0 END) as MiMappOrder"
		sqlStr = sqlStr + " 	, sum(CASE WHEN isNULL(OrgOrderserial,'')='' and IsNull(j.extJungsanType, 'C') = 'C'  THEN 1 ELSE 0 END) as MiMappOrder_C"

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_jungsan.dbo.tbl_xSite_JungsanData j WITH(NOLOCK)"
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "

		sqlStr = sqlStr + addSqlStr

		if (FRectGroupGubun = "sellsite") then
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	j.sellsite "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	j.sellsite "
		else
			sqlStr = sqlStr + " group by "
			sqlStr = sqlStr + " 	j.extMeachulDate "
			sqlStr = sqlStr + " order by "
			sqlStr = sqlStr + " 	j.extMeachulDate desc "
		end if

		' response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CExtJungsanStatisticItem

				if (FRectGroupGubun = "sellsite") then
					FItemList(i).Fsellsite				= rsget("sellsite")
				else
					FItemList(i).FextMeachulDate		= rsget("extMeachulDate")
				end if

				FItemList(i).FtotExtTenMeachulPriceProduct	= rsget("totExtTenMeachulPriceProduct")
				FItemList(i).FtotExtCommPriceProduct			= rsget("totExtCommPriceProduct")
				FItemList(i).FtotExtTenJungsanPriceProduct	= rsget("totExtTenJungsanPriceProduct")

				FItemList(i).FtotExtTenMeachulPriceDeliver	= rsget("totExtTenMeachulPriceDeliver")
				FItemList(i).FtotExtCommPriceDeliver			= rsget("totExtCommPriceDeliver")
				FItemList(i).FtotExtTenJungsanPriceDeliver	= rsget("totExtTenJungsanPriceDeliver")

				FItemList(i).FtotExtTenMeachulPriceEtc			= rsget("totExtTenMeachulPriceEtc")
				FItemList(i).FtotExtCommPriceEtc				= rsget("totExtCommPriceEtc")
				FItemList(i).FtotExtTenJungsanPriceEtc			= rsget("totExtTenJungsanPriceEtc")

				FItemList(i).FtotExtTenMeachulPrice			= rsget("totExtTenMeachulPrice")
				FItemList(i).FtotExtCommPrice					= rsget("totExtCommPrice")
				FItemList(i).FtotExtTenJungsanPrice			= rsget("totExtTenJungsanPrice")
				FItemList(i).FextMiMapping						= rsget("MiMapp")
				FItemList(i).FextRowCount						= rsget("RowCnt")

                FItemList(i).FextMiMapping_C					= rsget("MiMapp_C")
				FItemList(i).FextRowCount_C						= rsget("RowCnt_C")

				FItemList(i).FMiMappOrder					= rsget("MiMappOrder")
				FItemList(i).FMiMappOrder_C						= rsget("MiMappOrder_C")

				''2018/06/26
				FItemList(i).FtotExtitemCostProduct	    = rsget("totExtitemCostProduct")
				FItemList(i).FtotExtReducedPriceProduct	    = rsget("totExtReducedPriceProduct")
				FItemList(i).FtotExtOwnCouponPriceProduct	= rsget("totExtOwnCouponPriceProduct")
				FItemList(i).FtotExtTenCouponPriceProduct	= rsget("totExtTenCouponPriceProduct")

				FItemList(i).FtotExtitemCostDeliver	    = rsget("totExtitemCostDeliver")
				FItemList(i).FtotExtReducedPriceDeliver	    = rsget("totExtReducedPriceDeliver")
				FItemList(i).FtotExtOwnCouponPriceDeliver	= rsget("totExtOwnCouponPriceDeliver")
				FItemList(i).FtotExtTenCouponPriceDeliver	= rsget("totExtTenCouponPriceDeliver")

				FItemList(i).FtotExtitemCost            = rsget("totExtitemCost")
				FItemList(i).FtotExtReducedPrice            = rsget("totExtReducedPrice")
                FItemList(i).FtotExtOwnCouponPrice          = rsget("totExtOwnCouponPrice")
                FItemList(i).FtotExtTenCouponPrice          = rsget("totExtTenCouponPrice")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0

		FdiffnoSum = 0
		FdiffsumSum = 0
		FErrAsignSum = 0
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

End Class

%>
