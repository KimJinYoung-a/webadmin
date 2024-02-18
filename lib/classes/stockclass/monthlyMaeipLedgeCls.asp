<%

Class CMonthlyMaeipLedgeItem
    public FisJungsan
    public Fyyyymm
	public FstockPlace
	public Fshopid
	public FtargetGbn
	public Fitemgubun
    public Flastmwdiv
    public FMakerid

	public Fitemid
	public FitemName
	public Fitemoption
	public FitemoptionName

	public FprevSysStockNo
	public FprevSysStockSum

	public FIpgoNo
	public FIpgoSum
	public FMoveNo
	public FMoveSum
	public FSellNo
	public FSellSum
	public FOffChulNo
	public FOffChulSum
	public FEtcChulNo
	public FEtcChulSum
	public FCsNo
	public FCsSum
	public FLossChulNo
	public FLossChulSum
	public FcurSysStockNo
	public FcurSysStockSum
	public FcurErrRealCheckNo
	public FcurErrRealCheckSum
	public FpurchaseType
	public FpurchaseTypeName

'    public function IsMoveItem()
'        IsMoveItem = false
'
'        if (FisJungsan) then
'            if (Fyyyymm>="2012-01") and (Fyyyymm<"2012-10") and (LCASE(Fmakerid)="ithinkso") then
'                IsMoveItem = true
'            end if
'        else
'            if (Fyyyymm>="2012-01") and (Fyyyymm<"2012-10") and (LCASE(Fmakerid)="ithinkso") then
'                IsMoveItem = true
'            end if
'        end if
'    end function

    public function getTotErrNo()
        getTotErrNo = getDiffNo*-1
    end function

    public function getTotErrSum()
        getTotErrSum = getDiffSum*-1
    end function

    public function getDiffNo()
        getDiffNo = FprevSysStockNo + getIpgoNo + getMoveNo + FSellNo + FOffChulNo + FEtcChulNo + FCsNo + FLossChulNo - FcurSysStockNo
    end function

    public function getDiffSum()
        getDiffSum = FprevSysStockSum + getIpgoSum + getMoveSum + FSellSum + FOffChulSum + FEtcChulSum + FCsSum + FLossChulSum - FcurSysStockSum
    end function

    public function getIpgoNo()
        getIpgoNo = FIpgoNo
    end function

    public function getIpgoSum()
        if isNULL(FIpgoSum) then
            getIpgoSum = 0
        else
            getIpgoSum = FIpgoSum
        end if
    end function

    public function getMoveNo()
        getMoveNo = FMoveNo
    end function

    public function getMoveSum()
        getMoveSum = FMoveSum
    end function

    public function getStockPlaceOrShopid
        if (Fshopid<>"") then
            getStockPlaceOrShopid = Fshopid
        else
            getStockPlaceOrShopid = FstockPlace
        end if
    end function

    public function getBusiName()
        getBusiName=""
        Exit function

        if (FtargetGbn="ON") then
		    getBusiName      = "�¶���"
		elseif (FtargetGbn="OF") then
		    getBusiName      = "��������"
		elseif (FtargetGbn="AC") then
		    getBusiName      = "��ī����"
		elseif (FtargetGbn="IT") then
		    getBusiName      = "���̶��(��)"
		elseif (FtargetGbn="ET") then
		    getBusiName      = "���"
	    elseif (FtargetGbn="EG") then
		    getBusiName      = "EG"
		else
		    getBusiName      = "-"
	    end if
    end function

    public function getItemGubunName()
        if Fitemgubun="10" then
			getITemGubunName = "�Ϲ�"
		elseif Fitemgubun="90" then
			getITemGubunName = "��������"
		elseif Fitemgubun="60" then
			getITemGubunName = "��Ÿ"
		elseif Fitemgubun="70" then
			getITemGubunName = "�Ҹ�ǰ"
		elseif Fitemgubun="80" then
			getITemGubunName = "����ǰ"
		elseif Fitemgubun="85" then
			getITemGubunName = "����ǰ"
		elseif Fitemgubun="97" then
			getITemGubunName = "����"
		elseif Fitemgubun="98" then
			getITemGubunName = "DIY"
		elseif Fitemgubun="99" then
			getITemGubunName = "�Ϲ�"
		elseif Fitemgubun="95" then
			getITemGubunName = "��Ÿ"
		else
			getITemGubunName = "��Ÿ" ''Fitemgubun
		end if
    end function

    public function getMeaipTypeName()

        if Flastmwdiv="M" then
			getMeaipTypeName = "�԰�и���"
		elseif Flastmwdiv="S" then
			getMeaipTypeName = "�Ǹźи���"
		elseif Flastmwdiv="C" then
			getMeaipTypeName = "���и���"
		elseif Flastmwdiv="E" then
			getMeaipTypeName = "��Ÿ����"
		elseif Flastmwdiv="W" then
			getMeaipTypeName = "�԰�и���(W)"
		elseif Flastmwdiv="U" then
			getMeaipTypeName = "��ü<br>(U)"
		elseif Flastmwdiv="Z" then
			getMeaipTypeName = "-<br>(Z)"
		elseif Flastmwdiv="J" then
			getMeaipTypeName = "�Ǹ�(���)�и���"
		elseif Flastmwdiv="B011" then
			getMeaipTypeName = "��Ź�Ǹ�<br>(B011)"
		elseif Flastmwdiv="B012" then
			getMeaipTypeName = "��ü��Ź<br>(B012)"
		elseif Flastmwdiv="B013" then
			getMeaipTypeName = "�����Ź"
		elseif Flastmwdiv="B021" then
			getMeaipTypeName = "��������"
		elseif Flastmwdiv="B022" then
			getMeaipTypeName = "�������"
		elseif Flastmwdiv="B023" then
			getMeaipTypeName = "����������"
		elseif Flastmwdiv="B031" then
			getMeaipTypeName = "������"
		elseif Flastmwdiv="B032" then
			getMeaipTypeName = "���͸���"
		else
			getMeaipTypeName = Flastmwdiv
		end if

    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

End Class

Class CMonthlyMaeipLedge

	public FItemList()
	public FOneItem

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

    public FRectYYYY
	public FRectYYYYMM
	public FRectStockPlace
	public FRectShopid
	public FRectMakerid
    public FRectBySuplyPrice
    public FRectMeaipTp
    public FRectItemgubun
    public FRectTargetGbn

    public FRectSubGrpType
	public FRectShowShopid
	public FRectShowItem
    public FRectOnlyIpgoMeaip

	public FRectShowDiff
	public FRectPriceGubun
	public FRectShowUpbae
	public FRectShowPurchaseType
	public FRectPurchaseType
    Public FRectShowPoint
	public farrlist

    '' ��� ��ġ T �ΰ�� �������� (������ ���� ���ԵǾ� ����) ==> ���Դ� �ٷ� ���� (��꼭 ����ݾ�<>�����԰�)
    function getCaseStrNo(iyyyymm,ifieldNm)
        dim AddCASEStr
        ''�԰� �̵� �и� ==>��� �÷��� ����.
        if (ifieldNm="stIpgoNo") or (ifieldNm="totItemNo") then
			''AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
			AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
        elseif (ifieldNm="stIpgoMoveNo") then
			''AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
			AddCASEStr = " and ((isNULL(m.isMove,0)<>0)  or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"

            ifieldNm = "stIpgoNo"
        end if

        if (ifieldNm<>"curSysStockNo") and (FRectYYYY<>"") then ''�⵵�� �հ�.
            getCaseStrNo = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then IsNull("+ifieldNm+",0) else 0 end"
        else
            getCaseStrNo = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then IsNull("+ifieldNm+",0) else 0 end"
        end if
    end function

    function getCaseStrPrice(iyyyymm,ifieldNmNo,ifieldNmPrc)
        dim AddCASEStr, AddCASENoStr
        ''�԰� �̵� �и�
        if (ifieldNmNo="stIpgoNo") then
			''AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
			AddCASEStr = " and NOT ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
        elseif (ifieldNmNo="stIpgoMoveNo") then
			''AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace<>'T' and m.makerid in ('ithinkso','grandmintfestival','beautifulmintlife') and m.yyyymm>='2012-01' and  m.yyyymm<'2012-10') or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
			AddCASEStr = " and ((isNULL(m.isMove,0)<>0) or (m.stockPlace='S' and i.ipgomwdiv is not NULL and i.lastcentermwdiv is not NULL and IsNull(i.ipgomwdiv, '') = IsNull(i.lastcentermwdiv, '') and i.ipgomwdiv = 'M'))"
            ifieldNmNo = "stIpgoNo"
        end if

		AddCASENoStr = ifieldNmNo
		if (FRectPriceGubun = "V") and (ifieldNmNo = "stIpgoNo" and ifieldNmPrc = "totBuyCash") then
			AddCASENoStr = "1"  ''2014/10/13

			'AddCASENoStr = "i.totitemno" ''2016/06/02
			'ifieldNmPrc = ifieldNmPrc&"/(CASE WHEN "&AddCASENoStr&"=0 THEN 1 ELSE "&AddCASENoStr&" END)" ''2016/06/02

			AddCASENoStr = "(CASE WHEN i.totitemno=0 THEN 1 ELSE i.totitemno END)" ''2016/06/02 ����� ������ 0 �̳� ���԰��� ������ ����..
			ifieldNmPrc = "(CASE WHEN i.totitemno=0 THEN "&ifieldNmPrc&" ELSE "&ifieldNmPrc&"/"&AddCASENoStr&" END)" ''2016/06/02 �����
		end if

        if (ifieldNmNo<>"curSysStockNo") and (FRectYYYY<>"") then
            if (FRectBySuplyPrice=1) then ''Round ���� ���� ����
				if (FRectPriceGubun = "V") then
					'��ո��԰��� �ܰ��� ���� ��ո��԰� ����
					' ==> ����, �Ǹ�(���)�� ������ ��� �Ѿ׿��� ���ް� ����, skyer9, 2015-05-06

					'' getCaseStrPrice = " case "
					'' getCaseStrPrice = " 	when m.stockPlace in ('L', 'S') and LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) "
					'' getCaseStrPrice = " 	when m.stockPlace not in ('L', 'S') and LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then Round("+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0), 0) ELSE IsNull("+ifieldNmPrc+",0) END)*10/11, 0) "
					'' getCaseStrPrice = " 	else 0 end "
					getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				else
					getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0)*10/11) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				end if
            else
               getCaseStrPrice = "case when LEFT(m.yyyymm,4)='"+LEFT(iyyyymm,4)+"' "&AddCASEStr&" then "+AddCASENoStr+"*IsNull("+ifieldNmPrc+",0) else 0 end"
            end if
        else
            if (FRectBySuplyPrice=1) then ''Round ���� ���� ����
				if (FRectPriceGubun = "V") then
					'��ո��԰��� �ܰ��� ���� ��ո��԰� ����
					'' getCaseStrPrice = " case "
					'' getCaseStrPrice = " 	when m.stockPlace in ('L', 'S') and m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) "
					'' getCaseStrPrice = " 	when m.stockPlace not in ('L', 'S') and m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then Round("+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0) ELSE IsNull("+ifieldNmPrc+",0) END)*10/11), 0) "
					'' getCaseStrPrice = " 	else 0 end "
					getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN Round((IsNull("+ifieldNmPrc+",0)*10/11), 0) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				else
					getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*(CASE WHEN m.lastVatinclude='Y' THEN (IsNull("+ifieldNmPrc+",0)*10/11) ELSE IsNull("+ifieldNmPrc+",0) END) else 0 end"
				end if
            else
               getCaseStrPrice = "case when m.yyyymm='"+CStr(iyyyymm)+"' "&AddCASEStr&" then "+AddCASENoStr+"*IsNull("+ifieldNmPrc+",0) else 0 end"
            end if
        end if
    end function

    public Function GetMaeipJungsanSumSubDetail
        FRectSubGrpType = "makerid"

		if (FRectMakerid <> "") then
			FRectSubGrpType = "itemid"
		end if

        call GetMaeipJungsanSum
    end Function

    public Sub GetCurrentStockList()
        dim i,sqlStr

		sqlStr = " select " & vbCrLf
		sqlStr = sqlStr & "     top " & CStr(FPageSize * FCurrPage) & vbCrLf
		sqlStr = sqlStr & " 	isnull(a.warehouseCd,'BLK') as [���Ӽ�] " & vbCrLf
		sqlStr = sqlStr & " 	, left(isnull(isnull(os.RackcodeByOption,i.itemrackcode),''),1) as [���ڵ���ڸ�] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(isnull(os.RackcodeByOption,i.itemrackcode),'') as [���ڵ�] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(os.subRackcodeByOption,'') as [�������ڵ�] " & vbCrLf
		sqlStr = sqlStr & " 	, IsNull(i.makerid, si.makerid) as [�귣��] " & vbCrLf
		sqlStr = sqlStr & " 	, s.itemgubun as [����] " & vbCrLf
		sqlStr = sqlStr & " 	, s.itemid as [��ǰ�ڵ�] " & vbCrLf
		sqlStr = sqlStr & " 	, s.itemoption as [�ɼ��ڵ�] " & vbCrLf
		sqlStr = sqlStr & " 	, db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�] " & vbCrLf
		sqlStr = sqlStr & " 	, replace(isNULL(i.itemname,si.shopitemname),char(9),' ') as [��ǰ��] " & vbCrLf
		sqlStr = sqlStr & " 	, replace(isNULL(isNULL(o.optionname,si.shopitemoptionname),''),char(9),' ') as [�ɼǸ�] " & vbCrLf
		sqlStr = sqlStr & " 	, '' as [�����԰���(����)] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(s.totsysstock,0) as [�ý������] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(s.totsysstock,0)-isnull(a.agvstock,0) as [�ý������(BLK)] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(a.agvstock,0) as [�ý������(AGV)] " & vbCrLf
		sqlStr = sqlStr & " 	, IsNull(i.buycash, si.shopbuyprice) as [������԰�] " & vbCrLf
		sqlStr = sqlStr & " 	, (s.totsysstock*IsNull(i.buycash, si.shopbuyprice)) as [�հ�] " & vbCrLf
		sqlStr = sqlStr & " 	, s.errrealcheckno as [��������] " & vbCrLf
		sqlStr = sqlStr & " 	, s.errbaditemno as [�����ҷ�] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(s.totsysstock,0)+s.errrealcheckno+s.errbaditemno as [�ǻ���ȿ���] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(s.totsysstock,0)+s.errrealcheckno+s.errbaditemno-isnull(a.agvstock,0) as [�ǻ���ȿ���(BLK)] " & vbCrLf
		sqlStr = sqlStr & " 	, isnull(a.agvstock,0) as [�ǻ���ȿ���(AGV)] " & vbCrLf
		sqlStr = sqlStr & " 	, '' as [1�����ĺ���] " & vbCrLf
		sqlStr = sqlStr & " 	, '' as [1�����Ŀ���] " & vbCrLf
		sqlStr = sqlStr & " 	, '' as ������� " & vbCrLf
		sqlStr = sqlStr & " 	, '' as �ǻ翩�� " & vbCrLf
		sqlStr = sqlStr & " 	, '' as [���] " & vbCrLf
		sqlStr = sqlStr & " from " & vbCrLf
		sqlStr = sqlStr & "     [db_summary].[dbo].[tbl_current_logisstock_summary] s with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item i with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun='10' " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid=i.itemid " & vbCrLf
		sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun='10' " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid=o.itemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption=o.itemoption " & vbCrLf
		sqlStr = sqlStr & " 	left join db_item.dbo.tbl_item_option_stock os with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and os.itemgubun='10' " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid = os.itemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption = os.itemoption " & vbCrLf
		sqlStr = sqlStr & " 	left join db_summary.dbo.tbl_current_agvstock_summary as a with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun = a.itemgubun " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid = a.itemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption = a.itemoption " & vbCrLf
		sqlStr = sqlStr & " 	left join [db_shop].[dbo].tbl_shop_item si with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun<>'10' " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun=si.itemgubun " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid=si.shopitemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption=si.itemoption " & vbCrLf
		sqlStr = sqlStr & " 	left join db_summary.dbo.tbl_not_inc_SummaryStock exc with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemgubun=exc.itemgubun " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid=exc.itemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption=exc.itemoption " & vbCrLf
        sqlStr = sqlStr & " 	left join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary ss with(noLock) " & vbCrLf
		sqlStr = sqlStr & " 	on " & vbCrLf
		sqlStr = sqlStr & " 		1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 		and ss.yyyymm = convert(nvarchar(7),getdate(),121) " & vbCrLf
        sqlStr = sqlStr & " 		and s.itemgubun=ss.itemgubun " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemid=ss.itemid " & vbCrLf
		sqlStr = sqlStr & " 		and s.itemoption=ss.itemoption " & vbCrLf
		sqlStr = sqlStr & " where " & vbCrLf
		sqlStr = sqlStr & " 	1 = 1 " & vbCrLf
		sqlStr = sqlStr & " 	--and not (s.itemgubun='10' and s.itemid in (0,11406,6400)) " & vbCrLf
		sqlStr = sqlStr & " 	and (IsNULL(IsNull(i.mwdiv, si.centermwdiv),'Z')='M' or IsNULL(ss.lastmwdiv,'Z')='M') " & vbCrLf			'// ������� �����̰ų� ���������� ��ǰ��
		sqlStr = sqlStr & " 	and ( s.totsysstock<>0 or (s.realstock<>0)) " & vbCrLf
		sqlStr = sqlStr & " 	and exc.itemgubun is NULL " & vbCrLf
		sqlStr = sqlStr & " order by " & vbCrLf
		sqlStr = sqlStr & " 	  s.itemgubun,s.itemid,s.itemoption " & vbCrLf

		''response.write sqlStr & "<br>"
        ''response.end
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

    public Sub Getmonthlystock_notpaging()
        dim i,sqlStr

		SqlStr = "select top " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & "     isnull(a.warehouseCd,'BLK') as [���Ӽ�] "
		sqlStr = sqlStr & "     , left(isnull(isnull(os.RackcodeByOption,i.itemrackcode),''),1) as [���ڵ���ڸ�] "
		sqlStr = sqlStr & "     , isnull(isnull(os.RackcodeByOption,i.itemrackcode),'') as [���ڵ�] "
		sqlStr = sqlStr & "     , isnull(os.subRackcodeByOption,'') as [�������ڵ�] "
		sqlStr = sqlStr & "     , s.LastMakerid as [�귣��] "
		sqlStr = sqlStr & "     , s.itemgubun as [����] "
		sqlStr = sqlStr & "     , s.itemid as [��ǰ�ڵ�] "
		sqlStr = sqlStr & "     , s.itemoption as [�ɼ��ڵ�] "
		sqlStr = sqlStr & "     , db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�] "
		sqlStr = sqlStr & "     , replace(isNULL(i.itemname,si.shopitemname),char(9),' ') as [��ǰ��] "
		sqlStr = sqlStr & "     , replace(isNULL(isNULL(o.optionname,si.shopitemoptionname),''),char(9),' ') as [�ɼǸ�] "
		sqlStr = sqlStr & "     , s.lastIpgoDate  as [�����԰���(����)] "
		sqlStr = sqlStr & "     , isnull(s.totsysstock,0) as [�ý������] "
		sqlStr = sqlStr & "     , isnull(s.totsysstock,0)-isnull(a.agvstock,0) as [�ý������(BLK)] "
		sqlStr = sqlStr & "     , isnull(a.agvstock,0) as [�ý������(AGV)] "
		sqlStr = sqlStr & "     , round(CASE WHEN s.Lastvatinclude='Y' THEN s.avgipgoprice*10/11 ELSE s.avgipgoprice END,0) as [��ո��԰�(�ΰ�������)] "
		sqlStr = sqlStr & "     , (s.totsysstock*round(CASE WHEN s.Lastvatinclude='Y' THEN s.avgipgoprice*10/11 ELSE s.avgipgoprice END,0)) as [�հ�] "
		sqlStr = sqlStr & "     , s.errrealcheckno as [��������] "
		sqlStr = sqlStr & "     , s.errbaditemno as [�����ҷ�] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���(BLK)] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���(AGV)] "
		sqlStr = sqlStr & "     , '' as [1�����ĺ���] "
		sqlStr = sqlStr & "     , '' as [1�����Ŀ���] "
		sqlStr = sqlStr & "     , '' as ������� "
		sqlStr = sqlStr & "     , '' as �ǻ翩�� "
		sqlStr = sqlStr & "     , '' as [���] "
		sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s with(noLock)"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with(noLock)"
		sqlStr = sqlStr & " 	on  s.itemgubun='10' "
		sqlStr = sqlStr & " 	and s.itemid=i.itemid "
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option o with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_option_stock os with(noLock)"
		sqlStr = sqlStr & " 	on os.itemgubun='10'"
		sqlStr = sqlStr & " 	and s.itemid = os.itemid"
		sqlStr = sqlStr & " 	and s.itemoption = os.itemoption"
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_agvstock_summary as a with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun = a.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid = a.itemid"
		sqlStr = sqlStr & " 	and s.itemoption = a.itemoption"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_item si with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun<>'10'"
		sqlStr = sqlStr & " 	and s.itemgubun=si.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=si.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption=si.itemoption"
		sqlStr = sqlStr & " where s.yyyymm='"& FRectYYYYMM &"'"
		sqlStr = sqlStr & " and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
		sqlStr = sqlStr & " and IsNULL(s.lastmwdiv,'Z')='M'"
		sqlStr = sqlStr & " and s.targetGbn not in ('ET', 'EG') "
		sqlStr = sqlStr & " and ( s.totsysstock<>0 or (s.realstock<>0))"
		'sqlStr = sqlStr & " and i.itemid=149987"
		sqlStr = sqlStr & " order by s.itemgubun,s.itemid,s.itemoption"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

	' 1�� ���� ����ľ��� �Ұ�� �����ڷ� ���� ������
    public Sub Getmonthlystock_day1after_notpaging()
        dim i,sqlStr

		SqlStr = "select top " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & "     isnull(a.warehouseCd,'BLK') as [���Ӽ�] "
		sqlStr = sqlStr & "     , left(isnull(isnull(os.RackcodeByOption,i.itemrackcode),''),1) as [���ڵ���ڸ�] "
		sqlStr = sqlStr & "     , isnull(isnull(os.RackcodeByOption,i.itemrackcode),'') as [���ڵ�] "
		sqlStr = sqlStr & "     , isnull(os.subRackcodeByOption,'') as [�������ڵ�] "
		sqlStr = sqlStr & "     , s.LastMakerid as [�귣��] "
		sqlStr = sqlStr & "     , s.itemgubun as [����] "
		sqlStr = sqlStr & "     , s.itemid as [��ǰ�ڵ�] "
		sqlStr = sqlStr & "     , s.itemoption as [�ɼ��ڵ�] "
		sqlStr = sqlStr & "     , db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�] "
		sqlStr = sqlStr & "     , replace(isNULL(i.itemname,si.shopitemname),char(9),' ') as [��ǰ��] "
		sqlStr = sqlStr & "     , replace(isNULL(isNULL(o.optionname,si.shopitemoptionname),''),char(9),' ') as [�ɼǸ�] "
		sqlStr = sqlStr & "     , s.lastIpgoDate  as [�����԰���(����)] "
		sqlStr = sqlStr & "     , isnull(s.totsysstock,0) as [�ý������] "
		sqlStr = sqlStr & "     , isnull(s.totsysstock,0)-isnull(a.agvstock,0) as [�ý������(BLK)] "
		sqlStr = sqlStr & "     , isnull(a.agvstock,0) as [�ý������(AGV)] "
		sqlStr = sqlStr & "     , round(CASE WHEN s.Lastvatinclude='Y' THEN s.avgipgoprice*10/11 ELSE s.avgipgoprice END,0) as [��ո��԰�(�ΰ�������)] "
		sqlStr = sqlStr & "     , (s.totsysstock*round(CASE WHEN s.Lastvatinclude='Y' THEN s.avgipgoprice*10/11 ELSE s.avgipgoprice END,0)) as [�հ�] "
		sqlStr = sqlStr & "     , s.errrealcheckno as [��������] "
		sqlStr = sqlStr & "     , s.errbaditemno as [�����ҷ�] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���(BLK)] "
		sqlStr = sqlStr & "     , '' as [�ǻ���ȿ���(AGV)] "
		sqlStr = sqlStr & "     , isNULL(R.dfNo,0) as [1�����ĺ���] "
		sqlStr = sqlStr & "     , isNULL(R.errrealcheckno,0) as [1�����Ŀ���] "
		sqlStr = sqlStr & "     , '' as ������� "
		sqlStr = sqlStr & "     , '' as �ǻ翩�� "
		sqlStr = sqlStr & "     , '' as [���] "
		sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_accumulated_logisstock_summary s with(noLock)"
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item i with(noLock)"
		sqlStr = sqlStr & " 	on  s.itemgubun='10' "
		sqlStr = sqlStr & " 	and s.itemid=i.itemid "
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option o with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun='10' and s.itemid=o.itemid and s.itemoption=o.itemoption"
		sqlStr = sqlStr & " left join db_item.dbo.tbl_item_option_stock os with(noLock)"
		sqlStr = sqlStr & " 	on os.itemgubun='10'"
		sqlStr = sqlStr & " 	and s.itemid = os.itemid"
		sqlStr = sqlStr & " 	and s.itemoption = os.itemoption"
		sqlStr = sqlStr & " left join db_summary.dbo.tbl_current_agvstock_summary as a with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun = a.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid = a.itemid"
		sqlStr = sqlStr & " 	and s.itemoption = a.itemoption"
		sqlStr = sqlStr & " left join [db_shop].[dbo].tbl_shop_item si with(noLock)"
		sqlStr = sqlStr & " 	on s.itemgubun<>'10'"
		sqlStr = sqlStr & " 	and s.itemgubun=si.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=si.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption=si.itemoption"
		sqlStr = sqlStr & " left join ("
		sqlStr = sqlStr & " 	select  itemgubun,itemid,itemoption"
		sqlStr = sqlStr & " 	,SUM(totsysstock) as dfNo"
		sqlStr = sqlStr & " 	,SUM(errrealcheckno) as errrealcheckno"
		sqlStr = sqlStr & " 	from db_summary.dbo.tbl_daily_logisstock_summary with(noLock)"
        if (FRectYYYYMM = Left(Now, 7)) then
            '// �̹���
		    sqlStr = sqlStr & " 	where yyyymmdd>='" & Left(Now, 10) & "'  "		'--and sysstockno<>realstockno
        else
            '// ������
		    sqlStr = sqlStr & " 	where yyyymmdd>='"& dateadd("m",+1,FRectYYYYMM&"-01") &"'  "		'--and sysstockno<>realstockno
		    sqlStr = sqlStr & " 	and yyyymmdd<'"& dateadd("m",+1,FRectYYYYMM&"-11") &"'"
        end if
		sqlStr = sqlStr & " 	group by itemgubun,itemid,itemoption"
		sqlStr = sqlStr & " ) as R"
		sqlStr = sqlStr & " 	on s.itemgubun=r.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=r.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=r.itemoption"
		sqlStr = sqlStr & " where s.yyyymm='"& FRectYYYYMM &"'"
		'sqlStr = sqlStr & " and s.itemid = 219764"
		sqlStr = sqlStr & " and not (s.itemgubun='10' and s.itemid  in (0,11406,6400))"
		sqlStr = sqlStr & " and IsNULL(s.lastmwdiv,'Z')='M'"
		sqlStr = sqlStr & " and s.targetGbn not in ('ET', 'EG') "
		sqlStr = sqlStr & " and ( s.totsysstock<>0 or (s.realstock<>0) or (s.totsysstock+isNULL(R.dfNo,0))<>0)"
		'sqlStr = sqlStr & " and s.totsysstock<0"
		sqlStr = sqlStr & " order by s.itemgubun,s.itemid,s.itemoption"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

    public Sub GetCurrentShopstockList()
        dim i,sqlStr

		SqlStr = "select top " & CStr(FPageSize * FCurrPage)
        sqlStr = sqlStr & " s.LstMakerid as [�귣��]"
		sqlStr = sqlStr & " ,s.itemgubun as [����]"
		sqlStr = sqlStr & " ,s.itemid as [��ǰ�ڵ�]"
		sqlStr = sqlStr & " ,s.itemoption as [�ɼ��ڵ�]"
		sqlStr = sqlStr & " ,replace(i.shopitemname,char(9),' ') as [��ǰ��]"
		sqlStr = sqlStr & " ,replace(i.shopitemoptionname,char(9),' ') as [�ɼǸ�]"
		sqlStr = sqlStr & " ,s.lastIpgoDateLogics  as [�����԰���(����)]"
		sqlStr = sqlStr & " ,isnull(s.sysstockno,0) as [����(SYS)] "
		'sqlStr = sqlStr & " ,avgshopipgoprice"
		sqlStr = sqlStr & " , round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0) as [���ް�(��ո��԰�)] "
		sqlStr = sqlStr & " ,(s.sysstockno*round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0)) as [�հ�]"
		sqlStr = sqlStr & " ,db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�]"
		sqlStr = sqlStr & " ,s.errrealcheckno "
		sqlStr = sqlStr & " ,c.logischulgo [�����̵��߼���], c.logisreturn [�����ǰ�߼���]"
		sqlStr = sqlStr & " , '' as [1�� �������]	, '' as [1�� �ǻ����]"
        sqlStr = sqlStr & " , s.shopid as [����]"
		sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s with (nolock)"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item i with (nolock)"
		sqlStr = sqlStr & " 	on s.itemgubun=i.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=i.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption=i.itemoption"
		sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_shopstock_summary] c with (nolock)"   ' �̵�, ��ǰ�� ����.
		sqlStr = sqlStr & " 	on s.shopid=c.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=c.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=c.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=c.itemoption"
		sqlStr = sqlStr & " where 1=1"

        if (FRectShopID <> "") then
		    sqlStr = sqlStr & " and s.shopid='" & FRectShopID & "'"
        end if

		sqlStr = sqlStr & " and s.yyyymm=convert(nvarchar(7),getdate(),121)"
		sqlStr = sqlStr & " and s.LstComm_cd='B031'"
		sqlStr = sqlStr & " and (s.sysstockno<>0 or s.realstockno<>0 or isNULL(c.logischulgo,0)<>0 or  isNULL(c.logisreturn,0)<>0 )"
		sqlStr = sqlStr & " order by s.itemgubun,s.itemid,s.itemoption"

        '' response.write "�۾�����!!"
		'' response.write sqlStr & "<br>"
        '' response.end
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

    public Sub Getmonthlyshopstock_notpaging()
        dim i,sqlStr

		SqlStr = "select top " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.LstMakerid as [�귣��]"
		sqlStr = sqlStr & " ,s.itemgubun as [����]"
		sqlStr = sqlStr & " ,s.itemid as [��ǰ�ڵ�]"
		sqlStr = sqlStr & " ,s.itemoption as [�ɼ��ڵ�]"
		sqlStr = sqlStr & " ,replace(i.shopitemname,char(9),' ') as [��ǰ��]"
		sqlStr = sqlStr & " ,replace(i.shopitemoptionname,char(9),' ') as [�ɼǸ�]"
		sqlStr = sqlStr & " ,s.lastIpgoDateLogics  as [�����԰���(����)]"
		sqlStr = sqlStr & " ,isnull(s.sysstockno,0) as [����(SYS)] "
		'sqlStr = sqlStr & " ,avgshopipgoprice"
		sqlStr = sqlStr & " , round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0) as [���ް�(��ո��԰�)] "
		sqlStr = sqlStr & " ,(s.sysstockno*round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0)) as [�հ�]"
		sqlStr = sqlStr & " ,db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�]"
		sqlStr = sqlStr & " ,s.errrealcheckno "
		sqlStr = sqlStr & " ,c.logischulgo [�����̵��߼���], c.logisreturn [�����ǰ�߼���]"
		sqlStr = sqlStr & " , '' as [1�� �������]	, '' as [1�� �ǻ����]"
		sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s with (nolock)"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item i with (nolock)"
		sqlStr = sqlStr & " 	on s.itemgubun=i.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=i.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption=i.itemoption"
		sqlStr = sqlStr & " left join [db_summary].[dbo].[tbl_current_shopstock_summary] c with (nolock)"   ' �̵�, ��ǰ�� ����.
		sqlStr = sqlStr & " 	on s.shopid=c.shopid"
		sqlStr = sqlStr & " 	and s.itemgubun=c.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=c.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=c.itemoption"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and s.shopid='streetshop011'"
		sqlStr = sqlStr & " and s.yyyymm='"& FRectYYYYMM &"'"
		sqlStr = sqlStr & " and s.LstComm_cd='B031'"
		sqlStr = sqlStr & " and (s.sysstockno<>0 or s.realstockno<>0 or isNULL(c.logischulgo,0)<>0 or  isNULL(c.logisreturn,0)<>0 )"
		sqlStr = sqlStr & " order by s.itemgubun,s.itemid,s.itemoption"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

	' 1�� ���� ����ľ��� �Ұ�� �����ڷ� ���� ������
    public Sub Getmonthlyshopstock_day1after_notpaging()
        dim i,sqlStr

		SqlStr = "select top " & CStr(FPageSize * FCurrPage)
		sqlStr = sqlStr & " s.LstMakerid as [�귣��]"
		sqlStr = sqlStr & " ,s.itemgubun as [����]"
		sqlStr = sqlStr & " ,s.itemid as [��ǰ�ڵ�]"
		sqlStr = sqlStr & " ,s.itemoption as [�ɼ��ڵ�]"
		sqlStr = sqlStr & " ,replace(i.shopitemname,char(9),' ') as [��ǰ��]"
		sqlStr = sqlStr & " ,replace(i.shopitemoptionname,char(9),' ') as [�ɼǸ�]"
		sqlStr = sqlStr & " ,s.lastIpgoDateLogics  as [�����԰���(����)]"
		sqlStr = sqlStr & " ,isnull(s.realstockno,0) as [����(REAL)]"
		'sqlStr = sqlStr & " ,avgshopipgoprice"
		sqlStr = sqlStr & " , round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0) as [���ް�(��ո��԰�)] "
		sqlStr = sqlStr & " ,(s.realstockno*round(CASE WHEN s.Lstvatinclude='Y' THEN s.avgshopipgoprice*10/11 ELSE s.avgshopipgoprice END,0)) as [�հ�]"
		sqlStr = sqlStr & " ,db_storage.dbo.uf_getTenBarCodeType(s.itemgubun,s.itemid,s.itemoption) as [���ڵ�]"
		sqlStr = sqlStr & " ,s.errrealcheckno"
		sqlStr = sqlStr & " , '' as [�����̵��߼���], '' as [�����ǰ�߼���]"
		sqlStr = sqlStr & " , isNULL(R.dfNo,0) [1�� �������]"		', isNULL(R.errrealcheckno,0) [1�� �ǻ����]
		sqlStr = sqlStr & " from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary s with (nolock)"
		sqlStr = sqlStr & " left join db_shop.dbo.tbl_shop_item i with (nolock)"
		sqlStr = sqlStr & " 	on s.itemgubun=i.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=i.shopitemid"
		sqlStr = sqlStr & " 	and s.itemoption=i.itemoption"
		sqlStr = sqlStr & " left join ("
		sqlStr = sqlStr & " 	select  itemgubun,itemid,itemoption"
		sqlStr = sqlStr & " 	,SUM(sysstockno) as dfNo"
		sqlStr = sqlStr & " 	,SUM(errrealcheckno) as errrealcheckno"
		sqlStr = sqlStr & " 	from db_summary.dbo.tbl_daily_shopstock_summary with (nolock)"
		sqlStr = sqlStr & " 	where 1=1"
		sqlStr = sqlStr & " 	and shopid='streetshop011'"
		sqlStr = sqlStr & " 	and yyyymmdd>='"& dateadd("m",+1,FRectYYYYMM&"-01") &"'"  'and sysstockno<>realstockno
		sqlStr = sqlStr & " 	and yyyymmdd<'"& dateadd("m",+1,FRectYYYYMM&"-11") &"'"
		sqlStr = sqlStr & " 	group by itemgubun,itemid,itemoption"
		sqlStr = sqlStr & " ) as R"
		sqlStr = sqlStr & " 	on s.itemgubun=r.itemgubun"
		sqlStr = sqlStr & " 	and s.itemid=r.itemid"
		sqlStr = sqlStr & " 	and s.itemoption=r.itemoption"
		sqlStr = sqlStr & " where 1=1"
		sqlStr = sqlStr & " and s.shopid='streetshop011'"
		sqlStr = sqlStr & " and s.yyyymm='"& FRectYYYYMM &"'"
		sqlStr = sqlStr & " and s.LstComm_cd='B031'"
		sqlStr = sqlStr & " and (s.sysstockno<>0 or s.realstockno<>0"
		sqlStr = sqlStr & " 	or (s.sysstockno+isNULL(R.dfNo,0))<>0"
		sqlStr = sqlStr & " )"
		sqlStr = sqlStr & " order by s.itemgubun,s.itemid,s.itemoption"

		'response.write sqlStr & "<br>"
		rsget.pagesize = FPageSize
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsget.RecordCount
        FResultCount = rsget.RecordCount

        if FResultCount<1 then FResultCount=0

        if  not rsget.EOF  then
            farrlist = rsget.getrows()
        end if
        rsget.close
    end sub

    public Function GetMaeipJungsanSum
        dim sqlStr, addSql, i
		dim prevYYYYMM

        IF (FRectYYYY>="2014") then FRectYYYY="1998"

        IF (FRectYYYY<>"") then ''�⵵�� ��ȸ�ΰ��
            FRectYYYYMM=FRectYYYY+"-12"
            prevYYYYMM = Left(dateAdd("m",-1,FRectYYYY+"-01-01"),7)
			if FRectYYYY=cStr(year(date)) then
				FRectYYYYMM = Left(dateAdd("m",-1,left(date,7)+"-01"),7)
			end if
        else
		    prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)
	    end if

		''prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)

        addSql = " 	from db_summary.dbo.tbl_monthly_jungsanSum"
        addSql = addSql + " 	where 1=1"
        addSql = addSql + " 	and jGubun<>'CC'" '' ������� ǥ�þ���
        addSql = addSql + " 	and jTaxType<>'03'" '' ��õ¡�� ����  ''addSql = addSql + " 	and itemgubun<>'97'" '' ���´� ���� �ƴ�
        addSql = addSql + " 	and NOT (jmakerid in ('ithinkso','grandmintfestival','beautifulmintlife') and (yyyymm>='2012-01') and (yyyymm<'2013-11') and jMwdiv<>'M') " ''�ϴ� ���̶�� ���� ���� ���� ����.. 2012-01~2012-09 ������ ���곻������ ��������.


        if (FRectOnlyIpgoMeaip="on") then
            addSql = addSql + " 	and jMwdiv='M'"  '' �԰�� ������ ǥ�þ���
        else
            addSql = addSql + " 	and jMwdiv<>'M'"  '' �԰�� ������ ǥ�þ���
        end if

        addSql = addSql + " 	and yyyymm>'"&prevYYYYMM&"' and yyyymm<='"&FRectYYYYMM&"'"  ''������ = �� ����
        if (FRectMakerid<>"") then
            addSql = addSql + " 	and jmakerid='"&FRectMakerid&"'"
        end if

        if (FRectItemGubun<>"") then
            addSql = addSql + " 	and itemgubun='"&FRectItemGubun&"'"
        end if

		if (FRectShopid <> "") then
			addSql = addSql + " 	and jShopid='"&FRectShopid&"'"
		end if

        if (FRectTargetGbn<>"") then
            addSql = addSql + " 	and jtargetGbn='"&FRectTargetGbn&"'"
        end if

        if (FRectMeaipTp<>"") then
            addSql = addSql + " 	and jMwdiv='"&FRectMeaipTp&"'"
        end if

        ''if (FRectStockPlace="L") then
        ''    addSql = addSql + " and ((jtargetGbn<>'OF') or ( (jtargetGbn='OF') and (jMwdiv='M' and dtlgubuncd in ('B021')) ) )"
        ''elseif (FRectStockPlace="S") then
        ''    addSql = addSql + " and NOT ((jtargetGbn<>'OF') or ((jtargetGbn='OF') and (jMwdiv='M' and dtlgubuncd in ('B021'))) )"
        ''end if

        addSql = addSql + " 	group by itemgubun,jMwdiv"  '''jtargetGbn,
		if (FRectShowShopid <> "") then
			addSql = addSql + " 	,jshopid "
		end if
        IF (FRectSubGrpType="makerid") then
		    addSql = addSql + " 	, jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			addSql = addSql + " 	, jmakerid, itemid, itemoption "
		end if


        if (FRectSubGrpType<>"") then
    		sqlStr = " SELECT top 1 (COUNT(*) OVER ()) as CNT, CEILING(CAST((COUNT(*) OVER ()) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
    		sqlStr = sqlStr + addSql
    		''response.write sqlStr
    		''response.end

        	rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    			FTotalCount = rsget("cnt")
    			FTotalPage = rsget("totPg")
    		else
    			FTotalCount = 0
    			FTotalPage = 0
    		end if
        	rsget.Close

        	'������������ ��ü ���������� Ŭ �� �Լ�����
        	if CLng(FCurrPage)>CLng(FTotalPage) then
        		FResultCount = 0
        		exit function
        	end if
        end if

		sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
		sqlStr = sqlStr + " 	'" + CStr(FRectYYYYMM) + "' as yyyymm "
		sqlStr = sqlStr + " 	,'"&FRectStockPlace&"' as stockPlace "

		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,jshopid as shopid "
		end if

        '''sqlStr = sqlStr + " 	,jtargetGbn as targetGbn"
        sqlStr = sqlStr + " 	,itemgubun"
        IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	, jmakerid, itemid, itemoption "
		end if
        sqlStr = sqlStr + " 	,jMwdiv as lastmwdiv"
        sqlStr = sqlStr + " 	,0 as prevSysStockNo"
        sqlStr = sqlStr + " 	,0 as prevSysStockSum"
        if (FRectBySuplyPrice="1") then
            sqlStr = sqlStr + " 	,sum(jtotItemno) as IpgoNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END) as IpgoSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotItemno*-1 ELSE 0 END) as SellNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as SellSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotItemno*-1 ELSE 0 END) as OffChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as OffChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotItemno*-1 ELSE 0 END) as EtcChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as EtcChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotItemno*-1 ELSE 0 END) as CsNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as CsSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotItemno*-1 ELSE 0 END) as LossChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN (CASE WHEN jTaxType='02' THEN jtotBuycash ELSE jtotBuycash/11*10 END)*-1 ELSE 0 END) as LossChulSum"
        else
            sqlStr = sqlStr + " 	,sum(jtotItemno) as IpgoNo"
            sqlStr = sqlStr + " 	,sum(jtotBuycash) as IpgoSum"

            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotItemno*-1 ELSE 0 END) as SellNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='S' THEN jtotBuycash*-1 ELSE 0 END) as SellSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotItemno*-1 ELSE 0 END) as OffChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='C' THEN jtotBuycash*-1 ELSE 0 END) as OffChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotItemno*-1 ELSE 0 END) as EtcChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='E' THEN jtotBuycash*-1 ELSE 0 END) as EtcChulSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotItemno*-1 ELSE 0 END) as CsNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='T' THEN jtotBuycash*-1 ELSE 0 END) as CsSum"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotItemno*-1 ELSE 0 END) as LossChulNo"
            sqlStr = sqlStr + " 	,sum(CASE WHEN jMwdiv='L' THEN jtotBuycash*-1 ELSE 0 END) as LossChulSum"
        end if
        sqlStr = sqlStr + " 	,0 as curSysStockNo"
        sqlStr = sqlStr + " 	,0 as curSysStockSum"
        sqlStr = sqlStr + " 	,0 as curErrRealCheckNo"
        sqlStr = sqlStr + " 	,0 as curErrRealCheckSum"

		sqlStr = sqlStr + addSql

        sqlStr = sqlStr + " 	order by (CASE WHEN itemgubun='00' THEN '999' ELSE itemgubun END) asc ,jMwdiv desc"  '''(CASE WHEN jtargetGbn='TT' THEN 'AA' ELSE jtargetGbn END)  desc,
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,jshopid "
		end if
        IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,jmakerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	, jmakerid, itemid, itemoption "
		end if

		'rw 	sqlStr

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
				set FItemList(i) = new CMonthlyMaeipLedgeItem
                FItemList(i).FisJungsan         = true
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				'FItemList(i).FMoveNo     		= rsget("MoveNo")
				'FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("jmakerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("jmakerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
                end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    public Function GetMaeipLedgeSUMSubDetail
        FRectSubGrpType = "makerid"

		if (FRectMakerid <> "" or FRectShowItem<>"") then
			FRectSubGrpType = "itemid"
		end if

        call GetMaeipLedgeSUM
    end Function

    public Function GetMaeipLedgeSUM
		dim sqlStr, addSql, i
		dim prevYYYYMM

		'IF (FRectYYYY<>"") and (FRectYYYY>="2014") then FRectYYYY="1998"  ''�ּ�ó�� 2016/01/26 �����û

        if (FRectYYYY<>"") then ''�⵵�� ��ȸ�ΰ��
            FRectYYYYMM=FRectYYYY+"-12"
            prevYYYYMM = Left(dateAdd("m",-1,FRectYYYY+"-01-01"),7)
			if FRectYYYY=cStr(year(date)) then
				FRectYYYYMM = Left(dateAdd("m",-1,left(date,7)+"-01"),7)
			end if
        else
		    prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)
	    end if

		addSql = " from "
		addSql = addSql + " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail m with (nolock) "
		addSql = addSql + " 	left join db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i with (nolock) "
		addSql = addSql + " 	on "
		addSql = addSql + " 		1 = 1 "
		if (FRectSubGrpType<>"") then
			'�ӽ����̺� �̹� ������
			'addSql = addSql + " 		and i.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		end if
		addSql = addSql + " 		and m.yyyymm = i.yyyymm "
		addSql = addSql + " 		and m.stockPlace = i.stockPlace "
		addSql = addSql + " 		and m.shopid = i.shopid "
		addSql = addSql + " 		and m.itemgubun = i.itemgubun "
		addSql = addSql + " 		and m.itemid = i.itemid "
		addSql = addSql + " 		and m.itemoption = i.itemoption "
		addSql = addSql + " 		and i.ipgoMWDIV = 'M' "

		if FRectShowPurchaseType <> "" then
			addSql = addSql + " 	LEFT JOIN db_partner.dbo.tbl_partner p  with (nolock) on 1=1 "
			addSql = addSql + " 	and m.makerid = p.id "

			addSql = addSql + "		LEFT JOIN db_partner.dbo.tbl_partner_comm_code pc with (nolock) "
			addSql = addSql + "			on pc.pcomm_group = 'purchasetype' "
			addSql = addSql + "			and pc.pcomm_cd = convert(varchar(16),p.purchasetype) "
		end if

		if FRectShowUpbae <> "" or FRectSubGrpType = "itemid" then
			addSql = addSql + " 	left join [db_item].[dbo].[tbl_item] ii with (nolock) "
			addSql = addSql + " 	on "
			addSql = addSql + " 		1 = 1 "
			addSql = addSql + " 		and m.itemgubun = '10' "
			addSql = addSql + " 		and m.itemid = ii.itemid "
		end if

		if FRectSubGrpType = "itemid" then
			addSql = addSql + " 	left join db_item.dbo.tbl_item_option o with (nolock) "
			addSql = addSql + " 	on m.itemgubun = '10' "
			addSql = addSql + " 		and m.itemid = o.itemid "
			addSql = addSql + " 		and m.itemoption = o.itemoption "

			addSql = addSql + "		LEFT JOIN db_shop.dbo.tbl_shop_item si with (nolock) "
			addSql = addSql + "		on m.itemgubun <> '10' "
			addSql = addSql + "			and m.itemgubun = si.itemgubun "
			addSql = addSql + "			and m.itemid = si.shopitemid "
			addSql = addSql + "			and m.itemoption = si.itemoption "
		end if

		''addSql = addSql + " 	left join ( "
		''addSql = addSql + " 		select yyyymm, 'L' as stockPlace, '' as shopid, itemgubun, itemid, itemoption, sum(totItemNo) as totShopMoveItemNo "
		''addSql = addSql + " 		from "
		''addSql = addSql + " 		db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum "
		''addSql = addSql + " 		where yyyymm = '" + CStr(FRectYYYYMM) + "' and stockPlace = 'S' and ipgomwdiv = 'M' and ipgomwdiv = lastcentermwdiv "
		''addSql = addSql + " 		and shopid in ('streetshop011', 'streetshop014', 'streetshop018', 'streetshop019', 'streetshop099', 'streetshop100', 'streetshop101', 'streetshop900', 'streetshop999', 'wholesale1030') "
		''addSql = addSql + " 		group by yyyymm, stockPlace, itemgubun, itemid, itemoption "
		''addSql = addSql + " 	) T "
		''addSql = addSql + " 	on "
		''addSql = addSql + " 		1 = 1 "
		''addSql = addSql + " 		and m.yyyymm = T.yyyymm "
		''addSql = addSql + " 		and m.stockPlace = T.stockPlace "
		''addSql = addSql + " 		and m.shopid = T.shopid "
		''addSql = addSql + " 		and m.itemgubun = T.itemgubun "
		''addSql = addSql + " 		and m.itemid = T.itemid "
		''addSql = addSql + " 		and m.itemoption = T.itemoption "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and m.yyyymm>='"&prevYYYYMM&"' and m.yyyymm<='"&FRectYYYYMM&"'"
		addSql = addSql + " 	and m.targetGbn not in ('ET','EG')"
		addSql = addSql + " 	and m.etcjungsantype in (1,4)"
		''addSql = addSql + " 	and m.lastmwdiv not in ('B012')"

        addSql = addSql + " 	and NOT (m.lastmwdiv='B013' and m.targetGbn<>'IT')"               ''�����Ź�� IT��
        addSql = addSql + " 	and m.lastmwdiv not in ('W','B012','B011')"                     ''��ü��Ź ���� //����ڻ��� ������ �ƴ� CASE (����ڻ� ���·� �Ѹ����) W����

		if (FRectStockPlace <> "") then
			addSql = addSql + " 	and m.stockPlace = '" + CStr(FRectStockPlace) + "' "
		end if

        if (FRectMakerid<>"") then
            addSql = addSql + " 	and m.makerid='"&FRectMakerid&"'"
        end if

        if (FRectItemGubun<>"") then
            addSql = addSql + " 	and m.itemgubun='"&FRectItemGubun&"'"
        end if

		if (FRectShopid <> "") then
			addSql = addSql + " 	and m.shopid='"&FRectShopid&"'"
		end if

        if (FRectTargetGbn<>"") then
            addSql = addSql + " 	and m.targetGbn='"&FRectTargetGbn&"'"
        end if

        if (FRectMeaipTp<>"") then
            ''addSql = addSql + " 	and ((isNULL(m.lastmwdiv,'unknown')='"&FRectMeaipTp&"') or (isNULL(m.lastmwdiv,'unknown')='W' and '" + CStr(FRectMeaipTp) + "' = 'M')) "
        end if

		if FRectShowUpbae <> "" then
			addSql = addSql + " 	and ii.mwdiv = 'U' "
			''addSql = addSql + " 	and m.itemgubun = '10' "
		end if

		if FRectPurchaseType <> "" then
			if (FRectPurchaseType = "101") then
				addSql = addSql + " 	and p.PurchaseType <> '1' "
			else
				addSql = addSql + " 	and p.PurchaseType = '" & FRectPurchaseType & "' "
			end if
		end if

        if (FRectShowPoint<>"") then
            addSql = addSql + " 	and i.totitemno <> 0 "
            addSql = addSql + " 	and totBuyCash <> 0 "
            addSql = addSql + " 	and (totBuyCash/i.totitemno) <> Floor((totBuyCash/i.totitemno)) "
        end if

		addSql = addSql + " group by "
		addSql = addSql + " 	m.stockPlace "
		''addSql = addSql + " 	,m.targetGbn "
		addSql = addSql + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			addSql = addSql + " 	,m.shopid "
		end if

		''addSql = addSql + " 	,isNULL(lastmwdiv,'unknown') "
		IF (FRectSubGrpType="makerid") then
		    addSql = addSql + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			addSql = addSql + " 	,m.makerid, m.itemid, m.itemoption, isNull(ii.itemname,isNull(si.shopitemname,'')), isNull(o.optionname,'') "
		end if
		if FRectShowPurchaseType <> "" then
			addSql = addSql + " 	,p.purchaseType, pc.pcomm_name "
		end if

		'// having
		if (FRectShowDiff <> "") then
			'addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' then IpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			if (FRectYYYY<>"") then
				addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when LEFT(m.yyyymm,4) = '" + CStr(FRectYYYY) + "' and m.IpgoNo <> 0 then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			else
				addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' and m.IpgoNo <> 0 then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
		    end if
		end if


		sqlStr = " SELECT top 1 (COUNT(*) OVER ()) as CNT, CEILING(CAST((COUNT(*) OVER ()) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + addSql
	''response.write sqlStr
	''response.end

        if (FRectSubGrpType<>"") then
        	rsget.CursorLocation = adUseClient
            rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    		if  not rsget.EOF  then
    			FTotalCount = rsget("cnt")
    			FTotalPage = rsget("totPg")
    		else
    			FTotalCount = 0
    			FTotalPage = 0
    		end if
        	rsget.Close


        	'������������ ��ü ���������� Ŭ �� �Լ�����
        	if CLng(FCurrPage)>CLng(FTotalPage) then
        		FResultCount = 0
        		exit function
        	end if
        end if

        if (FRectSubGrpType="") then ''2016/05/09 �ӽ����̺� ������� ����.. by eastone

            addSql = replace(addSql,"db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum","#Stock_IpgoLedger_Sum_TMP")

            if (FRectYYYY<>"") then
				'����
				sqlStr = " SET NOCOUNT ON ;select i.yyyymm, i.stockPlace, i.shopid, i.itemgubun, i.itemid, i.itemoption, i.ipgoMWDIV, i.lastcentermwdiv "
				sqlStr = sqlStr + " 	,sum(totitemno) as totitemno "
				sqlStr = sqlStr + " 	,sum(totBuyCash) as totBuyCash "
				sqlStr = sqlStr + " INTO #Stock_IpgoLedger_Sum_TMP "
				sqlStr = sqlStr + " FROM db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i "
				sqlStr = sqlStr + " WHERE 1 = 1 "
				sqlStr = sqlStr + " AND i.yyyymm >= '"&prevYYYYMM&"' AND i.yyyymm <= '"&FRectYYYYMM&"' "
				sqlStr = sqlStr + " AND i.ipgoMWDIV = 'M' "
				sqlStr = sqlStr + " group by i.yyyymm, i.stockPlace, i.shopid, i.itemgubun, i.itemid, i.itemoption, i.ipgoMWDIV, i.lastcentermwdiv "
			else
				'����
				sqlStr = " SET NOCOUNT ON ;select * into #Stock_IpgoLedger_Sum_TMP "
				sqlStr = sqlStr + " from  db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i "
				sqlStr = sqlStr + " where 1=1 "
				sqlStr = sqlStr + " and i.yyyymm = '"&FRectYYYYMM&"' "
				sqlStr = sqlStr + " and i.ipgoMWDIV = 'M' " &VBCRLF
			end if

            ''2016/08/08
            if (FRectShowShopid <> "") then
                addSql = replace(addSql,"db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail","#Stock_MaeipLedger_Detail_TMP")
                sqlStr = sqlStr + " ;select * into #Stock_MaeipLedger_Detail_TMP "
                sqlStr = sqlStr + " from  db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail m "
                sqlStr = sqlStr + " where 1=1 "
                sqlStr = sqlStr + " and m.yyyymm>='"&prevYYYYMM&"' and m.yyyymm<='"&FRectYYYYMM&"'"
                sqlStr = sqlStr + " 	and m.targetGbn not in ('ET','EG')"
    		    sqlStr = sqlStr + " 	and m.etcjungsantype in (1,4)"
    		    sqlStr = sqlStr + " 	and NOT (m.lastmwdiv='B013' and m.targetGbn<>'IT')"               ''�����Ź�� IT��
                sqlStr = sqlStr + " 	and m.lastmwdiv not in ('W','B012','B011')" &VBCRLF               ''��ü��Ź ���� //����ڻ��� ������ �ƴ� CASE (����ڻ� ���·� �Ѹ����) W����
            end if

            sqlStr = sqlStr + "; select  "
        else
			sqlStr = " select "
	    end if
		sqlStr = sqlStr + " 	'" + CStr(FRectYYYYMM) + "' as yyyymm "
		sqlStr = sqlStr + " 	,m.stockPlace "
		''sqlStr = sqlStr + " 	,m.targetGbn "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,m.makerid as makerid " ''�ٸ� �귣�� �ϳ��� ǥ��
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	, m.makerid, m.itemid, m.itemoption "
			sqlStr = sqlStr + "		, replace(replace(isNull(ii.itemname,isNull(si.shopitemname,'')),char(9),''),',','_') as itemname "
			sqlStr = sqlStr + "		, replace(replace(isNull(o.optionname,''),char(9),''),',','_') as itemOptionName "
		end if

		''2015-03-17, skyer9
		''sqlStr = sqlStr + " 	,'M' as lastmwdiv"
		sqlStr = sqlStr + " 	,(case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) as lastmwdiv "

		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") as prevSysStockNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"IpgoNo")&") as IpgoNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoMoveNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as MoveNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") as SellNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as OffChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") as EtcChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") as CsNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") as LossChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&") as curSysStockNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curErrRealCheckNo")&") as curErrRealCheckNo "

		if (FRectPriceGubun = "V") then
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"totItemNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","avgIpgoPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","totBuyCash")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","avgIpgoPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as MoveSum " '// totBuyCash : ����-���� ���Ա��� ������ ��� ������ո��԰�
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","avgIpgoPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","avgIpgoPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","avgIpgoPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","avgIpgoPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","avgIpgoPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","avgIpgoPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","avgIpgoPrice")&") as curErrRealCheckSum "
		else
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","lastbuyPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","lastbuyPrice")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","lastbuyPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as MoveSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","lastbuyPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","lastbuyPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","lastbuyPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","lastbuyPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","lastbuyPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","lastbuyPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","lastbuyPrice")&") as curErrRealCheckSum "
		end if

		if FRectShowPurchaseType <> "" then
			sqlStr = sqlStr + " ,p.purchaseType, isNull(pc.pcomm_name,'') as purchaseTypeName "
		end if

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	(case when (case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) = 'M' then 1 else 100 end) "
		sqlStr = sqlStr + " 	,m.stockPlace "
		''sqlStr = sqlStr + " 	,m.targetGbn desc "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		''sqlStr = sqlStr + " 	,lastmwdiv "
		IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	,m.itemid, m.itemoption, isNull(ii.itemname,isNull(si.shopitemname,'')), isNull(o.optionname,'') "
		end if

		'����¡
		sqlStr = sqlStr + " OFFSET " & CStr(FPageSize*(FCurrPage-1)) & " ROWS FETCH NEXT " & CStr(FPageSize) & " ROWS ONLY "

        if (FRectSubGrpType="") then
            sqlStr = sqlStr + " ; drop table #Stock_IpgoLedger_Sum_TMP "
            if (FRectShowShopid <> "") then
                sqlStr = sqlStr + "; drop table #Stock_MaeipLedger_Detail_TMP"
            end if
        end if
	 ''response.write sqlStr   ''IDX_tbl_monthly_Stock_IpgoLedger_Sum_yyyymm_ipgoMWDIV 2015/06/01
	 'response.end

''rw sqlStr
''dbget.close() : response.end

		dbget.CommandTimeout = 60*2   ' 2��
		rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			do until rsget.eof
				set FItemList(i) = new CMonthlyMaeipLedgeItem

                FItemList(i).FisJungsan         = false
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				FItemList(i).FMoveNo     		= rsget("MoveNo")
				FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("makerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("makerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).FitemName     	= rsget("itemName")
					FItemList(i).Fitemoption    = rsget("itemoption")
					FItemList(i).FitemoptionName    = rsget("itemoptionName")
                end if
				if FRectShowPurchaseType <> "" then
					FItemList(i).FpurchaseType	= rsget("purchaseType")
					FItemList(i).FpurchaseTypeName	= rsget("purchaseTypeName")
				end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	public Function GetMaeipLedgeSUM_20160509
		dim sqlStr, addSql, i
		dim prevYYYYMM

		'IF (FRectYYYY<>"") and (FRectYYYY>="2014") then FRectYYYY="1998"  ''�ּ�ó�� 2016/01/26 �����û

        if (FRectYYYY<>"") then ''�⵵�� ��ȸ�ΰ��
            FRectYYYYMM=FRectYYYY+"-12"
            prevYYYYMM = Left(dateAdd("m",-1,FRectYYYY+"-01-01"),7)
        else
		    prevYYYYMM = Left(dateAdd("m",-1,FRectYYYYMM+"-01"),7)
	    end if

		addSql = " from "
		addSql = addSql + " 	db_summary.dbo.tbl_monthly_Stock_MaeipLedger_Detail m "
		addSql = addSql + " 	left join db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum i "
		addSql = addSql + " 	on "
		addSql = addSql + " 		1 = 1 "
		addSql = addSql + " 		and i.yyyymm = '" + CStr(FRectYYYYMM) + "' "
		addSql = addSql + " 		and m.yyyymm = i.yyyymm "
		addSql = addSql + " 		and m.stockPlace = i.stockPlace "
		addSql = addSql + " 		and m.shopid = i.shopid "
		addSql = addSql + " 		and m.itemgubun = i.itemgubun "
		addSql = addSql + " 		and m.itemid = i.itemid "
		addSql = addSql + " 		and m.itemoption = i.itemoption "
		addSql = addSql + " 		and i.ipgoMWDIV = 'M' "
		''addSql = addSql + " 	left join ( "
		''addSql = addSql + " 		select yyyymm, 'L' as stockPlace, '' as shopid, itemgubun, itemid, itemoption, sum(totItemNo) as totShopMoveItemNo "
		''addSql = addSql + " 		from "
		''addSql = addSql + " 		db_summary.dbo.tbl_monthly_Stock_IpgoLedger_Sum "
		''addSql = addSql + " 		where yyyymm = '" + CStr(FRectYYYYMM) + "' and stockPlace = 'S' and ipgomwdiv = 'M' and ipgomwdiv = lastcentermwdiv "
		''addSql = addSql + " 		and shopid in ('streetshop011', 'streetshop014', 'streetshop018', 'streetshop019', 'streetshop099', 'streetshop100', 'streetshop101', 'streetshop900', 'streetshop999', 'wholesale1030') "
		''addSql = addSql + " 		group by yyyymm, stockPlace, itemgubun, itemid, itemoption "
		''addSql = addSql + " 	) T "
		''addSql = addSql + " 	on "
		''addSql = addSql + " 		1 = 1 "
		''addSql = addSql + " 		and m.yyyymm = T.yyyymm "
		''addSql = addSql + " 		and m.stockPlace = T.stockPlace "
		''addSql = addSql + " 		and m.shopid = T.shopid "
		''addSql = addSql + " 		and m.itemgubun = T.itemgubun "
		''addSql = addSql + " 		and m.itemid = T.itemid "
		''addSql = addSql + " 		and m.itemoption = T.itemoption "
		addSql = addSql + " where "
		addSql = addSql + " 	1 = 1 "
		addSql = addSql + " 	and m.yyyymm>='"&prevYYYYMM&"' and m.yyyymm<='"&FRectYYYYMM&"'"
		addSql = addSql + " 	and m.targetGbn not in ('ET','EG')"
		addSql = addSql + " 	and m.etcjungsantype in (1,4)"
		''addSql = addSql + " 	and m.lastmwdiv not in ('B012')"

        addSql = addSql + " 	and NOT (m.lastmwdiv='B013' and m.targetGbn<>'IT')"               ''�����Ź�� IT��
        addSql = addSql + " 	and m.lastmwdiv not in ('W','B012','B011')"                     ''��ü��Ź ���� //����ڻ��� ������ �ƴ� CASE (����ڻ� ���·� �Ѹ����) W����

		if (FRectStockPlace <> "") then
			addSql = addSql + " 	and m.stockPlace = '" + CStr(FRectStockPlace) + "' "
		end if

        if (FRectMakerid<>"") then
            addSql = addSql + " 	and m.makerid='"&FRectMakerid&"'"
        end if

        if (FRectItemGubun<>"") then
            addSql = addSql + " 	and m.itemgubun='"&FRectItemGubun&"'"
        end if

		if (FRectShopid <> "") then
			addSql = addSql + " 	and m.shopid='"&FRectShopid&"'"
		end if

        if (FRectTargetGbn<>"") then
            addSql = addSql + " 	and m.targetGbn='"&FRectTargetGbn&"'"
        end if

        if (FRectMeaipTp<>"") then
            ''addSql = addSql + " 	and ((isNULL(m.lastmwdiv,'unknown')='"&FRectMeaipTp&"') or (isNULL(m.lastmwdiv,'unknown')='W' and '" + CStr(FRectMeaipTp) + "' = 'M')) "
        end if



		addSql = addSql + " group by "
		addSql = addSql + " 	m.stockPlace "
		''addSql = addSql + " 	,m.targetGbn "
		addSql = addSql + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			addSql = addSql + " 	,m.shopid "
		end if

		''addSql = addSql + " 	,isNULL(lastmwdiv,'unknown') "
		IF (FRectSubGrpType="makerid") then
		    addSql = addSql + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			addSql = addSql + " 	,m.makerid, m.itemid, m.itemoption "
		end if

		'// having
		if (FRectShowDiff <> "") then
			'addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' then IpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			if (FRectYYYY<>"") then
			    addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when LEFT(m.yyyymm,4) = '" + CStr(FRectYYYY) + "' then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
			else
			    addSql = addSql + " 	having (sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") + sum(case when m.yyyymm = '" + CStr(FRectYYYYMM) + "' then stIpgoNo else 0 end) + sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&")) <> 0 "
		    end if
		end if


		sqlStr = " SELECT top 1 (COUNT(*) OVER ()) as CNT, CEILING(CAST((COUNT(*) OVER ()) AS FLOAT)/" + CStr(FPageSize) + ") as totPg "
		sqlStr = sqlStr + addSql
		''response.write sqlStr
		''response.end

        if (FRectSubGrpType<>"") then
        	rsget.Open sqlStr,dbget,1
    		if  not rsget.EOF  then
    			FTotalCount = rsget("cnt")
    			FTotalPage = rsget("totPg")
    		else
    			FTotalCount = 0
    			FTotalPage = 0
    		end if
        	rsget.Close


        	'������������ ��ü ���������� Ŭ �� �Լ�����
        	if CLng(FCurrPage)>CLng(FTotalPage) then
        		FResultCount = 0
        		exit function
        	end if
        end if

		if (FRectSubGrpType="") then ''2016/05/09
            sqlStr = " select  "
        else
		    sqlStr = " select top " + CStr(FPageSize*FCurrPage) + " "
	    end if
		sqlStr = sqlStr + " 	'" + CStr(FRectYYYYMM) + "' as yyyymm "
		sqlStr = sqlStr + " 	,m.stockPlace "
		''sqlStr = sqlStr + " 	,m.targetGbn "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,m.makerid as makerid " ''�ٸ� �귣�� �ϳ��� ǥ��
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	,m.makerid, m.itemid, m.itemoption "
		end if

		''2015-03-17, skyer9
		''sqlStr = sqlStr + " 	,'M' as lastmwdiv"
		sqlStr = sqlStr + " 	,(case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) as lastmwdiv "

		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(prevYYYYMM,"curSysStockNo")&") as prevSysStockNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"IpgoNo")&") as IpgoNo "
		''sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoMoveNo")&") + sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as MoveNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"SellNo")&") as SellNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"OffChulNo")&") - sum("&getCaseStrNo(FRectYYYYMM,"OffChulMoveNo")&") as OffChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"EtcChulNo")&") as EtcChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"CsNo")&") as CsNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"LossChulNo")&") as LossChulNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curSysStockNo")&") as curSysStockNo "
		sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"curErrRealCheckNo")&") as curErrRealCheckNo "

		if (FRectPriceGubun = "V") then
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"totItemNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","avgIpgoPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","totBuyCash")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","avgIpgoPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as MoveSum " '// totBuyCash : ����-���� ���Ա��� ������ ��� ������ո��԰�
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","avgIpgoPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","avgIpgoPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","avgIpgoPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","avgIpgoPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","avgIpgoPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","avgIpgoPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","avgIpgoPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","avgIpgoPrice")&") as curErrRealCheckSum "
		else
			sqlStr = sqlStr + " 	,sum("&getCaseStrNo(FRectYYYYMM,"stIpgoNo")&") as IpgoNo "

			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(prevYYYYMM,"curSysStockNo","lastbuyPrice")&") as prevSysStockSum "
			''sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"1","IpgoSum")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoNo","lastbuyPrice")&") as IpgoSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"stIpgoMoveNo","lastbuyPrice")&") + IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as MoveSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"SellNo","lastbuyPrice")&") as SellSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"OffChulNo","lastbuyPrice")&") - IsNull(sum("&getCaseStrPrice(FRectYYYYMM,"OffChulMoveNo","lastbuyPrice")&"),0) as OffChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"EtcChulNo","lastbuyPrice")&") as EtcChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"CsNo","lastbuyPrice")&") as CsSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"LossChulNo","lastbuyPrice")&") as LossChulSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curSysStockNo","lastbuyPrice")&") as curSysStockSum "
			sqlStr = sqlStr + " 	,sum("&getCaseStrPrice(FRectYYYYMM,"curErrRealCheckNo","lastbuyPrice")&") as curErrRealCheckSum "
		end if

		sqlStr = sqlStr + addSql

		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	(case when (case when m.stockPlace in ('L', 'S') then 'M' else 'J' end) = 'M' then 1 else 100 end) "
		sqlStr = sqlStr + " 	,m.stockPlace "
		''sqlStr = sqlStr + " 	,m.targetGbn desc "
		sqlStr = sqlStr + " 	,m.itemgubun "
		if (FRectShowShopid <> "") then
			sqlStr = sqlStr + " 	,m.shopid "
		end if
		''sqlStr = sqlStr + " 	,lastmwdiv "
		IF (FRectSubGrpType="makerid") then
		    sqlStr = sqlStr + " 	,m.makerid "
		elseif (FRectSubGrpType = "itemid") then
			sqlStr = sqlStr + " 	,m.makerid, m.itemid, m.itemoption "
		end if

	 'response.write sqlStr   ''IDX_tbl_monthly_Stock_IpgoLedger_Sum_yyyymm_ipgoMWDIV 2015/06/01
	 'response.end

'rw sqlStr
'response.end

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
				set FItemList(i) = new CMonthlyMaeipLedgeItem

                FItemList(i).FisJungsan         = false
				FItemList(i).Fyyyymm     		= rsget("yyyymm")
				FItemList(i).FstockPlace    	= rsget("stockPlace")
				''FItemList(i).FtargetGbn     	= rsget("targetGbn")
				FItemList(i).Fitemgubun     	= rsget("itemgubun")
                FItemList(i).Flastmwdiv         = rsget("lastmwdiv")

				if (FRectShowShopid <> "") then
					FItemList(i).Fshopid         = rsget("shopid")
				end if

				FItemList(i).FprevSysStockNo    = rsget("prevSysStockNo")
				FItemList(i).FprevSysStockSum   = rsget("prevSysStockSum")

				FItemList(i).FIpgoNo     		= rsget("IpgoNo")
				FItemList(i).FIpgoSum     		= rsget("IpgoSum")
				FItemList(i).FMoveNo     		= rsget("MoveNo")
				FItemList(i).FMoveSum     		= rsget("MoveSum")
				FItemList(i).FSellNo     		= rsget("SellNo")
				FItemList(i).FSellSum     		= rsget("SellSum")
				FItemList(i).FOffChulNo     	= rsget("OffChulNo")
				FItemList(i).FOffChulSum     	= rsget("OffChulSum")
				FItemList(i).FEtcChulNo     	= rsget("EtcChulNo")
				FItemList(i).FEtcChulSum     	= rsget("EtcChulSum")
				FItemList(i).FCsNo     			= rsget("CsNo")
				FItemList(i).FCsSum     		= rsget("CsSum")
				FItemList(i).FLossChulNo     	= rsget("LossChulNo")
				FItemList(i).FLossChulSum     	= rsget("LossChulSum")

				FItemList(i).FcurSysStockNo     = rsget("curSysStockNo")
				FItemList(i).FcurSysStockSum    = rsget("curSysStockSum")

				FItemList(i).FcurErrRealCheckNo = rsget("curErrRealCheckNo")
				FItemList(i).FcurErrRealCheckSum	= rsget("curErrRealCheckSum")

                IF (FRectSubGrpType="makerid") then
                    FItemList(i).FMakerid = rsget("makerid")
				elseif (FRectSubGrpType = "itemid") then
					FItemList(i).FMakerid 		= rsget("makerid")
					FItemList(i).Fitemid     	= rsget("itemid")
					FItemList(i).Fitemoption    = rsget("itemoption")
                end if
				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close

	end function

	Private Sub Class_Initialize()
		redim FItemList(0)

		FCurrPage = 1
		FPageSize = 100
		FResultCount = 0
		FScrollCount = 10
		FTotalCount = 0
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

end class

%>
