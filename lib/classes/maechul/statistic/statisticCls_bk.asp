<%
class cStaticTotalClass_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub


	public FRegdate
	public FMaechulPlus
	public FMaechulMinus
	public FCountPlus
	public FCountMinus
	public FSubtotalprice
	public FMiletotalprice
	public FTotalcheckprice
	public FMinDate
	public FMaxDate
	public FWeek
	public FMonth
	public Facct200			'��ġ��
	public Facct900			'����Ʈī��
	public Facct100			'�ſ�ī��
	public Facct20			'�ǽð���ü
	public Facct7			'������
	public Facct400			'�޴���
	public Facct560			'����Ƽ��
	public Facct550			'������
	public Facct110			'OK+�ſ�
	public Facct80			'�þ�
	public Facct50			'������
	public FDifferent
	public FTotalSum
	public FCountOrder
	public FSiteName
	public FTenCardSpend
	public FAllAtDiscountprice
	public FMaechul
	public FItemNO
	public FOrgitemCost
	public FItemcostCouponNotApplied
	public FItemCost
	public FBuyCash
	public FMaechulProfit
	public FMaechulProfitPer
	public FMaechulProfitPer2
	public FTotItemCost
	public FMakerID
	public FCategoryName
	public FCateL
	public FCateM
	public FCateS
	public Fsellbizcd
 	 
	public FReducedPrice
    public FPurchasetype
  
	function getSellbizName
	    getSellbizName = ""
	    if isNULL(Fsellbizcd) then Exit Function

	    if (Fsellbizcd="0000000101") then
	        getSellbizName = "�¶���"
	    elseif (Fsellbizcd="0000000201") then
	        getSellbizName = "��������"
	    elseif (Fsellbizcd="0000000301") then
	        getSellbizName = "���̶��"
	    else
	        getSellbizName = Fsellbizcd
	    end if
	end function
	
	' �������. ��񿡼� �ϰ��� �����ؼ� ���� ������.
	public Function getPurchasetypeName()
    	IF FPurchasetype = "1" then
    	    getPurchasetypeName = "�Ϲ�����" 
    	ELSEIF FPurchasetype = "3" then
    	    getPurchasetypeName = "PB" 
    	ELSEIF FPurchasetype = "4" then
    	    getPurchasetypeName = "����" 
    	ELSEIF FPurchasetype = "5" then
    	    getPurchasetypeName = "OFF����" 
    	ELSEIF FPurchasetype = "6" then
    	    getPurchasetypeName = "����" 
    	ELSEIF FPurchasetype = "7" then
    	    getPurchasetypeName = "�귣�����"
        ELSEIF FPurchasetype = "8" then
    	    getPurchasetypeName = "����" 
    	END IF           
    end Function
end class
class cStaticTotalClass_list
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public FList
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage
	public FRectDateGijun
	public FRectStartdate
	public FRectEndDate
	public FRectSiteName
	public FRectSort
	public FRectCateL
	public FRectCateM
	public FRectCateS
	public FRectIsBanPum
	public FRectMakerID
	public FRectCateGubun
	public FRectPurchasetype
	public FRectBizSectionCd
  public FRectMwDiv 
  public FRectChannelDiv
  public FRectSellChannelDiv
  public FRectDispCate
 public FRectInc3pl
 public FTotItemCost
 
	public function fStatistic_dailylist			'�Ϻ��������
	dim i , sql
	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),m." & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL('10x10::'+m.rdsite,m.sitename) = '" & FRectSiteName & "' "
	    end if
	End If
	
	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	sql = sql & " GROUP BY Convert(varchar(10),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY Convert(varchar(10),m." & FRectDateGijun & ",120) DESC "
	rsget.open sql,dbget,1
'rw 	sql
	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate			= rsget("yyyymmdd")
				FList(i).FCountPlus 		= rsget("countplus")
				FList(i).FCountMinus      	= rsget("countminus")
				FList(i).FMaechulPlus 		= rsget("maechulplus")
				FList(i).FMaechulMinus     	= rsget("maechulminus")
				FList(i).FSubtotalprice     = rsget("subtotalprice")
				FList(i).FMiletotalprice	= rsget("miletotalprice")

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function


	public function fStatistic_weeklist			'�ֺ��������
	dim i , sql
	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, DATEPART(ww,m." & FRectDateGijun & ") as weekdt,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

	If FRectSiteName <> "" Then
		sql = sql & " AND m.sitename = '" & FRectSiteName & "' "
	End If

	sql = sql & " GROUP BY DATEPART(ww,m." & FRectDateGijun & ") "
	sql = sql & " ORDER BY Convert(varchar(10),max(m." & FRectDateGijun & "),120) DESC "
	rsget.open sql,dbget,1
'rw 	sql 
'response.end

	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= rsget("mindate")
				FList(i).FMaxDate			= rsget("maxdate")
				FList(i).FWeek				= rsget("weekdt")
				FList(i).FCountPlus 		= rsget("countplus")
				FList(i).FCountMinus      	= rsget("countminus")
				FList(i).FMaechulPlus 		= rsget("maechulplus")
				FList(i).FMaechulMinus     	= rsget("maechulminus")
				FList(i).FSubtotalprice     = rsget("subtotalprice")
				FList(i).FMiletotalprice	= rsget("miletotalprice")

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function



	public function fStatistic_monthlist			'�����������
	dim i , sql
	sql = "SELECT "
	sql = sql & " 	Convert(varchar(10),min(m." & FRectDateGijun & "),120) AS mindate, Convert(varchar(10),max(m." & FRectDateGijun & "),120) AS maxdate, Convert(varchar(7),m." & FRectDateGijun & ",120) AS regmonth,"
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then 1 Else 0 End),0) AS countplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv not in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulplus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then 1 Else 0 End),0) AS countminus, "
	sql = sql & " 	isNull(SUM(Case When m.jumundiv in ('9','6') Then m.subtotalprice+isNULL(m.miletotalprice,0) Else 0 End),0) AS maechulminus, "
	sql = sql & " 	isNull(SUM(m.subtotalprice),0) AS subtotalprice, "
	sql = sql & " 	isNull(SUM(m.miletotalprice),0) AS miletotalprice "
	sql = sql & " FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & " WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' AND '" & DateAdd("d",1,FRectEndDate) & "' "
	sql = sql & " AND m.ipkumdiv>3 AND m.cancelyn='N'"

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	sql = sql & " GROUP BY Convert(varchar(7),m." & FRectDateGijun & ",120) "
	sql = sql & " ORDER BY Convert(varchar(7),m." & FRectDateGijun & ",120) DESC "
	rsget.open sql,dbget,1
'rw 	sql
	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FMinDate			= rsget("mindate")
				FList(i).FMaxDate			= rsget("maxdate")
				FList(i).FMonth				= rsget("regmonth")
				FList(i).FCountPlus 		= rsget("countplus")
				FList(i).FCountMinus      	= rsget("countminus")
				FList(i).FMaechulPlus 		= rsget("maechulplus")
				FList(i).FMaechulMinus     	= rsget("maechulminus")
				FList(i).FSubtotalprice     = rsget("subtotalprice")
				FList(i).FMiletotalprice	= rsget("miletotalprice")

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function



	public function fStatistic_checkmethod			'������ĺ� �������
	dim i , sql

	sql = "SELECT "
	sql = sql & "	A.yyyymmdd, isNull(A.miletotalprice,0) AS miletotalprice, "
	sql = sql & "	isNull(B.acct200,0) AS acct200, isNull(B.acct900,0) AS acct900, "
	sql = sql & "	isNull(A.acct100,0)+ isNull(A.acct110,0)-isNull(b.acct110,0) AS acct100, isNull(A.acct20,0) AS acct20, isNull(A.acct7,0) AS acct7, isNull(A.acct400,0) AS acct400, " ''isNull(A.acct100,0)==> isNull(A.acct100,0)+ isNull(A.acct110,0)-isNull(b.acct110,0)
	sql = sql & "	isNull(A.acct560,0) AS acct560, isNull(A.acct550,0) AS acct550, isNull(b.acct110,0) AS acct110, isNull(A.acct80,0) AS acct80, isNull(A.acct50,0) AS acct50, "        ''isNull(A.acct110,0)==> isNull(b.acct110,0)
	sql = sql & "	(A.sumpaymentEtc-b.acct200-b.acct900) AS different "
	sql = sql & "FROM "
	sql = sql & "( "
	sql = sql & "	select "
	sql = sql & "		convert(varchar(10),m." & FRectDateGijun & ",21) as yyyymmdd, "
	sql = sql & "		sum(m.miletotalprice) as miletotalprice, "
	sql = sql & "		sum(CASE WHEN accountdiv='100' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct100, "
	sql = sql & "		sum(CASE WHEN accountdiv='20' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct20, "
	sql = sql & "		sum(CASE WHEN accountdiv='7' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct7, "
	sql = sql & "		sum(CASE WHEN accountdiv='400' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct400, "
	sql = sql & "		sum(CASE WHEN accountdiv='560' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct560, "
	sql = sql & "		sum(CASE WHEN accountdiv='550' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct550, "
	sql = sql & "		sum(CASE WHEN accountdiv='110' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct110, "
	sql = sql & "		sum(CASE WHEN accountdiv='80' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct80, "
	sql = sql & "		sum(CASE WHEN accountdiv='50' THEN m.subtotalPrice-isNULL(m.sumpaymentEtc,0) ELSE 0 END) as acct50, "
	sql = sql & "		sum(m.sumpaymentEtc) as sumpaymentEtc "
	sql = sql & "	from [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & "	where m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' and m.cancelyn='N' and m.ipkumdiv>3 "

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if
	
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	sql = sql & "	group by convert(varchar(10),m." & FRectDateGijun & ",21) "
	sql = sql & ") A "
	sql = sql & "LEFT JOIN "
	sql = sql & "( "
	sql = sql & "	select "
	sql = sql & "		convert(varchar(10),m." & FRectDateGijun & ",21) as yyyymmdd, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='200' then realpayedsum else 0 end ) as acct200, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='900' then realpayedsum else 0 end ) as acct900, "
	sql = sql & "		sum(CASE WHEN e.acctdiv='110' then realpayedsum else 0 end ) as acct110 "  ''2013/05/27 �߰�
	sql = sql & "	from [db_order].[dbo].[tbl_order_master] as m "
	sql = sql & "		inner Join [db_order].[dbo].[tbl_order_paymentEtc] as E on M.orderserial=E.orderserial and E.acctdiv in ('200','900','110') "
	sql = sql & "	where m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' and m.cancelyn='N' and m.ipkumdiv>3 and (m.sumpaymentEtc<>0 or m.accountdiv='110') "

	If FRectSiteName <> "" Then
		sql = sql & " AND m.sitename = '" & FRectSiteName & "' "
	End If
	
	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if
    
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If

	sql = sql & "	group by convert(varchar(10),m." & FRectDateGijun & ",21) "
	sql = sql & ") B ON A.yyyymmdd = B.yyyymmdd "
	sql = sql & "ORDER BY A.yyyymmdd DESC "
	rsget.open sql,dbget,1
'rw 	sql
	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate			= rsget("yyyymmdd")
				FList(i).FMiletotalprice	= rsget("miletotalprice")
				FList(i).Facct200			= rsget("acct200")
				FList(i).Facct900			= rsget("acct900")
				FList(i).Facct100			= rsget("acct100")
				FList(i).Facct20			= rsget("acct20")
				FList(i).Facct7				= rsget("acct7")
				FList(i).Facct400			= rsget("acct400")
				FList(i).Facct560			= rsget("acct560")
				FList(i).Facct550			= rsget("acct550")
				FList(i).Facct110			= rsget("acct110")
				FList(i).Facct80			= rsget("acct80")
				FList(i).Facct50			= rsget("acct50")
				FList(i).FTotalSum			= rsget("miletotalprice") + rsget("acct200") + rsget("acct900") + rsget("acct100") + rsget("acct20") + rsget("acct7") + rsget("acct400") + rsget("acct560") + rsget("acct550") + rsget("acct110") + rsget("acct80") + rsget("acct50")
				FList(i).FDifferent			= rsget("different")

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function


	public function fStatistic_sitename			'�Ǹ�ó�� �������
	dim i , sql

	sql = "SELECT "
	sql = sql & "		count(m.orderserial) as ordercnt, "
	sql = sql & "		isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename) as sitename, "
	''sql = sql & "		isNULL('10x10::'+ m.rdsite, m.sitename) as sitename, "
	sql = sql & "		isNull(SUM(m.totalsum),0) as totalsum, "
	sql = sql & "		isNull(SUM(m.tencardspend),0) as tencardspend, "
	sql = sql & "		isNull(SUM(m.allatdiscountprice),0) as allatdiscountprice, "
	sql = sql & "		isNull(SUM(m.miletotalprice),0) as miletotalprice, "
	sql = sql & "		isNull(SUM(m.subtotalprice),0) as subtotalprice, "
	sql = sql & "		p.sellbizcd"
	sql = sql & "	FROM [db_order].[dbo].[tbl_order_master] as m "
	sql = sql & "       left Join db_partner.dbo.tbl_partner p"
	sql = sql & "       on m.sitename=p.id"
	sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	If FRectSiteName <> "" Then
		sql = sql & " AND m.sitename = '" & FRectSiteName & "' "
	End If
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if
	sql = sql & "	GROUP BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename), p.sellbizcd "
	sql = sql & "	ORDER BY isNULL('10x10::'+ case when left(m.rdsite,6) = 'nvshop' then 'nvshop' when left(m.rdsite,13) = 'mobile_nvshop' then 'mobile_nvshop' else m.rdsite end, m.sitename) ASC "
	'sql = sql & "	GROUP BY isNULL('10x10::'+ m.rdsite, m.sitename), p.sellbizcd "
	'sql = sql & "	ORDER BY isNULL('10x10::'+ m.rdsite, m.sitename) ASC "
	
	'sql = sql & "	GROUP BY m.sitename, p.sellbizcd "
	'sql = sql & "	ORDER BY m.sitename ASC "

	
	'response.Write sql
	rsget.open sql,dbget,1

	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FCountOrder			= rsget("ordercnt")
				FList(i).FSiteName				= rsget("sitename")
				FList(i).FTotalSum				= rsget("totalsum")
				FList(i).FTenCardSpend			= rsget("tencardspend")
				FList(i).FAllAtDiscountprice	= rsget("allatdiscountprice")
				FList(i).FMaechul				= rsget("subtotalprice") + rsget("miletotalprice")
				FList(i).FMiletotalprice		= rsget("miletotalprice")
				FList(i).FSubtotalprice			= rsget("subtotalprice")

                FList(i).Fsellbizcd = rsget("sellbizcd")
		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function


	public function fStatistic_daily_item			'��ǰ������-�Ϻ�
	dim i , sql

	sql = "SELECT "
	sql = sql & "		Convert(varchar(10),m." & FRectDateGijun & ",120) AS yyyymmdd, "
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
	sql = sql & "	FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & "		INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
		If FRectPurchasetype <> "" Then
			sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as pp on d.makerid = pp.id "
		End IF
	sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100"

	If FRectSiteName <> "" Then
		sql = sql & " AND m.sitename = '" & FRectSiteName & "' "
	End If
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectPurchasetype <> "" Then
		sql = sql & " and pp.purchasetype = '" & FRectPurchasetype &"'"
	End IF
	if (FRectMakerID<>"") then
	    sql = sql & " and d.makerid='"&FRectMakerID&"'"
	end if
	if (FRectMwDiv<>"") then
        sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
    end if

	sql = sql & "	GROUP BY Convert(varchar(10),m."&FRectDateGijun&",120) "
	sql = sql & "	ORDER BY yyyymmdd DESC "

	rsget.open sql,dbget,1

	FTotalCount = rsget.recordcount

	redim FList(FTotalCount)
	i = 0
	If Not rsget.Eof Then
		Do Until rsget.Eof
			set FList(i) = new cStaticTotalClass_oneitem
				FList(i).FRegdate					= rsget("yyyymmdd")
				FList(i).FItemNO					= rsget("itemno")
				FList(i).FOrgitemCost				= rsget("orgitemcost")
				FList(i).FItemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
				FList(i).FItemCost					= rsget("itemcost")
				FList(i).FBuyCash					= rsget("buycash")
				FList(i).FReducedPrice				= rsget("reducedprice")
				FList(i).FMaechulProfit				= rsget("itemcost") - rsget("buycash")
				FList(i).FMaechulProfitPer			= Round(((rsget("itemcost") - rsget("buycash"))/CHKIIF(rsget("itemcost")=0,1,rsget("itemcost")))*100,2)
				FList(i).FMaechulProfitPer2			= Round(((rsget("reducedprice") - rsget("buycash"))/CHKIIF(rsget("reducedprice")=0,1,rsget("reducedprice")))*100,2)

		rsget.movenext
		i = i + 1
		Loop
	End If

	rsget.close
	end function


	public function fStatistic_brand			'�귣�庰����
	dim i , sql

	sql = "SELECT "
	sql = sql & "		d.makerid, p.purchasetype,  "
	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt ����. ����
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
	If FRectSort = "profit" Then
		sql = sql & "	,(isNull(sum(d.itemcost*d.itemno),0)-isNull(sum(d.buycash*d.itemno),0)) as profit "
	End If 
	sql = sql & "	FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & "		INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
	If FRectCateL <> "" Then
	    sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item] as i ON d.itemid = i.itemid "
	end if
		 
		sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
		 
		IF FRectDispCate<>"" THEN	'2014-02-27 ������ ����ī�װ� �˻� �߰� 
		sql = sql & " INNER JOIN db_item.dbo.tbl_display_cate_item as dc on d.itemid = dc.itemid and dc.catecode like '" + FRectDispCate + "%' and dc.isDefault='y'"
		END IF
	sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100"

	If FRectSiteName <> "" Then
	    if (FRectSiteName="mobileAll") then
	        sql = sql & " AND left(m.rdsite,6)='mobile'"
	    else
		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
	    end if
	End If

if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if
	
	If FRectCateL <> "" Then
		sql = sql & " AND i.cate_large = '" & FRectCateL & "' "
	End If
	If FRectCateM <> "" Then
		sql = sql & " AND i.cate_mid = '" & FRectCateM & "' "
	End If
	If FRectCateS <> "" Then
		sql = sql & " AND i.cate_small = '" & FRectCateS & "' "
	End If 
	
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectPurchasetype <> "" Then
		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
	End IF

    if (FRectMwDiv<>"") then
        sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
    end if

	sql = sql & "	GROUP BY d.makerid , p.purchasetype"
	sql = sql & "	ORDER BY " & FRectSort & " DESC " 
	if (TRUE) then
    	rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
    	If Not rsget.Eof Then
    		Do Until rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    				FList(i).FMakerID					= rsget("makerid")
    				FList(i).Fpurchasetype				= rsget("purchasetype")
    				FList(i).FCountOrder				= rsget("ordercnt")
    				FList(i).FItemNO					= rsget("itemno")
    				FList(i).FOrgitemCost				= rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= rsget("itemcost")
    				FList(i).FBuyCash					= rsget("buycash")
    				FList(i).FReducedPrice				= rsget("reducedprice")
    				FList(i).FMaechulProfit				= rsget("itemcost") - rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((rsget("itemcost") - rsget("buycash"))/CHKIIF(rsget("itemcost")=0,1,rsget("itemcost")))*100,2)
    				FList(i).FMaechulProfitPer2			= Round(((rsget("reducedprice") - rsget("buycash"))/CHKIIF(rsget("reducedprice")=0,1,rsget("reducedprice")))*100,2)

    		rsget.movenext
    		i = i + 1
    		Loop
    	End If

    	rsget.close
    else
        db3_rsget.open sql,db3_dbget,1

    	FTotalCount = db3_rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
    	If Not db3_rsget.Eof Then
    		Do Until db3_rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    				FList(i).FMakerID					= db3_rsget("makerid")
    				FList(i).Fpurchasetype				= db3_rsget("purchasetype")
    				FList(i).FCountOrder				= db3_rsget("ordercnt")
    				FList(i).FItemNO					= db3_rsget("itemno")
    				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= db3_rsget("itemcost")
    				FList(i).FBuyCash					= db3_rsget("buycash")
    				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)

    		db3_rsget.movenext
    		i = i + 1
    		Loop
    	End If

    	db3_rsget.close
    end if
	end function


	public function fStatistic_category			'ī�װ�������
	dim i , sql

	sql = "SELECT "
		If FRectCateGubun = "L" Then
			sql = sql & " l.code_large, '' as code_mid, '' as code_small, l.code_nm, l.orderNo, "
		ElseIf FRectCateGubun = "M" Then
			sql = sql & " mi.code_large, mi.code_mid, '' as code_small, mi.code_nm, mi.orderNo, "
		ElseIf FRectCateGubun = "S" Then
			sql = sql & " s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo, "
		End If
	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt ����. ����
	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
	sql = sql & "	FROM [db_order].[dbo].[tbl_order_master] as m "
	if (FRectBizSectionCd<>"") then
	    sql = sql & " Join db_partner.dbo.tbl_partner p"
	    sql = sql & " on m.sitename=p.id"
	    sql = sql & " and isNULL(p.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
	end if
	sql = sql & "		INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
	sql = sql & "		INNER JOIN [db_item].[dbo].[tbl_item_Category] as i ON d.itemid = i.itemid AND i.code_div='D' "
		If FRectCateGubun = "L" Then
			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_large] as l ON i.code_large = l.code_large "
		ElseIf FRectCateGubun = "M" Then
			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_mid] as mi ON i.code_large = mi.code_large AND i.code_mid = mi.code_mid "
		ElseIf FRectCateGubun = "S" Then
			sql = sql & " INNER JOIN [db_item].[dbo].[tbl_Cate_small] as s ON i.code_large = s.code_large AND i.code_mid = s.code_mid AND i.code_small = s.code_small "
		End If
		If FRectPurchasetype <> "" Then
			sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
		End IF
	sql = sql & "	WHERE m." & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,FRectEndDate) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100"

	If FRectSiteName <> "" Then
		sql = sql & " AND m.sitename = '" & FRectSiteName & "' "
	End If
	If FRectCateL <> "" Then
		sql = sql & " AND i.code_large = '" & FRectCateL & "' "
	End If
	If FRectCateM <> "" Then
		sql = sql & " AND i.code_mid = '" & FRectCateM & "' "
	End If
	If FRectCateS <> "" Then
		sql = sql & " AND i.code_small = '" & FRectCateS & "' "
	End If
	If FRectIsBanPum <> "all" Then
		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
	End If
	If FRectPurchasetype <> "" Then
		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
	End IF
    if (FRectMwDiv<>"") then
        sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
    end if
	if (FRectSellChannelDiv<>"") then
    	sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    end if

	If FRectCateGubun = "L" Then
		sql = sql & " GROUP BY l.code_large, l.code_nm, l.orderNo ORDER BY l.orderNo ASC "
	ElseIf FRectCateGubun = "M" Then
		sql = sql & " GROUP BY mi.code_large, mi.code_mid, mi.code_nm, mi.orderNo ORDER BY mi.orderNo ASC "
	ElseIf FRectCateGubun = "S" Then
		sql = sql & " GROUP BY s.code_large, s.code_mid, s.code_small, s.code_nm, s.orderNo ORDER BY s.orderNo ASC "
	End If

	if (TRUE) then
    	rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly

    	FTotalCount = rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
    	If Not rsget.Eof Then
    		Do Until rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    				FList(i).FCategoryName				= rsget("code_nm")
    				FList(i).FCateL						= rsget("code_large")
    				FList(i).FCateM						= rsget("code_mid")
    				FList(i).FCateS						= rsget("code_small")
    				FList(i).FCountOrder				= rsget("ordercnt")
    				FList(i).FItemNO					= rsget("itemno")
    				FList(i).FOrgitemCost				= rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= rsget("itemcost")
    				FList(i).FBuyCash					= rsget("buycash")
    				FList(i).FReducedPrice				= rsget("reducedprice")
    				FList(i).FMaechulProfit				= rsget("itemcost") - rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((rsget("itemcost") - rsget("buycash"))/CHKIIF(rsget("itemcost")=0,1,rsget("itemcost")))*100,2)
    				FList(i).FMaechulProfitPer2			= Round(((rsget("reducedprice") - rsget("buycash"))/CHKIIF(rsget("reducedprice")=0,1,rsget("reducedprice")))*100,2)

    		rsget.movenext
    		i = i + 1
    		Loop
    	End If

    	rsget.close
	else
    	db3_rsget.open sql,db3_dbget,1

    	FTotalCount = db3_rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
    	If Not db3_rsget.Eof Then
    		Do Until db3_rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    				FList(i).FCategoryName				= db3_rsget("code_nm")
    				FList(i).FCateL						= db3_rsget("code_large")
    				FList(i).FCateM						= db3_rsget("code_mid")
    				FList(i).FCateS						= db3_rsget("code_small")
    				FList(i).FCountOrder				= db3_rsget("ordercnt")
    				FList(i).FItemNO					= db3_rsget("itemno")
    				FList(i).FOrgitemCost				= db3_rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= db3_rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= db3_rsget("itemcost")
    				FList(i).FBuyCash					= db3_rsget("buycash")
    				FList(i).FMaechulProfit				= db3_rsget("itemcost") - db3_rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((db3_rsget("itemcost") - db3_rsget("buycash"))/CHKIIF(db3_rsget("itemcost")=0,1,db3_rsget("itemcost")))*100,2)

    		db3_rsget.movenext
    		i = i + 1
    		Loop
    	End If

    	db3_rsget.close
    end if
	end function


    public function fStatistic_DispCategory  '���� ī�װ�������
        dim i , sql, vDB
        dim DispCateCode : DispCateCode = FRectCateL&FRectCateM&FRectCateS  ''���� ����� ����
        dim grpLen : grpLen = 3+Len(DispCateCode)
        dim icateCode
 
    	vDB = " [db_order].[dbo].[tbl_order_master] as m INNER JOIN [db_order].[dbo].[tbl_order_detail] as d ON m.orderserial = d.orderserial "
     

        if (FRectDateGijun="beasongdate") then
            FRectDateGijun = "d."&FRectDateGijun
        else
            FRectDateGijun = "m."&FRectDateGijun
        end if

    	sql = "SELECT "

    	sql = sql & "  isNULL(l.catecode,'999') as cateCode"
        sql = sql & " , isNULL(l.cateName,'������') as cateName"
        sql = sql & " , isNULL(l.sortno,999) as sortno, "
    	sql = sql & "		0 AS ordercnt, " ''count(distinct m.orderserial) AS ordercnt ����. ����
    	sql = sql & "		isNull(sum(d.itemno),0) AS itemno, "
    	sql = sql & "		isNull(sum(d.orgitemcost*d.itemno),0) AS orgitemcost, "
    	sql = sql & "		isNull(sum(d.itemcostCouponNotApplied*d.itemno),0) AS itemcostCouponNotApplied, "
    	sql = sql & "		isNull(sum(d.itemcost*d.itemno),0) AS itemcost, "
    	sql = sql & "		isNull(sum(d.buycash*d.itemno),0) as buycash, "
    	sql = sql & "		isNull(sum(d.reducedPrice*d.itemno),0) as reducedprice "
    	sql = sql & "	FROM " & vDB & " "
    	sql = sql & "       left join db_partner.dbo.tbl_partner p2"
	    sql = sql & "       on m.sitename=p2.id "
    	sql = sql & "		LEFT JOIN db_item.[dbo].tbl_display_cate_item as i ON d.itemid = i.itemid AND i.isDefault='y' "
    	sql = sql & "       LEFT JOIN db_item.[dbo].tbl_display_cate as l ON Left(i.catecode,"&grpLen&")=l.catecode"

    		If FRectPurchasetype <> "" Then
    			sql = sql & " INNER JOIN [db_partner].[dbo].[tbl_partner] as p on d.makerid = p.id "
    		End IF

    		if (FRectBizSectionCd<>"") then
        	    sql = sql & " Join db_partner.dbo.tbl_partner p3"
        	    sql = sql & " on m.sitename=p3.id"
        	    sql = sql & " and isNULL(p3.sellbizcd,'0000000101')='"&FRectBizSectionCd&"'"
        	end if

    	sql = sql & "	WHERE " & FRectDateGijun & " BETWEEN '" & FRectStartdate & "' and '" & DateAdd("d",1,GetValidDate(FRectEndDate)) & "' AND m.ipkumdiv>3 AND m.cancelyn='N' AND d.cancelyn<>'Y' AND d.itemid<>0 AND d.itemid<>100"

        ''2014/01/15�߰�
        if (FRectInc3pl<>"") then
            if (FRectInc3pl="A") then

            else
                sql = sql & " and isNULL(p2.tplcompanyid,'')<>''"
            end if
        else
            sql = sql & " and isNULL(p2.tplcompanyid,'')=''"
        end if

    	If FRectSiteName <> "" Then
    	    if (FRectSiteName="mobileAll") then
    	        sql = sql & " AND left(m.rdsite,6)='mobile'"
    	    else
    		    sql = sql & " AND isNULL(m.sitename,m.rdsite) = '" & FRectSiteName & "' "
    	    end if
    	End If

		if (FRectSellChannelDiv<>"") then
       		sql = sql & " and m.beadaldiv in ("&getChannelvalue2ArrIDx(FRectSellChannelDiv)&")"
    	end if

    	if (DispCateCode<>"") then
            sql = sql & " and Left(l.catecode,"&Len(DispCateCode)&")='"&DispCateCode&"'"
        end if

    	If FRectIsBanPum <> "all" Then
    		sql = sql & " AND m.jumundiv" & FRectIsBanPum & "9 "
    	End If
    	If FRectPurchasetype <> "" Then
    		sql = sql & " and p.purchasetype = '" & FRectPurchasetype &"'"
    	End IF
    	if (FRectMwDiv<>"") then
            sql = sql & " and d.omwdiv = '" & FRectMwDiv &"'"
        end if

        sql = sql & " GROUP BY l.catecode, l.cateName, l.sortno ORDER BY l.sortno  , l.catecode"

  'rw   sql 
    	rsget.CursorLocation = adUseClient
        rsget.Open sql,dbget,adOpenForwardOnly, adLockReadOnly
 
    	FTotalCount = rsget.recordcount

    	redim FList(FTotalCount)
    	i = 0
 			FTotItemCost = 0
    	If Not rsget.Eof Then
    		Do Until  rsget.Eof
    			set FList(i) = new cStaticTotalClass_oneitem
    			    icateCode = CStr(rsget("cateCode"))

    				FList(i).FCategoryName				= rsget("cateName")
    				FList(i).FCategoryName              = replace(FList(i).FCategoryName,"^^","&gt;")
    				FList(i).FCateL						= Left(icateCode,3)
    				FList(i).FCateM						= Mid(icateCode,4,3)
    				FList(i).FCateS						= Mid(icateCode,7,3)
    				FList(i).FCountOrder				= rsget("ordercnt")
    				FList(i).FItemNO					= rsget("itemno")
    				FList(i).FOrgitemCost				= rsget("orgitemcost")
    				FList(i).FItemcostCouponNotApplied	= rsget("itemcostCouponNotApplied")
    				FList(i).FItemCost					= rsget("itemcost")
    				FList(i).FBuyCash					= rsget("buycash")
    				FList(i).FReducedPrice				= rsget("reducedprice")
    				FList(i).FMaechulProfit				= rsget("itemcost") - rsget("buycash")
    				FList(i).FMaechulProfitPer			= Round(((rsget("itemcost") - rsget("buycash"))/CHKIIF(rsget("itemcost")=0,1,rsget("itemcost")))*100,2)
    				FList(i).FMaechulProfitPer2			= Round(((rsget("reducedprice") - rsget("buycash"))/CHKIIF(rsget("reducedprice")=0,1,rsget("reducedprice")))*100,2)
					FTotItemCost 		=  FTotItemCost + FList(i).FItemCost	'�����Ѿ� �߰� - 2014-03-27 ������
    		rsget.movenext
    		i = i + 1
    		Loop

    	End If

    	rsget.close
    end function

end class



Function DateToWeekName(d)
	SELECT CASE d
		CASE "1" : DateToWeekName = "<font color=""red"">��</font>"
		CASE "2" : DateToWeekName = "��"
		CASE "3" : DateToWeekName = "ȭ"
		CASE "4" : DateToWeekName = "��"
		CASE "5" : DateToWeekName = "��"
		CASE "6" : DateToWeekName = "��"
		CASE "7" : DateToWeekName = "<font color=""blue"">��</font>"
	END SELECT
End Function
%>
