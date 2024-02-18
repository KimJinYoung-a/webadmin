<%
'###########################################################
' Description :  ����Ʈ ī�� ���� ��� Ŭ����
' History : 2012.11.08 �ѿ�� ����
'###########################################################

class cgiftcardsum_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fYYYYMM
	public fyyyymmdd
	public freforeremaincash
	public fsellCash
	public fuseCash
	public frefundCash
	public fuseroutCash
	public fdelcash
	public fremaincash
	public fgiftOrderSerial
	public fmasterCardCode
	public fsubtotalPrice
	public fcancelyn
	public fuserid
	public fbuyname
	public faccountname
	public fipkumdiv
	public fjukyocd
	public fjukyo
	public forderserial
	public fdeleteYn
	public fsitename
	public fshopid
	public fcanceldate
	public fpaydateid
end class

class cgiftcardsum_list
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

	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

	public FRectStartdate
	public FRectEndDate
	public FRectonoffgubun
	public frectaccountdiv
	public frectjukyocd
	public frectcancelyn

	'/admin/maechul/managementsupport/giftcardsum_month.asp
	public function fgiftcardsum_sell_month
		dim i , sql , sqlsearch, sqlsearch2, sqlsearch3

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectStartdate<>"" then
			sqlsearch2 = sqlsearch2 + " and convert(varchar(10),o.ipkumdate,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch2 = sqlsearch2 + " and convert(varchar(10),o.ipkumdate,21) <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectStartdate<>"" then
			sqlsearch3 = sqlsearch3 + " and convert(varchar(10),o.canceldate,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch3 = sqlsearch3 + " and convert(varchar(10),o.canceldate,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymm"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " , db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--�̿��ܾ�
			sql = sql & " , isnull(TotgiftCard,0) as sellCash"		'/--������(�Ǹų���)
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(TotgiftCard,0)"
			sql = sql & " 	+(isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0))"
			sql = sql & " 	+(isnull(t.useCash400,0)+isnull(t.useCash400tong,0))"
			sql = sql & " 	+isnull(t.useCash900,0)"
			sql = sql & " ) as remaincash"		'/--�ܾ�
			sql = sql & " , (isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0)) as useCash"		'/--������
			sql = sql & " , (isnull(t.useCash400,0)+isnull(t.useCash400tong,0)) as refundCash"		'/--ȯ��
			sql = sql & " , isnull(t.useCash900,0) as useroutCash"		'/--ȸ��Ż��

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " , db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--�̿��ܾ�
			sql = sql & " , 0 as sellCash"		'/--������(�Ǹų���)
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(t.useCash200off,0)"
			sql = sql & " ) as remaincash"		'/--�ܾ�
			sql = sql & " , (isnull(t.useCash200off,0)) as useCash"		'/--������
			sql = sql & " , 0 as refundCash"		'/--ȯ��
			sql = sql & " , 0 as useroutCash"		'/--ȸ��Ż��

		else
			sql = sql & " , db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--�̿��ܾ�
			sql = sql & " , isnull(TotgiftCard,0) as sellCash"		'/--������(�Ǹų���)
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_giftcard_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(TotgiftCard,0)"
			sql = sql & " 	+(isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0)+isnull(t.useCash200off,0))"
			sql = sql & " 	+(isnull(t.useCash400,0)+isnull(t.useCash400tong,0))"
			sql = sql & " 	+isnull(t.useCash900,0)"
			sql = sql & " ) as remaincash"		'/--�ܾ�
			sql = sql & " , (isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0)+isnull(t.useCash200off,0)) as useCash"		'/--������
			sql = sql & " , (isnull(t.useCash400,0)+isnull(t.useCash400tong,0)) as refundCash"		'/--ȯ��
			sql = sql & " , isnull(t.useCash900,0) as useroutCash"		'/--ȸ��Ż��
		end if

		sql = sql & " ,0 as delcash"		'/--�Ҹ�
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(7),L.fixyyyymmdd,21) as yyyymm"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='200' and jukyo='��ǰ����' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200"		'/--[ON ��ǰ����]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='200' and jukyo='��ǰȯ��' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200ban"		'/--[��ǰȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='300' and jukyo='��ǰ���� ��� ȯ��' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash300"		'/--[��ǰ���� ��� ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ��ġ�� ȯ��') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash400"		'/--[Giftī�� ��ġ�� ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ������ ȯ��') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash400tong"		'/--[Giftī�� ������ ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='S' and jukyocd='200' and jukyo='��ǰ����') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200off"		'/--[OFF ��ǰ����]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd in ('900','9999') and jukyo='ȸ��Ż��') THEN L.useCash"  ''9999 �߰�
		sql = sql & " 			ELSE 0 END) as useCash900"		'/--ȸ��Ż��
		sql = sql & " 	from db_user.dbo.tbl_giftcard_log L"
		sql = sql & " 	Left Join ("
		sql = sql & " 		select"
		sql = sql & " 		mg.gubun , mg.mastercardcode, O.giftorderserial, subtotalprice,totalsum"
		sql = sql & " 		, accountdiv, giftCardGbn"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order O"
		sql = sql & " 		left Join db_order.dbo.tbl_mobile_gift mg"
		sql = sql & " 			on mg.mastercardCode=O.mastercardCode"
		sql = sql & " 			and mg.isPay='Y'"
		sql = sql & " 		where O.cancelyn='N'"
		sql = sql & " 		and O.ipkumdiv>3"
		sql = sql & " 	) T"
		sql = sql & " 		on L.orderserial=T.giftOrderserial"
		sql = sql & " 	where L.deleteyn='N'"
		sql = sql & " 	and l.fixyyyymmdd is not null " & sqlsearch
		sql = sql & " 	group by convert(varchar(7),L.fixyyyymmdd,21)"
		sql = sql & " ) as t"
		sql = sql & " left join ("
		sql = sql & " 	select"
		sql = sql & " 	a.yyyymm"
		sql = sql & " 	, isnull(a.TotgiftCard,0) - isnull(b.TotgiftCard,0) as TotgiftCard"
		sql = sql & " 	from ("
		sql = sql & " 		select"
		sql = sql & " 		convert(varchar(7),o.ipkumdate,21) as yyyymm"
		sql = sql & " 		,sum(o.subtotalPrice) as TotgiftCard"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order o"
		sql = sql & " 		where o.ipkumdiv>3 " & sqlsearch2
		sql = sql & " 		group by convert(varchar(7),o.ipkumdate,21)"
		sql = sql & " 	) A"
		sql = sql & " 	left join ("
		sql = sql & " 		select"
		sql = sql & " 		convert(varchar(7),o.canceldate,21) as yyyymm"
		sql = sql & " 		,sum(o.subtotalPrice) as TotgiftCard"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order o"
		sql = sql & " 		where o.ipkumdiv>3"
		sql = sql & " 		and o.canceldate is Not NULL"
		sql = sql & " 		and o.ipkumdate  is Not NULL " & sqlsearch3
		sql = sql & " 		group by convert(varchar(7),o.canceldate,21)"
		sql = sql & " 	) B"
		sql = sql & " 		on A.yyyymm=B.yyyymm"
		sql = sql & " ) as s"
		sql = sql & " 	on t.yyyymm=s.yyyymm"
		sql = sql & " order by t.yyyymm asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cgiftcardsum_oneitem

				FItemList(i).fYYYYMM			= rsget("YYYYMM")
				FItemList(i).freforeremaincash			= rsget("reforeremaincash")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")
				FItemList(i).fremaincash			= rsget("remaincash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'/admin/maechul/managementsupport/giftcardsum_day.asp
	public function fgiftcardsum_sell_day
		dim i , sql , sqlsearch, sqlsearch2, sqlsearch3

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectStartdate<>"" then
			sqlsearch2 = sqlsearch2 + " and convert(varchar(10),o.ipkumdate,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch2 = sqlsearch2 + " and convert(varchar(10),o.ipkumdate,21) <'" + CStr(FRectEndDate) + "'"
		end if
		if FRectStartdate<>"" then
			sqlsearch3 = sqlsearch3 + " and convert(varchar(10),o.canceldate,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch3 = sqlsearch3 + " and convert(varchar(10),o.canceldate,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymmdd"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " , (isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0)) as useCash"		'/--������
			sql = sql & " , isnull(TotgiftCard,0) as sellCash"		'/--������(�Ǹų���)
			sql = sql & " , (isnull(t.useCash400,0)+isnull(t.useCash400tong,0)) as refundCash"		'/--ȯ��
			sql = sql & " , isnull(t.useCash900,0) as useroutCash"		'/--ȸ��Ż��

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " , (isnull(t.useCash200off,0)) as useCash"		'/--������
			sql = sql & " , 0 as sellCash"		'/--������(�Ǹų���)
			sql = sql & " , 0 as refundCash"		'/--ȯ��
			sql = sql & " , 0 as useroutCash"		'/--ȸ��Ż��

		else
			sql = sql & " , (isnull(t.useCash200,0)+isnull(t.useCash200ban,0)+isnull(t.useCash300,0)+isnull(t.useCash200off,0)) as useCash"		'/--������
			sql = sql & " , isnull(TotgiftCard,0) as sellCash"		'/--������(�Ǹų���)
			sql = sql & " , (isnull(t.useCash400,0)+isnull(t.useCash400tong,0)) as refundCash"		'/--ȯ��
			sql = sql & " , isnull(t.useCash900,0) as useroutCash"		'/--ȸ��Ż��
		end if

		sql = sql & " ,0 as delcash"		'/--�Ҹ�
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(10),d.lunar_date,21) as yyyymmdd"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='200' and jukyo='��ǰ����' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200"		'/--[ON ��ǰ����]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='200' and jukyo='��ǰȯ��' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200ban"		'/--[��ǰȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN L.siteDiv='T' and jukyocd='300' and jukyo='��ǰ���� ��� ȯ��' THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash300"		'/--[��ǰ���� ��� ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ��ġ�� ȯ��') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash400"		'/--[Giftī�� ��ġ�� ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ������ ȯ��') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash400tong"		'/--[Giftī�� ������ ȯ��]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='S' and jukyocd='200' and jukyo='��ǰ����') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash200off"		'/--[OFF ��ǰ����]
		sql = sql & " 	,sum(CASE"
		sql = sql & " 			WHEN (L.siteDiv='T' and jukyocd in ('900','9999') and jukyo='ȸ��Ż��') THEN L.useCash"
		sql = sql & " 			ELSE 0 END) as useCash900"		'/--ȸ��Ż��
		sql = sql & " 	from db_sitemaster.dbo.LunarToSolar d"
		sql = sql & " 	left join db_user.dbo.tbl_giftcard_log L"
		sql = sql & " 		on d.lunar_date=L.fixyyyymmdd"
		sql = sql & " 	Left Join ("
		sql = sql & " 		select"
		sql = sql & " 		mg.gubun , mg.mastercardcode, O.giftorderserial, subtotalprice,totalsum"
		sql = sql & " 		, accountdiv, giftCardGbn"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order O"
		sql = sql & " 		left Join db_order.dbo.tbl_mobile_gift mg"
		sql = sql & " 			on mg.mastercardCode=O.mastercardCode"
		sql = sql & " 			and mg.isPay='Y'"
		sql = sql & " 		where O.cancelyn='N'"
		sql = sql & " 		and O.ipkumdiv>3"
		sql = sql & " 	) T"
		sql = sql & " 		on L.orderserial=T.giftOrderserial"
		sql = sql & " 	where L.deleteyn='N'"
		sql = sql & " 	and l.fixyyyymmdd is not null " & sqlsearch
		sql = sql & " 	group by convert(varchar(10),d.lunar_date,21)"
		sql = sql & " ) as t"
		sql = sql & " left join ("
		sql = sql & " 	select"
		sql = sql & " 	a.yyyymmdd"
		sql = sql & " 	, isnull(a.TotgiftCard,0) - isnull(b.TotgiftCard,0) as TotgiftCard"
		sql = sql & " 	from ("
		sql = sql & " 		select"
		sql = sql & " 		convert(varchar(10),o.ipkumdate,21) as yyyymmdd"
		sql = sql & " 		,sum(o.subtotalPrice) as TotgiftCard"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order o"
		sql = sql & " 		where o.ipkumdiv>3 " & sqlsearch2
		sql = sql & " 		group by convert(varchar(10),o.ipkumdate,21)"
		sql = sql & " 	) A"
		sql = sql & " 	left join ("
		sql = sql & " 		select"
		sql = sql & " 		convert(varchar(10),o.canceldate,21) as yyyymmdd"
		sql = sql & " 		,sum(o.subtotalPrice) as TotgiftCard"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order o"
		sql = sql & " 		where o.ipkumdiv>3"
		sql = sql & " 		and o.canceldate is Not NULL"
		sql = sql & " 		and o.ipkumdate  is Not NULL " & sqlsearch3
		sql = sql & " 		group by convert(varchar(10),o.canceldate,21)"
		sql = sql & " 	) B"
		sql = sql & " 		on A.yyyymmdd=B.yyyymmdd"
		sql = sql & " ) as s"
		sql = sql & " 	on t.yyyymmdd=s.yyyymmdd"
		sql = sql & " order by t.yyyymmdd asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cgiftcardsum_oneitem

				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'//admin/maechul/managementsupport/giftcardsum_sell_list.asp
	public function fgiftcardsum_sell_list
		dim i , sql , sqlsearch

		if FRectonoffgubun="ONLINE" then
			if frectaccountdiv = "INICardSum" then		'/ī��
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='100'"
			elseif frectaccountdiv = "INIMooSum" then		'/������"
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='7'"
			elseif frectaccountdiv = "INISilSum" then		'/�ǽð�"
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='20'"
			elseif frectaccountdiv = "gifttingChangeSum" then		'/��������ȯ
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='550'"
			elseif frectaccountdiv = "gifticonChangeSuM" then		'/����Ƽ����ȯ
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='560'"
			elseif frectaccountdiv = "etcSum" then		'/����ǰ��
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='10'"
			elseif frectaccountdiv = "NonEgiftCard" then		'/����ǰ(vip)��
				sqlsearch = sqlsearch + " and o.giftCardGbn<>0"
			end if

		elseif FRectonoffgubun="OFFLINE" then

			sqlsearch = sqlsearch + " and o.giftOrderSerial='0'"		'//����Ʈī���̴Ͻý� �Ǹų��� ���� �ӽ÷� �ȳ����� �ҷ��� 0ó��
		else
			if frectaccountdiv = "INICardSum" then		'/ī��
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='100'"
			elseif frectaccountdiv = "INIMooSum" then		'/������"
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='7'"
			elseif frectaccountdiv = "INISilSum" then		'/�ǽð�"
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='20'"
			elseif frectaccountdiv = "gifttingChangeSum" then		'/��������ȯ
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='550'"
			elseif frectaccountdiv = "gifticonChangeSuM" then		'/����Ƽ����ȯ
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='560'"
			elseif frectaccountdiv = "etcSum" then		'/����ǰ��
				sqlsearch = sqlsearch + " and o.giftCardGbn=0 and o.accountdiv='10'"
			elseif frectaccountdiv = "NonEgiftCard" then		'/����ǰ(vip)��
				sqlsearch = sqlsearch + " and o.giftCardGbn<>0"
			end if
		end if
		if frectcancelyn<>"" then
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and o.canceldate >='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and o.canceldate <'" + CStr(FRectEndDate) + "'"
			end if

			sqlsearch = sqlsearch + " and o.cancelyn='"&frectcancelyn&"' and o.canceldate is Not NULL and o.ipkumdate  is Not NULL"
		else
			if FRectStartdate<>"" then
				sqlsearch = sqlsearch + " and o.ipkumdate >='" + CStr(FRectStartdate) + "'"
			end if
			if FRectEndDate<>"" then
				sqlsearch = sqlsearch + " and o.ipkumdate <'" + CStr(FRectEndDate) + "'"
			end if
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " o.giftOrderSerial, o.masterCardCode, o.ipkumdate as yyyymmdd, o.cancelyn, o.ipkumdiv"
		sql = sql & " , o.userid, o.buyname, o.canceldate"

		if frectcancelyn="Y" then
			sql = sql & " , isnull((case when o.cancelyn='Y' and o.subtotalPrice>=0 then o.subtotalPrice*-1 else o.subtotalPrice end),0) as subtotalPrice"
		else
			sql = sql & " , isnull(o.subtotalPrice,0) as subtotalPrice"
		end if

		sql = sql & " ,(CASE"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='100' then 'ī��'"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='7' then '������'"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='20' then '�ǽð�'"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='550' then '��������ȯ'"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='560' then '����Ƽ����ȯ'"
		sql = sql & " 		WHEN o.giftCardGbn=0 and o.accountdiv='10' then '����ǰ��'"
		sql = sql & " 		WHEN o.giftCardGbn<>0 THEN '����ǰ(vip)��'"
		sql = sql & " 		end) as accountname"
		sql = sql & " from db_order.dbo.tbl_giftcard_order o"
		sql = sql & " where o.ipkumdiv>3 " & sqlsearch
		sql = sql & " order by o.cancelyn asc, yyyymmdd asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cgiftcardsum_oneitem

				FItemList(i).fgiftOrderSerial			= rsget("giftOrderSerial")
				FItemList(i).fmasterCardCode			= rsget("masterCardCode")
				FItemList(i).fsubtotalPrice			= rsget("subtotalPrice")
				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")
				FItemList(i).fcancelyn			= rsget("cancelyn")
				FItemList(i).fuserid			= rsget("userid")
				FItemList(i).fbuyname			= rsget("buyname")
				FItemList(i).faccountname			= rsget("accountname")
				FItemList(i).fipkumdiv			= rsget("ipkumdiv")
				FItemList(i).fcanceldate			= rsget("canceldate")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'//admin/maechul/managementsupport/giftcardsum_use_list.asp
	public function fgiftcardsum_use_list
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd >='" + CStr(FRectStartdate) + "'"
			'sqlsearch = sqlsearch + " and l.regdate >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd <'" + CStr(FRectEndDate) + "'"
			'sqlsearch = sqlsearch + " and l.regdate <'" + CStr(FRectEndDate) + "'"
		end if

		if frectjukyocd = "useCash" then	'/������
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(L.siteDiv='T' and jukyocd='200' and jukyo='��ǰ����')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='200' and jukyo='��ǰȯ��')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='300' and jukyo='��ǰ���� ��� ȯ��')"
				sqlsearch = sqlsearch + " )"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and (L.siteDiv='S' and jukyocd='200' and jukyo='��ǰ����')"

			else
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(L.siteDiv='T' and jukyocd='200' and jukyo='��ǰ����')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='200' and jukyo='��ǰȯ��')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='300' and jukyo='��ǰ���� ��� ȯ��')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='S' and jukyocd='200' and jukyo='��ǰ����')"
				sqlsearch = sqlsearch + " )"
			end if

		elseif frectjukyocd = "refundCash" then		'/ȯ��
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ��ġ�� ȯ��')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ������ ȯ��')"
				sqlsearch = sqlsearch + " )"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"	'//�Ⱥ��̱� ����, 0���� �ӽ�ó��

			else
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ��ġ�� ȯ��')"
				sqlsearch = sqlsearch + " 	or (L.siteDiv='T' and jukyocd='400' and jukyo='Giftī�� ������ ȯ��')"
				sqlsearch = sqlsearch + " )"
			end if

		elseif frectjukyocd = "useroutCash" then		'/ȸ��Ż��
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and (L.siteDiv='T' and jukyocd in ('900','9999') and jukyo='ȸ��Ż��')"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"	'//�Ⱥ��̱� ����, 0���� �ӽ�ó��

			else
				sqlsearch = sqlsearch + " and (L.siteDiv='T' and jukyocd in ('900','9999') and jukyo='ȸ��Ż��')"
			end if

		elseif frectjukyocd = "delcash" then		'/�Ҹ�
			sqlsearch = sqlsearch + " and jukyocd='0'"	'//�Ⱥ��̱� ����, 0���� �ӽ�ó��
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " L.userid, L.useCash, L.jukyocd, L.jukyo, L.orderserial, L.deleteYn"
		sql = sql & " , L.fixyyyymmdd as yyyymmdd"
		'sql = sql & " , L.regdate as yyyymmdd"
		sql = sql & " , isnull(m.sitename,lm.sitename) as sitename, sm.shopid"
		sql = sql & " from db_user.dbo.tbl_giftcard_log L"
		sql = sql & " Left Join ("
		sql = sql & " 		select"
		sql = sql & " 		mg.gubun , mg.mastercardcode, O.giftorderserial, subtotalprice,totalsum"
		sql = sql & " 		, accountdiv, giftCardGbn"
		sql = sql & " 		from db_order.dbo.tbl_giftcard_order O"
		sql = sql & " 		left Join db_order.dbo.tbl_mobile_gift mg"
		sql = sql & " 			on mg.mastercardCode=O.mastercardCode"
		sql = sql & " 			and mg.isPay='Y'"
		sql = sql & " 		where O.cancelyn='N'"
		sql = sql & " 		and O.ipkumdiv>3"
		sql = sql & " ) T"
		sql = sql & " 		on L.orderserial=T.giftOrderserial"
		sql = sql & " left join db_order.dbo.tbl_order_master m"
		sql = sql & " 	on l.orderserial=m.orderserial"
		sql = sql & " left join db_log.dbo.tbl_old_order_master_2003 lm"
		sql = sql & " 	on l.orderserial=lm.orderserial"
		sql = sql & " left join db_shop.dbo.tbl_shopjumun_master sm"
		sql = sql & " 	on l.orderserial=sm.orderno"
		sql = sql & " where L.deleteyn='N' " & sqlsearch
		sql = sql & " and l.fixyyyymmdd is not null"
		sql = sql & " order by yyyymmdd"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cgiftcardsum_oneitem

				FItemList(i).fsitename			= rsget("sitename")
				FItemList(i).fshopid			= rsget("shopid")
				FItemList(i).fuserid			= rsget("userid")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).fjukyocd			= rsget("jukyocd")
				FItemList(i).fjukyo			= rsget("jukyo")
				FItemList(i).forderserial			= rsget("orderserial")
				FItemList(i).fdeleteYn			= rsget("deleteYn")
				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

end class

'//�ǸŰ�������
function drawgiftcardaccountdiv(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>��ü</option>
		<option value="INICardSum" <% if selVal="INICardSum" then response.write " selected" %>>ī��</option>
		<option value="INIMooSum" <% if selVal="INIMooSum" then response.write " selected" %>>������</option>
		<option value="INISilSum" <% if selVal="INISilSum" then response.write " selected" %>>�ǽð�</option>
		<option value="gifttingChangeSum" <% if selVal="gifttingChangeSum" then response.write " selected" %>>��������ȯ</option>
		<option value="gifticonChangeSuM" <% if selVal="gifticonChangeSuM" then response.write " selected" %>>����Ƽ����ȯ</option>
		<option value="etcSum" <% if selVal="etcSum" then response.write " selected" %>>����ǰ��</option>
		<option value="NonEgiftCard" <% if selVal="NonEgiftCard" then response.write " selected" %>>����ǰ(vip)��</option>

	</select>
<%
end function

'//�ǸŰ�������
function getgiftcardaccountdiv(selVal)

	if selVal = "" then exit function

	if selVal = "INICardSum" then
		getgiftcardaccountdiv = "ī��"
	elseif selVal = "INIMooSum" then
		getgiftcardaccountdiv = "������"
	elseif selVal = "INISilSum" then
		getgiftcardaccountdiv = "�ǽð�"
	elseif selVal = "gifttingChangeSum" then
		getgiftcardaccountdiv = "��������ȯ"
	elseif selVal = "gifticonChangeSuM" then
		getgiftcardaccountdiv = "����Ƽ����ȯ"
	elseif selVal = "etcSum" then
		getgiftcardaccountdiv = "����ǰ��"
	elseif selVal = "NonEgiftCard" then
		getgiftcardaccountdiv = "����ǰ(vip)��"
	end if
end function

'//���Ű�������
function drawgiftcardjukyocd(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>��ü</option>
		<option value="useCash" <% if selVal="useCash" then response.write " selected" %>>������</option>
		<option value="refundCash" <% if selVal="refundCash" then response.write " selected" %>>ȯ��</option>
		<option value="useroutCash" <% if selVal="useroutCash" then response.write " selected" %>>ȸ��Ż��</option>
		<option value="delcash" <% if selVal="delcash" then response.write " selected" %>>�Ҹ�</option>
	</select>
<%
end function

'//���Ű�������
function getgiftcardjukyocd(selVal)

	if selVal = "" then exit function

	if selVal = "useCash" then
		getgiftcardaccountdiv = "������"
	elseif selVal = "refundCash" then
		getgiftcardaccountdiv = "ȯ��"
	elseif selVal = "useroutCash" then
		getgiftcardaccountdiv = "ȸ��Ż��"
	elseif selVal = "delcash" then
		getgiftcardaccountdiv = "�Ҹ�"
	end if
end function
%>
