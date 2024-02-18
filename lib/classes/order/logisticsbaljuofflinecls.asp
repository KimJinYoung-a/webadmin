<%

''한글 한글 한글

public function GetLogicsSiteSeq()
	GetLogicsSiteSeq = "10"		'텐텐
end function

Class COfflineBaljuItem
	public FSiteSeq

	public FBaljuNum
	public FBaljuDate
	public FTargetId
	public FBaljuId
	public FDivcode
	public FTargetName
	public FBaljuName
	public Fbaljucode
	public Fstatecd

	public FMakerid
	public Fprtidx
	public Fmaeipdiv

	public FItemGubun
	public FItemId
	public FItemoption
	public FItemName
	public FItemOptionname
	public FImgSmall
	public FSellCash

	public FSellYn
	public FLimitYn
	public FLimitNo
	public FLimitSold

	public FRealBoxNo
	public FBoxSongjangNo

	public Ftotalbaljuno    '발주수량 = baljuitemno
	public Ftotalupcheno    '업체배송 = (case when TT.maeipdiv = 'U' then TT.baljuitemno else 0 end)
	public Ftotalofflineno  '오프라인 = (case when TT.maeipdiv = '9' then TT.baljuitemno else 0 end)
	public Ftotaltenbaeno   '텐텐배송 = (case when (TT.maeipdiv in ('M','W') then TT.baljuitemno else 0 end)

	public Ftotalnopackno
	public Ftotalpackno
	public Ftotaldeliverno
	public Ftotalconfirmno
	public Ftotaletcno
	public Ftotalprepackno

	Public FIsFinished

	public Fcomment

    public function getStateNameHTML()
        if Fstatecd="0" then
            getStateNameHTML = "주문접수"
        elseif Fstatecd="1" then
		    getStateNameHTML = "주문확인"
		elseif Fstatecd="2" then
		    getStateNameHTML = "입금대기"
		elseif Fstatecd="5" then
		    getStateNameHTML = "배송준비"
		elseif Fstatecd="6" then
		    getStateNameHTML = "출고대기"
		elseif Fstatecd="7" then
		    getStateNameHTML = "<font color='#CC3333'>출고완료</font>"
		else
		    getStateNameHTML = Fstatecd
		end if
    end Function

    public function getIsFinishedName()
        if FIsFinished="N" then
            getIsFinishedName = "출고작업중"
        elseif FIsFinished="W" then
		    getIsFinishedName = "출고대기"
		elseif FIsFinished="Y" then
		    getIsFinishedName = "<font color='#CC3333'>출고완료</font>"
		else
		    getIsFinishedName = FIsFinished
		end if
    end function

	public function IsSoldOut
		IsSoldOut = ((FSellYn="N") or ((FLimitYn="Y") and (GetLimitEa=0)))
	end function

	public function GetIsSlodOutText
		if IsSoldOut then
			GetIsSlodOutText = "SoldOut"
		else
			GetIsSlodOutText = ""
		end if
	end function

	public function GetIsLimitText
		if (FLimitYn="Y") then
			GetIsLimitText = "한정(" + CStr(GetLimitEa) + ")"
		else
			GetIsLimitText = ""
		end if
	end function

	public function GetLimitEa
		GetLimitEa = FLimitNo-FLimitSold
		if GetLimitEa<1 then GetLimitEa=0
	end function

        public function GetDivCodeName()
                if Fdivcode="101" then
                        GetDivCodeName = "가맹점용 개별매입"
                elseif Fdivcode="111" then
                        GetDivCodeName = "가맹점용 개별위탁"
                elseif Fdivcode="121" then
                        GetDivCodeName = "온라인위탁->가맹점위탁"
                elseif Fdivcode="131" then
                        GetDivCodeName = "온라인위탁->가맹점매입"
                elseif Fdivcode="201" then
                        GetDivCodeName = "온라인매입->가맹점매입"
                elseif Fdivcode="251" then
                        GetDivCodeName = "매입반품->오프재고"
                elseif Fdivcode="261" then
                        GetDivCodeName = "오프재고->가맹점출고"
                elseif Fdivcode="300" then
                        GetDivCodeName = "온라인주문"
                elseif Fdivcode="301" then
                        GetDivCodeName = "온라인매입"
                elseif Fdivcode="302" then
                        GetDivCodeName = "온라인위탁"
                elseif Fdivcode="501" then
                        GetDivCodeName = "직영샵주문"
                elseif Fdivcode="502" then
                        GetDivCodeName = "수수료샵"
                elseif Fdivcode="503" then
                        GetDivCodeName = "프랜차이즈"
                else
                        GetDivCodeName = ""
                end if
        end function

        public function GetDivCodeColor()
                if Fdivcode="101" then
                        GetDivCodeColor = "#0000AA"
                elseif Fdivcode="111" then
                        GetDivCodeColor = "#AA0000"
                elseif Fdivcode="121" then
                        GetDivCodeColor = "#AA00AA"
                elseif Fdivcode="131" then
                        GetDivCodeColor = "#00AAAA"
                elseif Fdivcode="201" then
                        GetDivCodeColor = "#AAAA00"
                elseif Fdivcode="300" then
                        GetDivCodeColor = "#FF0000"
                elseif Fdivcode="501" then
                        GetDivCodeColor = "#0000FF"
                elseif Fdivcode="502" then
                        GetDivCodeColor = "#00FF00"
                elseif Fdivcode="503" then
                        GetDivCodeColor = "#AAFFAA"
                else
                        GetDivCodeColor = "#000000"
                end if
        end function

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class COfflineBalju
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

		public FRectSiteSeq
        public FRectBaljuNum
        public FRectBaljuId
        public FRectSelectedOnly
        public FRectBoxNo
        public FRectBoxHasStock
        public FRectStartDate
        public FRectEndDate
        public FRectBaljuDate
        public FRectBaljuCode
        public FRectOnlyNoPackItem
		Public FRectIsFinished

        public FRectDivCode
        public FRectStatecd
        public FRectBaljuname
        public FRectTargetid
        public FRectTargetName

        '샵 발주 상품 리스트
        public Sub GetOfflineBaljuItemList()
                dim i,sqlStr

                sqlStr = " SELECT " + VbCrLf
                sqlStr = sqlStr + " 	b.baljukey as baljunum " + VbCrLf
                sqlStr = sqlStr + " 	, m.siteseq " + VbCrLf
                sqlStr = sqlStr + " 	, bm.baljudate " + VbCrLf
                sqlStr = sqlStr + " 	, '' as targetid " + VbCrLf
                sqlStr = sqlStr + " 	, m.shopid as baljuid " + VbCrLf
                sqlStr = sqlStr + " 	, '' as targetname " + VbCrLf
                sqlStr = sqlStr + " 	, m.shopname as baljuname " + VbCrLf
                sqlStr = sqlStr + " 	, m.ordercode as baljucode " + VbCrLf
                sqlStr = sqlStr + " 	, d.brandname as makerid " + VbCrLf
                sqlStr = sqlStr + " 	, 'M' as maeipdiv " + VbCrLf
                sqlStr = sqlStr + " 	, d.itemgubun " + VbCrLf
                sqlStr = sqlStr + " 	, d.itemid " + VbCrLf
                sqlStr = sqlStr + " 	, d.itemoption " + VbCrLf
                sqlStr = sqlStr + " 	, d.itemname " + VbCrLf
                sqlStr = sqlStr + " 	, d.itemoptionname " + VbCrLf
				sqlStr = sqlStr + " 	, isnull(i.imagesmall, '') as imgsmall " + VbCrLf
                sqlStr = sqlStr + " 	, d.packingstate as boxno " + VbCrLf
                sqlStr = sqlStr + " 	, isnull(d.songjangno,'0') as boxsongjangno " + VbCrLf
                sqlStr = sqlStr + " 	, d.requestedno as totalbaljuno " + VbCrLf
                sqlStr = sqlStr + " 	, 0 as totalupcheno " + VbCrLf
                sqlStr = sqlStr + " 	, 0 as totalofflineno " + VbCrLf
                sqlStr = sqlStr + " 	, d.requestedno as totaltenbaeno " + VbCrLf
                sqlStr = sqlStr + " 	, (d.requestedno - d.fixedno) as totalnopackno " + VbCrLf
                sqlStr = sqlStr + " 	, (case when (isnull(d.songjangno,'0') = '0') then d.fixedno else 0 end) as totalpackno " + VbCrLf
                sqlStr = sqlStr + " 	, (case when (isnull(d.songjangno,'0') <> '0') then d.fixedno else 0 end) as totaldeliverno " + VbCrLf
                sqlStr = sqlStr + " 	, (case when m.beasongdate is null then d.fixedno else 0 end) as totalconfirmno " + VbCrLf
                sqlStr = sqlStr + " 	, isnull(p.itemno,0) as totalprepackno " + VbCrLf
                sqlStr = sqlStr + " 	, d.comment " + VbCrLf
				sqlStr = sqlStr + " 	, bm.IsFinished " + VbCrLf

                sqlStr = sqlStr + GetFromWhere + " 	and m.beasongdate is null "

                sqlStr = sqlStr + " ORDER BY " + VbCrLf
                sqlStr = sqlStr + " 	d.brandname, d.itemgubun, d.itemid, d.itemoption, m.ordercode " + VbCrLf

                'response.write sqlStr
                'response.end

                rsget_Logistics.Open sqlStr, dbget_Logistics, 1

                FResultCount = rsget_Logistics.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget_Logistics.EOF  then
                        i = 0
                        do until rsget_Logistics.eof
                                set FItemList(i) = new COfflineBaljuItem

								FItemList(i).FSiteSeq           = rsget_Logistics("siteseq")

                                FItemList(i).FBaljuNum          = rsget_Logistics("baljunum")
                                FItemList(i).FBaljuDate         = rsget_Logistics("baljudate")
                                FItemList(i).FTargetId          = rsget_Logistics("targetid")
                                FItemList(i).FBaljuId           = rsget_Logistics("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget_Logistics("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget_Logistics("baljuname"))

                                FItemList(i).Fbaljucode         = rsget_Logistics("baljucode")
                                FItemList(i).FMakerid           = rsget_Logistics("makerid")
                                FItemList(i).Fmaeipdiv          = rsget_Logistics("maeipdiv")
                                FItemList(i).FItemGubun         = rsget_Logistics("ItemGubun")
                                FItemList(i).FItemId            = rsget_Logistics("ItemId")
                                FItemList(i).FItemoption        = rsget_Logistics("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget_Logistics("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget_Logistics("ItemOptionname"))

								IF (CStr(FItemList(i).FSiteSeq) = "10") THEN
								    FItemList(i).Fimgsmall      = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).FItemID) + "/" + rsget_Logistics("imgsmall")
								ELSE
								    FItemList(i).Fimgsmall      = rsget_Logistics("imgsmall")
							    END IF

                                FItemList(i).FRealBoxNo         = rsget_Logistics("boxno")
                                FItemList(i).FBoxSongjangNo     = rsget_Logistics("boxsongjangno")

                                FItemList(i).Ftotalbaljuno      = rsget_Logistics("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget_Logistics("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget_Logistics("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget_Logistics("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget_Logistics("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget_Logistics("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget_Logistics("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget_Logistics("totalconfirmno")

                                FItemList(i).Ftotalprepackno    = rsget_Logistics("totalprepackno")

                                FItemList(i).Fcomment    		= rsget_Logistics("comment")
								FItemList(i).FIsFinished   		= rsget_Logistics("IsFinished")

                                rsget_Logistics.MoveNext
                                i = i + 1
                        loop
                end if
                rsget_Logistics.close
        end sub

        '샵 발주 리스트
        public Sub GetOfflineBaljuList()
                dim i,sqlStr

                sqlStr = " SELECT " + VbCrLf
                sqlStr = sqlStr + " 	bm.baljukey as baljunum " + VbCrLf
                sqlStr = sqlStr + " 	, m.siteseq " + VbCrLf
                sqlStr = sqlStr + " 	, bm.baljudate " + VbCrLf
                sqlStr = sqlStr + " 	, '' as targetid " + VbCrLf
                sqlStr = sqlStr + " 	, m.shopid as baljuid " + VbCrLf
                sqlStr = sqlStr + " 	, '' as targetname " + VbCrLf
                sqlStr = sqlStr + " 	, m.shopname as baljuname " + VbCrLf
                sqlStr = sqlStr + " 	, 'M' as maeipdiv " + VbCrLf
                sqlStr = sqlStr + " 	, sum(d.requestedno) as totalbaljuno " + VbCrLf
                sqlStr = sqlStr + " 	, 0 as totalupcheno " + VbCrLf
                sqlStr = sqlStr + " 	, 0 as totalofflineno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(d.requestedno) as totaltenbaeno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(d.requestedno - d.fixedno) as totalnopackno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(case when (isnull(d.songjangno,'0') = '0') then d.fixedno else 0 end) as totalpackno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(case when (isnull(d.songjangno,'0') <> '0') then d.fixedno else 0 end) as totaldeliverno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(case when m.beasongdate is not null then d.fixedno else 0 end) as totalconfirmno " + VbCrLf
                sqlStr = sqlStr + " 	, sum(isnull(p.itemno,0)) as totalprepackno " + VbCrLf
				sqlStr = sqlStr + " 	, bm.isFinished " + VbCrLf

                sqlStr = sqlStr + GetFromWhere

                sqlStr = sqlStr + " GROUP BY " + VbCrLf
                sqlStr = sqlStr + " 	m.siteseq, bm.baljukey, bm.baljudate, m.shopid, m.shopname, bm.isFinished " + VbCrLf

                sqlStr = sqlStr + " ORDER BY " + VbCrLf
                sqlStr = sqlStr + " 	bm.baljukey desc, m.shopname " + VbCrLf

                'response.write sqlStr
                'response.end

                rsget_Logistics.Open sqlStr, dbget_Logistics, 1

                FResultCount = rsget_Logistics.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget_Logistics.EOF  then
                        i = 0
                        do until rsget_Logistics.eof
                                set FItemList(i) = new COfflineBaljuItem

								FItemList(i).FBaljuNum          = rsget_Logistics("baljunum")

								FItemList(i).FSiteSeq           = rsget_Logistics("siteseq")

                                FItemList(i).FBaljuDate         = rsget_Logistics("baljudate")
                                FItemList(i).FTargetId          = rsget_Logistics("targetid")
                                FItemList(i).FBaljuId           = rsget_Logistics("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget_Logistics("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget_Logistics("baljuname"))

                                FItemList(i).Ftotalbaljuno      = rsget_Logistics("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget_Logistics("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget_Logistics("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget_Logistics("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget_Logistics("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget_Logistics("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget_Logistics("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget_Logistics("totalconfirmno")

                                FItemList(i).Ftotalprepackno    = rsget_Logistics("totalprepackno")

								FItemList(i).FisFinished    	= rsget_Logistics("isFinished")

                                rsget_Logistics.MoveNext
                                i = i + 1
                        loop
                end if
                rsget_Logistics.close
        end sub

        '샵 리스트
        public Sub GetOfflineBaljuShopList()
                dim i,sqlStr

                sqlStr = " SELECT distinct " + VbCrLf
                sqlStr = sqlStr + " 	m.shopid as baljuid " + VbCrLf
                sqlStr = sqlStr + " 	, m.shopname as baljuname " + VbCrLf

                sqlStr = sqlStr + GetFromWhere + " 	and m.beasongdate is null "

                sqlStr = sqlStr + " ORDER BY " + VbCrLf
                sqlStr = sqlStr + " 	m.shopname " + VbCrLf
                'response.write sqlStr
                rsget_Logistics.Open sqlStr, dbget_Logistics, 1

                FResultCount = rsget_Logistics.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget_Logistics.EOF  then
                        i = 0
                        do until rsget_Logistics.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuId           = rsget_Logistics("baljuid")
                                FItemList(i).FBaljuName         = db2html(rsget_Logistics("baljuname"))

                                rsget_Logistics.MoveNext
                                i = i + 1
                        loop
                end if
                rsget_Logistics.close
        end sub

		Function GetMaxBoxNo
			dim i,sqlStr
			dim tmp

			sqlStr = " select "
            sqlStr = sqlStr + " 	Max(IsNull(d.packingstate, '0')) as boxno "

			'모든 박스
			tmp = FRectBoxNo
			FRectBoxNo = ""
			sqlStr = sqlStr + GetFromWhere + " 	and m.beasongdate is null "
			FRectBoxNo = tmp

			if (FRectBaljuId <> "") and (FRectBaljuDate <> "") then
				rsget_Logistics.Open sqlStr, dbget_Logistics, 1
				if  not rsget_Logistics.EOF  then
					GetMaxBoxNo    = rsget_Logistics("boxno")
				else
					GetMaxBoxNo    = "0"
				end if
				rsget_Logistics.close
			else
				GetMaxBoxNo    = "0"
			end if
		End Function

		Function GetBoxStateArray
			dim i,sqlStr
			dim result				' "0,0|1,1234|2|0"
			dim tmp

			sqlStr = " select "
			sqlStr = sqlStr + " 	IsNull(d.packingstate, '0') as boxno, IsNull(d.songjangno,'0') as boxsongjangno "

			'모든 박스
			tmp = FRectBoxNo
			FRectBoxNo = ""
			sqlStr = sqlStr + GetFromWhere + " 	and m.beasongdate is null " + VbCrLf
			FRectBoxNo = tmp

			sqlStr = sqlStr + " group by IsNull(d.packingstate, '0'), IsNull(d.songjangno,'0') "
			sqlStr = sqlStr + " order by IsNull(d.packingstate, '0'), IsNull(d.songjangno,'0') "
			'response.write sqlStr
			rsget_Logistics.Open sqlStr, dbget_Logistics, 1

            if  not rsget_Logistics.EOF  then
                i = 0
                do until rsget_Logistics.eof
                    result = result & "|" & rsget_Logistics("boxno") & "," & rsget_Logistics("boxsongjangno")

                    rsget_Logistics.MoveNext
                    i = i + 1
                loop
            end if
            rsget_Logistics.close

			GetBoxStateArray = result

		End Function

		public function GetFromWhere()
			dim tmpsql

			tmpsql = " from " + VbCrLf
			tmpsql = tmpsql + " 	db_aLogistics.dbo.tbl_Logistics_offline_baljumaster bm " + VbCrLf
			tmpsql = tmpsql + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_baljudetail b " + VbCrLf
			tmpsql = tmpsql + " 	on " + VbCrLf
			tmpsql = tmpsql + " 		bm.baljukey = b.baljukey " + VbCrLf
			tmpsql = tmpsql + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_order_master m " + VbCrLf
			tmpsql = tmpsql + " 	on " + VbCrLf
			tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
			tmpsql = tmpsql + " 		and m.siteSeq = 10 " + VbCrLf
			tmpsql = tmpsql + " 		and b.ordercode = m.ordercode " + VbCrLf
			tmpsql = tmpsql + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail d " + VbCrLf
			tmpsql = tmpsql + " 	on " + VbCrLf
			tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
			tmpsql = tmpsql + " 		and m.siteSeq = d.siteSeq " + VbCrLf
			tmpsql = tmpsql + " 		and m.ordercode = d.ordercode " + VbCrLf
			tmpsql = tmpsql + " 	left join [db_aLogistics].[dbo].tbl_Logistics_offline_item i " + VbCrLf
			tmpsql = tmpsql + " 	on " + VbCrLf
			tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
			tmpsql = tmpsql + " 		and d.siteseq = i.siteseq " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemgubun = i.siteitemgubun " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemid = i.siteitemid " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemoption = i.siteitemoption " + VbCrLf
			tmpsql = tmpsql + " 	left join [db_aLogistics].[dbo].tbl_Logistics_offline_tmppacking p " + VbCrLf
			tmpsql = tmpsql + " 	on " + VbCrLf
			tmpsql = tmpsql + " 		1 = 1 " + VbCrLf
			tmpsql = tmpsql + " 		and d.siteseq = p.siteseq " + VbCrLf
			tmpsql = tmpsql + " 		and d.ordercode = p.ordercode " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemgubun = p.itemgubun " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemid = p.itemid " + VbCrLf
			tmpsql = tmpsql + " 		and d.itemoption = p.itemoption " + VbCrLf
			tmpsql = tmpsql + " where " + VbCrLf
			tmpsql = tmpsql + " 	1 = 1 " + VbCrLf
			tmpsql = tmpsql + " 	and d.cancelyn <> 'Y' " + VbCrLf
			''tmpsql = tmpsql + " 	and m.beasongdate is null " + VbCrLf

            if (FRectSiteSeq <> "") then
            	tmpsql = tmpsql + " 	and m.SiteSeq = " & FRectSiteSeq & " " + VbCrLf
            end if

            if (FRectBoxNo <> "") and (FRectBoxNo <> "0") then
            	tmpsql = tmpsql + " 	and ((d.packingstate = " & FRectBoxNo & ") or (p.boxno = " & FRectBoxNo & ")) " + VbCrLf
            end if

            if (FRectSelectedOnly <> "N") then
            	tmpsql = tmpsql + " 	and m.shopid = '" & FRectBaljuId & "' " + VbCrLf
            end if

            if (FRectBaljuNum <> "") then
                    tmpsql = tmpsql + "                 and b.baljukey = '" + CStr(FRectBaljuNum) + "' "
            end if

            if (FRectBaljuDate <> "") then
                tmpsql = tmpsql + " 	and bm.baljudate >= '" & FRectBaljuDate & "' " + VbCrLf
                tmpsql = tmpsql + " 	and bm.baljudate < '" & CStr(Left(dateadd("d",1,FRectBaljuDate),10)) & "' " + VbCrLf
            end if

            if (FRectStartDate <> "") then
                tmpsql = tmpsql + " 	and bm.baljudate >= '" & FRectStartDate & "' " + VbCrLf
            end if

            if (FRectEndDate <> "") then
                tmpsql = tmpsql + " 	and bm.baljudate < '" & FRectEndDate & "' " + VbCrLf
            end if

            if (FRectOnlyNoPackItem = "Y") then
                    tmpsql = tmpsql + " and (((d.requestedno - d.fixedno) > 0) or (IsNull(d.comment, '') <> '')) "
            end If

			if (FRectIsFinished <> "") then
            	tmpsql = tmpsql + " 	and bm.isFinished = '" & FRectIsFinished & "' " + VbCrLf
            end if

            GetFromWhere = tmpsql

		end function

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
                HasPreScroll = StarScrollPage > 1
        end Function

        public Function HasNextScroll()
                HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
        end Function

        public Function StarScrollPage()
                StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
        end Function

end Class

'샵 리스트
public Function GetShopnameByShopid(baljudate, baljuid)
        dim i,sqlStr, result

        sqlStr = " SELECT top 1 " + VbCrLf
        sqlStr = sqlStr + " 	m.shopname as baljuname " + VbCrLf
	    sqlStr = sqlStr + " from "
	    sqlStr = sqlStr + " 	[db_aLogistics].[dbo].tbl_Logistics_offline_baljudetail b "
	    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_order_master m "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		1 = 1 "
	    sqlStr = sqlStr + " 		and b.siteseq = m.siteseq "
	    sqlStr = sqlStr + " 		and b.ordercode = m.ordercode "
	    sqlStr = sqlStr + " 	join [db_aLogistics].[dbo].tbl_Logistics_offline_order_detail d "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		1 = 1 "
	    sqlStr = sqlStr + " 		and m.siteseq = d.siteseq "
	    sqlStr = sqlStr + " 		and m.ordercode = d.ordercode "
	    sqlStr = sqlStr + " 	left join db_aLogistics.dbo.tbl_Logistics_offline_tmppacking p "
	    sqlStr = sqlStr + " 	on "
	    sqlStr = sqlStr + " 		1 = 1 "
	    sqlStr = sqlStr + " 		and p.siteseq = m.siteseq "
	    sqlStr = sqlStr + " 		and p.ordercode = m.ordercode "
	    sqlStr = sqlStr + " 		and p.itemgubun = d.itemgubun "
	    sqlStr = sqlStr + " 		and p.itemid = d.itemid "
	    sqlStr = sqlStr + " 		and p.itemoption = d.itemoption "
	    sqlStr = sqlStr + " where "
	    sqlStr = sqlStr + " 	1 = 1 "
	    sqlStr = sqlStr + " and b.baljuKey in ( "
	    sqlStr = sqlStr + " 	select "
	    sqlStr = sqlStr + "     	baljuKey "
	    sqlStr = sqlStr + " 	from "
	    sqlStr = sqlStr + "         [db_aLogistics].[dbo].tbl_Logistics_offline_baljumaster "
	    sqlStr = sqlStr + "     where "
	    sqlStr = sqlStr + "     	1 = 1 "
	    sqlStr = sqlStr + "         and baljudate >= '" + CStr(baljudate) + "' "
	    sqlStr = sqlStr + "         and baljudate < '" + CStr(Left(dateadd("d",1,baljudate),10)) + "' "
	    sqlStr = sqlStr + " ) "
	    sqlStr = sqlStr + " and m.shopid = '" + CStr(baljuid) + "' "
        'response.write sqlStr
        rsget_Logistics.Open sqlStr, dbget_Logistics, 1

        if  not rsget_Logistics.EOF  then
            result = db2html(rsget_Logistics("baljuname"))
        end if
        rsget_Logistics.close

        GetShopnameByShopid = result
end Function

%>
