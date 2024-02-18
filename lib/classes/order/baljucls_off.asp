<%

Class COfflineBaljuItem
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
    end function
    
        public function GetDivCodeName()
                if Fdivcode="101" then
                        GetDivCodeName = "가맹점용 개별매입"
                elseif Fdivcode="111" then
                        GetDivCodeName = "가맹점용 개별특정"
                elseif Fdivcode="121" then
                        GetDivCodeName = "온라인특정->가맹점특정"
                elseif Fdivcode="131" then
                        GetDivCodeName = "온라인특정->가맹점매입"
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
                        GetDivCodeName = "온라인특정"
                elseif Fdivcode="501" then
                        GetDivCodeName = "직영샾주문"
                elseif Fdivcode="502" then
                        GetDivCodeName = "수수료샾"
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

        public FRectBaljuNum
        public FRectBaljuId
        public FRectBoxNo
        public FRectBoxHasStock
        public FRectStartDate
        public FRectEndDate
        public FRectBaljuDate
        public FRectBaljuCode
        public FRectOnlyNoPackItem

        public FRectDivCode
        public FRectStatecd
        public FRectBaljuname
        public FRectTargetid
        public FRectTargetName

        '샆 발주 리스트
        public Sub GetOfflineBaljuList()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + "         T.baljunum, T.baljudate, T.targetid, T.baljuid, T.targetname, T.baljuname, T.statecd, "
                sqlStr = sqlStr + "         sum(T.totalbaljuno) as totalbaljuno, "
                sqlStr = sqlStr + "         sum(T.totalupcheno) as totalupcheno, "
                sqlStr = sqlStr + "         sum(T.totaltenbaeno) as totaltenbaeno, "
                sqlStr = sqlStr + "         sum(T.totalofflineno) as totalofflineno, "
                sqlStr = sqlStr + "         sum(T.totalnopackno) as totalnopackno, "
                sqlStr = sqlStr + "         sum(T.totalpackno) as totalpackno, "
                sqlStr = sqlStr + "         sum(T.totaldeliverno) as totaldeliverno, "
                sqlStr = sqlStr + "         sum(T.totalconfirmno) as totalconfirmno "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + "         select "
                sqlStr = sqlStr + "         b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, m.statecd, "
                sqlStr = sqlStr + "         d.baljuitemno as totalbaljuno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + "         (d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + "         (case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + "         from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + "         left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + "         where 1 = 1 "
                sqlStr = sqlStr + "         and m.idx = d.masteridx "
                sqlStr = sqlStr + "         and m.deldt is null "
                sqlStr = sqlStr + "         and d.deldt is null "
                sqlStr = sqlStr + "         and b.baljucode = m.baljucode "
                sqlStr = sqlStr + "         and m.divcode in ('501','503') "

                if (FRectStartDate <> "") then
                        sqlStr = sqlStr + "         and b.baljudate >= '" + CStr(FRectStartDate) + "' "
                end if

                if (FRectEndDate <> "") then
                        sqlStr = sqlStr + "         and b.baljudate < '" + CStr(FRectEndDate) + "' "
                end if

                sqlStr = sqlStr + " ) T "
                sqlStr = sqlStr + " group by T.baljunum, T.baljudate, T.targetid, T.baljuid, T.targetname, T.baljuname, T.statecd "
                sqlStr = sqlStr + " order by T.baljunum desc, T.targetid, T.baljuid "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")
                                
                                FItemList(i).Fstatecd           = rsget("statecd")
                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 발주 상품 리스트
        public Sub GetOfflineBaljuItemList()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + " b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, "
                sqlStr = sqlStr + " m.baljucode, d.makerid, c.prtidx, isnull(i.mwdiv,'9') as maeipdiv, "
                sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, isnull(i.smallimage, '') as imgsmall, d.packingstate as boxno, isnull(d.boxsongjangno,'0') as boxsongjangno, "
                sqlStr = sqlStr + " d.baljuitemno as totalbaljuno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + " (d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + " (case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + " (case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + " (case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_user].[dbo].tbl_user_c c, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','503') "
                sqlStr = sqlStr + " and d.makerid = c.userid "
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(FRectBaljuId) + "' "

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "                 and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                'if (FRectBaljuId <> "") then
                '        sqlStr = sqlStr + "                 and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                'end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                if (FRectBoxNo <> "") then
                        sqlStr = sqlStr + " and d.packingstate = " + CStr(FRectBoxNo) + " "
                end if

                sqlStr = sqlStr + " order by b.baljunum desc, d.packingstate, d.itemgubun, isnull(i.mwdiv,'9') desc, c.prtidx, d.makerid, d.itemid, d.itemoption, m.baljucode "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).Fbaljucode         = rsget("baljucode")
                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).Fprtidx            = rsget("prtidx")
                                FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
                                FItemList(i).FItemGubun         = rsget("ItemGubun")
                                FItemList(i).FItemId            = rsget("ItemId")
                                FItemList(i).FItemoption        = rsget("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget("ItemOptionname"))
                                FItemList(i).FImgSmall          = rsget("imgsmall")
                                if FItemList(i).Fimgsmall<>"" then
                                        FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                                end if

                                FItemList(i).FRealBoxNo         = rsget("boxno")
                                FItemList(i).FBoxSongjangNo     = rsget("boxsongjangno")

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 발주 상품 리스트(관리자 미배송 입력용)
        public Sub GetOfflineBaljuItemListForFinish()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + " b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, "
                sqlStr = sqlStr + " d.makerid, c.prtidx, isnull(i.mwdiv,'9') as maeipdiv, "
                sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, isnull(i.smallimage, '') as imgsmall, "
                sqlStr = sqlStr + " sum(d.baljuitemno) as totalbaljuno, "
                sqlStr = sqlStr + " sum(case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + " sum(case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + " sum(case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + " sum(d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + " sum(case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + " sum(case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + " sum(case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_user].[dbo].tbl_user_c c, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','503') "
                sqlStr = sqlStr + " and d.makerid = c.userid "
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(FRectBaljuId) + "' "

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "                 and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                'if (FRectBaljuId <> "") then
                '        sqlStr = sqlStr + "                 and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                'end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                if (FRectBoxNo <> "") then
                        sqlStr = sqlStr + " and d.packingstate = " + CStr(FRectBoxNo) + " "
                end if

                sqlStr = sqlStr + " group by b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, d.makerid, c.prtidx, isnull(i.mwdiv,'9'), d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, isnull(i.smallimage, '') "
                sqlStr = sqlStr + " order by b.baljunum desc, d.itemgubun, isnull(i.mwdiv,'9') desc, c.prtidx, d.makerid, d.itemid, d.itemoption "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).Fprtidx            = rsget("prtidx")
                                FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
                                FItemList(i).FItemGubun         = rsget("ItemGubun")
                                FItemList(i).FItemId            = rsget("ItemId")
                                FItemList(i).FItemoption        = rsget("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget("ItemOptionname"))
                                FItemList(i).FImgSmall          = rsget("imgsmall")
                                if FItemList(i).Fimgsmall<>"" then
                                        FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                                end if

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 발주 상품 리스트(프린트용)
        public Sub GetOfflineBaljuItemListForPrint()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + " T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.makerid, T.prtidx, T.maeipdiv, T.imgsmall, T.sellcash, "
                sqlStr = sqlStr + " sum(T.totalbaljuno) as totalbaljuno, sum(T.totalupcheno) as totalupcheno, sum(T.totalofflineno) as totalofflineno, sum(T.totaltenbaeno) as totaltenbaeno, "
                sqlStr = sqlStr + " sum(T.totalnopackno) as totalnopackno, sum(T.totalpackno) as totalpackno, sum(T.totaldeliverno) as totaldeliverno, sum(T.totalconfirmno) as totalconfirmno "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + "         select "
                sqlStr = sqlStr + "         d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
                sqlStr = sqlStr + "         d.makerid, c.prtidx, isnull(i.mwdiv,'9') as maeipdiv, isnull(i.smallimage, '') as imgsmall, d.sellcash, "
                sqlStr = sqlStr + "         d.baljuitemno as totalbaljuno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + "         (d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + "         (case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + "         from [db_storage].[dbo].tbl_shopbalju b, [db_user].[dbo].tbl_user_c c, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + "         left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + "         where 1 = 1 "
                sqlStr = sqlStr + "         and m.idx = d.masteridx "
                sqlStr = sqlStr + "         and m.deldt is null "
                sqlStr = sqlStr + "         and d.deldt is null "
                sqlStr = sqlStr + "         and b.baljucode = m.baljucode "
                sqlStr = sqlStr + "         and m.divcode in ('501','503') "
                sqlStr = sqlStr + "         and d.makerid = c.userid "

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "         and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                if (FRectBaljuId <> "") then
                        sqlStr = sqlStr + "         and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "         and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "         and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                sqlStr = sqlStr + " ) T "
                sqlStr = sqlStr + " group by T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.makerid, T.prtidx, T.maeipdiv, T.imgsmall, T.sellcash "

                if (FRectOnlyNoPackItem = "Y") then
                        sqlStr = " select TT.* from ( " + sqlStr + " ) TT where totalnopackno > 0 "
                        sqlStr = sqlStr + " order by TT.prtidx, TT.makerid, TT.itemgubun, TT.maeipdiv desc, TT.itemid, TT.itemoption "
                else
                        sqlStr = sqlStr + " order by T.prtidx, T.makerid, T.itemgubun, T.maeipdiv desc, T.itemid, T.itemoption "
                end if
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).Fprtidx            = rsget("prtidx")
                                FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
                                FItemList(i).FItemGubun         = rsget("ItemGubun")
                                FItemList(i).FItemId            = rsget("ItemId")
                                FItemList(i).FItemoption        = rsget("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget("ItemOptionname"))

                                FItemList(i).FImgSmall          = rsget("imgsmall")
                                if FItemList(i).Fimgsmall<>"" then
                                        FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                                end if

                                FItemList(i).FSellCash          = rsget("sellcash")

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub


        '샆 패킹 상품 리스트(바코드입력)
        public Sub GetOfflineBaljuPackItemList()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + " b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, "
                sqlStr = sqlStr + " m.baljucode, d.makerid, c.prtidx, isnull(i.mwdiv,'9') as maeipdiv, "
                sqlStr = sqlStr + " d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, isnull(i.smallimage, '') as imgsmall, d.packingstate as boxno, d.boxsongjangno, "
                sqlStr = sqlStr + " d.baljuitemno as totalbaljuno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + " (case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + " (d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + " (case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + " (case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + " (case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_user].[dbo].tbl_user_c c, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','503') "
                sqlStr = sqlStr + " and d.makerid = c.userid "

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "                 and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                if (FRectBaljuId <> "") then
                        sqlStr = sqlStr + "                 and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                if (FRectBoxNo <> "") then
                        sqlStr = sqlStr + " and d.packingstate = " + CStr(FRectBoxNo) + " "
                end if

                sqlStr = " select T.*, isnull(p.itemno,0) as totalprepackno from ( " + sqlStr + " ) T "
                sqlStr = sqlStr + " left join [db_storage].[dbo].tbl_shoppacking p "
                sqlStr = sqlStr + " on T.baljunum = p.baljunum and T.baljuid = p.baljuid and T.baljucode = p.baljucode and T.itemgubun = p.itemgubun and T.itemid = p.itemid and T.itemoption = p.itemoption "
                sqlStr = sqlStr + " order by T.baljunum desc, T.itemgubun, T.maeipdiv desc, T.prtidx, T.makerid, T.itemid, T.itemoption, T.baljucode "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).Fbaljucode         = rsget("baljucode")
                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).Fprtidx            = rsget("prtidx")
                                FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
                                FItemList(i).FItemGubun         = rsget("ItemGubun")
                                FItemList(i).FItemId            = rsget("ItemId")
                                FItemList(i).FItemoption        = rsget("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget("ItemOptionname"))
                                FItemList(i).FImgSmall          = rsget("imgsmall")
                                if FItemList(i).Fimgsmall<>"" then
                                        FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                                end if

                                FItemList(i).FRealBoxNo         = rsget("boxno")
                                FItemList(i).FBoxSongjangNo     = rsget("boxsongjangno")

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalprepackno    = rsget("totalprepackno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 발주 상품 리스트(박스 프린트용)
        public Sub GetOfflineBaljuItemListForPrintBox()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + " T.boxno, T.baljucode, T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.makerid, T.prtidx, T.maeipdiv, T.imgsmall, "
                sqlStr = sqlStr + " sum(T.totalbaljuno) as totalbaljuno, sum(T.totalupcheno) as totalupcheno, sum(T.totalofflineno) as totalofflineno, sum(T.totaltenbaeno) as totaltenbaeno, "
                sqlStr = sqlStr + " sum(T.totalnopackno) as totalnopackno, sum(T.totalpackno) as totalpackno, sum(T.totaldeliverno) as totaldeliverno, sum(T.totalconfirmno) as totalconfirmno "
                sqlStr = sqlStr + " from "
                sqlStr = sqlStr + " ( "
                sqlStr = sqlStr + "         select "
                sqlStr = sqlStr + "         isnull(d.packingstate,0) as boxno, m.baljucode, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
                sqlStr = sqlStr + "         d.makerid, c.prtidx, isnull(i.mwdiv,'9') as maeipdiv, isnull(i.smallimage, '') as imgsmall, "
                sqlStr = sqlStr + "         d.baljuitemno as totalbaljuno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = 'U' then d.baljuitemno else 0 end) as totalupcheno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') = '9' then d.baljuitemno else 0 end) as totalofflineno, "
                sqlStr = sqlStr + "         (case when isnull(i.mwdiv,'9') in ('M','W') then d.baljuitemno else 0 end) as totaltenbaeno, "
                sqlStr = sqlStr + "         (d.baljuitemno - d.realitemno) as totalnopackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') = '0') then d.realitemno else 0 end) as totalpackno, "
                sqlStr = sqlStr + "         (case when (isnull(d.boxsongjangno,'0') <> '0') then d.realitemno else 0 end) as totaldeliverno, "
                sqlStr = sqlStr + "         (case when (m.statecd in ('6','7')) then d.realitemno else 0 end) as totalconfirmno "
                sqlStr = sqlStr + "         from [db_storage].[dbo].tbl_shopbalju b, [db_user].[dbo].tbl_user_c c, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + "         left join [db_item].[dbo].tbl_item i on d.itemgubun = '10' and d.itemid = i.itemid "
                sqlStr = sqlStr + "         where 1 = 1 "
                sqlStr = sqlStr + "         and m.idx = d.masteridx "
                sqlStr = sqlStr + "         and m.deldt is null "
                sqlStr = sqlStr + "         and d.deldt is null "
                sqlStr = sqlStr + "         and b.baljucode = m.baljucode "
                sqlStr = sqlStr + "         and m.divcode in ('501','503') "
                sqlStr = sqlStr + "         and d.makerid = c.userid "

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "         and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                if (FRectBaljuId <> "") then
                        sqlStr = sqlStr + "         and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "         and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "         and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                if (FRectBoxNo <> "") then
                        sqlStr = sqlStr + " and d.packingstate = " + CStr(FRectBoxNo) + " "
                end if

                sqlStr = sqlStr + " ) T "
                sqlStr = sqlStr + " group by T.boxno, T.baljucode, T.itemgubun, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.makerid, T.prtidx, T.maeipdiv, T.imgsmall "

                if (FRectOnlyNoPackItem = "Y") then
                        sqlStr = " select TT.* from ( " + sqlStr + " ) TT where totalnopackno > 0 "
                        sqlStr = sqlStr + " order by TT.itemgubun, TT.maeipdiv desc, TT.prtidx, TT.makerid, TT.itemid, TT.itemoption "
                else
                        sqlStr = sqlStr + " order by T.itemgubun, T.maeipdiv desc, T.prtidx, T.makerid, T.itemid, T.itemoption "
                end if
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).Fbaljucode         = rsget("baljucode")

                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).Fprtidx            = rsget("prtidx")
                                FItemList(i).Fmaeipdiv          = rsget("maeipdiv")
                                FItemList(i).FItemGubun         = rsget("ItemGubun")
                                FItemList(i).FItemId            = rsget("ItemId")
                                FItemList(i).FItemoption        = rsget("Itemoption")
                                FItemList(i).FItemName          = db2html(rsget("ItemName"))
                                FItemList(i).FItemOptionname    = db2html(rsget("ItemOptionname"))

                                FItemList(i).FImgSmall          = rsget("imgsmall")
                                if FItemList(i).Fimgsmall<>"" then
                                        FItemList(i).Fimgsmall     = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + FItemList(i).Fimgsmall
                                end if

                                FItemList(i).FRealBoxNo         = rsget("boxno")

                                FItemList(i).Ftotalbaljuno      = rsget("totalbaljuno")
                                FItemList(i).Ftotalupcheno      = rsget("totalupcheno")
                                FItemList(i).Ftotaltenbaeno     = rsget("totaltenbaeno")
                                FItemList(i).Ftotalofflineno    = rsget("totalofflineno")

                                FItemList(i).Ftotalnopackno     = rsget("totalnopackno")
                                FItemList(i).Ftotalpackno       = rsget("totalpackno")
                                FItemList(i).Ftotaldeliverno    = rsget("totaldeliverno")
                                FItemList(i).Ftotalconfirmno    = rsget("totalconfirmno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 출고완료 리스트
        public Sub GetOfflineJumunFinishList()
                dim i,sqlStr

                sqlStr = " select distinct "
                sqlStr = sqlStr + "         b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, m.baljucode, m.statecd, d.boxsongjangno, d.packingstate "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and d.realitemno <> 0 "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','503') "
                sqlStr = sqlStr + " and m.statecd in ('6','7') "

                if (FRectBaljuId <> "") then
                        sqlStr = sqlStr + "                 and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                end if

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "                 and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                if (FRectStartDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectStartDate) + "' "
                end if

                if (FRectEndDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(FRectEndDate) + "' "
                end if

                sqlStr = sqlStr + " order by b.baljunum desc, m.baljuid, m.baljucode desc, d.packingstate "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).Fbaljucode         = rsget("baljucode")
                                FItemList(i).Fstatecd           = rsget("statecd")
                                FItemList(i).FBoxSongjangNo     = rsget("boxsongjangno")
                                FItemList(i).FRealBoxNo         = rsget("packingstate")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

        '샆 출고완료 상세보기
        public Sub GetOfflineJumunFinishView()
                dim i,sqlStr

                sqlStr = " select "
                sqlStr = sqlStr + "         b.baljunum, b.baljudate, m.targetid, m.baljuid, m.targetname, m.baljuname, m.baljucode, m.statecd, "
                sqlStr = sqlStr + "         d.makerid, d.itemgubun, d.itemid, d.itemoption, d.itemname, d.itemoptionname, "
                sqlStr = sqlStr + "         d.packingstate as boxno, isnull(d.boxsongjangno,'0') as boxsongjangno, d.baljuitemno, d.realitemno "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','503') "
                sqlStr = sqlStr + " and m.statecd in ('6','7') "
                sqlStr = sqlStr + " and m.baljucode = '" + CStr(FRectBaljuCode) + "' "

                if (FRectBaljuId <> "") then
                        sqlStr = sqlStr + "                 and b.baljuid = '" + CStr(FRectBaljuId) + "' "
                end if

                if (FRectBaljuNum <> "") then
                        sqlStr = sqlStr + "                 and b.baljunum = '" + CStr(FRectBaljuNum) + "' "
                end if

                if (FRectBaljuDate <> "") then
                        sqlStr = sqlStr + "                 and b.baljudate >= '" + CStr(FRectBaljuDate) + "' "
                        sqlStr = sqlStr + "                 and b.baljudate < '" + CStr(Left(dateadd("d",1,FRectBaljuDate),10)) + "' "
                end if

                sqlStr = sqlStr + " order by d.makerid, d.itemgubun, d.itemid, d.itemoption "
                'response.write sqlStr
                'dbget.close()	:	response.End
                rsget.Open sqlStr, dbget, 1

                FResultCount = rsget.RecordCount

                redim preserve FItemList(FResultCount)
                if  not rsget.EOF  then
                        i = 0
                        do until rsget.eof
                                set FItemList(i) = new COfflineBaljuItem

                                FItemList(i).FBaljuNum          = rsget("baljunum")
                                FItemList(i).FBaljuDate         = rsget("baljudate")
                                FItemList(i).FTargetId          = rsget("targetid")
                                FItemList(i).FBaljuId           = rsget("baljuid")
                                FItemList(i).FTargetName        = db2html(rsget("targetname"))
                                FItemList(i).FBaljuName         = db2html(rsget("baljuname"))

                                FItemList(i).Fbaljucode         = rsget("baljucode")
                                FItemList(i).Fstatecd           = rsget("statecd")

                                FItemList(i).FMakerid           = rsget("makerid")
                                FItemList(i).FItemGubun         = rsget("itemgubun")
                                FItemList(i).FItemId            = rsget("itemid")
                                FItemList(i).FItemoption        = rsget("itemoption")
                                FItemList(i).FItemName          = db2html(rsget("itemname"))
                                FItemList(i).FItemOptionname    = db2html(rsget("itemoptionname"))

                                FItemList(i).FRealBoxNo         = rsget("boxno")
                                FItemList(i).FBoxSongjangNo     = rsget("boxsongjangno")
                                FItemList(i).Ftotalbaljuno      = rsget("baljuitemno")
                                FItemList(i).Ftotalconfirmno    = rsget("realitemno")

                                rsget.MoveNext
                                i = i + 1
                        loop
                end if
                rsget.close
        end sub

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
%>