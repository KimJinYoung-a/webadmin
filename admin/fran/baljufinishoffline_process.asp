<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
' 사용안하는듯
response.end

dim mode
dim baljunum, baljuid, baljudate, itemgubun, itemid, itemoption, comment
dim i,cnt,sqlStr, errstring
dim masteridxlist, baljuname, baljucodelist, divcode, vatinclude, targetid, targetname, baljucode, brandlist, obaljucode

dim refer
refer = request.ServerVariables("HTTP_REFERER")

mode = request("mode")
baljunum = request("baljunum")
baljuid = request("baljuid")
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")
comment = request("comment")

dim itemexists, iid, newbaljucode, itemAlreadyExists, tmp

if mode="chulgoproc" then

        '잘못된 입력(박스번호가 0 이면서 송장번호가 있는경우 or realitemno 가 있으면서, 박스번호가 없는경우)체크
        sqlStr = " select d.itemname,d.itemoptionname "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "
        sqlStr = sqlStr + " and m.deldt is null "
        sqlStr = sqlStr + " and d.deldt is null "
        sqlStr = sqlStr + " and b.baljucode = m.baljucode "
        sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + " and (((isnull(d.packingstate,0) = 0) and (isnull(d.boxsongjangno,'0') <> '0')) or ((d.realitemno > 0) and (isnull(d.boxsongjangno,'0') = '0'))) "

        if (baljuid <> "") then
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
        end if

        if (baljunum <> "") then
                sqlStr = sqlStr + " and b.baljunum = '" + CStr(baljunum) + "' "
        end if

        if (baljudate <> "") then
                sqlStr = sqlStr + " and b.baljudate >= '" + CStr(baljudate) + "' "
                sqlStr = sqlStr + " and b.baljudate < '" + CStr(Left(dateadd("d",1,baljudate),10)) + "' "
        end if

        rsget.Open sqlStr, dbget, 1
        if  not rsget.EOF  then
                do until rsget.eof
                        if (trim(errstring) = "") then
                                errstring = rsget("itemname") + "(" + rsget("itemoptionname") + ")"
                        else
                                errstring = errstring + ", " + rsget("itemname") + "(" + rsget("itemoptionname") + ")"
                        end if

                        rsget.MoveNext
                loop
        else
                errstring = ""
        end if
        rsget.close

        if (errstring <> "") then
                response.write "<script>alert('잘못된 입력이 있습니다. 송장번호 또는 상품코드를 삭제후 다시 입력하세요.\n\n" + errstring + "');</script>"
                response.write "<script>history.back();</script>"
                dbget.close()	:	response.End
        end if


        '해당 발주코드/발주아이디에 대한 masteridx 를 구한다.
        sqlStr = " select distinct d.masteridx "
        sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
        sqlStr = sqlStr + " where 1 = 1 "
        sqlStr = sqlStr + " and m.idx = d.masteridx "
        sqlStr = sqlStr + " and m.deldt is null "
        sqlStr = sqlStr + " and d.deldt is null "
        sqlStr = sqlStr + " and b.baljucode = m.baljucode "
        sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
        sqlStr = sqlStr + " and m.statecd <> '7' "
        sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
        sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
        rsget.Open sqlStr,dbget,1

        masteridxlist = ""
        if  not rsget.EOF  then
                do until rsget.eof
                        if (masteridxlist <> "") then
                                masteridxlist = masteridxlist + "," + CStr(rsget("masteridx"))
                        else
                                masteridxlist = CStr(rsget("masteridx"))
                        end if

                        rsget.MoveNext
                loop
        end if
        rsget.close

        if (masteridxlist = "") then
                response.write "<script>alert('해당 발주번호/샆 에 대한 출고처리를 할수 없습니다.\n이미 출고완료된 주문서는 재출고될수 없습니다.');</script>"
                response.write "<script>history.back();</script>"
                dbget.close()	:	response.End
        end if


		'기본 master 정보
		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
		sqlStr = sqlStr + " set beasongdate='" + Left(now(), 10) + "'" + vbCrlf
		sqlStr = sqlStr + " where idx in (" + masteridxlist + ") "
		rsget.Open sqlStr, dbget, 1
		'response.write sqlStr


        '미배송주문상품에 대한 코맨트 입력
        if (Trim(comment) <> "") then
        	itemgubun = split(itemgubun,"|")
        	itemid = split(itemid,"|")
        	itemoption = split(itemoption,"|")
        	comment = split(comment,"|")
        	cnt = ubound(itemgubun)

        	for i=0 to cnt
    	        if (Trim(comment(i)) <> "") then
    	                '1. 미배송 상품에 대한 코맨트 입력
    	                sqlstr = " update [db_storage].[dbo].tbl_ordersheet_detail "
    	                sqlstr = sqlstr + " set comment = '" + Trim(comment(i)) + "' "
    	                sqlstr = sqlstr + " where itemgubun = '" + Trim(itemgubun(i)) + "' "
    	                sqlstr = sqlstr + " and itemid = " + Trim(itemid(i)) + " "
    	                sqlstr = sqlstr + " and itemoption = '" + Trim(itemoption(i)) + "' "
    	                sqlstr = sqlstr + " and masteridx in (" + CStr(masteridxlist) + ") "
    	                sqlstr = sqlstr + " and baljuitemno <> realitemno "
    	                sqlstr = sqlstr + " and deldt is null "
    	                rsget.Open sqlStr, dbget, 1
                    end if
            next


        	'2. 미배송주문 내역 체크(재주문 대상 상품 검색)
        	sqlStr = " select count(d.idx) as cnt  from [db_storage].[dbo].tbl_ordersheet_detail d "
        	sqlStr = sqlStr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
        	sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
        	sqlStr = sqlStr + " and d.comment='5일내출고' "
        	sqlStr = sqlStr + " and deldt is null "
                'response.write sqlStr
        	rsget.Open sqlStr, dbget, 1
        	itemexists = (rsget("cnt")>0)
        	rsget.Close

        	sqlStr = " select count(idx) as cnt from  [db_storage].[dbo].tbl_ordersheet_master"
        	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
        	sqlStr = sqlStr + " and clinkcode  is not null "
        	sqlStr = sqlStr + " and clinkcode<>'' "
        	rsget.Open sqlStr, dbget, 1
        	itemAlreadyExists = (rsget("cnt")>0)
        	rsget.Close

        	if Not itemexists then
        		'response.write "<script>alert('재 주문할 내역이 없습니다.');</script>"
        	elseif itemAlreadyExists then
        		'response.write "<script>alert('재 주문서가 이미 작성되어 있습니다. 작성할 수 없습니다.');</script>"
        	else
            	'미배송 주문서 작성
            	sqlStr = " select top 1 * from [db_storage].[dbo].tbl_ordersheet_master"
            	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
            	rsget.Open sqlStr, dbget, 1
        		targetid = rsget("targetid")
        		targetname = rsget("targetname")
        		divcode = rsget("divcode")
        		vatinclude = rsget("vatinclude")
            	rsget.Close

                '해당 발주코드/발주아이디에 대한 기본정보를 구한다.
                sqlStr = " select distinct m.baljuname, m.baljucode "
                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                sqlStr = sqlStr + " where 1 = 1 "
                sqlStr = sqlStr + " and m.idx = d.masteridx "
                sqlStr = sqlStr + " and m.deldt is null "
                sqlStr = sqlStr + " and d.deldt is null "
                sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
                sqlStr = sqlStr + " and m.statecd <> '7' "
                sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
                sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
            	sqlStr = sqlStr + " and d.baljuitemno <> d.realitemno "
            	sqlStr = sqlStr + " and d.comment='5일내출고' "
            	sqlStr = sqlStr + " and d.deldt is null "
                rsget.Open sqlStr,dbget,1

                baljuname = ""
                baljucode = ""
                baljucodelist = ""
                if  not rsget.EOF  then
                        baljuname = CStr(rsget("baljuname"))
                        baljucode = CStr(rsget("baljucode"))
                        baljucodelist = CStr(rsget("baljucode"))

                        rsget.MoveNext

                        do until rsget.eof
                                baljucodelist = baljucodelist + "," + CStr(rsget("baljucode"))
                                rsget.MoveNext
                        loop
                end if
                rsget.close



            	sqlStr = " select * from [db_storage].[dbo].tbl_ordersheet_master where 1=0 "
            	rsget.Open sqlStr,dbget,1,3
            	rsget.AddNew
            	rsget("targetid") = targetid
            	rsget("targetname") = targetname
            	rsget("baljuid") = baljuid
            	rsget("baljuname") = baljuname
            	rsget("reguser") = session("ssBctId")
            	rsget("regname") = session("ssBctCname")
            	rsget("divcode") = divcode
            	rsget("vatinclude") = vatinclude
            	rsget("scheduledate") = Left(now(), 10)
            	rsget("statecd") = "0"
            	rsget("comment") = baljucodelist + " 미배송건 재작성"

            	rsget.update
            	iid = rsget("idx")
            	rsget.close

            	baljucode = "RJ" + Format00(6,Right(CStr(iid),6))

            	''디테일 저장
            	sqlStr = " insert into [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
            	sqlStr = sqlStr + " (masteridx,itemgubun,makerid,itemid,itemoption," + vbCrlf
            	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
            	sqlStr = sqlStr + " baljuitemno,realitemno,baljudiv)"  + vbCrlf
            	sqlStr = sqlStr + " select " + CStr(iid) + ",itemgubun,makerid,itemid,itemoption," + vbCrlf
            	sqlStr = sqlStr + " itemname,itemoptionname,sellcash,suplycash,buycash," + vbCrlf
            	sqlStr = sqlStr + " sum(baljuitemno-realitemno),sum(baljuitemno-realitemno),baljudiv" + vbCrlf
            	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
            	sqlStr = sqlStr + " where masteridx in (" + CStr(masteridxlist) + ") "
            	sqlStr = sqlStr + " and baljuitemno <> realitemno "
            	sqlStr = sqlStr + " and comment='5일내출고'"
            	sqlStr = sqlStr + " and deldt is null"
            	sqlStr = sqlStr + " group by itemgubun,makerid,itemid,itemoption,itemname,itemoptionname,sellcash,suplycash,buycash,baljudiv "
            	rsget.Open sqlStr, dbget, 1


            	''서머리 저장
            	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
            	sqlStr = sqlStr + " set jumunsellcash=IsNULL(T.totsell,0)" + vbCrlf
            	sqlStr = sqlStr + " ,jumunsuplycash=IsNULL(T.totsupp,0)" + vbCrlf
            	sqlStr = sqlStr + " ,jumunbuycash=IsNULL(T.totbuy,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.realtotsell,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.realtotsupp,0)" + vbCrlf
            	sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.realtotbuy,0)" + vbCrlf
            	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
            	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
            	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
            	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
            	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
            	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
            	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
            	sqlStr = sqlStr + " where masteridx="  + CStr(iid) + vbCrlf
            	sqlStr = sqlStr + " and deldt is null" + vbCrlf
            	sqlStr = sqlStr + " ) as T" + vbCrlf
            	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(iid)
            	rsget.Open sqlStr, dbget, 1


            	''브랜드 리스트
            	brandlist = ""
            	sqlStr = " select distinct makerid from [db_storage].[dbo].tbl_ordersheet_detail"
            	sqlStr = sqlStr + " where masteridx=" + CStr(iid)
            	rsget.Open sqlStr, dbget, 1
            		do until rsget.eof
            			brandlist = brandlist + rsget("makerid") + ","
            			rsget.movenext
            		loop
            	rsget.close

            	if brandlist<>"" then
            		brandlist = Left(brandlist,Len(brandlist)-1)
            		brandlist = Left(brandlist,255)
            	end if

            	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
            	sqlStr = sqlStr + " set baljucode='" + baljucode + "'" + VbCrlf
            	'sqlStr = sqlStr + " , obaljucode='" + obaljucode + "'" + VbCrlf
            	sqlStr = sqlStr + " , brandlist='" + brandlist + "'"
            	sqlStr = sqlStr + " where idx=" + CStr(iid)
            	rsget.Open sqlStr, dbget, 1


            	''원발주서에 링크코드 저장.
            	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
            	sqlStr = sqlStr + " set clinkcode='" + baljucode + "'" + VbCrlf
            	sqlStr = sqlStr + " where idx in (" + CStr(masteridxlist) + ") "
            	rsget.Open sqlStr, dbget, 1

               	'response.write "<script>alert('재 주문서가 작성되어 있습니다.');</script>"
        	end if

        end if



        '각 주문코드별 처리(출고데이타 생성 등)
        tmp = split(masteridxlist,",")
        for i=0 to UBound(tmp)
                if (Trim(tmp(i)) <> "") then

                	'확정수량에 대한 합계금액 계산
                	sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
                	sqlStr = sqlStr + " set jumunsellcash=IsNull(T.totsell,0)" + vbCrlf
                	sqlStr = sqlStr + " ,jumunsuplycash=IsNull(T.totsupp,0)" + vbCrlf
                	sqlStr = sqlStr + " ,jumunbuycash=IsNull(T.totbuy,0)" + vbCrlf
                	sqlStr = sqlStr + " ,totalsellcash=IsNull(T.realtotsell,0)" + vbCrlf
                	sqlStr = sqlStr + " ,totalsuplycash=IsNull(T.realtotsupp,0)" + vbCrlf
                	sqlStr = sqlStr + " ,totalbuycash=IsNull(T.realtotbuy,0)" + vbCrlf
                	sqlStr = sqlStr + " from (select sum(sellcash*baljuitemno) as totsell, " + vbCrlf
                	sqlStr = sqlStr + " sum(suplycash*baljuitemno) as totsupp, " + vbCrlf
                	sqlStr = sqlStr + " sum(buycash*baljuitemno) as totbuy, " + vbCrlf
                	sqlStr = sqlStr + " sum(sellcash*realitemno) as realtotsell, " + vbCrlf
                	sqlStr = sqlStr + " sum(suplycash*realitemno) as realtotsupp, " + vbCrlf
                	sqlStr = sqlStr + " sum(buycash*realitemno) as realtotbuy " + vbCrlf
                	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail" + vbCrlf
                	sqlStr = sqlStr + " where masteridx="  + CStr(Trim(tmp(i))) + vbCrlf
                	sqlStr = sqlStr + " and deldt is null" + vbCrlf
                	sqlStr = sqlStr + " ) as T" + vbCrlf
                	sqlStr = sqlStr + " where [db_storage].[dbo].tbl_ordersheet_master.idx=" + CStr(Trim(tmp(i)))
                	rsget.Open sqlStr, dbget, 1


                    '해당 발주코드/발주아이디에 대한 기본정보를 구한다.
                    sqlStr = " select distinct m.baljuname, m.baljucode "
                    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_shopbalju b, [db_storage].[dbo].tbl_ordersheet_master m, [db_storage].[dbo].tbl_ordersheet_detail d "
                    sqlStr = sqlStr + " where 1 = 1 "
                    sqlStr = sqlStr + " and m.idx = d.masteridx "
                    sqlStr = sqlStr + " and m.deldt is null "
                    sqlStr = sqlStr + " and d.deldt is null "
                    sqlStr = sqlStr + " and b.baljucode = m.baljucode "
                    sqlStr = sqlStr + " and m.divcode in ('501','502','503') "
                    sqlStr = sqlStr + " and m.statecd <> '7' "
                    sqlStr = sqlStr + " and b.baljuid = '" + CStr(baljuid) + "' "
                    sqlStr = sqlStr + " and b.baljunum = " + CStr(baljunum) + " "
                    sqlStr = sqlStr + " and m.idx = " + CStr(Trim(tmp(i))) + " "
                    rsget.Open sqlStr,dbget,1

                    baljuname = ""
                    baljucode = ""
                    if  not rsget.EOF  then
                            baljuname = CStr(rsget("baljuname"))
                            baljucode = CStr(rsget("baljucode"))
                    end if
                    rsget.close


                	''출고 마스타에 입력. *-1
                	sqlStr = "select count(idx) as cnt from [db_storage].[dbo].tbl_ordersheet_detail d "
                	sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
                	sqlStr = sqlStr + " and d.deldt is null "
                	sqlStr = sqlStr + " and d.realitemno <> 0 "
                	rsget.Open sqlStr, dbget, 1
                        itemexists = rsget("cnt")>0
                	rsget.close

                	if itemexists then
                		'1.온라인 출고 마스타
                		sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0 "
                		rsget.Open sqlStr,dbget,1,3
                		rsget.AddNew
                		rsget("code") = ""
                		rsget("socid") = baljuid
                		rsget("socname") = baljuname
                		rsget("chargeid") = session("ssBctId")
                		rsget("divcode") = "006"
                		rsget("vatcode") = "008"
                		rsget("comment") = baljucode + " 주문 자동출고처리"
                		rsget("chargename") = session("ssBctCname")
                		rsget("ipchulflag") = "S"

                		rsget.update
                		iid = rsget("id")
                		rsget.close

                		newbaljucode = "SO" + Format00(6,Right(CStr(iid),6))

                		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
                		sqlStr = sqlStr + " set code='" + newbaljucode + "'" + VBCrlf
                		sqlStr = sqlStr + " where id=" + CStr(iid)
                		rsget.Open sqlStr,dbget,1


                		'2.온라인 출고 디테일 입력
                		sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail "
                                sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,itemno, "
                                sqlStr = sqlStr + " buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) "
                                sqlStr = sqlStr + " select '" + newbaljucode + "',d.itemid, d.itemoption, d.sellcash, d.suplycash, "
                                sqlStr = sqlStr + " sum(d.realitemno*-1) as itemno, d.buycash,d.ipgoflag,d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
                                sqlStr = sqlStr + " from [db_storage].[dbo].tbl_ordersheet_detail d "
                                sqlStr = sqlStr + " where d.masteridx = " + CStr(Trim(tmp(i))) + " "
                                sqlStr = sqlStr + " and deldt is null "
                                sqlStr = sqlStr + " and d.realitemno<>0 "
                                sqlStr = sqlStr + " group by d.itemid, d.itemoption, d.sellcash, d.suplycash, d.buycash,d.ipgoflag, "
                                sqlStr = sqlStr + " d.itemgubun,d.itemname,d.itemoptionname,d.makerid "
                		rsget.Open sqlStr,dbget,1


                		'3.온라인 출고 마스타 업데이트
                		sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
                		sqlStr = sqlStr + " set executedt='" + Left(now(), 10) + "'" + VBCrlf
                		sqlStr = sqlStr + " ,scheduledt='" + Left(now(), 10) + "'" + VBCrlf
                		sqlStr = sqlStr + " ,totalsellcash=IsNULL(T.totsell,0)" + VBCrlf
                		sqlStr = sqlStr + " ,totalsuplycash=IsNULL(T.totsupp,0)" + VBCrlf
                		sqlStr = sqlStr + " ,totalbuycash=IsNULL(T.totbuy,0)" + VBCrlf
                		sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
                		sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
                		sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
                		sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
                		sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
                		sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
                		sqlStr = sqlStr + " where mastercode='"  + CStr(newbaljucode) + "'" + vbCrlf
                		sqlStr = sqlStr + " and deldt is null" + vbCrlf
                		sqlStr = sqlStr + " ) as T"
                		sqlStr = sqlStr + " where id=" + CStr(iid)
                		rsget.Open sqlStr,dbget,1


                		'4.상태변경
                		sqlStr = " update [db_storage].[dbo].tbl_ordersheet_master" + vbCrlf
                		sqlStr = sqlStr + " set statecd='7'" + vbCrlf
                		sqlStr = sqlStr + " ,ipgodate='" + Left(now(), 10) + "'" + vbCrlf
                		sqlStr = sqlStr + " ,alinkcode='"  + CStr(newbaljucode) + "'" + vbCrlf
                		sqlStr = sqlStr + " where idx = " + CStr(Trim(tmp(i))) + " "
                		rsget.Open sqlStr, dbget, 1

                        '' 입/출고 재고반영
                        sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & newbaljucode & "','','',0,'',''"
	                    dbget.Execute sqlStr


                        '5.출고된 내역 한정판매 재설정("주문수-확정수" 만큼 빼준다.) -> 오프 상품준비 전환시 뺌.
                        'sqlstr = " update [db_item].[dbo].tbl_item "
                        'sqlstr = sqlstr + " set limitsold=limitsold - T.itemno "
                        'sqlstr = sqlstr + " from ( "
                        'sqlstr = sqlstr + " select sum(d.baljuitemno - d.realitemno) as itemno, d.itemid "
                        'sqlstr = sqlstr + " from [db_storage].[dbo].tbl_ordersheet_detail d, [db_item].[dbo].tbl_item i "
                        'sqlstr = sqlstr + " where d.masteridx in (" + CStr(masteridxlist) + ") "
                        'sqlstr = sqlstr + " and d.itemid=i.itemid "
                        'sqlstr = sqlstr + " and d.deldt is null "
                        'sqlstr = sqlstr + " and d.baljuitemno <> d.realitemno "
                        'sqlstr = sqlstr + " and d.itemgubun = '10' "
                        'sqlstr = sqlstr + " and i.limityn='Y' "
                        'sqlstr = sqlstr + " group by d.itemid "
                        'sqlstr = sqlstr + " ) as T "
                        'sqlstr = sqlstr + " where [db_item].[dbo].tbl_item.itemid=T.itemid "
                        'rsget.Open sqlStr, dbget, 1

                    end if
                end if
        next

        '' 오프 접수수량 재계산
        sqlstr = " exec [db_summary].[dbo].sp_ten_RealtimeStock_offjupsuAll"
        dbget.Execute sqlStr

end if

%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('baljulistoffline.asp');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
