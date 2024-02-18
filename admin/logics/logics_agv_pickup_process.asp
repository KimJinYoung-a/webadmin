<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : 이상구 생성
'           2020.05.12 정태훈 수정
'           2020.05.20 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim finishid, finishname, menupos, title, pickingStationCd, comment, mode, agvitemnoarr, code, agvitemno
dim itemgubunarr, itemidarr, itemoptionarr, itemnoarr, masteridx, itemgubun, itemid, itemoption, itemno, chk, didxarr
dim i,cnt,sqlStr, didx, AssignedRows, refer, refergubun, refergubunname, chkarr, requestNo, mastercode
	code = requestcheckvar(request("code"),10)
    refergubun = requestcheckvar(request("refergubun"),32)
    mode = RequestCheckVar(request("mode"), 32)
    masteridx = RequestCheckVar(request("masteridx"), 32)
    menupos =  request("menupos")
    finishid = session("ssBctid")
    finishname = html2db(session("ssBctCname"))
    title = html2db(request("title"))
    pickingStationCd = request("pickingStationCd")
    comment = html2db(request("comment"))
    chk = request("chk")
    itemgubunarr = request("itemgubunarr")
    itemidarr = request("itemidarr")
    itemoptionarr = request("itemoptionarr")
    itemnoarr = request("itemnoarr")
    didxarr = request("didxarr")
    mastercode = request("mastercode")

    if (mastercode = "") and (request("code") <> "") then
        mastercode = request("code")
    end if

refer = request.ServerVariables("HTTP_REFERER")

select case mode
    '/////////// 수정시 mode값(agvregarr) 와 엑셀일괄등록(imgstatic/linkweb/item/agv/upload_AGV_item_excel.asp) 과 로직스(/V2/onLine/logics_agv_pickup_process.asp mode값 agvregcs)도 같이 수정하셔야 합니다./////////
    case "write"
        itemgubunarr = split(itemgubunarr, "|")
        itemidarr = split(itemidarr, "|")
        itemoptionarr = split(itemoptionarr, "|")
        itemnoarr = split(itemnoarr, "|")

        '// 신규저장
		sqlStr = " select * from [db_aLogistics].[dbo].[tbl_agv_pickup_master] where 1=0"
		rsget_Logistics.Open sqlStr,dbget_Logistics,1,3
		rsget_Logistics.AddNew
		rsget_Logistics("reguserid") = finishid
		rsget_Logistics("title") = title
		rsget_Logistics("comment") = comment
        rsget_Logistics("stationCd") = pickingStationCd
        rsget_Logistics("status") = 0

        ''idx, reguserid, title, comment, status, stationCd, regdate
        ''idx, masteridx, makerid, itemgubun, itemid, itemoption, skuCd, itemname, itemoptionname, itemno, pickupno, regdate, updt, deldt

		rsget_Logistics.update
		    masteridx = rsget_Logistics("idx")
		rsget_Logistics.close

		for i=0 to UBound(itemgubunarr) - 1
			if (trim(itemgubunarr(i)) <> "") then
				itemgubun = trim(itemgubunarr(i))
				itemid = trim(itemidarr(i))
				itemoption = trim(itemoptionarr(i))
				itemno = CInt(trim(itemnoarr(i)))

				sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
				sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
				sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(itemno) & ") " & VBCrlf
				rsget_Logistics.Open sqlStr,dbget_Logistics,1
			end if
		next

        ' 상품정보업데이트
        call iteminforeg(masteridx)

	    response.write "<script>alert('저장되었습니다.');</script>"
	    response.write "<script>location.replace('logics_agv_pickupList.asp?menupos="+menupos+"')</script>"
    case "deldetail"
        chk= chk + ",,"
	    chk = Split(chk, ",")

        for i=0 to UBound(chk) - 1
            if (trim(chk(i)) <> "") then
                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_detail] "
                sqlStr = sqlStr + " set deldt = getdate() "
                sqlStr = sqlStr + " where idx = " & trim(chk(i)) & " and masteridx = " & masteridx
                dbget_Logistics.Execute sqlStr
            end if
        next

	    response.write "<script>alert('삭제 되었습니다.');</script>"
	    response.write "<script>location.replace('" + refer + "')</script>"
    case "modi"
        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set updt = getdate(), reguserid = '" & finishid & "', title = '" & title & "', stationCd = '" & pickingStationCd & "', comment = '" & comment & "' "
        sqlStr = sqlStr + " where idx = " & masteridx
        dbget_Logistics.Execute sqlStr

	    didxarr = Split(didxarr, ",")
        itemnoarr = Split(itemnoarr, ",")

        for i=0 to UBound(didxarr)
            if (trim(didxarr(i)) <> "") then
                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_detail] "
                sqlStr = sqlStr + " set itemno = " & trim(itemnoarr(i))
                sqlStr = sqlStr + " where idx = " & trim(didxarr(i)) & " and masteridx = " & masteridx
                ''response.write sqlStr & "<br />"
                dbget_Logistics.Execute sqlStr
            end if
        next

	    response.write "<script>alert('저장 되었습니다.');</script>"
	    response.write "<script>location.replace('" + refer + "')</script>"
    case "delmaster"
        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set deldt = getdate() "
        sqlStr = sqlStr + " where idx = " & masteridx
        dbget_Logistics.Execute sqlStr

	    response.write "<script>alert('삭제 되었습니다.');</script>"
	    response.write "<script>location.replace('logics_agv_pickupList.asp?menupos=" + menupos + "')</script>"
    case "adddetail"
        itemgubunarr = split(itemgubunarr, "|")
        itemidarr = split(itemidarr, "|")
        itemoptionarr = split(itemoptionarr, "|")
        itemnoarr = split(itemnoarr, "|")

		for i=0 to UBound(itemgubunarr)
			if (trim(itemgubunarr(i)) <> "") then
				itemgubun = trim(itemgubunarr(i))
				itemid = trim(itemidarr(i))
				itemoption = trim(itemoptionarr(i))
				itemno = CInt(trim(itemnoarr(i)))

                sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_detail] "
                sqlStr = sqlStr + " set itemno = " & itemno & " + (case when deldt is not NULL then 0 else itemno end), deldt = NULL, updt=getdate() "
                sqlStr = sqlStr + " where masteridx = " & masteridx & " and itemgubun = '" & itemgubun & "' and itemid = " & itemid & " and itemoption = '" & itemoption & "' "
                ''response.write sqlStr & "<br />"
                dbget_Logistics.Execute sqlStr, AssignedRows

                if (AssignedRows < 1) then
				    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
				    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
				    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(itemno) & ") " & VBCrlf
				    rsget_Logistics.Open sqlStr,dbget_Logistics,1
                end if
			end if
		next

        ' 상품정보업데이트
        call iteminforeg(masteridx)

	    ''response.write "<script>alert('저장 되었습니다.');</script>"
	    response.write "<script>location.replace('" + refer + "')</script>"

    '/////////// 수정시 mode값(write) 와 엑셀일괄등록(imgstatic/linkweb/item/agv/upload_AGV_item_excel.asp) 과 로직스(/V2/onLine/logics_agv_pickup_process.asp mode값 agvregcs)도 같이 수정하셔야 합니다./////////
    ' AGV인터페이스 에 저장 처리
    case "agvregarr"
        ' 주문서관리
        if ucase(refergubun)="A" then
            itemgubunarr = request("itemgubunarr")
            itemidarr = request("itemidarr")
            itemoptionarr = request("itemoptionarr")
            agvitemnoarr = request("agvitemnoarr")

            if isnull(itemgubunarr) or replace(itemgubunarr,"|","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('선택된 상품이 없습니다.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            itemgubunarr = split(itemgubunarr, "|")
            itemidarr = split(itemidarr, "|")
            itemoptionarr = split(itemoptionarr, "|")
            agvitemnoarr = split(agvitemnoarr, "|")

            refergubunname="주문서"

        ' 출고리스트
        elseif ucase(refergubun)="B" then
            chkarr= request("chk") + ",,"
            itemgubunarr= request("itemgubun") + ",,"
            itemidarr= request("itemid") + ",,"
            itemoptionarr= request("itemoption") + ",,"
            agvitemnoarr= request("itemno") + ",,"

            if isnull(chkarr) or replace(chkarr,",","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('선택된 상품이 없습니다.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

	        chkarr = split(chkarr, ",")
            itemgubunarr = split(itemgubunarr, ",")
            itemidarr = split(itemidarr, ",")
            itemoptionarr = split(itemoptionarr, ",")
            agvitemnoarr = split(agvitemnoarr, ",")

            refergubunname="출고"

        ' 미출고상품
        elseif ucase(refergubun)="C" then
            chkarr  = request("chk")+ ","

            if isnull(chkarr) or replace(chkarr,",","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('선택된 상품이 없습니다.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            chkarr = Split(chkarr, ",")

            refergubunname="미출고상품"

        ' 브랜드별 재고
        elseif ucase(refergubun)="BRANDSTOCK" then
            itemgubunarr = request("itemgubunarr")
            itemidarr = request("itemidarr")
            itemoptionarr = request("itemoptionarr")
            agvitemnoarr = request("itemnoarr")

            if isnull(itemgubunarr) or replace(itemgubunarr,"|","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('선택된 상품이 없습니다.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            itemgubunarr = split(itemgubunarr, ",")
            itemidarr = split(itemidarr, ",")
            itemoptionarr = split(itemoptionarr, ",")
            agvitemnoarr = split(agvitemnoarr, ",")

            refergubunname="브랜드별 재고"

        ' 출고리스트
        else
            response.write "<script type='text/javascript'>"
            response.write "	alert('입출고 구분코드가 없습니다.');"
            response.write "	location.replace('"& refer &"');"
            response.write "</script>"

            response.end
        end if

        '// 마스터 신규저장
        sqlStr = " select * from [db_aLogistics].[dbo].[tbl_agv_pickup_master] where 1=0"
        rsget_Logistics.Open sqlStr,dbget_Logistics,1,3
        rsget_Logistics.AddNew
        rsget_Logistics("reguserid") = finishid
        rsget_Logistics("title") = refergubunname & " " & code
        rsget_Logistics("comment") = NULL
        rsget_Logistics("stationCd") = NULL		' 스테이션
        rsget_Logistics("status") = 0

        ''idx, reguserid, title, comment, status, stationCd, regdate
        ''idx, masteridx, makerid, itemgubun, itemid, itemoption, skuCd, itemname, itemoptionname, itemno, pickupno, regdate, updt, deldt

        rsget_Logistics.update
            masteridx = rsget_Logistics("idx")
        rsget_Logistics.close

        if ucase(refergubun)="A" then
            '주문서관리
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - 업체반품(" & mastercode & ")"
        elseif ucase(refergubun)="B" then
            '출고리스트
            requestNo = "ETC - 기타출고(" & mastercode & ")"
        elseif ucase(refergubun)="C" then
            '미출고상품
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - 미배"
        elseif ucase(refergubun)="BRANDSTOCK" then
            '브랜드별 재고
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - 재고"
        else
            '
        end if

        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set updt = getdate(), requestNo = '" & requestNo & "' "
        sqlStr = sqlStr + " where idx = " & masteridx
        dbget_Logistics.Execute sqlStr

        ' 주문서관리
        if ucase(refergubun)="A" then
            for i=0 to UBound(itemgubunarr)
                if (trim(itemgubunarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(i))
                    itemid = trim(itemidarr(i))
                    itemoption = trim(itemoptionarr(i))
                    agvitemno = CLng(trim(agvitemnoarr(i)))*-1

                    ' 디테일 신규저장
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' 출고리스트
        elseif ucase(refergubun)="B" then
            for i=0 to UBound(chkarr) - 1
                if (trim(chkarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(chkarr(i)))
                    itemid = trim(itemidarr(chkarr(i)))
                    itemoption = trim(itemoptionarr(chkarr(i)))
                    agvitemno = trim(agvitemnoarr(CInt(chkarr(i))))*-1

                    ' 디테일 신규저장
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' 미출고상품
        elseif ucase(refergubun)="C" then
            for i = 0 to UBound(chkarr)
                if (trim(chkarr(i)) <> "") then
                    itemgubun = request("itemgubun" + CStr(trim(chkarr(i))))
                    itemid = request("itemid" + CStr(trim(chkarr(i))))
                    itemoption = request("itemoption" + CStr(trim(chkarr(i))))
                    agvitemno = request("agvitemno" + CStr(trim(chkarr(i))))

                    ' 디테일 신규저장
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next
        '브랜드별 재고
        elseif ucase(refergubun)="BRANDSTOCK" then
            for i=0 to UBound(itemgubunarr)
                if (trim(itemgubunarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(i))
                    itemid = trim(itemidarr(i))
                    itemoption = trim(itemoptionarr(i))
                    agvitemno = CLng(trim(agvitemnoarr(i)))

                    ' 디테일 신규저장
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' 출고리스트
        end if

        ' 상품정보업데이트
        call iteminforeg(masteridx)

        response.write "<script type='text/javascript'>"
        response.write "	alert('저장 되었습니다.');"
        ' 저장 경로가 팝업으로 넘어온 케이스는 부모창이 인터페이스리스트 페이지 이므로 부모창 리로드
        response.write "	if((typeof(opener) != 'undefined') && (typeof(opener.location) != 'undefined')) {"
        response.write "	    alert(opener);"
        response.write "	    opener.location.reload();"
        response.write "	    self.close();"
        ' 저장경로가 팝업이 아닌경우에는 인터페이스리스트 페이지를 팝업으로 띄워준다.
        response.write "	}else{"
        response.write "	    var popwin = window.open('/admin/logics/logics_agv_pickupList.asp','addreg','width=1280,height=960,scrollbars=yes,resizable=yes');"
        response.write "	    popwin.focus();"
        response.write "	}"
        response.write "	location.replace('"& refer &"');"
        response.write "</script>"
    case else
        response.write "잘못된 접속입니다."
end select

function iteminforeg(masteridx)
    if masteridx="" or isnull(masteridx) then exit function

    ' 온라인상품업데이트
    sqlStr = " update d" & vbcrlf
    sqlStr = sqlStr + " set d.skuCd = i.skuCd" & vbcrlf
    sqlStr = sqlStr + " , d.makerid = isnull(i.brandCd,oi.makerid)" & vbcrlf
    sqlStr = sqlStr + " , d.itemname = isnull(i.productName,oi.itemname)" & vbcrlf
    sqlStr = sqlStr + " , d.itemoptionname = isnull(i.optionName,ooi.optionname)" & vbcrlf
    sqlStr = sqlStr + " , d.updt=getdate()" & vbcrlf
    sqlStr = sqlStr + " from [db_aLogistics].[dbo].[tbl_agv_pickup_detail] as d with(noLock)" & vbcrlf
    sqlStr = sqlStr + " join tendb.db_item.dbo.tbl_item as oi with(noLock)" & vbcrlf
    sqlStr = sqlStr + " 	on d.itemgubun='10'" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemid=oi.itemid" & vbcrlf
    sqlStr = sqlStr + " 	and oi.itemid in ( select itemid from [db_aLogistics].[dbo].[tbl_agv_pickup_detail] with(noLock) where deldt is NULL and masteridx = "& masteridx &" )" & vbcrlf
    sqlStr = sqlStr + " left join tendb.db_item.dbo.tbl_item_option as ooi with(noLock)" & vbcrlf
    sqlStr = sqlStr + " 	on d.itemgubun='10'" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemid=ooi.itemid" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemoption = isnull(ooi.itemoption,'0000')" & vbcrlf
    sqlStr = sqlStr + " 	and ooi.itemid in ( select itemid from [db_aLogistics].[dbo].[tbl_agv_pickup_detail] with(noLock) where deldt is NULL and masteridx = "& masteridx &" )" & vbcrlf
    sqlStr = sqlStr + " left join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] as i with(noLock)" & vbcrlf
    sqlStr = sqlStr + " 	on d.itemgubun = i.itemgubun" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemid = i.itemid" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemoption = i.itemoption" & vbcrlf
    sqlStr = sqlStr + " where d.deldt is NULL" & vbcrlf
    sqlStr = sqlStr + " and d.masteridx = "& masteridx &"" & vbcrlf
    sqlStr = sqlStr + " and d.itemgubun='10'" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr

    ' 오프라인상품업데이트
    sqlStr = " update d" & vbcrlf
    sqlStr = sqlStr + " set d.skuCd = i.skuCd" & vbcrlf
    sqlStr = sqlStr + " , d.makerid = isnull(i.brandCd,fi.makerid)" & vbcrlf
    sqlStr = sqlStr + " , d.itemname = isnull(i.productName,fi.shopitemname)" & vbcrlf
    sqlStr = sqlStr + " , d.itemoptionname = isnull(i.optionName,fi.shopitemoptionname)" & vbcrlf
    sqlStr = sqlStr + " , d.updt=getdate()" & vbcrlf
    sqlStr = sqlStr + " from [db_aLogistics].[dbo].[tbl_agv_pickup_detail] as d with(noLock)" & vbcrlf
    sqlStr = sqlStr + " join tendb.db_shop.dbo.tbl_shop_item as fi with(noLock)" & vbcrlf
    sqlStr = sqlStr + " 	on fi.itemgubun <> '10'" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemgubun = fi.itemgubun" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemid = fi.shopitemid" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemoption = fi.itemoption" & vbcrlf
    sqlStr = sqlStr + " left join [db_aLogistics].[dbo].[tbl_agv_sendState_ItemInfo] as i with(noLock)" & vbcrlf
    sqlStr = sqlStr + " 	on d.itemgubun = i.itemgubun" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemid = i.itemid" & vbcrlf
    sqlStr = sqlStr + " 	and d.itemoption = i.itemoption" & vbcrlf
    sqlStr = sqlStr + " where d.deldt is NULL" & vbcrlf
    sqlStr = sqlStr + " and d.masteridx = "& masteridx &"" & vbcrlf
    sqlStr = sqlStr + " and d.itemgubun<>'10'" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr
end function

%>
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
