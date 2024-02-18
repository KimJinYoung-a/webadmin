<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : �̻� ����
'           2020.05.12 ������ ����
'           2020.05.20 �ѿ�� ����
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
    '/////////// ������ mode��(agvregarr) �� �����ϰ����(imgstatic/linkweb/item/agv/upload_AGV_item_excel.asp) �� ������(/V2/onLine/logics_agv_pickup_process.asp mode�� agvregcs)�� ���� �����ϼž� �մϴ�./////////
    case "write"
        itemgubunarr = split(itemgubunarr, "|")
        itemidarr = split(itemidarr, "|")
        itemoptionarr = split(itemoptionarr, "|")
        itemnoarr = split(itemnoarr, "|")

        '// �ű�����
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

        ' ��ǰ����������Ʈ
        call iteminforeg(masteridx)

	    response.write "<script>alert('����Ǿ����ϴ�.');</script>"
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

	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
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

	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>location.replace('" + refer + "')</script>"
    case "delmaster"
        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set deldt = getdate() "
        sqlStr = sqlStr + " where idx = " & masteridx
        dbget_Logistics.Execute sqlStr

	    response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
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

        ' ��ǰ����������Ʈ
        call iteminforeg(masteridx)

	    ''response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
	    response.write "<script>location.replace('" + refer + "')</script>"

    '/////////// ������ mode��(write) �� �����ϰ����(imgstatic/linkweb/item/agv/upload_AGV_item_excel.asp) �� ������(/V2/onLine/logics_agv_pickup_process.asp mode�� agvregcs)�� ���� �����ϼž� �մϴ�./////////
    ' AGV�������̽� �� ���� ó��
    case "agvregarr"
        ' �ֹ�������
        if ucase(refergubun)="A" then
            itemgubunarr = request("itemgubunarr")
            itemidarr = request("itemidarr")
            itemoptionarr = request("itemoptionarr")
            agvitemnoarr = request("agvitemnoarr")

            if isnull(itemgubunarr) or replace(itemgubunarr,"|","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('���õ� ��ǰ�� �����ϴ�.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            itemgubunarr = split(itemgubunarr, "|")
            itemidarr = split(itemidarr, "|")
            itemoptionarr = split(itemoptionarr, "|")
            agvitemnoarr = split(agvitemnoarr, "|")

            refergubunname="�ֹ���"

        ' �����Ʈ
        elseif ucase(refergubun)="B" then
            chkarr= request("chk") + ",,"
            itemgubunarr= request("itemgubun") + ",,"
            itemidarr= request("itemid") + ",,"
            itemoptionarr= request("itemoption") + ",,"
            agvitemnoarr= request("itemno") + ",,"

            if isnull(chkarr) or replace(chkarr,",","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('���õ� ��ǰ�� �����ϴ�.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

	        chkarr = split(chkarr, ",")
            itemgubunarr = split(itemgubunarr, ",")
            itemidarr = split(itemidarr, ",")
            itemoptionarr = split(itemoptionarr, ",")
            agvitemnoarr = split(agvitemnoarr, ",")

            refergubunname="���"

        ' ������ǰ
        elseif ucase(refergubun)="C" then
            chkarr  = request("chk")+ ","

            if isnull(chkarr) or replace(chkarr,",","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('���õ� ��ǰ�� �����ϴ�.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            chkarr = Split(chkarr, ",")

            refergubunname="������ǰ"

        ' �귣�庰 ���
        elseif ucase(refergubun)="BRANDSTOCK" then
            itemgubunarr = request("itemgubunarr")
            itemidarr = request("itemidarr")
            itemoptionarr = request("itemoptionarr")
            agvitemnoarr = request("itemnoarr")

            if isnull(itemgubunarr) or replace(itemgubunarr,"|","")="" then
                response.write "<script type='text/javascript'>"
                response.write "	alert('���õ� ��ǰ�� �����ϴ�.');"
                response.write "	location.replace('"& refer &"');"
                response.write "</script>"
            end if

            itemgubunarr = split(itemgubunarr, ",")
            itemidarr = split(itemidarr, ",")
            itemoptionarr = split(itemoptionarr, ",")
            agvitemnoarr = split(agvitemnoarr, ",")

            refergubunname="�귣�庰 ���"

        ' �����Ʈ
        else
            response.write "<script type='text/javascript'>"
            response.write "	alert('����� �����ڵ尡 �����ϴ�.');"
            response.write "	location.replace('"& refer &"');"
            response.write "</script>"

            response.end
        end if

        '// ������ �ű�����
        sqlStr = " select * from [db_aLogistics].[dbo].[tbl_agv_pickup_master] where 1=0"
        rsget_Logistics.Open sqlStr,dbget_Logistics,1,3
        rsget_Logistics.AddNew
        rsget_Logistics("reguserid") = finishid
        rsget_Logistics("title") = refergubunname & " " & code
        rsget_Logistics("comment") = NULL
        rsget_Logistics("stationCd") = NULL		' �����̼�
        rsget_Logistics("status") = 0

        ''idx, reguserid, title, comment, status, stationCd, regdate
        ''idx, masteridx, makerid, itemgubun, itemid, itemoption, skuCd, itemname, itemoptionname, itemno, pickupno, regdate, updt, deldt

        rsget_Logistics.update
            masteridx = rsget_Logistics("idx")
        rsget_Logistics.close

        if ucase(refergubun)="A" then
            '�ֹ�������
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - ��ü��ǰ(" & mastercode & ")"
        elseif ucase(refergubun)="B" then
            '�����Ʈ
            requestNo = "ETC - ��Ÿ���(" & mastercode & ")"
        elseif ucase(refergubun)="C" then
            '������ǰ
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - �̹�"
        elseif ucase(refergubun)="BRANDSTOCK" then
            '�귣�庰 ���
            requestNo = "ETC - " & Left(Now(), 10) & " - " & masteridx & " - ���"
        else
            '
        end if

        sqlStr = " update [db_aLogistics].[dbo].[tbl_agv_pickup_master] "
        sqlStr = sqlStr + " set updt = getdate(), requestNo = '" & requestNo & "' "
        sqlStr = sqlStr + " where idx = " & masteridx
        dbget_Logistics.Execute sqlStr

        ' �ֹ�������
        if ucase(refergubun)="A" then
            for i=0 to UBound(itemgubunarr)
                if (trim(itemgubunarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(i))
                    itemid = trim(itemidarr(i))
                    itemoption = trim(itemoptionarr(i))
                    agvitemno = CLng(trim(agvitemnoarr(i)))*-1

                    ' ������ �ű�����
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' �����Ʈ
        elseif ucase(refergubun)="B" then
            for i=0 to UBound(chkarr) - 1
                if (trim(chkarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(chkarr(i)))
                    itemid = trim(itemidarr(chkarr(i)))
                    itemoption = trim(itemoptionarr(chkarr(i)))
                    agvitemno = trim(agvitemnoarr(CInt(chkarr(i))))*-1

                    ' ������ �ű�����
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' ������ǰ
        elseif ucase(refergubun)="C" then
            for i = 0 to UBound(chkarr)
                if (trim(chkarr(i)) <> "") then
                    itemgubun = request("itemgubun" + CStr(trim(chkarr(i))))
                    itemid = request("itemid" + CStr(trim(chkarr(i))))
                    itemoption = request("itemoption" + CStr(trim(chkarr(i))))
                    agvitemno = request("agvitemno" + CStr(trim(chkarr(i))))

                    ' ������ �ű�����
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next
        '�귣�庰 ���
        elseif ucase(refergubun)="BRANDSTOCK" then
            for i=0 to UBound(itemgubunarr)
                if (trim(itemgubunarr(i)) <> "") then
                    itemgubun = trim(itemgubunarr(i))
                    itemid = trim(itemidarr(i))
                    itemoption = trim(itemoptionarr(i))
                    agvitemno = CLng(trim(agvitemnoarr(i)))

                    ' ������ �ű�����
                    sqlStr = " insert into [db_aLogistics].[dbo].[tbl_agv_pickup_detail] " + VBCrlf
                    sqlStr = sqlStr + " (masteridx, itemgubun, itemid, itemoption, itemno) " + VBCrlf
                    sqlStr = sqlStr + " values('" & masteridx & "', '" & itemgubun & "'," & itemid & ", '" & itemoption & "', " & CStr(agvitemno) & ") " & VBCrlf

                    'response.write sqlStr & "<br>"
                    dbget_Logistics.execute sqlStr
                end if
            next

        ' �����Ʈ
        end if

        ' ��ǰ����������Ʈ
        call iteminforeg(masteridx)

        response.write "<script type='text/javascript'>"
        response.write "	alert('���� �Ǿ����ϴ�.');"
        ' ���� ��ΰ� �˾����� �Ѿ�� ���̽��� �θ�â�� �������̽�����Ʈ ������ �̹Ƿ� �θ�â ���ε�
        response.write "	if((typeof(opener) != 'undefined') && (typeof(opener.location) != 'undefined')) {"
        response.write "	    alert(opener);"
        response.write "	    opener.location.reload();"
        response.write "	    self.close();"
        ' �����ΰ� �˾��� �ƴѰ�쿡�� �������̽�����Ʈ �������� �˾����� ����ش�.
        response.write "	}else{"
        response.write "	    var popwin = window.open('/admin/logics/logics_agv_pickupList.asp','addreg','width=1280,height=960,scrollbars=yes,resizable=yes');"
        response.write "	    popwin.focus();"
        response.write "	}"
        response.write "	location.replace('"& refer &"');"
        response.write "</script>"
    case else
        response.write "�߸��� �����Դϴ�."
end select

function iteminforeg(masteridx)
    if masteridx="" or isnull(masteridx) then exit function

    ' �¶��λ�ǰ������Ʈ
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

    ' �������λ�ǰ������Ʈ
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
