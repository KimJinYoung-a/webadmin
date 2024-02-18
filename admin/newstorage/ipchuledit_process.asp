<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ �߰� �˾�
' History : 		   �ʱ������ ��
'			2016.03.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<%
Function GetTextFromUrl(url)
  Dim oXMLHTTP
  Dim strStatusTest

  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

  oXMLHTTP.Open "GET", url, False
  oXMLHTTP.Send

  If oXMLHTTP.Status = 200 Then
    GetTextFromUrl = oXMLHTTP.responseText
  End If
End Function

function fnParseJsonResult(jsonResult, ByRef resultCode, ByRef resultMessage)
    dim resultJson, resultData, totalQty, item

    totalQty = 0

    Set resultJson = New aspJson
    resultJson.loadJSON(jsonResult)

    resultCode = resultJson.data("resultCode")
    resultMessage = resultJson.data("resultMessage")
end function

dim tmp, retMsg, resultCode, resultMessage
dim mode,masterid,scheduledt,executedt,baljuid,targetid,targetname,reguser,divcode,comment,regname,baljuname, rackipgoyn
dim itemgubunarr, itemarr, itemoptionarr, sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr, mwdivarr
dim itemgubun, item, itemoption, sellcash, suplycash, buycash, itemno, designer, mwdiv
dim itemname, itemoptionname, didx, uniqregdate, errMSG, i,cnt,sqlStr, chk, result, existsmakerid
dim iid, ipgocode, code, STOCKBASEDATE, yyyymmdd, tmpitemid, tmpitemgubun, tmpitemoption, tmpitemno, isdeleted, AssignedRows
dim row, rows, dataValid
dim HTTP_Object, SiteURL, regbad, itemexists, baljucode, baljucodeArr
'dim itemnamearr,itemoptionnamearr
	mode     = request("mode")
	masterid = request("masterid")
	scheduledt = request("scheduledt")
	executedt = request("executedt")
	baljuid = request("baljuid")
	targetid = request("targetid")
	reguser = request("reguser")
	divcode = request("divcode")
	comment = html2db(request("comment"))
	regname = html2db(request("regname"))
	baljuname = html2db(request("baljuname"))
	rackipgoyn = request("rackipgoyn")
	uniqregdate = request("uniqregdate")
	code = request("code")

	itemgubunarr = request("itemgubunarr")
	itemarr = request("itemarr")
	itemoptionarr = request("itemoptionarr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	buycasharr = request("buycasharr")
	itemnoarr = request("itemnoarr")
	designerarr = request("designerarr")
	mwdivarr= request("mwdivarr")
	'itemnamearr= request("itemnamearr")
	'itemoptionnamearr= request("itemoptionnamearr")

itemgubunarr = split(itemgubunarr, "|")
itemarr = split(itemarr, "|")
itemoptionarr = split(itemoptionarr, "|")
sellcasharr = split(sellcasharr, "|")
suplycasharr = split(suplycasharr, "|")
buycasharr = split(buycasharr, "|")
itemnoarr = split(itemnoarr, "|")
designerarr = split(designerarr, "|")
mwdivarr = split(mwdivarr, "|")

'itemnamearr = split(itemnamearr, "|")
'itemoptionnamearr = split(itemoptionnamearr, "|")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

isdeleted = false

'response.write mode

if mode="delmaster" then
	''�԰��� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.code, m.executedt, m.deldt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + masterid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master" + VbCrlf
	sqlStr = sqlStr + " set deldt=getdate()" + VbCrlf
	sqlStr = sqlStr + " where id=" + CStr(masterid)
	rsget.Open sqlStr, dbget, 1

	if (not isdeleted) then
		''QuickUpdateNewIpgoDetailSummary code, true
		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
		dbget.Execute sqlStr, AssignedRows

		if (AssignedRows>0) then
		    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
		end if
	end if

	response.write "<script>alert('�����Ǿ����ϴ�.')</script>"
	response.write "<script>location.replace('/admin/newstorage/ipgolist.asp?menupos=539')</script>"
	dbget.close()	:	response.End

elseif mode="addipgo" then
	'1.�¶��� �԰� ����Ÿ

	'// ========================================================================
	if (uniqregdate <> "") then
		'// ����� ���̵� + �ð��� ������ �ߺ��Է� üũ
		sqlStr = "select top 1 id from [db_storage].[dbo].tbl_acount_storage_master "
		sqlStr = sqlStr + " where indt = '" + CStr(uniqregdate) + "' and chargeid = '" + CStr(reguser) + "' "

		errMSG = ""
		rsget.Open sqlStr, dbget, 1
		if Not rsget.Eof then
			errMSG = "�̹� �԰����� ����Ǿ����ϴ�.(�ߺ��Է�)"
		end if
		rsget.close

		if (errMSG <> "") then
			response.write "<script>alert('" + CStr(errMSG) + "');</script>"
			response.write errMSG
			dbget.close()	:	response.End
		end if
	end if

	'��ü�� �˻�
	sqlStr = " select top 1 socname_kor, userid from [db_user].[dbo].tbl_user_c"
	sqlStr = sqlStr + " where userid='" + trim(targetid) + "'"

	'response.write sqlStr & "<br>"
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		targetname = rsget("socname_kor")
		existsmakerid = trim(rsget("userid"))
	end if
	rsget.close

	' ����ó ���̵� üũ��. ��ü���� ��ҹ��� �����ؼ� ��Ȯ�ϰ� �Է�
	if existsmakerid = "" then
		response.write "<script type='text/javascript'>alert('�ش�Ǵ� ����ó�� �����ϴ�.');</script>"
		dbget.close()	:	response.End
	end if
	targetid = existsmakerid

	'// ========================================================================
	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = targetid
	rsget("socname") = targetname
	rsget("chargeid") = reguser
	rsget("chargename") = regname
	rsget("divcode") = divcode ''001-����, 002-��Ź
	rsget("vatcode") = "008"   ''�ΰ���.(�̰͸� �޴´�.)
	rsget("comment") = comment
	rsget("scheduledt") = scheduledt
	rsget("executedt") = executedt
	rsget("ipchulflag") = "I"

	rsget.update
	iid = rsget("id")
	rsget.close

	ipgocode = "ST" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + ipgocode + "'" + VBCrlf

	if (uniqregdate <> "") then
		sqlStr = sqlStr + " ,indt='" + CStr(uniqregdate) + "' " + VBCrlf
	else
		sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	end if

	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	'''2.�¶��� �԰� ������ �Է�
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			item = trim(itemarr(i))
			itemoption = trim(itemoptionarr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))
			itemno = trim(itemnoarr(i))
			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = ""
			itemoptionname = ""

            if (mwdiv = "U") then
                '// �����ΰ��, ����ø��Ա����� ������. 2021-12-30, skyer9
                if (divcode = "001") then
                    '// ����
                    mwdiv = "M"
                elseif (divcode = "002") then
                    '// ��Ź
                    mwdiv = "W"
                end if
            end if

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + ipgocode + "'," + item + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1

		end if
	next


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

    ''�������� ��ǰ��, �ɼ�
    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=T.shopitemname" + vbCrlf
	sqlStr = sqlStr + " ,iitemoptionname=IsNULL(T.shopitemoptionname,'')" + vbCrlf
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item T " + vbCrlf
	sqlStr = sqlStr + " where mastercode='" + CStr(ipgocode) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=T.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=T.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=T.itemoption"
	dbget.Execute sqlStr

    ''// �����ǰ ���Ա��� ����
    '// 1. ���͸��Ա��� �����̸� ����
    '// 2. ������� ���Ա��� �����̸� ����
    '// 3. �ֹ��� ���� ���Ա��� �����̸� ����
	sqlStr = " update d "
	sqlStr = sqlStr + " set d.mwgubun = (case "
	sqlStr = sqlStr + " 		when si.shopitemid is not NULL and si.centermwdiv = 'M' then 'M' "
	sqlStr = sqlStr + " 		when a.itemid is not NULL and a.lastmwdiv = 'M' then 'M' "
	sqlStr = sqlStr + " 		when m.divcode in ('001', '801') then 'M' "
	sqlStr = sqlStr + " 		else d.mwgubun end) "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_acount_storage_detail d "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		m.code = d.mastercode "
	sqlStr = sqlStr + " 	join [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.iitemgubun = '10' "
	sqlStr = sqlStr + " 		and d.itemid = i.itemid "
	sqlStr = sqlStr + " 	left join [db_shop].[dbo].[tbl_shop_item] si "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.iitemgubun = si.itemgubun "
	sqlStr = sqlStr + " 		and d.itemid = si.shopitemid "
	sqlStr = sqlStr + " 		and d.itemoption = si.itemoption "
	sqlStr = sqlStr + " 	left join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] a "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and d.iitemgubun = a.itemgubun "
	sqlStr = sqlStr + " 		and d.itemid = a.itemid "
	sqlStr = sqlStr + " 		and d.itemoption = a.itemoption "
	sqlStr = sqlStr + " 		and a.yyyymm = convert(varchar(7), m.executedt, 121) "
	sqlStr = sqlStr + " 	 "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and d.mastercode = '" + CStr(ipgocode) + "' "
	sqlStr = sqlStr + " 	and i.mwdiv = 'U' "
    rsget.Open sqlStr,dbget,1

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(ipgocode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1


	'' ��� ���Ӹ� ������Ʈ
	'''UpdateIpchulgoSummaryByCode(ipgocode)
	sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & ipgocode & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

    '// AGV�� ��ǰ���� ����
    Set HTTP_Object = Server.CreateObject("MSXML2.ServerXMLHTTP")

    IF application("Svr_Info")="Dev" THEN
        SiteURL = "http://testwapi.10x10.co.kr/agv/api.asp?mode=senditeminfo&ordertype=ipgo&baljucode=" & ipgocode
    else
        SiteURL = "http://wapi.10x10.co.kr/agv/api.asp?mode=senditeminfo&ordertype=ipgo&baljucode=" & ipgocode
    end if

    With HTTP_Object
        .SetTimeouts 30000, 30000, 30000, 30000
        .Open "POST", SiteURL, False
        .SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        .Send ""
        .WaitForResponse 60
    End With

    Set HTTP_Object = Nothing

	refer = refer + "&idx=" + CStr(iid)
	response.write "<script language='javascript'>"
	response.write "alert('���� �Ǿ����ϴ�.');"
	response.write "location.replace('" + refer + "');"
	response.write "</script>"
	dbget.close()	:	response.End
elseif mode="editmaster" then
	iid = request("masterid")

	''�԰��� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.code, m.executedt, m.deldt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		code = rsget("code")
		yyyymmdd = rsget("executedt")
		isdeleted = not IsNULL(rsget("deldt"))

	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)

'	if (yyyymmdd<>CStr(executedt)) and (not isdeleted) then
'	    ''�԰��� ����� - �������� �ݿ�.
'		QuickUpdateNewIpgoDetailSummary code, true
'	end if

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set scheduledt='" + scheduledt + "' " + VBCrlf
	sqlStr = sqlStr + " ,executedt='" + executedt + "' " + VBCrlf
	sqlStr = sqlStr + " ,rackipgoyn='" + rackipgoyn + "' " + VBCrlf
	sqlStr = sqlStr + " ,divcode='" + divcode + "' " + VBCrlf
	sqlStr = sqlStr + " ,comment='" + comment + "' " + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)

	rsget.Open sqlStr,dbget,1

	''���ݿ�
	if (yyyymmdd<>CStr(executedt)) and (code<>"") and (not isdeleted)  then
		''QuickUpdateNewIpgoDetailSummary code, false
		sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'','" & yyyymmdd & "'"
    	dbget.Execute sqlStr, AssignedRows

    	if (AssignedRows>0) then
    	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
    	end if

	end if

elseif mode="editrackipgoyn" then
	iid = request("masterid")

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set rackipgoyn='" + rackipgoyn + "' " + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)

	rsget.Open sqlStr,dbget,1

elseif mode="adddetail" then
	iid = request("masterid")

	''�԰��� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		yyyymmdd = rsget("executedt")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)


	'''2.�¶��� �԰� ������ �߰�
	for i=0 to UBound(itemgubunarr) - 1
		if (trim(itemgubunarr(i)) <> "") then
			itemgubun = trim(itemgubunarr(i))
			item = trim(itemarr(i))
			itemoption = trim(itemoptionarr(i))
			sellcash = trim(sellcasharr(i))
			suplycash = trim(suplycasharr(i))
			buycash = trim(buycasharr(i))
			itemno = trim(itemnoarr(i))
			designer = trim(designerarr(i))
			mwdiv = trim(mwdivarr(i))
			itemname = "" 'trim(itemnamearr(i))
			itemoptionname = "" 'trim(itemoptionnamearr(i))

			sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash, " + VBCrlf
			sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid) " + VBCrlf
			sqlStr = sqlStr + " values('" + request("code") + "'," + item + ", '" + itemoption + "', " + sellcash + ", " + suplycash + ", " + itemno + ", getdate(), getdate(), " + buycash + ", '" + mwdiv + "', '" + itemgubun + "', '" + itemname + "', '" + itemoptionname + "', '" + designer + "') " + VBCrlf
			rsget.Open sqlStr,dbget,1

			''���ݿ�
			''QuickUpdateItemIpgoSummary  yyyymmdd, itemgubun, item, itemoption, itemno,(itemno<0)

		end if
	next


	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_item].[dbo].tbl_item.itemname"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun='10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item.itemid"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemname=[db_shop].[dbo].tbl_shop_item.shopitemname"
	sqlStr = sqlStr + " from [db_shop].[dbo].tbl_shop_item "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun<>'10'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.iitemgubun=[db_shop].[dbo].tbl_shop_item.itemgubun"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_shop].[dbo].tbl_shop_item.shopitemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_shop].[dbo].tbl_shop_item.itemoption"
	rsget.Open sqlStr, dbget, 1

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " set iitemoptionname=IsNULL([db_item].[dbo].tbl_item_option.optionname,'')"
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option "
	sqlStr = sqlStr + " where mastercode='" + CStr(request("code")) + "'"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemid=[db_item].[dbo].tbl_item_option.itemid"
	sqlStr = sqlStr + " and [db_storage].[dbo].tbl_acount_storage_detail.itemoption=[db_item].[dbo].tbl_item_option.itemoption"
	rsget.Open sqlStr, dbget, 1

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

    '' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

	response.write "<script language='javascript'>"
	response.write "location.replace('" + refer + "');"
	response.write "</script>"
	dbget.close()	:	response.End
elseif mode="editdetail" then
	iid = request("masterid")

	chk= request("cksel") + ",,"
	chk = split(chk, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	itemno = request("itemno") + ",,"
	itemno = split(itemno, ",")

	sellcash = request("sellcash") + ",,"
	sellcash = split(sellcash, ",")

	suplycash = request("suplycash") + ",,"
	suplycash = split(suplycash, ",")


	if request("buycash")="" then
		buycash = request("suplycash") + ",,"
	else
		buycash = request("buycash") + ",,"
	end if
	buycash = split(buycash, ",")


	''�԰��� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		yyyymmdd = rsget("executedt")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)


	for i=0 to UBound(chk) - 1
		if (trim(chk(i)) <> "") then
			tmpitemgubun = ""
			tmpitemid = ""
			tmpitemoption = ""
			tmpitemno = 0

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close

			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
			sqlStr = sqlStr + " ,itemno=" + CStr(itemno(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,sellcash=" + CStr(sellcash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,suplycash=" + CStr(suplycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			dbget.Execute(sqlStr)

			''��� �ݿ�
			'if (tmpitemgubun<>"") and (CStr(tmpitemno)<>CStr(itemno(CInt(chk(i))))) then
			'	QuickUpdateItemIpgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, (itemno(CInt(chk(i)))-tmpitemno),(tmpitemno<0)
			'end if
		end if
	next

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	'' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

elseif mode="deldetail" then
	iid = request("masterid")
	chk= request("cksel") + ",,"
	chk = split(chk, ",")

	didx = request("didx") + ",,"
	didx = split(didx, ",")

	''�԰��� - �ֱ� 2�� ������ ���� ��. �԰�¥ ������� ����.
	sqlStr = "select top 1  m.executedt"
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_master m"
	sqlStr = sqlStr + " where m.id=" + iid + ""

	rsget.Open sqlStr,dbget,1
	if not rsget.Eof then
		yyyymmdd = rsget("executedt")
	end if
	rsget.close
	if IsNULL(yyyymmdd) then yyyymmdd=""
	yyyymmdd = Left(CStr(yyyymmdd),10)


	for i=0 to UBound(chk) - 1
		if (trim(chk(i)) <> "") then

			tmpitemgubun = ""
			tmpitemid = ""
			tmpitemoption = ""
			tmpitemno = 0

			sqlStr = " select iitemgubun, itemid, itemoption, itemno " + VBCrlf
			sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))

			rsget.Open sqlStr,dbget,1
			if Not rsget.Eof then
				 tmpitemgubun 	= rsget("iitemgubun")
				 tmpitemid		= rsget("itemid")
				 tmpitemoption	= rsget("itemoption")
				 tmpitemno		= rsget("itemno")
			end if
			rsget.close


			sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			sqlStr = sqlStr + " set deldt=getdate()" + VBCrlf
			sqlStr = sqlStr + " where id=" + CStr(didx(CInt(chk(i))))
			rsget.Open sqlStr,dbget,1

			''��� �ݿ�
			'if (tmpitemgubun<>"") then
			'	QuickUpdateItemIpgoSummary  yyyymmdd, tmpitemgubun, tmpitemid, tmpitemoption, tmpitemno*-1,(tmpitemno<0)
			'end if
		end if
	next

	'''2.�¶��� �԰� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(request("code")) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	rsget.Open sqlStr,dbget,1

	'' ��� �ݿ�
    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & code & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if
elseif mode="agvipgoitemdivisionorder" then

	i = Request.Form("itemgubun").Count
	redim chk(i)
	redim itemgubun(i)
	redim itemid(i)
	redim itemoption(i)
	redim agvitemno(i)

	for i = 0 to Request.Form("itemgubun").Count - 1
		if (Request.Form("cksel").Count >= (i + 1)) then
			chk(i) = Request.Form("cksel")(i + 1)
		else
			chk(i) = ""
		end if

		itemgubun(i) = Request.Form("itemgubun")(i + 1)
		itemid(i) = Request.Form("itemid")(i + 1)
		itemoption(i) = Request.Form("itemoption")(i + 1)
		agvitemno(i) = Request.Form("agvitemno")(i + 1)
	next

	for i=0 to UBound(chk) - 1
		tmp = trim(chk(i))
		if (tmp <> "") then
			sqlStr = "if not exists(select itemid from [db_aLogistics].[dbo].tbl_agv_scheduleditems where itemgubun='"&trim(itemgubun(CInt(tmp)))&"' and itemid="&trim(itemid(CInt(tmp)))&" and itemoption='"&trim(itemoption(CInt(tmp)))&"' and requestMaster='STOCKIN("&code&")' and isusing = 'Y')"
			sqlStr = sqlStr & "	begin"
			sqlStr = sqlStr & "		insert into [db_aLogistics].[dbo].tbl_agv_scheduleditems(itemgubun,itemid,itemoption,realstock,rackCode,requestMaster,displayOrderTypeCD)" & vbCrlf
			sqlStr = sqlStr & "		values('" & trim(itemgubun(CInt(tmp))) & "'," & trim(itemid(CInt(tmp))) & ",'" & trim(itemoption(CInt(tmp))) & "'," & trim(agvitemno(CInt(tmp))) & ",'','STOCKIN(" & code & ")','�԰�����')" & vbCrlf
			sqlStr = sqlStr & "	end"
			rsget_Logistics.Open sqlStr, dbget_Logistics, 1
		end if
	next

	IF application("Svr_Info")="Dev" THEN
		retMsg = GetTextFromUrl("http://testwapi.10x10.co.kr/agv/api.asp?mode=agvipgo&requestMaster=STOCKIN(" & code & ")")
	else
		retMsg = GetTextFromUrl("http://wapi.10x10.co.kr/agv/api.asp?mode=agvipgo&requestMaster=STOCKIN(" & code & ")")

		'retMsg = "!!" & Trim(retMsg) & "!!"
		'response.end
		Call fnParseJsonResult(retMsg, resultCode, resultMessage)
		if (resultCode <> "200") then
			retMsg = resultMessage
		else
			retMsg = ""
		end if
	end if

	response.write "<script language='javascript'>"
	response.write "alert('"&retMsg&"');"
	response.write "location.replace('"&refer&"');"
	response.write "</script>"
	dbget_Logistics.close()	:	response.End
elseif mode="agvipgoitemdivisionorderdelete" then

	IF application("Svr_Info")="Dev" THEN
		retMsg = GetTextFromUrl("http://testwapi.10x10.co.kr/agv/api.asp?mode=agvipgodel&requestMaster=STOCKIN(" & code & ")")
		resultCode="200"
	else
		retMsg = GetTextFromUrl("http://wapi.10x10.co.kr/agv/api.asp?mode=agvipgodel&requestMaster=STOCKIN(" & code & ")")

		''retMsg = "!!" & Trim(retMsg) & "!!"
		Call fnParseJsonResult(retMsg, resultCode, resultMessage)
		if (resultCode <> "200") then
			retMsg = resultMessage
		else
			retMsg = ""
		end if
	end if

	if resultCode="200" then
		sqlStr = "update [db_aLogistics].[dbo].tbl_agv_scheduleditems" & vbCrlf
		sqlStr = sqlStr & "	set isusing='N'" & vbCrlf
		sqlStr = sqlStr & "	where requestMaster='STOCKIN(" & code & ")'"
		sqlStr = sqlStr & "	and isusing='Y'"
		rsget_Logistics.Open sqlStr, dbget_Logistics, 1
	end if

	response.write "<script language='javascript'>"
	response.write "alert('"&retMsg&"');"
	response.write "location.replace('"&refer&"');"
	response.write "</script>"
	dbget_Logistics.close()	:	response.End

elseif mode="regchulgoreturn" then
    regbad = request("regbad")
    baljuid = request("socid")

    itemexists = False

    sqlStr = " select top 1 code "
    sqlStr = sqlStr & "	from "
    sqlStr = sqlStr & "	[db_storage].[dbo].tbl_acount_storage_master "
    sqlStr = sqlStr & "	where socid = '" & baljuid & "' and checkusersn = '" & code & "' and deldt is NULL "
    ''response.write sqlStr

	rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
    	code = rsget("code")
		itemexists = True
    end if
	rsget.Close

	if itemexists then
		response.write "<script>alert('��ǰ������ �����մϴ�.(" & code & ")');</script>"
		response.write "��ǰ������ �����մϴ�.(" & code & ")"
		dbget.close()	:	response.End
	end if

    comment = code & " ����ǰ"
    if (regbad = "Y") then
        comment = comment & "(�ҷ���ϿϷ�)"
    end if

	'1.�¶��� ��� ����Ÿ
	sqlStr = " select * from [db_storage].[dbo].tbl_acount_storage_master where 1=0"
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1,3
	rsget.AddNew
	rsget("code") = ""
	rsget("socid") = baljuid
	rsget("chargeid") = session("ssBctId")
    rsget("chargename") = session("ssBctCname")
	rsget("divcode") = ""
	rsget("vatcode") = ""
	rsget("comment") = comment
	rsget("ipchulflag") = "S"

	rsget.update
		iid = rsget("id")
	rsget.close

	baljucode = "SO" + Format00(6,Right(CStr(iid),6))

	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set code='" + baljucode + "'" + VBCrlf
	sqlStr = sqlStr + " where id=" + CStr(iid)
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	sqlStr = " update m1 "
	sqlStr = sqlStr + " set "
	sqlStr = sqlStr + " 	m1.divcode = m2.divcode, m1.socname = m2.socname, m1.vatcode = m2.vatcode, m1.checkusersn = '" & code & "', m1.statecd = m2.statecd, m1.finishid = '" & session("ssBctId") & "', m1.finishname = '" & session("ssBctCname") & "' "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m1 "
	sqlStr = sqlStr + " 	join [db_storage].[dbo].tbl_acount_storage_master m2 "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		1 = 1 "
	sqlStr = sqlStr + " 		and m1.code = '" & baljucode & "' "
	sqlStr = sqlStr + " 		and m2.code = '" & code & "' "
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	'''2.�¶��� ��� ������ �Է�
	sqlStr = " insert into [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
	sqlStr = sqlStr + " (mastercode,itemid,itemoption,sellcash,suplycash,"
	sqlStr = sqlStr + " itemno,indt,updt,buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid)"
	sqlStr = sqlStr + " select '" + baljucode + "', itemid,itemoption,sellcash,suplycash, "
	sqlStr = sqlStr + " itemno*-1,getdate(),getdate(),buycash,mwgubun,iitemgubun,iitemname,iitemoptionname,imakerid "
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail d "
	sqlStr = sqlStr + " where d.mastercode= '" & code & "' "
	sqlStr = sqlStr + " and deldt is null"
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

	'''2.�¶��� ��� ����Ÿ ������Ʈ
	sqlStr = " update [db_storage].[dbo].tbl_acount_storage_master " + VBCrlf
	sqlStr = sqlStr + " set executedt='" + executedt + "'" + VBCrlf
	sqlStr = sqlStr + " ,scheduledt='" + executedt + "'" + VBCrlf
	sqlStr = sqlStr + " ,totalsellcash=T.totsell" + VBCrlf
	sqlStr = sqlStr + " ,totalsuplycash=T.totsupp" + VBCrlf
	sqlStr = sqlStr + " ,totalbuycash=T.totbuy" + VBCrlf
	sqlStr = sqlStr + " ,indt=getdate()" + VBCrlf
	sqlStr = sqlStr + " ,updt=getdate()" + VBCrlf
	sqlStr = sqlStr + " from (select sum(sellcash*itemno) as totsell, " + vbCrlf
	sqlStr = sqlStr + " sum(suplycash*itemno) as totsupp, " + vbCrlf
	sqlStr = sqlStr + " sum(buycash*itemno) as totbuy " + vbCrlf
	sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail" + vbCrlf
	sqlStr = sqlStr + " where mastercode='"  + CStr(baljucode) + "'" + vbCrlf
	sqlStr = sqlStr + " and deldt is null" + vbCrlf
	sqlStr = sqlStr + " ) as T"
	sqlStr = sqlStr + " where id=" + CStr(iid)
	'response.write sqlStr &"<Br>"
	rsget.Open sqlStr,dbget,1

    sqlStr = "exec db_summary.dbo.sp_ten_recentIpChul_Update '" & baljucode & "','','',0,'',''"
	dbget.Execute sqlStr, AssignedRows

	'// ������� �ݿ�
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_Shop_Stock_RecentLogicsIpChul_Update] '" & baljuid & "', '" & baljucode & "', 'N' "
	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	if (AssignedRows>0) then
	    response.write "<script>alert('����� " & AssignedRows & "�� �ݿ��Ǿ����ϴ�.')</script>"
	end if

    if (regbad = "Y") then
        '// �ҷ����

	    ''���Էµ� ������ �߰��� (���̺� ��ǰ�� ������ ���)
	    sqlStr = " update [db_summary].[dbo].tbl_erritem_daily_summary"
	    sqlStr = sqlStr + " set errbaditemno=errbaditemno + IsNULL(T.itemno,0)*-1"
	    sqlStr = sqlStr + " from ( "
	    sqlStr = sqlStr + " select b.iitemgubun as itemgubun, b.itemid, b.itemoption, b.itemno"
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b, [db_summary].[dbo].tbl_erritem_daily_summary s"
	    sqlStr = sqlStr + " where s.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	    sqlStr = sqlStr + " and b.itemid=s.itemid"
	    sqlStr = sqlStr + " and b.itemoption=s.itemoption"
        sqlStr = sqlStr + " and b.mastercode = '" & baljucode & "' "
	    sqlStr = sqlStr + " ) T"
	    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=T.itemgubun"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=T.itemid"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=T.itemoption"

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    ''���Էµ� ������ �߰��� (���̺� ��ǰ�� ���� ���)
	    sqlStr = " insert into [db_summary].[dbo].tbl_erritem_daily_summary"
	    sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errbaditemno,reguser)"
	    sqlStr = sqlStr + " select "
	    sqlStr = sqlStr + " '" + executedt + "'"
	    sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1,'" + session("ssBctId") + "'"
	    sqlStr = sqlStr + " from ("
	    sqlStr = sqlStr + " select b.*, b.iitemgubun as itemgubun "
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	    sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_erritem_daily_summary s on s.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	    sqlStr = sqlStr + " and b.itemid=s.itemid"
	    sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	    sqlStr = sqlStr + " where s.itemid is null and b.mastercode = '" & baljucode & "' "
	    sqlStr = sqlStr + " ) T"

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    sqlStr = "update [db_summary].[dbo].tbl_erritem_daily_summary"
	    sqlStr = sqlStr + " set toterrno=errbaditemno+erretcno+errrealcheckno" ''errcsno�� �������̺��� �Ϲ� ������ ���亯��  errcsno+
	    sqlStr = sqlStr + " ,lastupdate=getdate()"
	    sqlStr = sqlStr + " ,modiuser='" + session("ssBctId") + "'"
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_erritem_daily_summary.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemgubun=b.iitemgubun"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemid=b.itemid"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_erritem_daily_summary.itemoption=b.itemoption"
        sqlStr = sqlStr + " and b.mastercode = '" & baljucode & "' "

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    ''�Ϻ� ���α׿� �߰�
	    sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	    sqlStr = sqlStr + " set errbaditemno=errbaditemno + b.itemno*-1"
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
        sqlStr = sqlStr + " and b.mastercode = '" & baljucode & "' "

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    sqlStr = " insert into [db_summary].[dbo].tbl_daily_logisstock_summary"
	    sqlStr = sqlStr + " (yyyymmdd,itemgubun,itemid,itemoption,errbaditemno)"
	    sqlStr = sqlStr + " select "
	    sqlStr = sqlStr + " '" + executedt + "'"
	    sqlStr = sqlStr + " ,T.itemgubun,T.itemid,T.itemoption,T.itemno*-1"
	    sqlStr = sqlStr + " from ("
	    sqlStr = sqlStr + " select b.*, b.iitemgubun as itemgubun "
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	    sqlStr = sqlStr + " left join [db_summary].[dbo].tbl_daily_logisstock_summary s on s.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and b.iitemgubun=s.itemgubun"
	    sqlStr = sqlStr + " and b.itemid=s.itemid"
	    sqlStr = sqlStr + " and b.itemoption=s.itemoption"
	    sqlStr = sqlStr + " where s.itemid is null and b.mastercode = '" & baljucode & "' "
	    sqlStr = sqlStr + " ) T"

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    ''���Ӹ�.
	    sqlStr = " update [db_summary].[dbo].tbl_daily_logisstock_summary"
	    sqlStr = sqlStr + " set toterrno=errbaditemno+erretcno+errrealcheckno"  ''errcsno+
	    sqlStr = sqlStr + " ,totsysstock=totipgono+totchulgono-totsellno+errcsno"
	    sqlStr = sqlStr + " ,availsysstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno"
	    sqlStr = sqlStr + " ,realstock=totipgono+totchulgono-totsellno+errcsno+errbaditemno+erretcno+errrealcheckno"
	    sqlStr = sqlStr + " ,lastupdate=getdate()"
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail b "
	    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_daily_logisstock_summary.yyyymmdd='" + executedt + "'"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemgubun=b.iitemgubun"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemid=b.itemid"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_daily_logisstock_summary.itemoption=b.itemoption"
        sqlStr = sqlStr + " and b.mastercode = '" & baljucode & "' "

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1

	    ''��������̺� ������Ʈ
	    sqlStr = "update [db_summary].[dbo].tbl_current_logisstock_summary"
	    sqlStr = sqlStr + " set errbaditemno=errbaditemno+IsNULL(T.itemno*-1,0)"
	    sqlStr = sqlStr + " ,toterrno=toterrno+IsNULL(T.itemno*-1,0)"  ''errcsno+
	    sqlStr = sqlStr + " ,availsysstock=availsysstock+IsNULL(T.itemno*-1,0)"
	    sqlStr = sqlStr + " ,realstock=realstock+IsNULL(T.itemno*-1,0)"
	    sqlStr = sqlStr + " ,shortageno = shortageno+IsNULL(T.itemno*-1,0)"
	    sqlStr = sqlStr + " ,lastupdate=getdate()"
	    sqlStr = sqlStr + " from [db_storage].[dbo].tbl_acount_storage_detail T "
	    sqlStr = sqlStr + " where [db_summary].[dbo].tbl_current_logisstock_summary.itemgubun=T.iitemgubun"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemid=T.itemid"
	    sqlStr = sqlStr + " and [db_summary].[dbo].tbl_current_logisstock_summary.itemoption=T.itemoption"
        sqlStr = sqlStr + " and T.mastercode = '" & baljucode & "' "

        ''response.write sqlStr & "<br />"
	    rsget.Open sqlStr,dbget,1
    end if
elseif mode="modichulgoprc" then
    '// ��� �ϰ�����
    rows = request("modlst")
    rows = Replace(rows, "'", "")
    rows = Replace(rows, "`", "")
    rows = Replace(rows, ",", "")
    rows = Split(rows, vbCrLf)

    yyyymmdd = request("yyyymmdd")

    baljucodeArr = ",X,"
	tmp = 0

    for i = 0 to UBOund(rows)
        row = rows(i)
        row = Split(row, vbTab)

        if UBOund(row) = 4 then
            baljucode	= row(0)
            itemgubun	= row(1)
            item		= row(2)
            itemoption	= row(3)
            suplycash	= row(4)

            dataValid = False
            if (InStr(baljucodeArr, "," & baljucode & ",") < 1) then
                sqlStr = " select top 1 code from [db_storage].[dbo].tbl_acount_storage_master "
                sqlStr = sqlStr + " where code = '" & baljucode & "' and executedt >= '" & yyyymmdd & "' "
                sqlStr = sqlStr + " 	and ipchulflag in ('S','E') "		'�����
                ''response.write sqlStr & "<br />"
	            rsget.Open sqlStr, dbget, 1
                if Not rsget.Eof then
    	            dataValid = True
                    baljucodeArr = baljucodeArr & baljucode & ","
                else
                    dataValid = False
                end if
	            rsget.Close
            else
                dataValid = True
            end if

            if (dataValid = True) then
			    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			    sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
			    sqlStr = sqlStr + " ,suplycash=" & suplycash & " " + VBCrlf
			    ''sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
			    sqlStr = sqlStr + " where 1=1"
                sqlStr = sqlStr + " and mastercode = '" & baljucode & "' "
                sqlStr = sqlStr + " and iitemgubun = '" & itemgubun & "' "
                sqlStr = sqlStr + " and itemid = '" & item & "' "
                sqlStr = sqlStr + " and itemoption = '" & itemoption & "' "
                ''response.write sqlStr & "<br />"
			    dbget.Execute(sqlStr)
				tmp = tmp+1
            end if
        end if
    next

    if Left(baljucodeArr, 1) = "," then
        baljucodeArr = Mid(baljucodeArr, 2, 1000)
    end if

    if Right(baljucodeArr, 1) = "," then
        baljucodeArr = Mid(baljucodeArr, 1, Len(baljucodeArr) - 1)
    end if

    baljucodeArr= Replace(baljucodeArr, ",", "','")

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.updt = getdate(), m.totalsuplycash = d.totsupp "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select mastercode, sum(suplycash*itemno) as totsupp "
	sqlStr = sqlStr + " 		from [db_storage].[dbo].tbl_acount_storage_detail "
	sqlStr = sqlStr + " 		where mastercode in ('" & baljucodeArr & "') and deldt is NULL "
	sqlStr = sqlStr + " 		group by mastercode "
	sqlStr = sqlStr + " 	) d on m.code = d.mastercode "
    ''response.write sqlStr & "<br />"
    dbget.Execute sqlStr

	Response.Write "<script type=""text/javascript"">"
	Response.Write "alert('���� �Ǿ����ϴ�.\n\n�� �� " & i & "�� �� " & tmp & "�� ó��');"
	Response.Write "opener.location.reload();"
	Response.Write "self.close();"
	Response.Write "</script>"
    dbget.close : response.end
elseif mode="modiStoredPrc" then
    '// �԰� �ϰ�����
    rows = request("modlst")
    rows = Replace(rows, "'", "")
    rows = Replace(rows, "`", "")
    rows = Replace(rows, ",", "")
    rows = Split(rows, vbCrLf)

    yyyymmdd = request("yyyymmdd")

    baljucodeArr = ",X,"
	tmp = 0

    for i = 0 to UBOund(rows)
        row = rows(i)
        row = Split(row, vbTab)

        if UBOund(row) = 4 then
            baljucode	= row(0)
            itemgubun	= row(1)
            item		= row(2)
            itemoption	= row(3)
            suplycash	= row(4)

            dataValid = False
            if (InStr(baljucodeArr, "," & baljucode & ",") < 1) then
                sqlStr = " select top 1 code from [db_storage].[dbo].tbl_acount_storage_master "
                sqlStr = sqlStr + " where code = '" & baljucode & "' and executedt >= '" & yyyymmdd & "' "
				sqlStr = sqlStr + " 	and ipchulflag='I'"		'�԰���
                ''response.write sqlStr & "<br />"
	            rsget.Open sqlStr, dbget, 1
                if Not rsget.Eof then
    	            dataValid = True
                    baljucodeArr = baljucodeArr & baljucode & ","
                else
                    dataValid = False
                end if
	            rsget.Close
            else
                dataValid = True
            end if

            if (dataValid = True) then
			    sqlStr = " update [db_storage].[dbo].tbl_acount_storage_detail " + VBCrlf
			    sqlStr = sqlStr + " set updt=getdate()" + VBCrlf
			    sqlStr = sqlStr + " ,suplycash=" & suplycash & " " + VBCrlf
			    ''sqlStr = sqlStr + " ,buycash=" + CStr(buycash(CInt(chk(i)))) + " " + VBCrlf
			    sqlStr = sqlStr + " where 1=1"
                sqlStr = sqlStr + " and mastercode = '" & baljucode & "' "
                sqlStr = sqlStr + " and iitemgubun = '" & itemgubun & "' "
                sqlStr = sqlStr + " and itemid = '" & item & "' "
                sqlStr = sqlStr + " and itemoption = '" & itemoption & "' "
                ''response.write sqlStr & "<br />"
			    dbget.Execute(sqlStr)
				tmp = tmp+1
            end if
        end if
    next

    if Left(baljucodeArr, 1) = "," then
        baljucodeArr = Mid(baljucodeArr, 2, 1000)
    end if

    if Right(baljucodeArr, 1) = "," then
        baljucodeArr = Mid(baljucodeArr, 1, Len(baljucodeArr) - 1)
    end if

    baljucodeArr= Replace(baljucodeArr, ",", "','")

	sqlStr = " update m "
	sqlStr = sqlStr + " set m.updt = getdate(), m.totalsuplycash = d.totsupp "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	[db_storage].[dbo].tbl_acount_storage_master m "
	sqlStr = sqlStr + " 	join ( "
	sqlStr = sqlStr + " 		select mastercode, sum(suplycash*itemno) as totsupp "
	sqlStr = sqlStr + " 		from [db_storage].[dbo].tbl_acount_storage_detail "
	sqlStr = sqlStr + " 		where mastercode in ('" & baljucodeArr & "') and deldt is NULL "
	sqlStr = sqlStr + " 		group by mastercode "
	sqlStr = sqlStr + " 	) d on m.code = d.mastercode "
    ''response.write sqlStr & "<br />"
    dbget.Execute sqlStr

	Response.Write "<script type=""text/javascript"">"
	Response.Write "alert('���� �Ǿ����ϴ�.\n\n�� �� " & i & "�� �� " & tmp & "�� ó��');"
	Response.Write "opener.location.reload();"
	Response.Write "self.close();"
	Response.Write "</script>"
    dbget.close : response.end
else
	'
end if
%>

<script type="text/javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
