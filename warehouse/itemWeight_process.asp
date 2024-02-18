<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 해외배송 상품 관리
' History : 서동석 생성
'			2016.05.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%
dim itemgubun, itemid, itemoption, itemWeight, isUpcheDeli, overSeaYn, pojangok, volX, volY, volZ, mode, multilangcnt, AssignedRow
dim linkPriceTypeusd, OptCnt, i, check, chdeliverOverseas
	itemgubun	= trim(request("itemgubun"))
	itemid		= trim(request("itemid"))
	itemoption	= trim(request("itemoption"))
	itemWeight	= trim(request("itemWeight"))
	mode		= trim(request("mode"))
	isUpcheDeli = trim(request("isUcDeli"))
	overSeaYn	= trim(request("overSeaYn"))
	pojangok	= trim(request("pojangok"))
	volX		= trim(request("volX"))
	volY		= trim(request("volY"))
	volZ		= trim(request("volZ"))
	check		= trim(request("check"))
	chdeliverOverseas		= requestcheckvar(request("chdeliverOverseas"),1)

linkPriceTypeusd=0
multilangcnt=0

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr

function DoSomethingForForeignSite(itemid)
    dim sqlStr

	'/ 사이트별 화폐단위
	sqlStr = "select" & vbcrlf
	sqlStr = sqlStr & "	e.sitename, e.currencyunit" & vbcrlf
	sqlStr = sqlStr & "	, (select top 1 exchangeRate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as exchangeRate" & vbcrlf
	sqlStr = sqlStr & "	, (select top 1 multiplerate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as multiplerate" & vbcrlf
	sqlStr = sqlStr & "	, (select top 1 linkPriceType from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as linkPriceType" & vbcrlf
	sqlStr = sqlStr & "	into #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
	sqlStr = sqlStr & "	from db_item.dbo.tbl_exchangeRate e" & vbcrlf
	sqlStr = sqlStr & "	where e.sitename='WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "	group by e.sitename, e.currencyunit" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/오프라인 상품
	sqlStr = "select" & vbcrlf
	sqlStr = sqlStr & "	itemgubun, shopitemid" & vbcrlf
	sqlStr = sqlStr & "	, min(orgsellprice) as orgsellprice, min(shopitemprice)	as shopitemprice" & vbcrlf		'/옵션별 추가 금액 제끼고 가져옴
	sqlStr = sqlStr & "	into #tmp_shop_item" & vbcrlf
	sqlStr = sqlStr & "	from db_shop.dbo.tbl_shop_item" & vbcrlf
	sqlStr = sqlStr & "	where itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & "	and isusing='Y'" & vbcrlf
	sqlStr = sqlStr & "	and shopitemid = "& itemid &"" & vbcrlf
	sqlStr = sqlStr & "	group by itemgubun, shopitemid" & vbcrlf
	sqlStr = sqlStr & "	CREATE NONCLUSTERED INDEX [tmp_shopitemid] ON #tmp_shop_item(shopitemid ASC)" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/상품 꼿고
	sqlStr = "insert into [db_item].[dbo].[tbl_item_multiSite_regItem]" & vbcrlf
	sqlStr = sqlStr & "		select" & vbcrlf
	sqlStr = sqlStr & "		i.itemid, 'WSLWEB', 'Y', 0, 0, getdate(), 'SYSTEM', getdate(), 'SYSTEM'" & vbcrlf
	sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
	sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
	sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
	sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "		where ri.itemid is null" & vbcrlf
	sqlStr = sqlStr & "		and i.itemid = "& itemid &"" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/언어팩 꼿고
	sqlStr = "insert into db_item.[dbo].[tbl_item_multiLang]" & vbcrlf
	sqlStr = sqlStr & "		select" & vbcrlf
	sqlStr = sqlStr & "		i.itemid, 'KR', i.itemname, isnull(ic.designercomment,'') as designercomment, '', ic.itemsource, ic.itemsize" & vbcrlf
	sqlStr = sqlStr & "		, ic.sourcearea, c.socname_kor, 'Y', getdate(), getdate(), ic.keywords, ''" & vbcrlf
	sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
	sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
	sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
	sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "		left join db_user.dbo.tbl_user_c c" & vbcrlf
	sqlStr = sqlStr & "			on i.makerid=c.userid" & vbcrlf
	sqlStr = sqlStr & "		left join db_item.[dbo].[tbl_item_multiLang] ml" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid=ml.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ml.countryCd='KR'" & vbcrlf
	sqlStr = sqlStr & "		left join db_item.dbo.tbl_item_Contents ic" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = ic.itemid" & vbcrlf
	sqlStr = sqlStr & "		where i.itemid = "& itemid &"" & vbcrlf
	sqlStr = sqlStr & "		and ml.itemid is null" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/언어팩 카운트 계산해서 꼿음
	sqlStr = "update ri" & vbcrlf
	sqlStr = sqlStr & "		set ri.multilangcnt = isnull(t.multilangcnt,0)" & vbcrlf
	sqlStr = sqlStr & "		, ri.lastupdate = getdate()" & vbcrlf
	sqlStr = sqlStr & "		from db_item.dbo.tbl_item_multiSite_regItem ri" & vbcrlf
	sqlStr = sqlStr & "		left join (" & vbcrlf
	sqlStr = sqlStr & "			select itemid, count(*) as multilangcnt" & vbcrlf
	sqlStr = sqlStr & "			from [db_item].[dbo].[tbl_item_multiLang]" & vbcrlf
	sqlStr = sqlStr & "			where useyn='Y'" & vbcrlf
	sqlStr = sqlStr & "			and itemid = "& itemid &"" & vbcrlf
	sqlStr = sqlStr & "			group by itemid" & vbcrlf
	sqlStr = sqlStr & "		) as t" & vbcrlf
	sqlStr = sqlStr & "			on ri.itemid = t.itemid" & vbcrlf
	sqlStr = sqlStr & "		where ri.sitename= 'WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "		and ri.itemid = "& itemid &"" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/옵션 꼿고
	sqlStr = "insert into db_item.[dbo].[tbl_item_multiLang_option](itemid, countryCd, itemoption, isusing, optionTypeName, optionname, regdate)" & vbcrlf
	sqlStr = sqlStr & "		select" & vbcrlf
	sqlStr = sqlStr & "		ri.itemid, 'KR', o.itemoption, 'Y', o.optionTypeName, o.optionname, getdate()" & vbcrlf
	sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
	sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "		join db_shop.dbo.tbl_shop_item ii" & vbcrlf
	sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
	sqlStr = sqlStr & "		left join db_item.dbo.tbl_item_option o" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = o.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ii.itemoption = o.itemoption" & vbcrlf
	sqlStr = sqlStr & "			and o.isusing='Y'" & vbcrlf
	sqlStr = sqlStr & "		left join db_item.[dbo].[tbl_item_multiLang_option] mo" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid=mo.itemid" & vbcrlf
	sqlStr = sqlStr & "			and o.itemoption = mo.itemoption" & vbcrlf
	sqlStr = sqlStr & "			and mo.countryCd='KR'" & vbcrlf
	sqlStr = sqlStr & "		where ri.itemid = "& itemid &"" & vbcrlf
	sqlStr = sqlStr & "		and mo.itemid is null" & vbcrlf
	sqlStr = sqlStr & "		and o.itemoption is not null" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	'/ 가격 꼿고
	sqlStr = "insert into db_item.dbo.tbl_item_multiLang_price" & vbcrlf
	sqlStr = sqlStr & "		select " & vbcrlf
	sqlStr = sqlStr & "		'WSLWEB' ,i.itemid, ee.currencyUnit" & vbcrlf
	'sqlStr = sqlStr & "	, (case" & vbcrlf
	'sqlStr = sqlStr & "		when ee.currencyUnit='WON' or ee.currencyUnit='KRW' then (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end)" & vbcrlf
	'sqlStr = sqlStr & "		else round((((( (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) *ee.multiplerate)/ee.exchangeRate)*100)/100) ,2)" & vbcrlf
	'sqlStr = sqlStr & "		end) as orgprice" & vbcrlf
	'sqlStr = sqlStr & "	, ((case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) * ee.multiplerate) as wonPrice" & vbcrlf
	sqlStr = sqlStr & "		, (case" & vbcrlf
	sqlStr = sqlStr & "			when ee.currencyUnit='WON' or ee.currencyUnit='KRW' then (case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end)" & vbcrlf
	sqlStr = sqlStr & "			else round((((( (case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end) *ee.multiplerate)/ee.exchangeRate)*100)/100) ,2)" & vbcrlf
	sqlStr = sqlStr & "			end) as orgprice" & vbcrlf
	sqlStr = sqlStr & "		, ((case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end) * ee.multiplerate) as wonPrice" & vbcrlf
	sqlStr = sqlStr & "		,NULL as mayDiscountPrice" & vbcrlf
	sqlStr = sqlStr & "		,ee.multiplerate" & vbcrlf
	sqlStr = sqlStr & "		,getdate()" & vbcrlf
	sqlStr = sqlStr & "		,getdate()" & vbcrlf
	sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
	sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
	sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
	sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
	sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
	sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
	sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
	sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
	sqlStr = sqlStr & "		join #tmp_exchangeRatecurrencyunitgroup ee" & vbcrlf
	sqlStr = sqlStr & "			on ri.sitename = ee.sitename" & vbcrlf
	sqlStr = sqlStr & "		left join db_item.dbo.tbl_item_multiLang_price p" & vbcrlf
	sqlStr = sqlStr & "			on ri.sitename=p.sitename" & vbcrlf
	sqlStr = sqlStr & "			and i.itemid=P.itemid" & vbcrlf
	sqlStr = sqlStr & "			and ee.currencyUnit=P.currencyUnit" & vbcrlf
	sqlStr = sqlStr & "		where ri.itemid = "& itemid &"" & vbcrlf
	sqlStr = sqlStr & "		and P.itemid is NULL" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr

	sqlStr = "drop table #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
	sqlStr = sqlStr & " drop table #tmp_shop_item" & vbcrlf

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr
end function

if (mode="ByWeightProc") then
	if itemgubun = "" then
		Response.Write "무게입력 불가 : 상품구분이 없음"
		Response.end
	end if

    'if (isUpcheDeli="False") then
    	if (itemWeight="") or (itemid="") then
    		response.write "<script>alert('상품코드나 무게가 입력되지 않았습니다.');</script>"
    		response.write "<script>history.back();</script>"
    		dbget.close()	:	response.End
    	end if

        if (itemgubun <> "10") then
            response.write "10코드 상품이 아닙니다." & mode
            dbget.close : response.end
        end if

    	sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
    	sqlStr = sqlStr + " set pojangok='" + Cstr(pojangok) + "'" + VbCrlf
    	sqlStr = sqlStr + " , deliverOverseas='" & overSeaYn & "' " + VbCrlf
		sqlStr = sqlStr + " , lastupdate = getdate()" & VbCrlf
    	sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
    	dbget.execute(sqlStr)

		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
        Call DoSomethingForForeignSite(itemid)
		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
    'else
    '   response.write "<script >alert('텐바이텐 배송 상품에만 무게를 입력할 수 있습니다.');</script>"
    '   response.write "<script >location.replace('" + refer + "');</script>"
    '    dbget.close()	:	response.End
    'end if

elseif (mode = "BySizeProc") then
	if itemgubun = "" then
		Response.Write "무게입력 불가 : 상품구분이 없음"
		Response.end
	end if

    'if (isUpcheDeli="False") then
    	if (volX="") or (volY="") or (volZ="") or (itemid="") then
    		response.write "<script>alert('상품코드나 사이즈가 입력되지 않았습니다.');</script>"
    		response.write "<script>history.back();</script>"
    		dbget.close()	:	response.End
    	end if

        if (itemgubun = "10") then
		    sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_Volumn] WHERE itemid = "& CStr(itemid) & " AND itemoption ='" + Trim(itemoption) + "')" & VbCrlf
		    sqlStr = sqlStr + "	BEGIN" & VbCrlf
		    sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_Volumn]" & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight=" & itemWeight & VbCrlf
			sqlStr = sqlStr + "		,volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,lastupdate=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemid=" & itemid & " AND itemoption=" & itemoption & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			sqlStr = sqlStr + "ELSE" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_Volumn](itemid, itemoption, itemWeight, volX, volY, volZ)" & VbCrlf
			sqlStr = sqlStr + "		VALUES (" & itemid & ",'" & itemoption & "'," & itemWeight & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

			'구형 테이블에도 입력
			sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_pack_Volumn] WHERE itemid = "& CStr(itemid) & ")" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_pack_Volumn]" & VbCrlf
			sqlStr = sqlStr + "		SET volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,lastupdt=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemid=" & itemid & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			sqlStr = sqlStr + "ELSE" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_pack_Volumn](itemid, volX, volY, volZ)" & VbCrlf
			sqlStr = sqlStr + "		VALUES (" & itemid & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

    		sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
	    	sqlStr = sqlStr + " set itemWeight='" + CStr(itemWeight) + "'" + VbCrlf
    		sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
    		dbget.execute(sqlStr)

			'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
        	Call DoSomethingForForeignSite(itemid)
			'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
		else
			sqlStr = sqlStr + "		UPDATE [db_shop].[dbo].[tbl_shop_item] " & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight='" + CStr(itemWeight) + "', volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,updt=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemgubun = '" & itemgubun & "' and shopitemid=" & itemid & " AND itemoption='" & itemoption & "'" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        end if

elseif (mode = "ByOptionSizeProc") then
	if itemgubun = "" then
		Response.Write "무게입력 불가 : 상품구분이 없음"
		Response.end
	end if

	OptCnt = request("oitemWeight").count

	if (itemid="") then
		response.write "<script>alert('상품코드가 입력되지 않았습니다.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

    if (itemgubun = "10") then
	    sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_logics_addinfo] WHERE itemid = "& CStr(itemid) & ")" & VbCrlf
        sqlStr = sqlStr + "	BEGIN" & VbCrlf
        sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_logics_addinfo]" & VbCrlf
        sqlStr = sqlStr + "		SET itemManageType='O'" & VbCrlf
	    sqlStr = sqlStr + "		WHERE itemid=" & itemid & VbCrlf
	    sqlStr = sqlStr + "	END" & VbCrlf
		sqlStr = sqlStr + "ELSE" & VbCrlf
	    sqlStr = sqlStr + "	BEGIN" & VbCrlf
	    sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_logics_addinfo](itemid, itemManageType)" & VbCrlf
	    sqlStr = sqlStr + "		VALUES (" & itemid & ", 'O')" & VbCrlf
	    sqlStr = sqlStr + "	END" & VbCrlf
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    end if

	For i=1 To OptCnt
		itemoption = Trim(request("itemoption")(i))
		itemWeight = Trim(request("oitemWeight")(i))
		if itemWeight="" then itemWeight="0"
		volX = Trim(request("ovolX")(i))
		volY = Trim(request("ovolY")(i))
		volZ = Trim(request("ovolZ")(i))
		if volX="" then volX="0"
		if volY="" then volY="0"
		if volZ="" then volZ="0"

        if (itemgubun = "10") then
			sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_Volumn] WHERE itemid = "& CStr(itemid) & " AND itemoption ='" + Trim(itemoption) + "')" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_Volumn]" & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight=" & itemWeight & VbCrlf
			sqlStr = sqlStr + "		,volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,lastupdate=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemid=" & itemid & " AND itemoption=" & itemoption & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			sqlStr = sqlStr + "ELSE" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_Volumn](itemid, itemoption, itemWeight, volX, volY, volZ)" & VbCrlf
			sqlStr = sqlStr + "		VALUES (" & itemid & ",'" & itemoption & "'," & itemWeight & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			if i=1 then
				'구형 테이블에도 입력
				sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_pack_Volumn] WHERE itemid = "& CStr(itemid) & ")" & VbCrlf
				sqlStr = sqlStr + "	BEGIN" & VbCrlf
				sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_pack_Volumn]" & VbCrlf
				sqlStr = sqlStr + "		SET volX=" & volX & VbCrlf
				sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
				sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
				sqlStr = sqlStr + "		,lastupdt=GETDATE()" & VbCrlf
				sqlStr = sqlStr + "		WHERE itemid=" & itemid & VbCrlf
				sqlStr = sqlStr + "	END" & VbCrlf
				sqlStr = sqlStr + "ELSE" & VbCrlf
				sqlStr = sqlStr + "	BEGIN" & VbCrlf
				sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_pack_Volumn](itemid, volX, volY, volZ)" & VbCrlf
				sqlStr = sqlStr + "		VALUES (" & itemid & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
				sqlStr = sqlStr + "	END" & VbCrlf
				rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			end if
		else
			sqlStr = sqlStr + "		UPDATE [db_shop].[dbo].[tbl_shop_item] " & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight='" + CStr(itemWeight) + "', volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,updt=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemgubun = '" & itemgubun & "' and shopitemid=" & itemid & " AND itemoption='" & itemoption & "'" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
		end if
	Next

    if (itemgubun = "10") then
	    sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
	    sqlStr = sqlStr + " set itemWeight='" + CStr(itemWeight) + "'" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    dbget.execute(sqlStr)

		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
        Call DoSomethingForForeignSite(itemid)
		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
    end if

elseif (mode = "ByOptionSameSizeProc") then
	if itemgubun = "" then
		Response.Write "무게입력 불가 : 상품구분이 없음"
		Response.end
	end if

	OptCnt = request("oitemWeight").count


	if (itemid="") then
		response.write "<script>alert('상품코드가 입력되지 않았습니다.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if

    if (itemgubun = "10") then
	    sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_logics_addinfo] WHERE itemid = "& CStr(itemid) & ")" & VbCrlf
        sqlStr = sqlStr + "	BEGIN" & VbCrlf
        sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_logics_addinfo]" & VbCrlf
        sqlStr = sqlStr + "		SET itemManageType='I'" & VbCrlf
	    sqlStr = sqlStr + "		WHERE itemid=" & itemid & VbCrlf
	    sqlStr = sqlStr + "	END" & VbCrlf
		sqlStr = sqlStr + "ELSE" & VbCrlf
	    sqlStr = sqlStr + "	BEGIN" & VbCrlf
	    sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_logics_addinfo](itemid, itemManageType)" & VbCrlf
	    sqlStr = sqlStr + "		VALUES (" & itemid & ", 'I')" & VbCrlf
	    sqlStr = sqlStr + "	END" & VbCrlf
	    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
    end if

	For i=1 To OptCnt
		itemoption = Trim(request("itemoption")(i))
		itemWeight = Trim(request("itemWeight"))
		volX = Trim(request("volX"))
		volY = Trim(request("volY"))
		volZ = Trim(request("volZ"))

        if (itemgubun = "10") then
			sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_Volumn] WHERE itemid = "& CStr(itemid) & " AND itemoption ='" + Trim(itemoption) + "')" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_Volumn]" & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight=" & itemWeight & VbCrlf
			sqlStr = sqlStr + "		,volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,lastupdate=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemid=" & itemid & " AND itemoption=" & itemoption & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			sqlStr = sqlStr + "ELSE" & VbCrlf
			sqlStr = sqlStr + "	BEGIN" & VbCrlf
			sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_Volumn](itemid, itemoption, itemWeight, volX, volY, volZ)" & VbCrlf
			sqlStr = sqlStr + "		VALUES (" & itemid & ",'" & itemoption & "'," & itemWeight & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
			sqlStr = sqlStr + "	END" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			if i=1 then
				'구형 테이블에도 입력
				sqlStr = "IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_pack_Volumn] WHERE itemid = "& CStr(itemid) & ")" & VbCrlf
				sqlStr = sqlStr + "	BEGIN" & VbCrlf
				sqlStr = sqlStr + "		UPDATE [db_item].[dbo].[tbl_item_pack_Volumn]" & VbCrlf
				sqlStr = sqlStr + "		SET volX=" & volX & VbCrlf
				sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
				sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
				sqlStr = sqlStr + "		,lastupdt=GETDATE()" & VbCrlf
				sqlStr = sqlStr + "		WHERE itemid=" & itemid & VbCrlf
				sqlStr = sqlStr + "	END" & VbCrlf
				sqlStr = sqlStr + "ELSE" & VbCrlf
				sqlStr = sqlStr + "	BEGIN" & VbCrlf
				sqlStr = sqlStr + "		INSERT INTO [db_item].[dbo].[tbl_item_pack_Volumn](itemid, volX, volY, volZ)" & VbCrlf
				sqlStr = sqlStr + "		VALUES (" & itemid & "," & volX & "," & volY & "," & volZ & ")" & VbCrlf
				sqlStr = sqlStr + "	END" & VbCrlf
				rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
			end if
		else
			sqlStr = sqlStr + "		UPDATE [db_shop].[dbo].[tbl_shop_item] " & VbCrlf
			sqlStr = sqlStr + "		SET itemWeight='" + CStr(itemWeight) + "', volX=" & volX & VbCrlf
			sqlStr = sqlStr + "		,volY=" & volY & VbCrlf
			sqlStr = sqlStr + "		,volZ=" & volZ & VbCrlf
			sqlStr = sqlStr + "		,updt=GETDATE()" & VbCrlf
			sqlStr = sqlStr + "		WHERE itemgubun = '" & itemgubun & "' and shopitemid=" & itemid & " AND itemoption='" & itemoption & "'" & VbCrlf
			rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
        end if
	Next

    if (itemgubun = "10") then
	    sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
	    sqlStr = sqlStr + " set itemWeight='" + CStr(itemWeight) + "'" + VbCrlf
	    sqlStr = sqlStr + " where itemid=" + CStr(itemid)  + VbCrlf
	    dbget.execute(sqlStr)

	    '//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
        Call DoSomethingForForeignSite(itemid)
	    '//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
    end if

elseif mode="chdeliverOverseas" then
	if chdeliverOverseas="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('일괄변경하실 해외배송여부를 선택해 주세요.');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if request.form("check").count>0 then
		for i=1 to request.form("check").count
			itemid = request.form("check")(i)

			sqlStr = "update [db_item].[dbo].tbl_item" & VbCrlf
			sqlStr = sqlStr & " set deliverOverseas='" & chdeliverOverseas & "' " & VbCrlf
			sqlStr = sqlStr & " , lastupdate = getdate() where" & VbCrlf
			sqlStr = sqlStr & " itemid="& trim(itemid) &"" & VbCrlf

			response.write sqlStr & "<Br>"
			dbget.execute sqlStr
		next
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('일괄변경하실 상품을 선택해 주세요.');"
		response.write "</script>"
		dbget.close() : response.end
	end if

	response.write "<script type='text/javascript'>"
	response.write "	alert('저장 되었습니다.');"
	response.write "	parent.location.reload();"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<script type="text/javascript">
	alert('등록 되었습니다.');
	location.replace('<%= refer %>&prcAfter=Y');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
