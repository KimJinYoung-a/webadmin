<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
response.charSet = "utf-8"
%>
<%
'####################################################
' Description :  온라인 다국어 상품 설명 입력 처리
' History : 2013.07.12 허진원 생성
'			2016.08.30 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim vItemID, vCountryCd, vItemName, vItemContent, vItemCopy, vItemSource
dim vItemSize, vMakerName, vSourceArea, useyn, i, sqlstr, keywords, mode
dim vitemoption, voptisusing, voptiontypename, voptionname
dim multilangcnt, linkPriceTypeusd, vQuery
	vItemID = requestCheckVar(Request("itemid"),10)
	vCountryCd = requestCheckVar(Request("countrycd"),32)
	mode = requestCheckVar(Request("mode"),32)

linkPriceTypeusd=0
multilangcnt=0

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

If vItemID = "" OR vCountryCd = "" Then
	Response.Write "<script type='text/javascript'>alert('잘못된 경로입니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : Response.End
End IF

'// 상품정보 받기
vItemName = chrbyte(Request("itemname"),60,"")
vItemContent = chrbyte(Request("itemcontent"),800,"")
vItemCopy = chrbyte(Request("itemcopy"),250,"")
vItemSource = chrbyte(Request("itemsource"),128,"")
vItemSize = chrbyte(Request("itemsize"),128,"")
vMakerName = chrbyte(Request("makername"),64,"")
vSourceArea = chrbyte(Request("sourcearea"),128,"")
useyn = chrbyte(Request("useyn"),1,"")
keywords = chrbyte(Request("keywords"),512,"")

'//옵션 받기 (배열처리)
redim vitemoption(Request("itemoption").Count)
redim voptisusing(Request("optisusing").Count)
redim voptiontypename(Request("optiontypename").Count)
redim voptionname(Request("optionname").Count)

for i=1 to Request("itemoption").Count
	vitemoption(i) = chrbyte(Request("itemoption")(i),4,"")
	voptisusing(i) = chrbyte(Request("optisusing")(i),1,"")
	voptiontypename(i) = chrbyte(Request("optiontypename")(i),32,"")
	voptionname(i) = chrbyte(Request("optionname")(i),96,"")
next

if vItemContent<>"" then
	if checkNotValidHTML(vItemContent) then
		Response.Write "<script type='text/javascript'>alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.');window.close()</script>"
		session.codePage = 949 : dbget.close() : Response.End
	end if
	
	vItemContent = replace(vItemContent,"'","""")
end if
if vItemName<>"" then vItemName = replace(vItemName,"'","""")
if vItemCopy<>"" then vItemCopy = replace(vItemCopy,"'","""")
if vItemSource<>"" then vItemSource = replace(vItemSource,"'","""")
if vItemSize<>"" then vItemSize = replace(vItemSize,"'","""")
if vMakerName<>"" then vMakerName = replace(vMakerName,"'","""")
if vSourceArea<>"" then vSourceArea = replace(vSourceArea,"'","""")
if keywords<>"" then keywords = replace(keywords,"'","""")

Select Case mode
	'//신규등록
	Case "new"
		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
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

'		'/오프라인 상품
'		sqlStr = "select" & vbcrlf
'		sqlStr = sqlStr & "	itemgubun, shopitemid" & vbcrlf
'		sqlStr = sqlStr & "	into #tmp_shop_item" & vbcrlf
'		sqlStr = sqlStr & "	from db_shop.dbo.tbl_shop_item" & vbcrlf
'		sqlStr = sqlStr & "	where itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "	group by itemgubun, shopitemid" & vbcrlf
'		sqlStr = sqlStr & "	CREATE NONCLUSTERED INDEX [tmp_shopitemid] ON #tmp_shop_item(shopitemid ASC)" & vbcrlf
'
'    	'response.write sqlStr & "<br>"
'    	dbget.execute sqlStr

		'/상품 꼿고
		sqlStr = "insert into [db_item].[dbo].[tbl_item_multiSite_regItem]" & vbcrlf
		sqlStr = sqlStr & "		select" & vbcrlf
		sqlStr = sqlStr & "		i.itemid, 'WSLWEB', 'Y', 0, 0, getdate(), 'SYSTEM', getdate(), 'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
'		sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
'		sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
		sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
		sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		where ri.itemid is null" & vbcrlf
		sqlStr = sqlStr & "		and i.itemid = "& vItemID &"" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		'//언어팩 등록
		vQuery = " IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_multiLang] WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "') "
		vQuery = vQuery & " BEGIN "
		vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang] SET"
		vQuery = vQuery & "			itemname = N'" & db2html(vItemName) & "', "
		vQuery = vQuery & "			itemcopy = N'" & db2html(vItemCopy) & "', "
		vQuery = vQuery & "			itemContent = N'" & db2html(vItemContent) & "', "
		vQuery = vQuery & "			itemsource = N'" & db2html(vItemSource) & "', "
		vQuery = vQuery & "			itemsize = N'" & db2html(vItemSize) & "', "
		vQuery = vQuery & "			makername = N'" & db2html(vMakerName) & "', "
		vQuery = vQuery & "			sourcearea = N'" & db2html(vSourceArea) & "', "
		vQuery = vQuery & "			useyn = N'" & useyn & "', "
		vQuery = vQuery & "			keywords = N'" & db2html(keywords) & "', "
		vQuery = vQuery & "			lastupdate = getdate()"
		vQuery = vQuery & "		WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' "
		vQuery = vQuery & " END "
		vQuery = vQuery & " ELSE "
		vQuery = vQuery & " BEGIN "
		vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang](itemid, countryCd, itemname, itemcopy, itemContent, itemsource, itemsize, makername, sourcearea, useyn, keywords) "
		vQuery = vQuery & "		VALUES(N'" & vItemID & "', N'" & vCountryCd & "', N'" & db2html(vItemName) & "', N'" & db2html(vItemCopy) & "', N'" & db2html(vItemContent) & "' "
		vQuery = vQuery & "		, N'" & db2html(vItemSource) & "', N'" & db2html(vItemSize) & "', N'" & db2html(vMakerName) & "', N'" & db2html(vSourceArea) & "', N'" & useyn & "'"
		vQuery = vQuery & "		, N'" & db2html(keywords) & "') "
		vQuery = vQuery & " END "

		'response.write vQuery &"<Br>"
		dbget.execute vQuery

		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = "insert into db_item.[dbo].[tbl_item_multiLang]" & vbcrlf
			vQuery = vQuery & " 	select" & vbcrlf
			vQuery = vQuery & " 	i.itemid, 'KR', i.itemname, isnull(ic.designercomment,'') as designercomment, '', ic.itemsource, ic.itemsize, ic.sourcearea, c.socname_kor" & vbcrlf
			vQuery = vQuery & " 	, 'Y', getdate(), getdate(), ic.keywords, ''" & vbcrlf
			vQuery = vQuery & " 	from db_item.dbo.tbl_item i" & vbcrlf
			vQuery = vQuery & " 	left join db_user.dbo.tbl_user_c c" & vbcrlf
			vQuery = vQuery & " 		on i.makerid=c.userid" & vbcrlf
			vQuery = vQuery & " 	left join db_item.[dbo].[tbl_item_multiLang] ml" & vbcrlf
			vQuery = vQuery & " 		on i.itemid=ml.itemid" & vbcrlf
			vQuery = vQuery & " 		and ml.countryCd='KR'" & vbcrlf
			vQuery = vQuery & " 	left join db_item.dbo.tbl_item_Contents ic" & vbcrlf
			vQuery = vQuery & " 		on i.itemid = ic.itemid" & vbcrlf
			vQuery = vQuery & " 	where ml.itemid is null" & vbcrlf
			vQuery = vQuery & " 	and i.itemid = " & vItemID & "" & vbcrlf
			vQuery = vQuery & " 	" & vbcrlf

			'response.write vQuery &"<Br>"
			dbget.execute vQuery
		end if

		'/언어팩 카운트 계산해서 꼿음
		sqlStr = "update ri" & vbcrlf
		sqlStr = sqlStr & "		set ri.multilangcnt = isnull(t.multilangcnt,0)" & vbcrlf
		sqlStr = sqlStr & "		, ri.lastupdate = getdate()" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item_multiSite_regItem ri" & vbcrlf
		sqlStr = sqlStr & "		left join (" & vbcrlf
		sqlStr = sqlStr & "			select itemid, count(*) as multilangcnt" & vbcrlf
		sqlStr = sqlStr & "			from [db_item].[dbo].[tbl_item_multiLang]" & vbcrlf
		sqlStr = sqlStr & "			where useyn='Y'" & vbcrlf
		sqlStr = sqlStr & "			and itemid = "& vItemID &"" & vbcrlf
		sqlStr = sqlStr & "			group by itemid" & vbcrlf
		sqlStr = sqlStr & "		) as t" & vbcrlf
		sqlStr = sqlStr & "			on ri.itemid = t.itemid" & vbcrlf
		sqlStr = sqlStr & "		where ri.sitename= 'WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		and ri.itemid = "& vItemID &"" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		'/옵션 꼿고
		if ubound(vitemoption)>0 then
			sqlstr = ""
			for i=1 to ubound(vitemoption)
				if vitemoption(i)<>"0000" then
					vQuery = " IF EXISTS(SELECT itemoption FROM [db_item].[dbo].[tbl_item_multiLang_option] WHERE" & vbCrLf
					vQuery = vQuery & " itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' AND itemoption = '" & requestCheckVar(vitemoption(i),4) & "') " & vbCrLf
					vQuery = vQuery & " BEGIN " & vbCrLf
					vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang_option]" & vbCrLf
					vQuery = vQuery & "		SET optiontypename = N'" & db2html(requestCheckVar(voptiontypename(i),32)) & "'" & vbCrLf
					vQuery = vQuery & "		, optionname = N'" & db2html(requestCheckVar(voptionname(i),96)) & "'" & vbCrLf
					vQuery = vQuery & "		, isusing = N'" & requestCheckVar(voptisusing(i),1) & "'" & vbCrLf
					vQuery = vQuery & "		, lastupdate=getdate()" & vbCrLf
					vQuery = vQuery & "		WHERE countryCD='"&vCountryCd&"' And itemid = '" & vItemID & "' AND itemoption = '" & requestCheckVar(vitemoption(i),4) & "'" & vbCrLf
					vQuery = vQuery & " END " & vbCrLf
					vQuery = vQuery & " ELSE " & vbCrLf
					vQuery = vQuery & " BEGIN " & vbCrLf
					vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang_option](itemid, countryCd, itemoption, optiontypename, optionname, isusing, regdate, lastupdate)" & vbCrLf
					vQuery = vQuery & "		VALUES(" & vbCrLf
					vQuery = vQuery & "		N'" & vItemID & "', N'" & vCountryCd & "', N'" & requestCheckVar(vitemoption(i),4) & "', N'" & db2html(requestCheckVar(voptiontypename(i),32)) & "'"
					vQuery = vQuery & "		, N'" & db2html(requestCheckVar(voptionname(i),96)) & "'" & vbCrLf
					vQuery = vQuery & "		, N'" & requestCheckVar(voptisusing(i),1) & "', getdate(), getdate()" & vbCrLf
					vQuery = vQuery & "		) " & vbCrLf
					vQuery = vQuery & " END " & vbCrLf

					'response.write vQuery & "<br>"
					dbget.execute vQuery
				end if
			next
		end if

		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = " insert into db_item.[dbo].[tbl_item_multiLang_option]" & vbCrLf
			vQuery = vQuery & "		select" & vbCrLf
			vQuery = vQuery & "		i.itemid, 'KR', o.itemoption, 'Y', o.optionTypeName, o.optionname, getdate(), getdate()" & vbCrLf
			vQuery = vQuery & "		from db_item.dbo.tbl_item i" & vbCrLf
			vQuery = vQuery & "		left join db_item.dbo.tbl_item_option o" & vbCrLf
			vQuery = vQuery & "			on i.itemid = o.itemid" & vbCrLf
			vQuery = vQuery & "		left join db_item.[dbo].[tbl_item_multiLang_option] mo" & vbCrLf
			vQuery = vQuery & "			on i.itemid=mo.itemid" & vbCrLf
			vQuery = vQuery & "			and o.itemoption = mo.itemoption" & vbCrLf
			vQuery = vQuery & "			and mo.countryCd='KR'" & vbCrLf
			vQuery = vQuery & "		where mo.itemid is null" & vbCrLf
			vQuery = vQuery & "		and o.itemoption is not null" & vbCrLf		'옵션없음은 넣지 안음.
			vQuery = vQuery & "		and i.itemid = " & vItemID & "" & vbCrLf

			'response.write vQuery & "<br>"
			dbget.execute vQuery
		end if

		'/ 가격 꼿고
		sqlStr = "insert into db_item.dbo.tbl_item_multiLang_price" & vbcrlf
		sqlStr = sqlStr & "		select " & vbcrlf
		sqlStr = sqlStr & "		'WSLWEB' ,i.itemid, ee.currencyUnit" & vbcrlf
		sqlStr = sqlStr & "		, (case" & vbcrlf
		sqlStr = sqlStr & "			when ee.currencyUnit='WON' or ee.currencyUnit='KRW' then (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end)" & vbcrlf
		sqlStr = sqlStr & "			else round((((( (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) *ee.multiplerate)/ee.exchangeRate)*100)/100) ,2)" & vbcrlf
		sqlStr = sqlStr & "			end) as orgprice" & vbcrlf
		sqlStr = sqlStr & "		, ((case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) * ee.multiplerate) as wonPrice" & vbcrlf
		sqlStr = sqlStr & "		,NULL as mayDiscountPrice" & vbcrlf
		sqlStr = sqlStr & "		,ee.multiplerate" & vbcrlf
		sqlStr = sqlStr & "		,getdate()" & vbcrlf
		sqlStr = sqlStr & "		,getdate()" & vbcrlf
		sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
'		sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
'		sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
		sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
		sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		join #tmp_exchangeRatecurrencyunitgroup ee" & vbcrlf
		sqlStr = sqlStr & "			on ri.sitename = ee.sitename" & vbcrlf
		sqlStr = sqlStr & "		left join db_item.dbo.tbl_item_multiLang_price p" & vbcrlf
		sqlStr = sqlStr & "			on ri.sitename=p.sitename" & vbcrlf
		sqlStr = sqlStr & "			and i.itemid=P.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ee.currencyUnit=P.currencyUnit" & vbcrlf
		sqlStr = sqlStr & "		where i.itemid = "& vItemID &"" & vbcrlf
		sqlStr = sqlStr & "		and P.itemid is NULL" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		if useyn="Y" then
			sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
			sqlStr = sqlStr & " set deliverOverseas='Y'" + VbCrlf
			sqlStr = sqlStr & " , lastupdate = getdate() where" + VbCrlf
			sqlStr = sqlStr & " itemid="& vItemID &"" + VbCrlf

			'response.write sqlStr & "<br>"
			dbget.execute sqlStr
		end if

		sqlStr = "drop table #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
'		sqlStr = sqlStr & " drop table #tmp_shop_item" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr
		'//////////////////////// 해외 판매 상품 자동 입력 ////////////////////////

	'//정보수정
	Case "modi"
		'//////////////////////// 해외 판매 상품 자동 입력	'/2016.05.31 한용민 생성 ////////////////////////
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

'		'/오프라인 상품
'		sqlStr = "select" & vbcrlf
'		sqlStr = sqlStr & "	itemgubun, shopitemid" & vbcrlf
'		sqlStr = sqlStr & "	into #tmp_shop_item" & vbcrlf
'		sqlStr = sqlStr & "	from db_shop.dbo.tbl_shop_item" & vbcrlf
'		sqlStr = sqlStr & "	where itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "	group by itemgubun, shopitemid" & vbcrlf
'		sqlStr = sqlStr & "	CREATE NONCLUSTERED INDEX [tmp_shopitemid] ON #tmp_shop_item(shopitemid ASC)" & vbcrlf
'
'    	'response.write sqlStr & "<br>"
'    	dbget.execute sqlStr

		'/상품 꼿고
		sqlStr = "insert into [db_item].[dbo].[tbl_item_multiSite_regItem]" & vbcrlf
		sqlStr = sqlStr & "		select" & vbcrlf
		sqlStr = sqlStr & "		i.itemid, 'WSLWEB', 'Y', 0, 0, getdate(), 'SYSTEM', getdate(), 'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
'		sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
'		sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
		sqlStr = sqlStr & "		left join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
		sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		where ri.itemid is null" & vbcrlf
		sqlStr = sqlStr & "		and i.itemid = "& vItemID &"" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		'//언어팩 등록
		vQuery = " IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_multiLang] WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "') "
		vQuery = vQuery & " BEGIN "
		vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang] SET"
		vQuery = vQuery & "			itemname = N'" & db2html(vItemName) & "', "
		vQuery = vQuery & "			itemcopy = N'" & db2html(vItemCopy) & "', "
		vQuery = vQuery & "			itemContent = N'" & db2html(vItemContent) & "', "
		vQuery = vQuery & "			itemsource = N'" & db2html(vItemSource) & "', "
		vQuery = vQuery & "			itemsize = N'" & db2html(vItemSize) & "', "
		vQuery = vQuery & "			makername = N'" & db2html(vMakerName) & "', "
		vQuery = vQuery & "			sourcearea = N'" & db2html(vSourceArea) & "', "
		vQuery = vQuery & "			useyn = N'" & useyn & "', "
		vQuery = vQuery & "			keywords = N'" & db2html(keywords) & "', "
		vQuery = vQuery & "			lastupdate = getdate()"
		vQuery = vQuery & "		WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' "
		vQuery = vQuery & " END "
		vQuery = vQuery & " ELSE "
		vQuery = vQuery & " BEGIN "
		vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang](itemid, countryCd, itemname, itemcopy, itemContent, itemsource, itemsize, makername, sourcearea, useyn, keywords) "
		vQuery = vQuery & "		VALUES(N'" & vItemID & "', N'" & vCountryCd & "', N'" & db2html(vItemName) & "', N'" & db2html(vItemCopy) & "', N'" & db2html(vItemContent) & "' "
		vQuery = vQuery & "		, N'" & db2html(vItemSource) & "', N'" & db2html(vItemSize) & "', N'" & db2html(vMakerName) & "', N'" & db2html(vSourceArea) & "', N'" & useyn & "'"
		vQuery = vQuery & "		, N'" & db2html(keywords) & "') "
		vQuery = vQuery & " END "

		'response.write vQuery &"<Br>"
		dbget.execute vQuery

		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = "insert into db_item.[dbo].[tbl_item_multiLang]" & vbcrlf
			vQuery = vQuery & " 	select" & vbcrlf
			vQuery = vQuery & " 	i.itemid, 'KR', i.itemname, isnull(ic.designercomment,'') as designercomment, '', ic.itemsource, ic.itemsize, ic.sourcearea, c.socname_kor" & vbcrlf
			vQuery = vQuery & " 	, 'Y', getdate(), getdate(), ic.keywords, ''" & vbcrlf
			vQuery = vQuery & " 	from db_item.dbo.tbl_item i" & vbcrlf
			vQuery = vQuery & " 	left join db_user.dbo.tbl_user_c c" & vbcrlf
			vQuery = vQuery & " 		on i.makerid=c.userid" & vbcrlf
			vQuery = vQuery & " 	left join db_item.[dbo].[tbl_item_multiLang] ml" & vbcrlf
			vQuery = vQuery & " 		on i.itemid=ml.itemid" & vbcrlf
			vQuery = vQuery & " 		and ml.countryCd='KR'" & vbcrlf
			vQuery = vQuery & " 	left join db_item.dbo.tbl_item_Contents ic" & vbcrlf
			vQuery = vQuery & " 		on i.itemid = ic.itemid" & vbcrlf
			vQuery = vQuery & " 	where ml.itemid is null" & vbcrlf
			vQuery = vQuery & " 	and i.itemid = " & vItemID & "" & vbcrlf
			vQuery = vQuery & " 	" & vbcrlf

			'response.write vQuery &"<Br>"
			dbget.execute vQuery
		end if

		'/언어팩 카운트 계산해서 꼿음
		sqlStr = "update ri" & vbcrlf
		sqlStr = sqlStr & "		set ri.multilangcnt = isnull(t.multilangcnt,0)" & vbcrlf
		sqlStr = sqlStr & "		, ri.lastupdate = getdate()" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item_multiSite_regItem ri" & vbcrlf
		sqlStr = sqlStr & "		left join (" & vbcrlf
		sqlStr = sqlStr & "			select itemid, count(*) as multilangcnt" & vbcrlf
		sqlStr = sqlStr & "			from [db_item].[dbo].[tbl_item_multiLang]" & vbcrlf
		sqlStr = sqlStr & "			where useyn='Y'" & vbcrlf
		sqlStr = sqlStr & "			and itemid = "& vItemID &"" & vbcrlf
		sqlStr = sqlStr & "			group by itemid" & vbcrlf
		sqlStr = sqlStr & "		) as t" & vbcrlf
		sqlStr = sqlStr & "			on ri.itemid = t.itemid" & vbcrlf
		sqlStr = sqlStr & "		where ri.sitename= 'WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		and ri.itemid = "& vItemID &"" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		'/옵션 꼿고
		if ubound(vitemoption)>0 then
			sqlstr = ""
			for i=1 to ubound(vitemoption)
				if vitemoption(i)<>"0000" then
					vQuery = " IF EXISTS(SELECT itemoption FROM [db_item].[dbo].[tbl_item_multiLang_option] WHERE" & vbCrLf
					vQuery = vQuery & " itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' AND itemoption = '" & requestCheckVar(vitemoption(i),4) & "') " & vbCrLf
					vQuery = vQuery & " BEGIN " & vbCrLf
					vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang_option]" & vbCrLf
					vQuery = vQuery & "		SET optiontypename = N'" & db2html(requestCheckVar(voptiontypename(i),32)) & "'" & vbCrLf
					vQuery = vQuery & "		, optionname = N'" & db2html(requestCheckVar(voptionname(i),96)) & "'" & vbCrLf
					vQuery = vQuery & "		, isusing = N'" & requestCheckVar(voptisusing(i),1) & "'" & vbCrLf
					vQuery = vQuery & "		, lastupdate=getdate()" & vbCrLf
					vQuery = vQuery & "		WHERE countryCD='"&vCountryCd&"' And itemid = '" & vItemID & "' AND itemoption = '" & requestCheckVar(vitemoption(i),4) & "'" & vbCrLf
					vQuery = vQuery & " END " & vbCrLf
					vQuery = vQuery & " ELSE " & vbCrLf
					vQuery = vQuery & " BEGIN " & vbCrLf
					vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang_option](itemid, countryCd, itemoption, optiontypename, optionname, isusing, regdate, lastupdate)" & vbCrLf
					vQuery = vQuery & "		VALUES(" & vbCrLf
					vQuery = vQuery & "		N'" & vItemID & "', N'" & vCountryCd & "', N'" & requestCheckVar(vitemoption(i),4) & "', N'" & db2html(requestCheckVar(voptiontypename(i),32)) & "'"
					vQuery = vQuery & "		, N'" & db2html(requestCheckVar(voptionname(i),96)) & "'" & vbCrLf
					vQuery = vQuery & "		, N'" & requestCheckVar(voptisusing(i),1) & "', getdate(), getdate()" & vbCrLf
					vQuery = vQuery & "		) " & vbCrLf
					vQuery = vQuery & " END " & vbCrLf

					'response.write vQuery & "<br>"
					dbget.execute vQuery
				end if
			next
		end if

		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = " insert into db_item.[dbo].[tbl_item_multiLang_option]" & vbCrLf
			vQuery = vQuery & "		select" & vbCrLf
			vQuery = vQuery & "		i.itemid, 'KR', o.itemoption, 'Y', o.optionTypeName, o.optionname, getdate(), getdate()" & vbCrLf
			vQuery = vQuery & "		from db_item.dbo.tbl_item i" & vbCrLf
			vQuery = vQuery & "		left join db_item.dbo.tbl_item_option o" & vbCrLf
			vQuery = vQuery & "			on i.itemid = o.itemid" & vbCrLf
			vQuery = vQuery & "		left join db_item.[dbo].[tbl_item_multiLang_option] mo" & vbCrLf
			vQuery = vQuery & "			on i.itemid=mo.itemid" & vbCrLf
			vQuery = vQuery & "			and o.itemoption = mo.itemoption" & vbCrLf
			vQuery = vQuery & "			and mo.countryCd='KR'" & vbCrLf
			vQuery = vQuery & "		where mo.itemid is null" & vbCrLf
			vQuery = vQuery & "		and o.itemoption is not null" & vbCrLf		'옵션없음은 넣지 안음.
			vQuery = vQuery & "		and i.itemid = " & vItemID & "" & vbCrLf

			'response.write vQuery & "<br>"
			dbget.execute vQuery
		end if

		'/ 가격 꼿고
		sqlStr = "insert into db_item.dbo.tbl_item_multiLang_price" & vbcrlf
		sqlStr = sqlStr & "		select " & vbcrlf
		sqlStr = sqlStr & "		'WSLWEB' ,i.itemid, ee.currencyUnit" & vbcrlf
		sqlStr = sqlStr & "		, (case" & vbcrlf
		sqlStr = sqlStr & "			when ee.currencyUnit='WON' or ee.currencyUnit='KRW' then (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end)" & vbcrlf
		sqlStr = sqlStr & "			else round((((( (case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) *ee.multiplerate)/ee.exchangeRate)*100)/100) ,2)" & vbcrlf
		sqlStr = sqlStr & "			end) as orgprice" & vbcrlf
		sqlStr = sqlStr & "		, ((case when ee.linkPriceType='1' then i.sellcash else i.orgPrice end) * ee.multiplerate) as wonPrice" & vbcrlf
		sqlStr = sqlStr & "		,NULL as mayDiscountPrice" & vbcrlf
		sqlStr = sqlStr & "		,ee.multiplerate" & vbcrlf
		sqlStr = sqlStr & "		,getdate()" & vbcrlf
		sqlStr = sqlStr & "		,getdate()" & vbcrlf
		sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		,'SYSTEM'" & vbcrlf
		sqlStr = sqlStr & "		from db_item.dbo.tbl_item i" & vbcrlf
'		sqlStr = sqlStr & "		join #tmp_shop_item ii" & vbcrlf
'		sqlStr = sqlStr & "			on ii.itemgubun='10'" & vbcrlf
'		sqlStr = sqlStr & "			and i.itemid = ii.shopitemid" & vbcrlf
		sqlStr = sqlStr & "		join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
		sqlStr = sqlStr & "			on i.itemid = ri.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ri.sitename='WSLWEB'" & vbcrlf
		sqlStr = sqlStr & "		join #tmp_exchangeRatecurrencyunitgroup ee" & vbcrlf
		sqlStr = sqlStr & "			on ri.sitename = ee.sitename" & vbcrlf
		sqlStr = sqlStr & "		left join db_item.dbo.tbl_item_multiLang_price p" & vbcrlf
		sqlStr = sqlStr & "			on ri.sitename=p.sitename" & vbcrlf
		sqlStr = sqlStr & "			and i.itemid=P.itemid" & vbcrlf
		sqlStr = sqlStr & "			and ee.currencyUnit=P.currencyUnit" & vbcrlf
		sqlStr = sqlStr & "		where i.itemid = "& vItemID &"" & vbcrlf
		sqlStr = sqlStr & "		and P.itemid is NULL" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr

		if useyn="Y" then
			sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
			sqlStr = sqlStr & " set deliverOverseas='Y'" + VbCrlf
			sqlStr = sqlStr & " , lastupdate = getdate() where" + VbCrlf
			sqlStr = sqlStr & " itemid="& vItemID &"" + VbCrlf

			'response.write sqlStr & "<br>"
			dbget.execute sqlStr
		end if

		sqlStr = "drop table #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
'		sqlStr = sqlStr & " drop table #tmp_shop_item" & vbcrlf

    	'response.write sqlStr & "<br>"
    	dbget.execute sqlStr
		'//////////////////////// 해외 판매 상품 자동 입력 ////////////////////////
End Select
%>

<script type="text/javascript">
	alert("저장되었습니다.");
	opener.document.location.reload();
	window.close();
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<% session.codePage = 949 %>