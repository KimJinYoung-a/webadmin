<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
response.charSet = "utf-8"
'####################################################
' Description :  온라인 환율 관리
' History : 2013.05.02 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%
dim idx, sitename, currencyUnit, currencyChar, exchangeRate, basedate ,sqlStr, userid, menupos, mode, makerid
dim countrylangcd, multipleRate, linkPriceType, existscount, referer
dim orgsitename, orgcountrylangcd, orgcurrencyUnit, orgmultipleRate, orgexchangeRate, orglinkPriceType
	idx   = requestCheckVar(getNumeric(request("idx")),10)
	sitename   = requestCheckVar(request("sitename"),32)
	makerid	= Trim(requestCheckVar(request("makerid"),32))
	currencyUnit   = requestCheckVar(request("currencyUnit"),16)
	currencyChar   = requestCheckVar(request("currencyChar"),50)
	exchangeRate   = requestCheckVar(request("exchangeRate"),20)
	basedate   = requestCheckVar(request("basedate"),10)
	userid = session("ssBctId")
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	mode = requestCheckVar(request("mode"),32)
	countrylangcd = requestCheckVar(request("countrylangcd"),32)
	multipleRate = requestCheckVar(request("multipleRate"),20)
	linkPriceType = requestCheckVar(request("linkPriceType"),10)

if (linkPriceType="") then linkPriceType="0"
existscount=0
referer = request.ServerVariables("HTTP_REFERER")

if mode = "exchangeRateedit" then
	if idx="0" or idx="" or isnull(idx) then
		sqlStr = "select count(idx) as existscount"
		sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate with (nolock)"
		sqlStr = sqlStr & " where sitename='"& sitename &"' and countrylangcd='"& countrylangcd &"' and currencyUnit='"& currencyUnit &"'"

		'response.write sqlStr & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		if not rsget.EOF  then
			existscount = rsget("existscount")
		end if
		rsget.close

		if existscount>0 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('이미 존재하는 사이트구분 대표언어 화폐단위 입니다.');"
			response.write "</script>"
			dbget.close() : response.end
		end if
	end if

	sqlStr = "if exists(" + VbCrlf
	sqlStr = sqlStr & "		select top 1 * from db_item.dbo.tbl_exchangeRate where idx='"&idx&"'" + VbCrlf
	sqlStr = sqlStr & " )" + VbCrlf
    sqlStr = sqlStr & " 	update db_item.dbo.tbl_exchangeRate" + VbCrlf
    sqlStr = sqlStr & " 	set currencyChar=N'" + currencyChar + "'" + VbCrlf
    sqlStr = sqlStr & " 	,countrylangcd=N'" + countrylangcd + "'" + VbCrlf
    sqlStr = sqlStr & " 	,exchangeRate=N'" + exchangeRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,multipleRate=N'" + multipleRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,linkPriceType=N'" + linkPriceType + "'" + VbCrlf
    sqlStr = sqlStr & " 	,basedate=N'" + basedate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,lastupdate=getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,lastuserid=N'" + userid + "'" + VbCrlf
	sqlStr = sqlStr & " 	,makerid=N'" + makerid + "'" + VbCrlf
    sqlStr = sqlStr & " 	where idx='"&idx&"'" + VbCrlf
	sqlStr = sqlStr & " else" + VbCrlf
	sqlStr = sqlStr & " 	insert into db_item.dbo.tbl_exchangeRate (" + VbCrlf
    sqlStr = sqlStr & " 	sitename, countrylangcd, multipleRate, currencyUnit ,currencyChar ,exchangeRate ,linkPriceType,basedate ,regdate, lastupdate, reguserid, lastuserid, makerid"+ VbCrlf
	sqlStr = sqlStr & " 	) values("
    sqlStr = sqlStr & " 	N'" + sitename + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + countrylangcd + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + multipleRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + currencyUnit + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + currencyChar + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + exchangeRate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + linkPriceType + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + basedate + "'" + VbCrlf
    sqlStr = sqlStr & " 	,getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,getdate()" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + userid + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + userid + "'" + VbCrlf
    sqlStr = sqlStr & " 	,N'" + makerid + "'" + VbCrlf
    sqlStr = sqlStr & " 	)"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

	'/수정일 경우에만
	if idx <> "" then
		'/사이트구분이 홀쎄일 일경우
		if sitename="WSLWEB" then
			orgsitename=""
			orgcountrylangcd=""
			orgcurrencyUnit=""
			orgmultipleRate=""
			orgexchangeRate=""
			orglinkPriceType=""
			'/수정전 데이터를 받아옴
			sqlStr = "select top 1 sitename, countrylangcd, currencyUnit, multipleRate, exchangeRate, linkPriceType"
			sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate"
			sqlStr = sqlStr & " where idx="& idx &""

			'response.write sqlStr & "<br>"
			rsget.Open sqlStr,dbget,1

			if not rsget.EOF  then
				orgsitename = rsget("sitename")
				orgcountrylangcd = rsget("countrylangcd")
				orgcurrencyUnit = rsget("currencyUnit")
				orgmultipleRate = rsget("multipleRate")
				orgexchangeRate = rsget("exchangeRate")
				orglinkPriceType = rsget("linkPriceType")
			end if
			rsget.close

			'/사이트구분과 화폐단위가 수정전과 수정후가 같은거
			if orgsitename=sitename and orgcurrencyUnit=currencyUnit then
				'/환율이나 대표배수가 수정 될경우
				if orgmultipleRate<>multipleRate or orgexchangeRate<>exchangeRate or orglinkPriceType<>linkPriceType then
					'/ 사이트별 화폐단위
					sqlStr = "select" & vbcrlf
					sqlStr = sqlStr & " e.sitename, e.currencyunit" & vbcrlf
					sqlStr = sqlStr & " , (select top 1 exchangeRate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as exchangeRate" & vbcrlf		' 환율
					sqlStr = sqlStr & " , (select top 1 multiplerate from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as multiplerate" & vbcrlf		' 배수
					sqlStr = sqlStr & " , (select top 1 linkPriceType from db_item.dbo.tbl_exchangeRate where e.sitename=sitename and e.currencyunit=currencyunit) as linkPriceType" & vbcrlf		' 적용가격
					sqlStr = sqlStr & " into #tmp_exchangeRatecurrencyunitgroup" & vbcrlf
					sqlStr = sqlStr & " from db_item.dbo.tbl_exchangeRate e" & vbcrlf
					sqlStr = sqlStr & " where e.sitename='"& sitename &"'" & vbcrlf
					sqlStr = sqlStr & " group by e.sitename, e.currencyunit" & vbcrlf

					'/ 오프라인 상품
					sqlStr = sqlStr & " select" & vbcrlf
					sqlStr = sqlStr & " itemgubun, shopitemid" & vbcrlf
					sqlStr = sqlStr & " , min(orgsellprice) as orgsellprice, min(shopitemprice)	as shopitemprice" & vbcrlf		' 옵션별 추가 금액 제끼고 가져옴
					sqlStr = sqlStr & " into #tmp_shop_item" & vbcrlf
					sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item" & vbcrlf
					sqlStr = sqlStr & " where itemgubun='10'" & vbcrlf
					'sqlStr = sqlStr & " and isusing='Y'" & vbcrlf
					sqlStr = sqlStr & " group by itemgubun, shopitemid" & vbcrlf

					'/ 일괄 적용
					sqlStr = sqlStr & " update p set" & vbcrlf
					sqlStr = sqlStr & " p.orgprice = (case" & vbcrlf
					sqlStr = sqlStr & " 		when ee.currencyUnit='WON' or ee.currencyUnit='KRW' then (case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end)" & vbcrlf
					sqlStr = sqlStr & " 		else CEILING((( (case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end) *ee.multiplerate)/ee.exchangeRate) *2)/2" & vbcrlf
					sqlStr = sqlStr & " 		end)" & vbcrlf
					sqlStr = sqlStr & " , p.wonprice = ((case when ee.linkPriceType='1' then ii.shopitemprice else ii.orgsellprice end) * ee.multiplerate)" & vbcrlf
					sqlStr = sqlStr & " , p.lastupdate = getdate()" & vbcrlf
					sqlStr = sqlStr & " , p.lastuserid = 'SYSTEM'" & vbcrlf
					sqlStr = sqlStr & " , p.lastexchangerate = ee.exchangeRate" & vbcrlf
					sqlStr = sqlStr & " from #tmp_shop_item ii" & vbcrlf
					sqlStr = sqlStr & " join [db_item].[dbo].[tbl_item_multiSite_regItem] ri" & vbcrlf
					sqlStr = sqlStr & " 	on ii.shopitemid = ri.itemid" & vbcrlf
					sqlStr = sqlStr & " 	and ri.sitename='"& sitename &"'" & vbcrlf
					sqlStr = sqlStr & " join #tmp_exchangeRatecurrencyunitgroup ee" & vbcrlf
					sqlStr = sqlStr & " 	on ri.sitename = ee.sitename" & vbcrlf
					sqlStr = sqlStr & " 	and currencyUnit = '"& currencyUnit &"'" & vbcrlf
					sqlStr = sqlStr & " join db_item.dbo.tbl_item_multiLang_price p" & vbcrlf
					sqlStr = sqlStr & " 	on ri.sitename=p.sitename" & vbcrlf
					sqlStr = sqlStr & " 	and ii.shopitemid=P.itemid" & vbcrlf
					sqlStr = sqlStr & " 	and ee.currencyUnit=P.currencyUnit" & vbcrlf
					sqlStr = sqlStr & " where ii.itemgubun='10'" & vbcrlf

					sqlStr = sqlStr & " drop table #tmp_exchangeRatecurrencyunitgroup"
					sqlStr = sqlStr & " drop table #tmp_shop_item"

					'response.write sqlStr &"<Br>"
				    dbget.Execute sqlStr
				end if
			end if
		end if
	end if

elseif mode = "exchangeRatedel" then
	sqlStr = "delete from db_item.dbo.tbl_exchangeRate" + VbCrlf
	sqlStr = sqlStr & " where idx='"&idx&"'"

	'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

else
	response.write "<script type='text/javascript'>"
	response.write "	alert(MODE 구분자가 지정되지 않았습니다.');"
	response.write "	location.replace('" & referer & "');"
	response.write "</script>"
end if
%>

<script type='text/javascript'>
	alert('OK');
	location.replace('<%=referer%>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<% session.codePage = 949 %>