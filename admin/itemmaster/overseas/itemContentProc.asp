<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
response.charSet = "utf-8"
'####################################################
' Description :  온라인 해외판매상품
' History : 2013.05.06 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->

<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
Dim vQuery, vItemID, vCountryCd, cOverSeas, vItemName, vItemContent, vItemCopy, vItemSource
dim vItemSize, vMakerName, vSourceArea, vRegUserID, useyn, i, sqlstr, keywords, areaCode11st
dim Siteisusing, sitename, multilangcnt
	vItemID = requestCheckVar(Request("itemid"),10)
	vCountryCd = requestCheckVar(Request("countrycd"),32)
	vRegUserID = session("ssBctId")

multilangcnt=0

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

If vItemID = "" OR vCountryCd = "" Then
	Response.Write "<script type='text/javascript'>alert('잘못된 경로입니다.');window.close()</script>"
	session.codePage = 949 : dbget.close() : Response.End
End IF
'response.write vCountryCd & "!!"
'response.end
vItemName = chrbyte(html2db(Request("itemname")),60,"")
vItemContent = chrbyte(Request("itemcontent"),800,"")
vItemCopy = chrbyte(html2db(Request("itemcopy")),250,"")
vItemSource = chrbyte(html2db(Request("itemsource")),128,"")
vItemSize = chrbyte(html2db(Request("itemsize")),128,"")
vMakerName = chrbyte(html2db(Request("makername")),64,"")
vSourceArea = chrbyte(html2db(Request("sourcearea")),128,"")
useyn = chrbyte(Request("useyn"),1,"")
keywords = chrbyte(Request("keywords"),512,"")
areaCode11st = chrbyte(Request("areaCode11st"),4,"")
Siteisusing = chrbyte(html2db(Request("Siteisusing")),1,"")
sitename = requestCheckVar(Request("sitename"),32)

if useyn = "" then useyn = "Y"

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

Dim vMakerIDX
If sitename = "CHNWEB" Then  ''중국사이트
	vQuery = ""
	vQuery = vQuery & " SELECT TOP 1 deliverytype FROM db_item.dbo.tbl_item with (nolock) WHERE itemid = '" & vItemID & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("deliverytype") = "7" Then		'업체배송
			Response.Write "<script>alert('업체배송 상품은 등록할 수 없습니다');window.close()</script>"
			session.codePage = 949 : dbget.close() : Response.End
		End If
	rsget.Close

	vQuery = "IF NOT EXISTS(select itemid from [db_item].[dbo].[tbl_kaffa_reg_item] where itemid = '" & vItemID & "') " & _
			 " BEGIN " & _
			 "INSERT INTO [db_item].[dbo].[tbl_kaffa_reg_item](itemid, regdate, kaffamakerid, reguserid, lastupdate, useyn) " & _
			 " VALUES(N'" & vItemID & "', getdate(), null, N'" & session("ssBctId") & "', getdate(), 'n')" & _
			 " END "
	dbget.execute vQuery
End If

'//언어팩 등록
if vCountryCd <> "X" then
	vQuery = ""
	vQuery = vQuery & " IF EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_multiLang] with (nolock) WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "') "
	vQuery = vQuery & " BEGIN "
	vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang] SET"
	vQuery = vQuery & "			itemname = N'" & vItemName & "', "
	vQuery = vQuery & "			itemcopy = N'" & vItemCopy & "', "
	vQuery = vQuery & "			itemContent = N'" & vItemContent & "', "
	vQuery = vQuery & "			itemsource = N'" & vItemSource & "', "
	vQuery = vQuery & "			itemsize = N'" & vItemSize & "', "
	vQuery = vQuery & "			makername = N'" & vMakerName & "', "
	vQuery = vQuery & "			sourcearea = N'" & vSourceArea & "', "
	vQuery = vQuery & "			useyn = N'" & useyn & "', "
	vQuery = vQuery & "			keywords = N'" & keywords & "', "
	vQuery = vQuery & "			areaCode11st = N'" & areaCode11st & "', "
	vQuery = vQuery & "			lastupdate = getdate()"
	vQuery = vQuery & "		WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' "
	vQuery = vQuery & "		UPDATE [db_temp].[dbo].[tbl_해외판매상품알바관리로그] SET reguserid = '" & vRegUserID & "', lastupdate = getdate() WHERE itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' "
	vQuery = vQuery & " END "
	vQuery = vQuery & " ELSE "
	vQuery = vQuery & " BEGIN "
	vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang](itemid, countryCd, itemname, itemcopy, itemContent, itemsource, itemsize, makername, sourcearea, useyn, keywords, areaCode11st) "
	vQuery = vQuery & "		VALUES(N'" & vItemID & "', N'" & vCountryCd & "', N'" & vItemName & "', N'" & vItemCopy & "', N'" & vItemContent & "' "
	vQuery = vQuery & "		, N'" & vItemSource & "', N'" & vItemSize & "', N'" & vMakerName & "', N'" & vSourceArea & "', N'" & useyn & "', N'" & keywords & "', N'" & areaCode11st & "') "
	vQuery = vQuery & "		INSERT INTO [db_temp].[dbo].[tbl_해외판매상품알바관리로그](itemid, countryCd, reguserid) VALUES('" & vItemID & "', '" & vCountryCd & "', '" & vRegUserID & "')  "
	vQuery = vQuery & " END "
	
	'response.write vQuery &"<Br>"
	dbget.execute vQuery

	If sitename = "WSLWEB" Then
		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = "insert into db_item.[dbo].[tbl_item_multiLang]" & vbcrlf
			vQuery = vQuery & " 	select" & vbcrlf
			vQuery = vQuery & " 	i.itemid, 'KR', i.itemname, isnull(ic.designercomment,'') as designercomment, '', ic.itemsource, ic.itemsize, ic.sourcearea, c.socname_kor" & vbcrlf
			vQuery = vQuery & " 	, 'Y', getdate(), getdate(), ic.keywords, ''" & vbcrlf
			vQuery = vQuery & " 	from db_item.dbo.tbl_item i with (nolock)" & vbcrlf
			vQuery = vQuery & " 	left join db_user.dbo.tbl_user_c c with (nolock)" & vbcrlf
			vQuery = vQuery & " 		on i.makerid=c.userid" & vbcrlf
			vQuery = vQuery & " 	left join db_item.[dbo].[tbl_item_multiLang] ml with (nolock)" & vbcrlf
			vQuery = vQuery & " 		on i.itemid=ml.itemid" & vbcrlf
			vQuery = vQuery & " 		and ml.countryCd='KR'" & vbcrlf
			vQuery = vQuery & " 	left join db_item.dbo.tbl_item_Contents ic with (nolock)" & vbcrlf
			vQuery = vQuery & " 		on i.itemid = ic.itemid" & vbcrlf
			vQuery = vQuery & " 	where ml.itemid is null" & vbcrlf
			vQuery = vQuery & " 	and i.itemid = " & vItemID & "" & vbcrlf
			vQuery = vQuery & " 	" & vbcrlf

			'response.write vQuery &"<Br>"
			dbget.execute vQuery
		end if
	end if
end if

'//상품의 언어팩의 갯수를 카운트 한다.
sqlstr = "select count(ml.itemid) as multilangcnt"
sqlstr = sqlstr & "	from [db_item].[dbo].[tbl_item_multiLang] ml with (nolock)"
sqlstr = sqlstr & "	where ml.itemid="& vItemID &""

'response.write sqlstr & "<br>"
rsget.CursorLocation = adUseClient
rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.EOF  then
		multilangcnt = rsget("multilangcnt")
	End If
rsget.Close

'//사이트 등록
sqlstr = " if exists(" + vbcrlf
sqlstr = sqlstr & "		select top 1 *" + vbcrlf
sqlstr = sqlstr & "		from db_item.dbo.tbl_item_multiSite_regItem with (nolock)" + vbcrlf
sqlstr = sqlstr & "		where itemid = '" & vItemID & "' AND sitename = '" & sitename & "'" + vbcrlf
sqlstr = sqlstr & "	)" + vbcrlf
sqlstr = sqlstr & "		update db_item.dbo.tbl_item_multiSite_regItem set" + vbcrlf
sqlstr = sqlstr & "		isusing=N'"& Siteisusing &"'" + vbcrlf
sqlstr = sqlstr & "		,lastupdate = getdate()" + vbcrlf
sqlstr = sqlstr & "		,lastuserid = N'"& vRegUserID &"'" + vbcrlf
sqlstr = sqlstr & "		,multilangcnt="& multilangcnt &"" + vbcrlf
sqlstr = sqlstr & "		WHERE itemid = '" & vItemID & "' AND sitename = '" & sitename & "' " + vbcrlf
sqlstr = sqlstr & " else" + vbcrlf
sqlstr = sqlstr & " 	insert into db_item.dbo.tbl_item_multiSite_regItem(" + vbcrlf
sqlstr = sqlstr & " 	itemid, sitename, isusing, multilangcnt, regdate, reguserid, lastupdate, lastuserid" + vbcrlf
sqlstr = sqlstr & "		) values (" + vbcrlf
sqlstr = sqlstr & "		N'" & vItemID & "', N'" & sitename & "', N'" & Siteisusing & "', "& multilangcnt &", getdate(), N'" & vRegUserID & "'" + vbcrlf
sqlstr = sqlstr & "		, getdate(), N'" & vRegUserID & "' " + vbcrlf
sqlstr = sqlstr & "		)"

'response.write sqlstr &"<Br>"
dbget.execute sqlstr

'가격저장
dim pricecount, currencyUnit, orgprice, wonprice, exchangeRate
	pricecount = Request("pricecount")

if pricecount > 0 then
	for i =0 to pricecount-1
		currencyUnit 	= Request("currencyUnit"&i)
		orgprice 	= Request("orgprice"&i)
		wonprice 	= Request("wonprice"&i)
        exchangeRate= Request("exchangeRate"&i)  ''저장당시 환율

		sqlstr = " if exists(" + vbcrlf
		sqlstr = sqlstr & "		select *" + vbcrlf
		sqlstr = sqlstr & "		from db_item.dbo.tbl_item_multiLang_price with (nolock)" + vbcrlf
		sqlstr = sqlstr & "		where itemid=N'" & vItemID & "'" + vbcrlf
		sqlstr = sqlstr & "		and sitename=N'" & sitename & "'" + vbcrlf
		sqlstr = sqlstr & "		and currencyUnit=N'"& currencyUnit &"'" + vbcrlf
		sqlstr = sqlstr & "	)" + vbcrlf
		sqlstr = sqlstr & "		update P " + vbcrlf
		sqlstr = sqlstr & "		set orgprice=N'"& orgprice &"'" + vbcrlf
		sqlstr = sqlstr & "		,wonprice=N'"& wonprice &"'" + vbcrlf
		sqlstr = sqlstr & "		,lastexchangeRate=(CASE WHEN P.orgprice<>N'"&orgprice&"' or P.wonprice<>N'"&wonprice&"' THEN N'"&exchangeRate&"' ELSE P.lastexchangeRate END)" + vbcrlf  ''금액이 바뀐경우만
		sqlstr = sqlstr & "		,lastupdate=getdate()" + vbcrlf
		sqlstr = sqlstr & "		,lastuserid=N'"& vRegUserID &"'" + vbcrlf
		sqlstr = sqlstr & "		From db_item.dbo.tbl_item_multiLang_price P"+ vbcrlf
		sqlstr = sqlstr & "		where P.itemid=N'" & vItemID & "'" + vbcrlf
		sqlstr = sqlstr & "		and P.sitename=N'" & sitename & "'" + vbcrlf
		sqlstr = sqlstr & "		and P.currencyUnit=N'"& currencyUnit &"'" + vbcrlf
		sqlstr = sqlstr & "	else" + vbcrlf
		sqlstr = sqlstr & "		insert into db_item.dbo.tbl_item_multiLang_price(" + vbcrlf
		sqlstr = sqlstr & "		sitename, itemid, currencyUnit, orgprice, wonprice, lastexchangeRate, regdate, lastupdate, reguserid, lastuserid" + vbcrlf
		sqlstr = sqlstr & "		) values (" + vbcrlf
		sqlstr = sqlstr & "		N'" & sitename & "', N'" & vItemID & "', N'"& currencyUnit &"', N'"& orgprice &"', N'"& wonprice &"', N'"& exchangeRate &"'" + vbcrlf
		sqlstr = sqlstr & "		, getdate(), getdate(), N'" & vRegUserID & "', N'" & vRegUserID & "'" + vbcrlf
		sqlstr = sqlstr & "		)"

		'rw sqlstr & "<Br>"
		dbget.execute sqlstr
	next
end if

'### 옵션저장.
Dim vOptionCount, vItemOption, vOptionTypeName, vOptionName, vOptIsUsing
vOptionCount = chrbyte(Request("optioncount"),3,"")

if vCountryCd <> "X" then
	vQuery = ""
	For i=0 To vOptionCount-1
		vItemOption 	= Request("itemoption"&i)
		vOptionTypeName	= html2db(Request("optiontypename"&i))
		vOptionName		= html2db(Request("optionname"&i))
		vOptIsUsing		= Request("optisusing"&i)
	
		if vItemOption<>"0000" then
			vQuery = " IF EXISTS(SELECT itemoption FROM [db_item].[dbo].[tbl_item_multiLang_option] with (nolock) WHERE" & vbCrLf
			vQuery = vQuery & " itemid = '" & vItemID & "' AND countryCd = '" & vCountryCd & "' AND itemoption = '" & vItemOption & "') " & vbCrLf
			vQuery = vQuery & " BEGIN " & vbCrLf
			vQuery = vQuery & "		UPDATE [db_item].[dbo].[tbl_item_multiLang_option]" & vbCrLf
			vQuery = vQuery & "		SET optiontypename = N'" & vOptionTypeName & "'" & vbCrLf
			vQuery = vQuery & "		, optionname = N'" & vOptionName & "'" & vbCrLf
			vQuery = vQuery & "		, isusing = N'" & vOptIsUsing & "'" & vbCrLf
			vQuery = vQuery & "		, lastupdate=getdate()" & vbCrLf
			vQuery = vQuery & "		WHERE countryCD='"&vCountryCd&"' And itemid = '" & vItemID & "' AND itemoption = '" & vItemOption & "'" & vbCrLf
			vQuery = vQuery & " END " & vbCrLf
			vQuery = vQuery & " ELSE " & vbCrLf
			vQuery = vQuery & " BEGIN " & vbCrLf
			vQuery = vQuery & "		INSERT INTO [db_item].[dbo].[tbl_item_multiLang_option](itemid, countryCd, itemoption, optiontypename, optionname, isusing, regdate, lastupdate)" & vbCrLf
			vQuery = vQuery & "		VALUES(" & vbCrLf
			vQuery = vQuery & "		N'" & vItemID & "', N'" & vCountryCd & "', N'" & vItemOption & "', N'" & vOptionTypeName & "', N'" & vOptionName & "'" & vbCrLf
			vQuery = vQuery & "		, N'" & vOptIsUsing & "', getdate(), getdate()" & vbCrLf
			vQuery = vQuery & "		) " & vbCrLf
			vQuery = vQuery & " END " & vbCrLf

			'response.write vQuery & "<br>"
			dbget.execute vQuery

			vItemOption		= ""
			vOptionTypeName	= ""
			vOptionName		= ""
			vOptIsUsing		= ""
		end if
	Next

	' 최초 등록시에 옵션 저장을 안하기 때문에. 해당 언어팩 옵션이 없을경우 저장함.
	vQuery = " insert into db_item.[dbo].[tbl_item_multiLang_option]" & vbCrLf
	vQuery = vQuery & "		select" & vbCrLf
	vQuery = vQuery & "		i.itemid, '"& vCountryCd &"', o.itemoption, 'Y', o.optionTypeName, o.optionname, getdate(), getdate()" & vbCrLf
	vQuery = vQuery & "		from db_item.dbo.tbl_item i with (nolock)" & vbCrLf
	vQuery = vQuery & "		left join db_item.dbo.tbl_item_option o with (nolock)" & vbCrLf
	vQuery = vQuery & "			on i.itemid = o.itemid" & vbCrLf
	vQuery = vQuery & "		left join db_item.[dbo].[tbl_item_multiLang_option] mo with (nolock)" & vbCrLf
	vQuery = vQuery & "			on i.itemid=mo.itemid" & vbCrLf
	vQuery = vQuery & "			and o.itemoption = mo.itemoption" & vbCrLf
	vQuery = vQuery & "			and mo.countryCd='"& vCountryCd &"'" & vbCrLf
	vQuery = vQuery & "		where mo.itemid is null" & vbCrLf
	vQuery = vQuery & "		and o.itemoption is not null" & vbCrLf		'옵션없음은 넣지 안음.
	vQuery = vQuery & "		and i.itemid = " & vItemID & "" & vbCrLf

	'response.write vQuery & "<br>"
	dbget.execute vQuery

	If sitename = "WSLWEB" Then
		' 한글 언어팩이 없는 경우 디폴트로 꽂음.
		if ucase(vCountryCd) <> "KR" then
			vQuery = " insert into db_item.[dbo].[tbl_item_multiLang_option]" & vbCrLf
			vQuery = vQuery & "		select" & vbCrLf
			vQuery = vQuery & "		i.itemid, 'KR', o.itemoption, 'Y', o.optionTypeName, o.optionname, getdate(), getdate()" & vbCrLf
			vQuery = vQuery & "		from db_item.dbo.tbl_item i with (nolock)" & vbCrLf
			vQuery = vQuery & "		left join db_item.dbo.tbl_item_option o with (nolock)" & vbCrLf
			vQuery = vQuery & "			on i.itemid = o.itemid" & vbCrLf
			vQuery = vQuery & "		left join db_item.[dbo].[tbl_item_multiLang_option] mo with (nolock)" & vbCrLf
			vQuery = vQuery & "			on i.itemid=mo.itemid" & vbCrLf
			vQuery = vQuery & "			and o.itemoption = mo.itemoption" & vbCrLf
			vQuery = vQuery & "			and mo.countryCd='KR'" & vbCrLf
			vQuery = vQuery & "		where mo.itemid is null" & vbCrLf
			vQuery = vQuery & "		and o.itemoption is not null" & vbCrLf		'옵션없음은 넣지 안음.
			vQuery = vQuery & "		and i.itemid = " & vItemID & "" & vbCrLf

			'response.write vQuery & "<br>"
			dbget.execute vQuery
		end if
	end if

'response.end
end if

if Siteisusing="Y" then
	vQuery = "update [db_item].[dbo].tbl_item" + VbCrlf
	vQuery = vQuery & " set deliverOverseas='Y'" + VbCrlf
	vQuery = vQuery & " , lastupdate = getdate() where" + VbCrlf
	vQuery = vQuery & " itemid="& vItemID &"" + VbCrlf
	
	'response.write sqlStr & "<br>"
	dbget.execute vQuery
end if

response.write "<script type='text/javascript'>"
response.write "	alert('저장되었습니다.');"
session.codePage = 949
response.write "	opener.document.location.reload();"
response.write "	window.close();"
response.write "</script>"
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>