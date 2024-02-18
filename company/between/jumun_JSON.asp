<%@  codepage="65001" language="VBScript" %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp" -->
<!-- #include virtual="/lib/util/JSON_UTIL_0.1.1.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function queryStrings
	Dim strSql
	strSql = ""
	strSql = strSql & " select "
	strSql = strSql & " m.orderserial, d.itemid, u.btwUserCd, m.userid, d.orgitemcost, d.itemno*d.itemcost as sellprice "
	strSql = strSql & " , m.regdate, m.canceldate, m.beadaldate, m.reqzipaddr, d.itemoption, o.optionname" & vbcrlf
	strSql = strSql & " , isnull(dd.requiredetailUTF8,d.requiredetail) as requiredetail"
	strSql = strSql & " ,Case WHEN m.accountdiv = '7' THEN '무통장' "
	strSql = strSql & "		WHEN m.accountdiv = '14' THEN '편의점결제' "
	strSql = strSql & "		WHEN m.accountdiv = '100' THEN '신용카드' "
	strSql = strSql & "		WHEN m.accountdiv = '20' THEN '실시간이체' "
	strSql = strSql & "		WHEN m.accountdiv = '30' THEN '포인트' "
	strSql = strSql & " 	WHEN m.accountdiv = '50' THEN '외부몰' "
	strSql = strSql & " 	WHEN m.accountdiv = '80' THEN 'All@카드' "
	strSql = strSql & " 	WHEN m.accountdiv = '90' THEN '상품권결제' "
	strSql = strSql & " 	WHEN m.accountdiv = '110' THEN 'OK+신용' "
	strSql = strSql & " 	WHEN m.accountdiv = '400' THEN '핸드폰결제' end as JumunMethodName "
	strSql = strSql & " ,d.itemno, m.buyname, m.reqname, m.comment, i.itemscore, i.regdate "
	strSql = strSql & " ,'http://webimage.10x10.co.kr/image/basic/' +  "
	strSql = strSql & " Case When len(convert(varchar(5), i.itemid / 10000)) < 2 Then '0'+convert(varchar(5), i.itemid / 10000) "
	strSql = strSql & " else convert(varchar(5), i.itemid / 10000) end "
	strSql = strSql & " + '/' + i.basicimage as basicimage "
	strSql = strSql & " ,i.brandname "
	strSql = strSql & " ,CASE WHEN i.deliverytype = '1' THEN '텐바이텐배송' "
	strSql = strSql & " 	WHEN i.deliverytype = '2' or i.deliverytype = '5' THEN '업체무료배송' "
	strSql = strSql & " 	WHEN i.deliverytype = '4' THEN '텐바이텐무료배송' "
	strSql = strSql & " 	WHEN i.deliverytype = '6' THEN '현장수령상품' "
	strSql = strSql & " 	WHEN i.deliverytype = '7' THEN '업체착불배송' "
	strSql = strSql & " 	WHEN i.deliverytype = '9' THEN '업체조건배송' "
	strSql = strSql & " Else '텐바이텐배송' end as DeliveryName "
	strSql = strSql & " ,Case When bi.catecode = '102' Then '소품/취미' "
	strSql = strSql & " 	When bi.catecode = '103' Then '디지털' "
	strSql = strSql & " 	When bi.catecode = '104' Then '키친/푸드' "
	strSql = strSql & " 	When bi.catecode = '105' Then '패션' "
	strSql = strSql & " 	When bi.catecode = '106' Then '뷰티' "
	strSql = strSql & " 	When bi.catecode = '107' Then 'SALE' end as catename "
	strSql = strSql & " ,CASE WHEN i.limityn = 'Y' THEN i.limitno - i.limitsold "
	strSql = strSql & " Else '9999' End as itemsu "
	strSql = strSql & " from db_order.dbo.tbl_order_master as m  "
	strSql = strSql & " join db_order.dbo.tbl_order_detail as d on m.orderserial = d.orderserial "
	strSql = strSql & " join db_etcmall.dbo.tbl_between_userInfo as u on u.usersn = m.rduserid "
	strSql = strSql & " join db_item.dbo.tbl_item as i on d.itemid = i.itemid "
	strSql = strSql & " left join db_item.dbo.tbl_item_option as o on d.itemid = o.itemid and d.itemoption = o.itemoption "
	strSql = strSql & " left join [192.168.0.78].db_outmall.dbo.tbl_between_cate_item as bi on i.itemid = bi.itemid and bi.catecode <> '101' "
	strSql = strSql & " LEFT JOIN db_order.dbo.tbl_order_require dd" & vbcrlf
	strSql = strSql & "    	ON d.idx = dd.detailidx" & vbcrlf
	strSql = strSql & " where m.beadaldiv='8' "
	strSql = strSql & " and m.sitename='10x10' "
	strSql = strSql & " and m.rdsite='betweenshop' "
	strSql = strSql & " and d.itemid <> '0' "
	strSql = strSql & " order by m.orderserial asc "
	queryStrings = strSql
End Function

Dim appPath : appPath = server.mappath("/company/between/") + "\"
Dim FileName: FileName = "/json_Order.txt"
Dim fso, tFile, FlushData

Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
		'QueryToJSON(dbget, queryStrings).Flush
		FlushData = QueryToJSON(dbget, queryStrings).jsString
		tFile.writeLine FlushData
		tFile.Close
	Set tFile = Nothing
Set fso = Nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->