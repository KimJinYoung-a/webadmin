<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim referer
referer = request.ServerVariables("HTTP_REFERER")

dim mode, itemgubun, itemid, itemoption
mode        = RequestCheckVar(request("mode"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
itemoption  = RequestCheckVar(request("itemoption"),4)

dim orgsellprice, shopitemprice, discountsellprice
dim shopsuplycash, shopbuyprice, extbarcode
dim isusing, shopitemname, shopitemoptionname
dim vatinclude, makerid, centermwdiv
dim cd1, cd2, cd3
dim sqlStr

orgsellprice    = RequestCheckVar(request("orgsellprice"),9)
shopitemprice   = RequestCheckVar(request("shopitemprice"),9)
discountsellprice   = RequestCheckVar(request("discountsellprice"),9)
shopsuplycash       = RequestCheckVar(request("shopsuplycash"),9)
shopbuyprice        = RequestCheckVar(request("shopbuyprice"),9)
extbarcode      = RequestCheckVar(request("extbarcode"),32)
isusing         = RequestCheckVar(request("isusing"),1)
shopitemname    = html2db(RequestCheckVar(request("shopitemname"),128))
shopitemoptionname  = html2db(RequestCheckVar(request("shopitemoptionname"),128))
vatinclude          = RequestCheckVar(request("vatinclude"),1)
makerid         = RequestCheckVar(request("makerid"),32)
centermwdiv     = RequestCheckVar(request("centermwdiv"),1)

cd1 = RequestCheckVar(request("cd1"),3)
cd2 = RequestCheckVar(request("cd2"),3)
cd3 = RequestCheckVar(request("cd3"),3)

if Not IsNumeric(orgsellprice) then orgsellprice =0
if Not IsNumeric(discountsellprice) then discountsellprice =0
if Not IsNumeric(shopsuplycash) then shopsuplycash =0
if Not IsNumeric(shopbuyprice) then shopbuyprice =0

if CStr(orgsellprice)="0" then orgsellprice=shopitemprice

dim extbarcodeAlreadyExists, stockitemexists
extbarcodeAlreadyExists = false
stockitemexists         = false

''범용바코드
if (extbarcode<>"") then
	sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
	sqlStr = sqlStr + " where barcode='" + trim(extbarcode) + "'" + VbCrlf
	sqlStr = sqlStr + " and not ("
	sqlStr = sqlStr + " 	itemgubun='" + itemgubun + "'" + VbCrlf
	sqlStr = sqlStr + " 	and itemid=" + CStr(itemid) + "" + VbCrlf
	sqlStr = sqlStr + " 	and itemoption='" + CStr(itemoption) + "'" + VbCrlf
	sqlStr = sqlStr + " ) "

	rsget.Open sqlStr, dbget, 1
	if Not rsget.EOF then
		extbarcodeAlreadyExists = true
	end if
	rsget.close
end if

if extbarcodeAlreadyExists then
	response.write "<script>alert('" + extbarcode + " : 이미 사용중인 범용 바코드 입니다.');</script>"
	response.write "<script>location.replace('" + Cstr(referer) + "');</script>"
else
	sqlStr = "update [db_shop].[dbo].tbl_shop_item" + vbCrlf
	sqlStr = sqlStr + " set shopitemprice=" + CStr(shopitemprice) + "" + vbCrlf
	sqlStr = sqlStr + " , orgsellprice=" + CStr(orgsellprice) + "" + vbCrlf
	sqlStr = sqlStr + " , shopsuplycash=" + CStr(shopsuplycash) + "" + vbCrlf
	sqlStr = sqlStr + " , extbarcode='" + CStr(extbarcode) + "'" + vbCrlf
	sqlStr = sqlStr + " , isusing='" + CStr(isusing) + "'" + vbCrlf
	sqlStr = sqlStr + " , shopitemname='" + html2db(shopitemname) + "'" + vbCrlf
	sqlStr = sqlStr + " , shopitemoptionname='" + html2db(shopitemoptionname) + "'" + vbCrlf
	sqlStr = sqlStr + " , discountsellprice=" + CStr(discountsellprice) + "" + vbCrlf
	sqlStr = sqlStr + " , shopbuyprice=" + CStr(shopbuyprice) + "" + vbCrlf
	sqlStr = sqlStr + " , updt=getdate()" + vbCrlf

	if cd1<>"" then
		sqlStr = sqlStr + " , catecdl='" + cd1 + "'" + vbCrlf
	end if

	if cd2<>"" then
		sqlStr = sqlStr + " , catecdm='" + cd2 + "'" + vbCrlf
	end if

	if cd3<>"" then
		sqlStr = sqlStr + " , catecdn='" + cd3 + "'" + vbCrlf
	end if
    
    if centermwdiv<>"" then
        sqlStr = sqlStr + " , centermwdiv='" + centermwdiv + "'" + vbCrlf
	end if
	
	sqlStr = sqlStr + " , vatinclude='" + vatinclude + "'" + vbCrlf
	sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + vbCrlf
	sqlStr = sqlStr + " and shopitemid=" + CStr(itemid) + vbCrlf
	sqlStr = sqlStr + " and itemoption='" + itemoption + "'"

'response.write sqlStr
	dbget.Execute sqlStr



	''바코드 테이블 확인

	if (extbarcode<>"") then
		sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" + VbCrlf
		sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
		sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
		sqlStr = sqlStr + " and itemoption='" + itemoption + "'" + VbCrlf
		rsget.Open sqlStr,dbget,1
			stockitemexists = (not rsget.Eof)
		rsget.close

		if (stockitemexists) then
			sqlStr = " update [db_item].[dbo].tbl_item_option_stock" + VbCrlf
			sqlStr = sqlStr + " set barcode='" + trim(extbarcode) + "'" + VbCrlf
			sqlStr = sqlStr + " where itemgubun='" + itemgubun + "'" + VbCrlf
			sqlStr = sqlStr + " and itemid=" + CStr(itemid) + "" + VbCrlf
			sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'" + VbCrlf

			rsget.Open sqlStr,dbget,1
		else
			sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" + VbCrlf
			sqlStr = sqlStr + " (itemgubun,itemid,itemoption,barcode)" + VbCrlf
			sqlStr = sqlStr + " values("
			sqlStr = sqlStr + " '" + itemgubun + "'," + VbCrlf
			sqlStr = sqlStr + " " + CStr(itemid) + "," + VbCrlf
			sqlStr = sqlStr + " '" + itemoption + "'," + VbCrlf
			sqlStr = sqlStr + " '" + trim(extbarcode) + "'" + VbCrlf
			sqlStr = sqlStr + " )" + VbCrlf
			rsget.Open sqlStr,dbget,1
		end if
	end if

	
	response.write "<script>alert('수정 되었습니다.');</script>"
	response.write "<script>location.replace('" + Cstr(referer) + "');</script>"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
