<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : OFFSHOP 정보
' History : 2012.07.30 한용민 생성
'           2014.06.25 허진원; 비번 SHA256 추가
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbhelper.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/order/clsems_serviceArea.asp" -->

<%
dim mode ,sqlStr ,i ,menupos
dim userpass ,shopname ,shopphone ,shopCountryCode ,shopzipcode ,shopaddr1 ,shopaddr2 ,manname
dim manhp ,manphone ,manemail ,shopdiv ,isusing ,stockbasedate ,shopsocno ,shopceoname
dim vieworder ,currencyUnit ,exchangeRate ,multipleRate ,decimalPointLen ,decimalPointCut ,pyeong
dim ismobileusing ,mobileshopname ,mobileworkhour ,mobileclosedate ,mobiletel ,mobileaddr
dim mobilebysubway ,mobilebybus ,mobilelatitude ,mobilelongitude , shopid ,groupid ,admindisplang
dim company_no, loginsite, countrylangcd, currencyUnit_Pos
dim ctropen, viewsort, engName, shopfax, engAddress
	countrylangcd   = RequestCheckVar(request("countrylangcd"),32)
	loginsite   = RequestCheckVar(request("loginsite"),32)
	groupid   = RequestCheckVar(request("groupid"),6)
	menupos   = RequestCheckVar(request("menupos"),10)
	shopid = RequestCheckVar(request("shopid"),32)
	mode   = RequestCheckVar(request("mode"),32)
	userpass    = request("userpass")
	shopname    = html2db(request("shopname"))
	shopphone   = request("shopphone")
	shopzipcode = request("shopzipcode")
	shopaddr1   = html2db(request("shopaddr1"))
	shopaddr2   = html2db(request("shopaddr2"))
	manname     = html2db(request("manname"))
	manhp       = request("manhp")
	manphone    = request("manphone")
	manemail    = html2db(request("manemail"))
	shopdiv     = request("shopdiv")
	isusing     = request("isusing")
	stockbasedate = request("stockbasedate")
	shopsocno   = request("shopsocno")
	shopceoname = html2db(request("shopceoname"))
	vieworder	= request("vieworder")
	currencyUnit = request("currencyUnit")
	currencyUnit_Pos = request("currencyUnit_Pos")
	multipleRate = request("multipleRate")
	pyeong    = request("pyeong")
	shopCountryCode = request("shopCountryCode")
    decimalPointLen = request("decimalPointLen")
    decimalPointCut = request("decimalPointCut")
    exchangeRate    = request("exchangeRate")
	ismobileusing    	= request("ismobileusing")
	mobileshopname    	= html2db(request("mobileshopname"))
	mobileworkhour    	= html2db(request("mobileworkhour"))
	mobileclosedate    	= html2db(request("mobileclosedate"))
	mobiletel    		= html2db(request("mobiletel"))
	mobileaddr    		= html2db(request("mobileaddr"))
	mobilebysubway    	= html2db(request("mobilebysubway"))
	mobilebybus    		= html2db(request("mobilebybus"))
	mobilelatitude    	= request("mobilelatitude")
	mobilelongitude    	= request("mobilelongitude")
	admindisplang    	= request("admindisplang")
	company_no    		= request("company_no")
   ctropen	= RequestCheckVar(request("ctropen"),1)
   viewsort	= RequestCheckVar(request("viewsort"),2)
   engName	= RequestCheckVar(request("engName"),32)
   shopfax	= RequestCheckVar(request("shopfax"),16)
    engAddress	= RequestCheckVar(request("engAddress"),128)
'' response.write mode &"<br>"
'' response.write groupid &"<br>"
'' response.end

'/신규등록
if mode = "new" then

	sqlStr = "select top 1 * from db_shop.dbo.tbl_shop_user where userid = '"&shopid&"'"

	'response.write sqlsrr & "<br>"
	rsget.Open sqlStr,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>alert('매장아이디가 이미 존재 합니다[1].'); history.go(-1);</script>"
		dbget.close()	:	response.End
	end if
	rsget.close

	sqlStr = "select top 1 * from [db_partner].[dbo].tbl_partner where id = '"&shopid&"'"

	'response.write sqlsrr & "<br>"
	rsget.Open sqlStr,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>alert('매장아이디가 이미 존재 합니다[2].'); history.go(-1);</script>"
		dbget.close()	:	response.End
	end if
	rsget.close

	sqlStr = "select top 1 * from [db_user].[dbo].tbl_user_c where userid = '"&shopid&"'"

	'response.write sqlsrr & "<br>"
	rsget.Open sqlStr,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>alert('매장아이디가 이미 존재 합니다[3].'); history.go(-1);</script>"
		dbget.close()	:	response.End
	end if
	rsget.close

	sqlStr = "select top 1 * from [db_user].[dbo].tbl_logindata where userid = '"&shopid&"'"

	'response.write sqlsrr & "<br>"
	rsget.Open sqlStr,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>alert('매장아이디가 이미 존재 합니다[4].'); history.go(-1);</script>"
		dbget.close()	:	response.End
	end if
	rsget.close

	sqlStr = "insert into db_shop.dbo.tbl_shop_user(" + VbCrlf
	sqlStr = sqlStr + " userid ,userpass, shopname ,shopphone ,shopCountryCode" + VbCrlf		' Enc_shoppass
	sqlStr = sqlStr + " ,shopzipcode ,shopaddr1 ,shopaddr2,manname ,manhp" + VbCrlf
	sqlStr = sqlStr + " ,manphone ,manemail ,shopdiv ,isusing ,stockbasedate" + VbCrlf
	sqlStr = sqlStr + " ,shopsocno ,shopceoname ,vieworder ,currencyUnit, currencyUnit_Pos ,exchangeRate" + VbCrlf
	sqlStr = sqlStr + " ,multipleRate ,decimalPointLen ,decimalPointCut ,pyeong ,ismobileusing" + VbCrlf
	sqlStr = sqlStr + " ,mobileshopname ,mobileworkhour ,mobileclosedate ,mobiletel ,mobileaddr" + VbCrlf
	sqlStr = sqlStr + " ,mobilebysubway ,mobilebybus ,mobilelatitude, admindisplang, loginsite, countrylangcd, ctropen, viewsort, engName, shopfax, engAddress" + VbCrlf
	sqlStr = sqlStr + " ) values (" + VbCrlf
	sqlStr = sqlStr + " '" + shopid + "','','" + shopname + "','" + shopphone + "','" + shopCountryCode + "'" + VbCrlf		' ,'" + md5(userpass) + "'
	sqlStr = sqlStr + " ,'" + shopzipcode + "','" + shopaddr1 + "','" + shopaddr2 + "','" + manname + "','" + manhp + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + manphone + "','" + manemail + "','" + shopdiv + "','" + isusing + "','" + stockbasedate + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + shopsocno + "','" + shopceoname + "','" + vieworder + "','" + currencyUnit + "','" + currencyUnit_POS + "','" + exchangeRate + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + multipleRate + "','" + decimalPointLen + "','" + decimalPointCut + "','" + pyeong + "','" + CStr(ismobileusing) + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + CStr(mobileshopname) + "','" + CStr(mobileworkhour) + "','" + CStr(mobileclosedate) + "','" + CStr(mobiletel) + "','" + CStr(mobileaddr) + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + CStr(mobilebysubway) + "','" + CStr(mobilebybus) + "','" + CStr(mobilelatitude) + "' ,'"&admindisplang&"','"&loginsite&"'" + VbCrlf
	sqlStr = sqlStr + " ,'" + countrylangcd + "', '"&ctropen&+ "', '"&viewsort&"', '"&engName&"', '"&shopfax&"', '"&engAddress&"'" + VbCrlf
	sqlStr = sqlStr + " )"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = "insert into [db_partner].[dbo].tbl_partner" + VbCrlf
	sqlStr = sqlStr + " (id, Enc_password, Enc_password64, company_name, userdiv, isusing, part_sn, posit_sn, level_sn, groupid, company_no)" + VbCrlf
	sqlStr = sqlStr + " values("

	'//직영매장
	if shopdiv = "1" then
		sqlStr = sqlStr + " '" + shopid + "','','" + SHA256(MD5(userpass)) + "','" + shopname + "','501','" + isusing + "',4,15,5, '"&groupid&"', '"&company_no&"'" + VbCrlf

	'//기타
	else
		sqlStr = sqlStr + " '" + shopid + "','','" + SHA256(MD5(userpass)) + "','" + shopname + "','503','" + isusing + "',5,15,6, '"&groupid&"','"&company_no&"'" + VbCrlf
	end if

	sqlStr = sqlStr + ")"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = "insert into [db_user].[dbo].tbl_user_c" + VbCrlf
	sqlStr = sqlStr + " (userid, socno, socname, isusing, isb2b, userdiv, maeipdiv,defaultmargine" + VbCrlf
	sqlStr = sqlStr + " , socname_kor, coname,prtidx, streetusing, extstreetusing, specialbrand, regdate)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + shopid + "','000-00-00000','" + shopname + "','" + isusing + "','N',21,'M',35" + VbCrlf
	sqlStr = sqlStr + " ,'" + shopname + "','" + shopname + "',9999,'N','N','N',getdate()"
	sqlStr = sqlStr + " )"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = "insert into [db_user].[dbo].tbl_logindata" + VbCrlf
	sqlStr = sqlStr + " (userid, userpass, userdiv, Enc_userpass, Enc_userpass64)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + shopid + "','',21,'','"&SHA256(Md5(userpass))&"'" + VbCrlf
	sqlStr = sqlStr + " )"

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	response.write "<script>alert('OK'); opener.location.reload(); self.close();</script>"

'/수정
elseif mode = "edit" then
    '' userpass='" + userpass + "'" + VbCrlf 주석처리

	sqlStr = "update [db_shop].[dbo].tbl_shop_user" + VbCrlf
	sqlStr = sqlStr + " set shopname='" + shopname + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopphone='" + shopphone + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopCountryCode='" + shopCountryCode + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopzipcode='" + shopzipcode + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopaddr1='" + shopaddr1 + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopaddr2='" + shopaddr2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,manname='" + manname + "'" + VbCrlf
	sqlStr = sqlStr + " ,manhp='" + manhp + "'" + VbCrlf
	sqlStr = sqlStr + " ,manphone='" + manphone + "'" + VbCrlf
	sqlStr = sqlStr + " ,manemail='" + manemail + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopdiv='" + shopdiv + "'" + VbCrlf
	sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
	sqlStr = sqlStr + " ,stockbasedate='" + stockbasedate + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopsocno='" + shopsocno + "'" + VbCrlf
	sqlStr = sqlStr + " ,shopceoname='" + shopceoname + "'" + VbCrlf
	sqlStr = sqlStr + " ,vieworder='" + vieworder + "'" + VbCrlf
	sqlStr = sqlStr + " ,currencyUnit='" + currencyUnit + "'" + VbCrlf
	sqlStr = sqlStr + " ,currencyUnit_POS='" + currencyUnit_POS + "'" + VbCrlf
	sqlStr = sqlStr + " ,exchangeRate=" + exchangeRate + "" + VbCrlf
	sqlStr = sqlStr + " ,multipleRate='" + multipleRate + "'" + VbCrlf
	sqlStr = sqlStr + " ,decimalPointLen=" + decimalPointLen + "" + VbCrlf
	sqlStr = sqlStr + " ,decimalPointCut=" + decimalPointCut + "" + VbCrlf
	sqlStr = sqlStr + " ,pyeong=" + pyeong + "" + VbCrlf
	sqlStr = sqlStr + " ,ismobileusing='" + CStr(ismobileusing) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileshopname='" + CStr(mobileshopname) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileworkhour='" + CStr(mobileworkhour) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileclosedate='" + CStr(mobileclosedate) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobiletel='" + CStr(mobiletel) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobileaddr='" + CStr(mobileaddr) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilebysubway='" + CStr(mobilebysubway) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilebybus='" + CStr(mobilebybus) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilelatitude='" + CStr(mobilelatitude) + "' " + VbCrlf
	sqlStr = sqlStr + " ,mobilelongitude='" + CStr(mobilelongitude) + "' " + VbCrlf
	sqlStr = sqlStr + " ,admindisplang = '"&admindisplang&"'" + VbCrlf
	sqlStr = sqlStr + " ,loginsite = '"&loginsite&"'" + VbCrlf
	sqlStr = sqlStr + " ,countrylangcd = '"&countrylangcd&"'" + VbCrlf
	sqlStr = sqlStr + " ,srLastupdate=getdate()" + VbCrlf
	sqlStr = sqlStr + " ,ctropen='"&ctropen&"'" + VbCrlf
	sqlStr = sqlStr + " ,viewsort='"&viewsort&"'" + VbCrlf
	sqlStr = sqlStr + " ,engName='"&engName&"'" + VbCrlf
	sqlStr = sqlStr + " ,shopfax='"&shopfax&"'" + VbCrlf
	sqlStr = sqlStr + " ,engAddress='"&engAddress&"'" + VbCrlf
	sqlStr = sqlStr + " where userid='" + shopid + "'" + VbCrlf

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = " update [db_partner].[dbo].tbl_partner "
	sqlStr = sqlStr + " set groupid = '" + CStr(groupid) + "' "
	sqlStr = sqlStr + " where id = '" + CStr(shopid) + "' and IsNull(groupid, '') <> '" + CStr(groupid) + "' "
	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = " update [db_partner].[dbo].tbl_partner "
	sqlStr = sqlStr + " set company_no = '" + CStr(company_no) + "' "
	sqlStr = sqlStr + " where id = '" + CStr(shopid) + "' and IsNull(company_no, '') <> '" + CStr(company_no) + "' "
	''response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	response.write "<script>alert('OK'); opener.location.reload(); self.close();</script>"
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
