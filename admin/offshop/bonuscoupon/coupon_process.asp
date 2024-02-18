<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  보너스 쿠폰
' History : 2011.05.12 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim menupos
dim idx,couponname,couponvalue,coupontype ,sqlStr , limityn ,shopid , lastupdateadminid
dim openfinishdate,isusing, etcstr ,isopenlistcoupon ,minbuyprice,startdate,expiredate
dim targetitemlist,targetbrandlist,couponmeaipprice , doublesaleyn,limitno ,validsitename ,i
dim exitemidlist, exbrandidlist
	menupos		= requestCheckVar(request("menupos"),10)
	idx			= requestCheckVar(request("idx"),10)
	couponname	= requestCheckVar(html2db(request("couponname")),10)
	couponvalue = requestCheckVar(request("couponvalue"),10)
	coupontype	= requestCheckVar(request("coupontype"),1)
	minbuyprice = requestCheckVar(request("minbuyprice"),30)
	startdate	= requestCheckVar(request("startdate"),20) & " " & requestCheckVar(request("startdatetime"),20)
	expiredate	= requestCheckVar(request("expiredate"),20) & " " & requestCheckVar(request("expiredatetime"),20)
	openfinishdate	= requestCheckVar(request("openfinishdate"),20) & " " & requestCheckVar(request("openfinishdatetime"),20)
	isusing			= requestCheckVar(request("isusing"),1)
	etcstr	= html2db(request("etcstr"))
	isopenlistcoupon = requestCheckVar(request("isopenlistcoupon"),1)
	validsitename = requestCheckVar(request("validsitename"),32)
	doublesaleyn = requestCheckVar(request("doublesaleyn"),1)
	limityn = requestCheckVar(request("limityn"),1)
	limitno = requestCheckVar(request("limitno"),10)
	shopid = requestCheckVar(request("shopid"),32)

	exitemidlist = requestCheckVar(Replace(request("exitemid"), " ", ""),512)
	exbrandidlist = requestCheckVar(Replace(request("exbrandid"), " ", ""),512)

	targetitemlist	= requestCheckVar(html2db(request("targetitemlist")),512)
	targetbrandlist	= requestCheckVar(html2db(request("targetbrandlist")),512)

	lastupdateadminid = session("ssBctId")

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if (Not IsNumeric(couponmeaipprice)) or (couponmeaipprice="") then couponmeaipprice=0

if (idx<>"" and idx<>"0") then
	if etcstr <> "" then
		if etcstr(contents_jupsu) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = "update [db_shop].dbo.tbl_shop_user_coupon_master" + VBCrlf
	sqlStr = sqlStr + " set couponname='" + couponname + "'" + VBCrlf
	sqlStr = sqlStr + " ,couponvalue=" + couponvalue + "" + VBCrlf
	sqlStr = sqlStr + " ,coupontype='" + coupontype + "'" + VBCrlf
	sqlStr = sqlStr + " ,minbuyprice=" + minbuyprice + "" + VBCrlf
	sqlStr = sqlStr + " ,startdate='" + startdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,expiredate='" + expiredate + "'" + VBCrlf
	sqlStr = sqlStr + " ,openfinishdate='" + openfinishdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " ,etcstr='" + etcstr + "'" + VBCrlf
	sqlStr = sqlStr + " ,isopenlistcoupon='" + isopenlistcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,doublesaleyn='" + doublesaleyn + "'" + VBCrlf
	sqlStr = sqlStr + " ,targetitemlist='" + targetitemlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,targetbrandlist='" + targetbrandlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,limityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " ,limitno='" + limitno + "'" + VBCrlf
	sqlStr = sqlStr + " ,couponmeaipprice=" + CStr(couponmeaipprice) + "" + VBCrlf

	sqlStr = sqlStr + " ,exitemidlist='" + exitemidlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,exbrandidlist='" + exbrandidlist + "'" + VBCrlf

	sqlStr = sqlStr + " ,lastupdateadminid='" + lastupdateadminid + "'" + VBCrlf

	IF (validsitename<>"") then
	    sqlStr = sqlStr + " ,validsitename='" + CStr(validsitename) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,validsitename=NULL" + VBCrlf
    END IF

	sqlStr = sqlStr + " where idx=" + CStr(idx)

	'// response.write sqlStr &"<Br>"
	'// response.end

	dbget.execute sqlStr

	shopid = split(shopid,",")

	'/기존내역 사용안함
	sqlstr = "update db_shop.dbo.tbl_shop_user_coupon_master_shoplist set" + vbcrlf
	sqlstr = sqlstr & " isusing='N'" + vbcrlf
	sqlstr = sqlstr & " where masteridx="&idx&"" + vbcrlf

	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	'/매장등록
	for i = 0 to ubound(shopid)
		'/등록
		sqlstr = "insert into db_shop.dbo.tbl_shop_user_coupon_master_shoplist" + vbcrlf
		sqlstr = sqlstr & " (shopid ,masteridx ,isusing, lastupdateadminid"
		sqlstr = sqlstr & " ) values (" + vbcrlf
		sqlstr = sqlstr & " '"&shopid(i)&"',"&idx&",'Y'" + vbcrlf
		sqlstr = sqlstr & " ,'"&lastupdateadminid&"'" + vbcrlf
		sqlstr = sqlstr & " )"

		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr
	next

else
	if etcstr <> "" then
		if etcstr(contents_jupsu) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	sqlStr = "insert into [db_shop].dbo.tbl_shop_user_coupon_master" + VBCrlf
	sqlStr = sqlStr + " (couponname, couponvalue, coupontype, minbuyprice, startdate" + VBCrlf
	sqlStr = sqlStr + " ,expiredate, openfinishdate, isusing, etcstr, isopenlistcoupon" + VBCrlf
	sqlStr = sqlStr + " ,doublesaleyn, targetitemlist, targetbrandlist, couponmeaipprice, limityn"
	sqlStr = sqlStr + " ,limitno, exitemidlist, exbrandidlist, lastupdateadminid, validsitename)" + VBCrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " '" + couponname + "'" + VBCrlf
	sqlStr = sqlStr + " ," + couponvalue + "" + VBCrlf
	sqlStr = sqlStr + " ,'" + coupontype + "'" + VBCrlf
	sqlStr = sqlStr + " ," + minbuyprice + "" + VBCrlf
	sqlStr = sqlStr + " ,'" + startdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + expiredate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + openfinishdate + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + etcstr + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + isopenlistcoupon + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + doublesaleyn + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + targetitemlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + targetbrandlist + "'" + VBCrlf
	sqlStr = sqlStr + " ," + Cstr(couponmeaipprice) + "" + VBCrlf
	sqlStr = sqlStr + " ,'" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + limitno + "'" + VBCrlf

	sqlStr = sqlStr + " ,'" + exitemidlist + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + exbrandidlist + "'" + VBCrlf

	sqlStr = sqlStr + " ,'" + lastupdateadminid + "'" + VBCrlf

	IF (validsitename<>"") then
	    sqlStr = sqlStr + " ,'" + CStr(validsitename) + "'" + VBCrlf
	ELSE
	    sqlStr = sqlStr + " ,NULL" + VBCrlf
    END IF

	sqlStr = sqlStr + " )"

	'response.write sqlStr &"<Br>"
	dbget.execute sqlStr

	shopid = split(shopid,",")

	'/매장등록
	for i = 0 to ubound(shopid)
		'/등록
		sqlstr = "insert into db_shop.dbo.tbl_shop_user_coupon_master_shoplist" + vbcrlf
		sqlstr = sqlstr & " (shopid ,masteridx ,isusing, lastupdateadminid)"
		sqlstr = sqlstr & " 	select top 1 '"&trim(shopid(i))&"' , idx ,'Y','"&lastupdateadminid&"'"
		sqlstr = sqlstr & " 	from db_shop.dbo.tbl_shop_user_coupon_master"
		sqlstr = sqlstr & " 	order by idx desc"

		'response.write sqlstr &"<Br>"
		dbget.execute sqlstr
	next
end if

%>

<script language="javascript">
	alert('저장되었습니다.');
	location.replace('/admin/offshop/bonuscoupon/couponlist.asp?menupos=<%=menupos%>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
