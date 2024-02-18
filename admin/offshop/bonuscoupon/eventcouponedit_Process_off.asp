<%@ language=vbscript %>
<% option explicit %>
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
dim useridarr,couponvalue,coupontype
dim couponname,minbuyprice,startdate,expiredate
dim targetitemlist, couponmeaipprice, reguserid
dim i
useridarr = request("useridarr")
couponvalue = requestCheckVar(request("couponvalue"),10)
coupontype = requestCheckVar(request("coupontype"),1)
couponname = requestCheckVar(request("couponname"),128)
minbuyprice = requestCheckVar(request("minbuyprice"),10)
startdate = requestCheckVar(request("startdate"),30)
expiredate = requestCheckVar(request("expiredate"),30)
targetitemlist = requestCheckVar(request("targetitemlist"),512)
couponmeaipprice = requestCheckVar(request("couponmeaipprice"),30)
reguserid  =  session("ssBctId")

if (Not IsNumeric(couponmeaipprice)) then couponmeaipprice=0
if (couponmeaipprice="") then couponmeaipprice=0

if Right(useridarr,1)="," then
	useridarr = Left(useridarr,Len(useridarr)-1)
end if

useridarr = replace(useridarr," ","")
useridarr = replace(useridarr,",","','")

dim sqlStr

if (useridarr <> "") then
	sqlStr = "insert into db_shop.dbo.tbl_shop_user_coupon"
	sqlStr = sqlStr + " (masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)"
	sqlStr = sqlStr + " select 0,userid,'" + Cstr(coupontype) + "','" + Cstr(couponvalue) + "','" + couponname + "','"  + Cstr(minbuyprice) + "',"
	sqlStr = sqlStr +  "'"  + startdate + "','"   + expiredate + "','" + targetitemlist + "'," + CStr(couponmeaipprice) + ",'" + reguserid + "'"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
	sqlStr = sqlStr + " where userid in ('" + useridarr + "')"

	rsget.Open sqlStr,dbget,1
end if

'dim refer
'refer = request.ServerVariables("HTTP_REFERER")
%>

<script type='text/javascript'>
	alert('저장 되었습니다.');
	location.replace('eventcouponlist_off.asp');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->