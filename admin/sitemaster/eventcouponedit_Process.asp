<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim useridarr,couponvalue,coupontype
dim couponname,minbuyprice,startdate,expiredate
dim targetitemlist, couponmeaipprice, reguserid
dim i

useridarr = request("useridarr")
couponvalue = request("couponvalue")
coupontype = request("coupontype")
couponname = request("couponname")
minbuyprice = request("minbuyprice")
startdate = request("startdate")
expiredate = request("expiredate")
targetitemlist = request("targetitemlist")
couponmeaipprice = request("couponmeaipprice")
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
	'일반회원
	sqlStr = "insert into [db_user].[dbo].tbl_user_coupon"
	sqlStr = sqlStr + " (masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)"
	sqlStr = sqlStr + " select 0,userid,'" + Cstr(coupontype) + "','" + Cstr(couponvalue) + "','" + couponname + "','"  + Cstr(minbuyprice) + "',"
	sqlStr = sqlStr +  "'"  + startdate + "','"   + expiredate + "','" + targetitemlist + "'," + CStr(couponmeaipprice) + ",'" + reguserid + "'"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
	sqlStr = sqlStr + " where userid in ('" + useridarr + "')"

	rsget.Open sqlStr,dbget,1

	'biz회원
	sqlStr = "insert into [db_user].[dbo].tbl_user_coupon"
	sqlStr = sqlStr + " (masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)"
	sqlStr = sqlStr + " select 0,userid,'" + Cstr(coupontype) + "','" + Cstr(couponvalue) + "','" + couponname + "','"  + Cstr(minbuyprice) + "',"
	sqlStr = sqlStr +  "'"  + startdate + "','"   + expiredate + "','" + targetitemlist + "'," + CStr(couponmeaipprice) + ",'" + reguserid + "'"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c"
	sqlStr = sqlStr + " where userid in ('" + useridarr + "')"

	rsget.Open sqlStr,dbget,1
end if

'dim refer
'refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('eventcouponlist.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->