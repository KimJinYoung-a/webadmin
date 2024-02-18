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
couponvalue = RequestCheckvar(request("couponvalue"),10)
coupontype = RequestCheckvar(request("coupontype"),10)
couponname = RequestCheckvar(request("couponname"),64)
minbuyprice = RequestCheckvar(request("minbuyprice"),10)
startdate = RequestCheckvar(request("startdate"),10)
expiredate = RequestCheckvar(request("expiredate"),10)
targetitemlist = request("targetitemlist")
couponmeaipprice = RequestCheckvar(request("couponmeaipprice"),10)
reguserid  =  session("ssBctId")
if useridarr <> "" then
	if checkNotValidHTML(useridarr) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if couponname <> "" then
	if checkNotValidHTML(couponname) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if targetitemlist <> "" then
	if checkNotValidHTML(targetitemlist) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if (Not IsNumeric(couponmeaipprice)) then couponmeaipprice=0
if (couponmeaipprice="") then couponmeaipprice=0

if Right(useridarr,1)="," then
	useridarr = Left(useridarr,Len(useridarr)-1)
end if

useridarr = replace(useridarr," ","")
useridarr = replace(useridarr,",","','")

dim sqlStr

if (useridarr <> "") then
	sqlStr = "insert into [db_user].[dbo].tbl_user_coupon"
	sqlStr = sqlStr + " (masteridx,userid,coupontype,couponvalue,couponname,minbuyprice,startdate,expiredate,targetitemlist,couponmeaipprice,validsitename,notvalid10x10,reguserid)"
	sqlStr = sqlStr + " select 0,userid,'" + Cstr(coupontype) + "','" + Cstr(couponvalue) + "','" + couponname + "','"  + Cstr(minbuyprice) + "',"
	sqlStr = sqlStr +  "'"  + startdate + "','"   + expiredate + "','" + targetitemlist + "'," + CStr(couponmeaipprice) + ",'academy','Y','" + reguserid + "'"
	sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_n"
	sqlStr = sqlStr + " where userid in ('" + useridarr + "')"

	rsget.Open sqlStr,dbget,1
end if

'dim refer
'refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('lecCouponlist.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->