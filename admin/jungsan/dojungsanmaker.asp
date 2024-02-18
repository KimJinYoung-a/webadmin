<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim refer
request.ServerVariables("HTTP_REFERER")

dim extsitename, orderserial, userid, buyname
dim totalsum, deasangsum, beasongpay, jungsansum
dim startdate, enddate, commission
dim totalcount, tot_totalsum, tot_beasongpay, tot_deasangsum
dim tot_jungsansum, txetc

extsitename = request("extsitename")
orderserial = request("orderserial")
userid = request("userid")
buyname = request("buyname")
totalsum = request("totalsum")
deasangsum = request("deasangsum")
beasongpay = request("beasongpay")
jungsansum = request("jungsansum")
startdate = request("startdate")
enddate = request("enddate")
commission = request("commission")
totalcount = request("totalcount")
tot_totalsum = request("tot_totalsum")
tot_beasongpay = request("tot_beasongpay")
tot_deasangsum = request("tot_deasangsum")
tot_jungsansum = request("tot_jungsansum")
txetc = html2DB(request("txetc"))

dim arr_orderserial, arr_userid, arr_buyname
dim arr_totalsum, arr_deasangsum, arr_beasongpay
dim arr_jungsansum

arr_orderserial = split(orderserial,"|")
arr_userid		= split(userid,"|")
arr_buyname		= split(buyname,"|")
arr_totalsum	= split(totalsum,"|")
arr_deasangsum	= split(deasangsum,"|")
arr_beasongpay	= split(beasongpay,"|")
arr_jungsansum	= split(jungsansum,"|")

dim sqlStr, iid, i, ttlcnt
sqlStr = " insert into [db_jungsan].[dbo].tbl_etcsite_jungsanmaster"
sqlStr = sqlStr + " (sitename,totalno,totalsum,totalbeasongpay,"
sqlStr = sqlStr + " totaldeasang,totaljungsansum,"
sqlStr = sqlStr + " comission,etcstr,startdate,enddate)"
sqlStr = sqlStr + " values("
sqlStr = sqlStr + " '" + extsitename + "',"
sqlStr = sqlStr + " " + CStr(totalcount) + ","
sqlStr = sqlStr + " " + CStr(tot_totalsum) + ","
sqlStr = sqlStr + " " + CStr(tot_beasongpay) + ","
sqlStr = sqlStr + " " + CStr(tot_deasangsum) + ","
sqlStr = sqlStr + " " + CStr(tot_jungsansum) + ","
sqlStr = sqlStr + " " + CStr(commission) + ","
sqlStr = sqlStr + " '" + txetc + "',"
sqlStr = sqlStr + " '" + startdate + "',"
sqlStr = sqlStr + " '" + enddate + "'"
sqlStr = sqlStr + " )"

''response.write sqlStr + "<br>"
rsget.Open sqlStr,dbget,1

sqlStr = " select ident_current('[db_jungsan].[dbo].tbl_etcsite_jungsanmaster') as ident"

rsget.Open sqlStr,dbget,1
iid = rsget("ident")
rsget.close

ttlcnt = UBound(arr_orderserial)
for i=1 to ttlcnt
	sqlStr = "insert into [db_jungsan].[dbo].tbl_etcsite_jungsandetail"
	sqlStr = sqlStr + " (masterid,orderserial,userid, buyname, totalsum,deasangsum,beasongpay,jungsansum)"
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " " + CStr(iid) + ","
	sqlStr = sqlStr + " '" + CStr(arr_orderserial(i)) + "',"
	sqlStr = sqlStr + " '" + CStr(arr_userid(i)) + "',"
	sqlStr = sqlStr + " '" + CStr(arr_buyname(i)) + "',"
	sqlStr = sqlStr + " " + CStr(arr_totalsum(i)) + ","
	sqlStr = sqlStr + " " + CStr(arr_deasangsum(i)) + ","
	sqlStr = sqlStr + " " + CStr(arr_beasongpay(i)) + ","
	sqlStr = sqlStr + " " + CStr(arr_jungsansum(i)) + ""
	sqlStr = sqlStr + " )"
	
	rsget.Open sqlStr,dbget,1
next
%>
<script >alert('저장 되었습니다.');</script>
<script >location.replace('/admin/jungsan/jungsanmaster_partner.asp');</script>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->