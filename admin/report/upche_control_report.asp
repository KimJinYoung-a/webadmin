<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/upche_control_reportcls.asp"-->
<%

dim page,shopid,i,ix
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate,toDate
dim pointyn

pointyn = request("pointyn")
if pointyn = "" then pointyn="Y"

shopid = request("shopid")
if shopid = "" then shopid = "streetshop001"

if session("ssBctDiv")="201" then
	shopid = "cafe002"
elseif session("ssBctDiv")="301" then
	shopid = "cafe003"
end if

page = request("page")
if page="" then page=1

yyyy1 = request("yyyy1")
mm1 = request("mm1")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = left(DateSerial(yyyy1,mm1,dd1),7)

toDate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),7)

dim ocount
set ocount = new COffShopSellReport
ocount.FRectShopID = shopid
ocount.FRectStartDay = fromDate
ocount.FRectEndDay = toDate
ocount.FRectPointYN = pointyn
ocount.GetWeeklySellCount

dim totalsellprice
dim tsellea1,tsellea2,tsellea3,tsellea4,tsellea5,tsellea6,tsellea7
dim avgsellsum1,avgsellea1,avgvisitsellprice1,avgpossessoin1
dim avgsellsum2,avgsellea2,avgvisitsellprice2,avgpossessoin2
dim avgsellsum3,avgsellea3,avgvisitsellprice3,avgpossessoin3
dim avgsellsum4,avgsellea4,avgvisitsellprice4,avgpossessoin4
dim avgsellsum5,avgsellea5,avgvisitsellprice5,avgpossessoin5
dim avgsellsum6,avgsellea6,avgvisitsellprice6,avgpossessoin6
dim avgsellsum7,avgsellea7,avgvisitsellprice7,avgpossessoin7
dim w_avgsellsum,w_avgsellea,w_avgvisitsellprice,w_avgpossessoin
dim we_avgsellsum,we_avgsellea,we_avgvisitsellprice,we_avgpossessoin
dim weeksellsum,weekavgsellea,weekavgvisitsellprice,weekavgpossessoin
dim weekendsum1,weekendsum2,weekendsum3,weekendsum4

%>