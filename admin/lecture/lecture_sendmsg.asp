<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->

<%
dim idx,orderserial
dim serial,OSerial
dim i,Fdate
dim msg

idx=request("idx")
if idx="" then idx=0

orderserial=request("orderserial")

msg=request("msg")

dim olec
set olec = new CLectureDetail
olec.GetLectureOne idx

dim itemid,odetail
itemid = olec.Flinkitemid
set odetail = new CLecture
odetail.FRectItemID = itemid
odetail.Corderserial=orderserial
odetail.GetLectureSelected

if  FormatDatetime(now()) < FormatDatetime(olec.Flecdate01) then
	Fdate=olec.Flecdate01
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate02) then
	Fdate=olec.Flecdate02
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate03) then
	Fdate=olec.Flecdate03
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate04) then
	Fdate=olec.Flecdate04
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate05) then
	Fdate=olec.Flecdate05
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate06) then
	Fdate=olec.Flecdate06
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate07) then
	Fdate=olec.Flecdate07
elseif FormatDatetime(now()) < FormatDatetime(olec.Flecdate08) then
	Fdate=olec.Flecdate08
end if

for i=0 to odetail.FResultcount-1

'msg ="[컬리지]" + odetail.FItemList(i).FBuyName + "님께서 등록하신 컬리지 강좌시작은 " + CStr(month(Fdate)) + "월 " + CStr(day(Fdate)) + "일" + Left(CStr(timevalue(olec.Flecdate01)),(len(timevalue(olec.Flecdate01))-3)) + "분 입니다."
dim sql
sql = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg )" + vbcrlf
sql = sql + "values(" + vbcrlf
sql = sql + "'" + odetail.FItemList(i).FBuyHp +"'," + vbcrlf
sql = sql + "'02-741-9070','1',getdate(),'" + msg + "')" + vbcrlf

response.write sql
dbget.close()	:	response.End
rsget.open sql,dbget,1
'response.write sql
next

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('메시지를 전송 하였습니다.');</script>"
response.write "<script>window.close();</script>"
dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->