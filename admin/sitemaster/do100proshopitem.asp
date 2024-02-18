<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim mode, masteridx, idx,eCode
dim itemid, isusing, startdate, enddate

dim couponname, couponvalue, coupontype
dim couponstartdate, couponexpiredate
dim minbuyprice

dim mdname1,mdcomment1
dim mdname2,mdcomment2
dim mdname3,mdcomment3


mode = request("mode")
masteridx = request("masteridx")
IF masteridx = "" THEN masteridx = 0
eCode = request("eCode")
IF eCode = "" THEN eCode = 0
	
IF masteridx = 0 and eCode = 0 THEN
%>
	<script language="javascript">
	<!--
		alert("데이터에 문제가 있습니다. 관리자에게 문의해주세요");
		history.back(-1);
	//-->
	</script>
<%	
dbget.close()	:	response.End
END IF	

idx = request("idx")
isusing = request("isusing")
itemid = request("itemid")

startdate = request("startdate")
enddate = request("enddate") + " 23:59:59"

couponname    	= html2db(request("couponname"))
couponvalue		= request("couponvalue")
coupontype 		= request("coupontype")
couponstartdate	= request("couponstartdate")
couponexpiredate = request("couponexpiredate")  + " 23:59:59"
minbuyprice     = request("minbuyprice")

mdname1			= html2db(request("mdname1"))
mdcomment1		= html2db(request("mdcomment1"))
mdname2			= html2db(request("mdname2"))
mdcomment2		= html2db(request("mdcomment2"))
mdname3			= html2db(request("mdname3"))
mdcomment3		= html2db(request("mdcomment3"))


'==============================================================================

dim sqlStr,iid

if (mode = "write") then

	sqlStr = " insert into [db_sitemaster].[dbo].tbl_100proshop (masteridx,itemid,startdate,enddate,"
	sqlStr = sqlStr + " couponname, couponvalue, coupontype, couponstartdate, couponexpiredate,"
	sqlStr = sqlStr + " minbuyprice,isusing,"
	sqlStr = sqlStr + " mdname1,mdcomment1,mdname2,mdcomment2,mdname3,mdcomment3,evt_code) "
	sqlStr = sqlStr + " values(" & masteridx & ""
	sqlStr = sqlStr + " ," + Cstr(itemid) + ""
	sqlStr = sqlStr + " ,'" + Cstr(startdate) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(enddate) + "'"
	sqlStr = sqlStr + " ,'" + couponname + "'"
	sqlStr = sqlStr + " ," + couponvalue + ""
	sqlStr = sqlStr + " ,'" + coupontype + "'"
	sqlStr = sqlStr + " ,'" + couponstartdate + "'"
	sqlStr = sqlStr + " ,'" + couponexpiredate + "'"
	sqlStr = sqlStr + " ," + Cstr(minbuyprice) + ""
	sqlStr = sqlStr + " ,'" + Cstr(isusing) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdname1) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdcomment1) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdname2) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdcomment2) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdname3) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(mdcomment3) + "'"
	sqlStr = sqlStr + " ,'" + Cstr(eCode) + "')"

	rsget.Open sqlStr, dbget, 1


	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].tbl_100proshop') as iid"
	rsget.Open sqlStr, dbget, 1
	If Not Rsget.Eof then
		iid = rsget("iid")
	end if
	rsget.close

end if


if (mode = "modify") then

	sqlStr = " update [db_sitemaster].[dbo].tbl_100proshop "
	sqlStr = sqlStr + "set itemid = " + Cstr(itemid) + " "
	sqlStr = sqlStr + " ,startdate = '" + Cstr(startdate) + "' "
	sqlStr = sqlStr + " ,enddate = '" + Cstr(enddate) + "' "
	sqlStr = sqlStr + " ,couponname = '" + Cstr(couponname) + "' "
	sqlStr = sqlStr + " ,couponvalue = " + Cstr(couponvalue) + " "
	sqlStr = sqlStr + " ,coupontype = '" + Cstr(coupontype) + "' "
	sqlStr = sqlStr + " ,couponstartdate = '" + Cstr(couponstartdate) + "' "
	sqlStr = sqlStr + " ,couponexpiredate = '" + Cstr(couponexpiredate) + "' "
	sqlStr = sqlStr + " ,minbuyprice = " + Cstr(minbuyprice) + " "
	sqlStr = sqlStr + " ,isusing = '" + Cstr(isusing) + "' "
	sqlStr = sqlStr + " ,mdname1 = '" + Cstr(mdname1) + "' "
	sqlStr = sqlStr + " ,mdcomment1 = '" + Cstr(mdcomment1) + "' "
	sqlStr = sqlStr + " ,mdname2 = '" + Cstr(mdname2) + "' "
	sqlStr = sqlStr + " ,mdcomment2 = '" + Cstr(mdcomment2) + "' "
	sqlStr = sqlStr + " ,mdname3 = '" + Cstr(mdname3) + "' "
	sqlStr = sqlStr + " ,mdcomment3 = '" + Cstr(mdcomment3) + "' "

	sqlStr = sqlStr + " where idx = " + CStr(idx) + " "

	'response.write sqlStr
	rsget.Open sqlStr, dbget, 1

	iid = idx

end if

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('/admin/sitemaster/100proshoplist.asp');</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
