<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.14 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid,username ,asid ,reqaddr1,reqaddr2,reqetc ,reqname,reqphone,reqhp,reqzip
dim sqlStr, resultRows
	asid = requestCheckVar(request("asid"),10)
	reqname = requestCheckVar(html2db(request("reqname")),32)
	reqphone = requestCheckVar(request("reqphone1"),4) & "-" & requestCheckVar(request("reqphone2"),4) & "-" & requestCheckVar(request("reqphone3"),4)
	reqhp = requestCheckVar(request("reqhp1"),4) & "-" & requestCheckVar(request("reqhp2"),4) & "-" & requestCheckVar(request("reqhp3"),4)
	reqzip = requestCheckVar(request("zipcode"),7)
	reqaddr1 = requestCheckVar(html2db(request("addr1")),128)
	reqaddr2 = requestCheckVar(html2db(request("addr2")),255)
	reqetc = requestCheckVar(html2db(request("reqetc")),512)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

dbget.begintrans

sqlStr = "update db_shop.dbo.tbl_shopbeasong_cs_delivery" + VbCrlf
sqlStr = sqlStr + " set reqname='" + reqname + "'" + VbCrlf
sqlStr = sqlStr + " ,reqphone='" + reqphone + "'" + VbCrlf
sqlStr = sqlStr + " ,reqhp='" + reqhp + "'" + VbCrlf
sqlStr = sqlStr + " ,reqzipcode='" + reqzip + "'" + VbCrlf
sqlStr = sqlStr + " ,reqzipaddr='" + reqaddr1 + "'" + VbCrlf
sqlStr = sqlStr + " ,reqetcaddr='" + reqaddr2 + "'" + VbCrlf
sqlStr = sqlStr + " ,reqetcstr='" + reqetc + "'" + VbCrlf
sqlStr = sqlStr + " where asid=" + asid

'response.write sqlStr &"<br>"
dbget.Execute sqlStr, resultRows

if (resultRows=0) then
	sqlStr = ""
    sqlStr = "insert into db_shop.dbo.tbl_shopbeasong_cs_delivery" + VbCrlf
    sqlStr = sqlStr + "(asid, reqname, reqphone, reqhp, reqzipcode, reqzipaddr" + VbCrlf
    sqlStr = sqlStr + " ,reqetcaddr, reqetcstr)" + VbCrlf
    sqlStr = sqlStr + " values(" + CStr(asid) + VbCrlf
    sqlStr = sqlStr + " ,'" + reqname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqphone + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqhp + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqzip + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqaddr1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqaddr2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqetc + "'" + VbCrlf
    sqlStr = sqlStr + " )"
	
	'response.write sqlStr &"<br>"
    dbget.Execute sqlStr, resultRows    
end if

if err.number = 0 then
	dbget.CommitTrans
	
	response.write "<script type='text/javascript'>alert('저장 되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
else
	dbget.RollBackTrans
	
    response.write "<script type='text/javascript'>"
    response.write "	alert('데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망');"
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End	
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->