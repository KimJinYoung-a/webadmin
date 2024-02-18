<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim userid,username ,asid ,reqaddr1,reqaddr2,reqetc ,reqname,reqphone,reqhp,reqzip ,reqemail
dim sqlStr, resultRows
	asid = requestCheckVar(request("asid"),10)
	reqname = requestCheckVar(request("reqname"),32)
	reqphone = requestCheckVar(request("reqphone1"),4) & "-" & requestCheckVar(request("reqphone2"),4) & "-" & requestCheckVar(request("reqphone3"),4)
	reqhp = requestCheckVar(request("reqhp1"),4) & "-" & requestCheckVar(request("reqhp2"),4) & "-" & requestCheckVar(request("reqhp3"),4)
	reqzip = requestCheckVar(request("zipcode"),7)
	reqaddr1 = requestCheckVar(request("addr1"),128)
	reqaddr2 = requestCheckVar(request("addr2"),255)
	reqetc = requestCheckVar(request("reqetc"),512)
	reqemail = requestCheckVar(request("reqemail"),128)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

	sqlStr = " if exists("     + VbCrlf
    sqlStr = sqlStr + " 	select top 1 * from db_shop.dbo.tbl_shopjumun_cs_delivery"     + VbCrlf
    sqlStr = sqlStr + " 	where asid = "&asid&""     + VbCrlf
    sqlStr = sqlStr + " )"     + VbCrlf
    sqlStr = sqlStr + " 	update db_shop.dbo.tbl_shopjumun_cs_delivery set"     + VbCrlf
    sqlStr = sqlStr + " 	reqname='" + html2db(reqname) + "'"   + VbCrlf
    sqlStr = sqlStr + " 	,reqphone = '" + html2db(reqphone) + "'"  + VbCrlf
    sqlStr = sqlStr + " 	,reqhp = '" + html2db(reqhp) + "'"        + VbCrlf
    sqlStr = sqlStr + " 	,reqzipcode = '" + html2db(reqzip) + "'"  + VbCrlf
    sqlStr = sqlStr + " 	,reqzipaddr = '" + html2db(reqaddr1) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqetcaddr = '" + html2db(reqaddr2) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqetcstr = '" + html2db(reqetc) + "'"    + VbCrlf
    sqlStr = sqlStr + " 	,reqemail = '" + html2db(reqemail) + "'"    + VbCrlf    
    sqlStr = sqlStr + " 	where asid='" + CStr(asid) + "'" + VbCrlf
    sqlStr = sqlStr + " else"     + VbCrlf
    sqlStr = sqlStr + " 	insert into db_shop.dbo.tbl_shopjumun_cs_delivery("     + VbCrlf
    sqlStr = sqlStr + " 	asid ,reqname ,reqphone ,reqhp ,reqzipcode ,reqzipaddr ,reqetcaddr ,reqetcstr "     + VbCrlf
    sqlStr = sqlStr + " 	,reqemail ,regdate) values"     + VbCrlf
    sqlStr = sqlStr + " 	("     + VbCrlf
    sqlStr = sqlStr + " 	"&asid&" ,'" + html2db(reqname) + "','" + html2db(reqphone) + "' ,'" + html2db(reqhp) + "'"     + VbCrlf
    sqlStr = sqlStr + " 	,'" + html2db(reqzip) + "','" + html2db(reqaddr1) + "' ,'" + html2db(reqaddr2) + "'"     + VbCrlf
    sqlStr = sqlStr + " 	,'" + html2db(reqetc) + "','" + html2db(reqemail) + "',getdate()"     + VbCrlf
    sqlStr = sqlStr + " 	)"

    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>alert('저장 되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
%>

<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->