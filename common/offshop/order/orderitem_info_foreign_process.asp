<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 주문서
' History : 2016.09.05 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode , addSql ,sqlStr, i, detailidxarr, baljucodearr, itemgubunarr, itemidarr, itemoptionarr, itemnamearr, itemoptionnamearr
	mode = requestcheckvar(Request("mode"),32)
	detailidxarr = Request("detailidxarr")
	baljucodearr = Request("baljucodearr")
	itemgubunarr = Request("itemgubunarr")
	itemidarr = Request("itemidarr")
	itemoptionarr = Request("itemoptionarr")
	itemnamearr = Request("itemnamearr")
	itemoptionnamearr = Request("itemoptionnamearr")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")
	
if mode = "itemedit" then
	detailidxarr = split(detailidxarr,",")
	baljucodearr = split(baljucodearr,",")
	itemgubunarr = split(itemgubunarr,",")
	itemidarr = split(itemidarr,",")
	itemoptionarr = split(itemoptionarr,",")
	itemnamearr = split(itemnamearr,",")
	itemoptionnamearr = split(itemoptionnamearr,",")

	for i = 0 to ubound(detailidxarr)-1	 
    	sqlStr = "update [db_storage].[dbo].tbl_ordersheet_detail set" & VbCrlf
    	sqlStr = sqlStr & " itemname = '"& trim(html2db(replace(itemnamearr(i),"!@#",","))) &"'" & VbCrlf

    	if trim(itemoptionarr(i)) <> "" then
    		sqlStr = sqlStr & " ,itemoptionname = '"& trim(html2db(replace(itemoptionnamearr(i),"!@#",","))) &"'" & VbCrlf
    	end if

    	sqlStr = sqlStr & " ,updt = getdate()" & VbCrlf
    	sqlStr = sqlStr & " where idx = "& detailidxarr(i) &"" & VbCrlf
    	sqlStr = sqlStr & " and itemgubun = '"& itemgubunarr(i) &"'" & VbCrlf
    	sqlStr = sqlStr & " and itemid = "& itemidarr(i) &"" & VbCrlf
    	sqlStr = sqlStr & " and itemoption = '"& itemoptionarr(i) &"'" & VbCrlf

		response.write sqlStr &"<br>"
		dbget.execute sqlStr
    next

	response.write "<script langauge='javascript'>alert('OK'); parent.location.reload(); location.href ='about:blank';</script>"
	dbget.close()	:	response.End

else
	response.write "<script type='text/javascript'>alert('구분자가 없습니다.'); location.href ='about:blank';</script>"
	dbget.close()	:	response.End
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
