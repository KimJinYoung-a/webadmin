<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###############################################
' PageName : main_manager.asp
' Discription : 오프라인 사이트 관리
' History : 2010.04.19 한용민 생성
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim poscode ,posname ,posVarname ,linktype ,fixtype ,imagewidth ,isusing
dim useSet ,imageheight ,sqlStr, ItemExists
	poscode   = requestCheckVar(request.Form("poscode"),10)
	posname   = requestCheckVar(html2Db(request.Form("posname")),128)
	posVarname= requestCheckVar(request.Form("posVarname"),32)
	linktype  = requestCheckVar(request.Form("linktype"),1)
	fixtype   = requestCheckVar(request.Form("fixtype"),1)
	imagewidth= requestCheckVar(request.Form("imagewidth"),16)
	isusing   = requestCheckVar(request.Form("isusing"),1)
	imageheight= requestCheckVar(request.Form("imageheight"),16)
	useSet= requestCheckVar(request.Form("useSet"),10)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

sqlStr = "select top 1 * from [db_sitemaster].[dbo].tbl_Offshopmain_contents_poscode"
sqlStr = sqlStr + " where poscode=" + CStr(poscode)

rsget.Open sqlStr,dbget,1
    ItemExists = Not rsget.Eof
rsget.Close

if (ItemExists) then
    sqlStr = " update [db_sitemaster].[dbo].tbl_Offshopmain_contents_poscode" + VbCrlf
    sqlStr = sqlStr + " set posname='" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,posVarname='" + posVarname + "'" + VbCrlf
    sqlStr = sqlStr + " ,linktype='" + linktype + "'" + VbCrlf
    sqlStr = sqlStr + " ,fixtype='" + fixtype + "'" + VbCrlf
    sqlStr = sqlStr + " ,imagewidth='" + imagewidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,imageheight='" + imageheight + "'" + VbCrlf
    sqlStr = sqlStr + " ,useSet=" + useSet + VbCrlf
    sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " where poscode=" + CStr(poscode) + VbCrlf
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
else
    sqlStr = " insert into [db_sitemaster].[dbo].tbl_Offshopmain_contents_poscode" + VbCrlf
    sqlStr = sqlStr + " (poscode,posname,posVarname,linktype,fixtype,imagewidth,imageheight,useSet,isusing)"+ VbCrlf
    sqlStr = sqlStr + " values("
    sqlStr = sqlStr + " " + CStr(poscode) + VbCrlf
    sqlStr = sqlStr + " ,'" + posname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + posVarname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + linktype + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + fixtype + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imagewidth + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + imageheight + "'" + VbCrlf
    sqlStr = sqlStr + " ," + useSet + VbCrlf
    sqlStr = sqlStr + " ,'" + isusing + "'" + VbCrlf
    sqlStr = sqlStr + " )" + VbCrlf
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr
end if

response.write "<script type='text/javascript'>alert('저장되었습니다.');</script>"
response.write "<script type='text/javascript'>location.replace('" + referer + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->