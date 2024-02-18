<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
'#############################
'endtype:1 - 엠디강제종료(정리대상처리)
'#############################

dim mode,i,makerid,subidx
dim sReturnURL
dim sqlStr, adminid

mode = requestCheckvar(request("mode"),20)
makerid = requestCheckvar(request("makerid"),32)
subidx = requestCheckvar(request("subidx"),10)

adminid = session("ssBctID")

SELECT CASE mode
CASE "delay30"
    sqlStr = "exec db_brand.dbo.[usp_Ten_BrandService_OutBrand_RequireDelay] '"&makerid&"',"&subidx&",30,'"&adminid&"'"
    dbget.Execute(sqlStr)

    response.write "<script>alert('OK');parent.location.reload();</script>"
CASE "delay90"
    sqlStr = "exec db_brand.dbo.[usp_Ten_BrandService_OutBrand_RequireDelay] '"&makerid&"',"&subidx&",90,'"&adminid&"'"
    dbget.Execute(sqlStr)

    response.write "<script>alert('OK');parent.location.reload();</script>"
CASE "soldoutitems"
    sqlStr = "exec db_brand.[dbo].[usp_Ten_BrandService_OutBrand_ExpireItemByBrand] '"&makerid&"',"&subidx&",'"&adminid&"'"
    dbget.Execute(sqlStr)

    response.write "<script>alert('OK');parent.location.reload();</script>"
CASE ELSE
	response.write "<script>alert('데이터 처리에 문제가 발생하였습니다.("&mode&")');</script>"
END SELECT

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->