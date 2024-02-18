<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim sqlStr, yyyymmdd, USD, CNY, MYR, SGD, iMsg, menupos
menupos		= Request("menupos")
yyyymmdd	= requestCheckvar(Request("yyyymmdd"),10)
USD			= Request("USD")
CNY			= Request("CNY")
MYR			= Request("MYR")
SGD			= Request("SGD")


If (isnumeric(USD) = False) OR (isnumeric(CNY) = False) OR (isnumeric(MYR) = False) OR (isnumeric(SGD) = False) Then
	iMsg = "환율은 숫자여야 합니다."	
End If

If iMsg = "" Then
    sqlStr = ""
    sqlStr = sqlStr & " IF NOT EXISTS(SELECT TOP 1 * FROM db_item.dbo.tbl_dayexchageRate WHERE yyyymmdd = '"&yyyymmdd&"') "
    sqlStr = sqlStr & " 	INSERT INTO db_item.dbo.tbl_dayexchageRate (yyyymmdd, USD, CNY, MYR, SGD, regdate, regUserid) VALUES "
    sqlStr = sqlStr & " 	('"&yyyymmdd&"', '"&USD&"', '"&CNY&"', '"&MYR&"', '"&SGD&"', getdate(), '"&session("ssBctID")&"'  ) "
    sqlStr = sqlStr & " ELSE "
    sqlStr = sqlStr & " 	UPDATE db_item.dbo.tbl_dayexchageRate SET "
    sqlStr = sqlStr & " 	USD = '"&USD&"' "
    sqlStr = sqlStr & " 	,CNY = '"&CNY&"' "
    sqlStr = sqlStr & " 	,MYR = '"&MYR&"' "
    sqlStr = sqlStr & " 	,SGD = '"&SGD&"' "
    sqlStr = sqlStr & " 	,lastUpdate = getdate() "
    sqlStr = sqlStr & " 	,lastUserid = '"&session("ssBctID")&"' "
    sqlStr = sqlStr & " 	WHERE yyyymmdd = '"&yyyymmdd&"'"
	rsget.Open sqlStr,dbget,1
	iMsg = "저장하였습니다."
End If 
%>
<script language="javascript">
<% If (iMsg <> "") Then %>
alert("<%=iMsg %>");
document.parent.reload();
<% End If %>

</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->