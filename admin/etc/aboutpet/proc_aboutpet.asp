<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
'// 저장 모드 접수
Dim cmdparam, sqlStr, iErrMsg, i
Dim itemarr, regedItemNamearr, regedOptionnamearr, regedItempricearr, aboutpetSellYnarr

cmdparam			= Trim(Request("cmdparam"))
itemarr				= Trim(Request("itemarr"))
regedItemNamearr	= Trim(Request("regedItemNamearr"))
regedOptionnamearr	= Trim(Request("regedOptionnamearr"))
regedItempricearr	= Trim(Request("regedItempricearr"))
aboutpetSellYnarr	= Trim(Request("aboutpetSellYnarr"))

' If Right(itemarr,2) = "||" Then itemarr = Left(itemarr, Len(itemarr) - 2)
' If Right(regedItemNamearr,2) = "||" Then regedItemNamearr = Left(regedItemNamearr, Len(regedItemNamearr) - 2)
' If Right(regedOptionnamearr,2) = "||" Then regedOptionnamearr = Left(regedOptionnamearr, Len(regedOptionnamearr) - 2)
' If Right(regedItempricearr,2) = "||" Then regedItempricearr = Left(regedItempricearr, Len(regedItempricearr) - 2)
' If Right(aboutpetSellYnarr,2) = "||" Then aboutpetSellYnarr = Left(aboutpetSellYnarr, Len(aboutpetSellYnarr) - 2)

Dim splititem, splitregedItemName, splitregedOptionname, splitregedItemprice, splitaboutpetSellYn
splititem				= Split(itemarr, "||")
splitregedItemName		= Split(regedItemNamearr, "||")
splitregedOptionname	= Split(regedOptionnamearr, "||")
splitregedItemprice		= Split(regedItempricearr, "||")
splitaboutpetSellYn		= Split(aboutpetSellYnarr, "||")

For i = 0 To Ubound(splititem)-1
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.dbo.tbl_aboutpet_regitem "
	sqlStr = sqlStr & " SET itemname = '"& html2db(splitregedItemName(i)) &"' "
	sqlStr = sqlStr & " ,sellprice = '"& splitregedItemprice(i) &"' "
	sqlStr = sqlStr & " ,optionname = '"& html2db(splitregedOptionname(i)) &"' "
	sqlStr = sqlStr & " ,aboutpetsellyn = '"& splitaboutpetSellYn(i) &"' "
	sqlStr = sqlStr & " WHERE idx = '"& splititem(i) &"' "
	dbget.execute sqlStr
Next
%>
<script language="javascript">
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->