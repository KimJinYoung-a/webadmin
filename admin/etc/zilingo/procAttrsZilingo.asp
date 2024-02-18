<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim mode, sqlStr, iErrMsg, joongBok, categbn
Dim itemid, itemoption, attrs, saveStr
Dim i, tmpArray, saveVal, attrsCnt
itemid = Request("itemid")
itemoption = Request("itemoption")
attrs = Request("attrs")

tmpArray = Split(attrs, ",")

For i = 0 To Ubound(tmpArray)
	saveVal = saveVal & tmpArray(i) & ","
Next
If Right(saveVal,1) = "," Then
	saveVal = Left(saveVal, Len(saveVal) - 1)
End If


sqlStr = ""
sqlStr = sqlStr & " IF Exists(SELECT * FROM db_etcmall.[dbo].[tbl_zilingo_attr_mapping] WHERE itemid = '"&itemid&"' and itemoption = '"&itemoption&"') "
sqlStr = sqlStr & " BEGIN "
sqlStr = sqlStr & " UPDATE db_etcmall.[dbo].[tbl_zilingo_attr_mapping] SET "
sqlStr = sqlStr & " attributes = '"&saveVal&"' "
sqlStr = sqlStr & " WHERE itemid = '"&itemid&"' and itemoption = '"&itemoption&"' "
sqlStr = sqlStr & " END ELSE "
sqlStr = sqlStr & " BEGIN "
sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_zilingo_attr_mapping] "
sqlStr = sqlStr & " (itemid, itemoption, attributes, regdate)" & VbCrlf
sqlStr = sqlStr & " VALUES('" & itemid & "', '"& itemoption &"', '"&saveVal&"', getdate()) "  & VbCrlf
sqlStr = sqlStr & " END "
dbget.execute(sqlStr)
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->