<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<%
dim gubun,ix,sqlStr,Fcontents,masterid

gubun = request("gubun")
masterid = request("masterid")

sqlStr = "select top 1 contents" + vbcrlf
sqlStr = sqlStr + " from [db_cs].dbo.tbl_qna_compliment" + vbcrlf
sqlStr = sqlStr + " where isusing = 'Y'" + vbcrlf
sqlStr = sqlStr + " and gubun = '" + gubun + "'" + vbcrlf
sqlStr = sqlStr + " and masterid = '" + masterid + "'" + vbcrlf
sqlStr = sqlStr + " order by newid()"

rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

if  not rsget.EOF  then
	 Fcontents = replace(db2html(rsget("contents")),vbcrlf,"<br>")
end if
rsget.close

%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html">
<script>

var source,convert,temp;

source = "<br>";
convert = "\n";
temp = '<% = Fcontents %>';

while (temp.indexOf(source)>-1) {
	 pos= temp.indexOf(source);
	 temp = "" + (temp.substring(0, pos) + convert + 
	 temp.substring((pos + source.length), temp.length));
}

	parent.TnChangeText2(temp);

</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
