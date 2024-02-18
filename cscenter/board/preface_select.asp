<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<%
dim gubun,ix,sqlStr,Fcontents,userid,masterid


gubun = request("gubun")
userid = request("userid")
masterid = request("masterid")

dim useridForShow : useridForShow = "고객"
if (userid <> "") then
	useridForShow = userid
end if


sqlStr = "select top 1 contents" + vbcrlf
sqlStr = sqlStr + " from [db_cs].[dbo].tbl_qna_preface" + vbcrlf
sqlStr = sqlStr + " where isusing = 'Y'" + vbcrlf
sqlStr = sqlStr + " and gubun = '" + gubun + "'" + vbcrlf
sqlStr = sqlStr + " and masterid = '" + masterid + "'" + vbcrlf
sqlStr = sqlStr + " order by newid()"

rsget.Open sqlStr,dbget,1

if  not rsget.EOF  then
	 Fcontents = replace(db2html(rsget("contents")),vbcrlf,"<br>")
	 Fcontents = replace(db2html(Fcontents),"(아이디)",useridForShow)
	 Fcontents = replace(db2html(Fcontents),"(이름)",session("ssBctCname"))
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

	parent.TnChangeText(temp);
	parent.frm.imsitxt.value = temp;

</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
