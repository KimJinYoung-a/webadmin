<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim itemid,userid,mode

itemid=request("itemid")
userid=request("userid")
mode= request("mode")


dim sql
if mode="best" then

	sql = " update db_contents.dbo.tbl_diary_event_evaluate "&_
				" set topYn='N' "&_
				" where itemid='" & itemid & "'" &_

				" update db_contents.dbo.tbl_diary_event_evaluate " &_
				" set topYn='Y'" &_
				" where userid='" & userid & "'" &_
				" and itemid='" & itemid & "'"


elseif mode="del" then
	sql = " delete db_contents.dbo.tbl_diary_event_evaluate " &_
				" where userid='" & userid & "'" &_
				" and itemid='" & itemid & "'"
end if

rsget.open sql,dbget,1
%>
<script language="javascript" type="text/javascript">
alert('적용 되었습니다.');
parent.document.location.reload(true);
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
