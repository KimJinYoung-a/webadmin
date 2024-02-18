<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/diary_collection/diary_collection_cls.asp" -->
<%

dim rank,diaryid,itemid

rank=request("rank")
diaryid= request("diaryid")
itemid = request("itemid")

rank =split(rank,",")
diaryid =split(diaryid,",")
itemid =split(itemid,",")


dim i,strSQL

dbget.beginTrans

strSQL =" DELETE FROM [db_diary_collection].dbo.tbl_diary_mdpick"

dbget.execute(strSQL)

for i=0 to Ubound(rank)
	strSQL = strSQL &_
			" INSERT INTO [db_diary_collection].dbo.tbl_diary_mdpick(diaryid,pickrank) " &_
			" values('" & trim(diaryid(i)) & "','" & trim(rank(i)) & "')"
next




dbget.execute(strSQL)



'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>"
		response.write	" alert('저장되었습니다.'); opener.location.reload(true);self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.go(-1);" &_
					"</script>"


	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->