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



'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('����Ǿ����ϴ�.'); opener.location.reload(true);self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.go(-1);" &_
					"</script>"


	End If

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->