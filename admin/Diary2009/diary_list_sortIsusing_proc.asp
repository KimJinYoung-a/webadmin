<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���̾ ����Ʈ ���� ���ù�ȣ,��뿩�� ó�� ������
' History : 2015.09.14 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/Diary2009/classes/DiaryCls.asp"-->
<%
Dim msg
Dim tmpSort, tmpIsusing
Dim cnt, i, sqlStr, idx, mode
Dim detailidxarr, isusingarr, sortnoarr
''	idx	= Request("idx")
	mode = Request("mode")
	sortnoarr 	= Request("sortnoarr")
	isusingarr = Request("isusingarr")
	detailidxarr = Request("detailidxarr")

dbget.beginTrans
if mode="sortisusingedit" then

	'�����̹��� �ľ�
	detailidxarr = split(detailidxarr,",")
	cnt = ubound(detailidxarr)
	sortnoarr	=  split(sortnoarr,",")
	isusingarr	=  split(isusingarr,",")


	For i = 0 to cnt
		tmpSort = sortnoarr(i)
		tmpIsusing = isusingarr(i)
		
		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_diary2010.dbo.tbl_DiaryMaster SET " & VBCRLF
		sqlStr = sqlStr & " isusing = '"&tmpIsusing&"'" & VBCRLF
		sqlStr = sqlStr & " ,mdpicksort = '"&tmpSort&"'" & VBCRLF
		sqlStr = sqlStr & " WHERE diaryID =" & detailidxarr(i)
		
		'response.write sqlStr & "<Br>"
		dbget.execute sqlStr
	Next

	msg = "���� �Ǿ����ϴ�"

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)
		Alert_move msg,"/admin/diary2009/"
		
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)
		Alert_return "ó���� ������ �߻��߽��ϴ�."
	End If
'	response.write "<script language='javascript'>"
'	response.write "	alert('����Ǿ����ϴ�');"
'	response.write "	location.replace('/admin/diary2009/index.asp);"
'	response.write "</script>"
else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End	
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->