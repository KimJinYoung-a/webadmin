<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*5
%>
<%
'###########################################################
' Description : csv���ϻ���
' History : �̻� ����
'			2020.08.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
' �̹�� ������. �˻������� ��� �߰� �Ǵµ� �ʵ忡 ������ �־���ؼ� ���������� ������ ����. ����� ������ ������ �� �ɸ�. �׸��� ������ ���� �ٿ�ε� ���� ����̳� �䱸 ������ �ִµ� �����ټ��� ����. ������ �ǿ뼺�� �������� ����� �Ǽ��� 2�ǻ��� �ȵ�.

dim sqlStr, ArrRows, intLoop
dim idx, columnNames, params, paramVals, Separator, Field
dim i, j, k

idx = RequestCheckVar(request("idx"), 100)
paramVals = RequestCheckVar(request("paramVals"), 500)

''idx = "1"
''paramVals = "2019-04-01|2019-05-01|S|disney10x10"

'// ��ŷ���
if Not IsNumeric(idx) then
	response.write "�߸��� �����Դϴ�."
	dbget.close() : response.end
end if

'// ��ŷ���
paramVals = Replace(paramVals, "'", "")

sqlStr = " select * from [db_temp].[dbo].[tbl_export_csv] where idx = " & idx
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	sqlStr = rsget("sqlStr")
	columnNames = rsget("columnNames")
	params = rsget("params")
END IF
rsget.close

if (params = "") then
	response.write "�߸��� �����Դϴ�."
	dbget.close() : response.end
end if

params = Split(params, "|")
paramVals = Split(paramVals, "|")

if UBound(params) <> UBound(paramVals) then
	response.write "�Ķ���Ͱ� �������� �ʽ��ϴ�."
	dbget.close() : response.end
end if


for i = 0 to UBound(params)
	sqlStr = Replace(sqlStr, params(i), paramVals(i))
next

intLoop=0
Response.Buffer = true    '���ۻ�뿩��
response.ContentType = "text/csv"
response.AddHeader "Content-Disposition", "attachment;filename=""data.csv"""

columnNames = """" & Replace(columnNames, "|", """,""") & """"
response.write columnNames & vbCrLf

rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	do until (rsget.EOF OR rsget.BOF)
		Separator = ""
		for i = 0 to rsget.Fields.Count - 1
			Field = rsget.Fields(i).Value & ""
            Field = """" & Replace(Field, """", """""") & """"
            Response.Write Separator & Field
            Separator = ","
        next

        if intLoop mod 3000 = 0 then
            Response.Flush		' ���۸��÷���
        end if
        Response.Write vbNewLine
		intLoop=intLoop+1
        rsget.MoveNext
	loop
END IF
rsget.close

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
