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
' Description : csv파일생성
' History : 이상구 생성
'			2020.08.06 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<% '<!-- include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
' 이방식 사용금지. 검색조건은 계속 추가 되는데 필드에 쿼리로 넣어야해서 복잡해지고 쓸수가 없음. 만드는 공수가 오히려 더 걸림. 그리고 각각의 엑셀 다운로드 마다 양식이나 요구 조건이 있는데 맞춰줄수가 없음. 다행히 실용성이 떨어져서 사용한 건수가 2건뿐이 안됨.

dim sqlStr, ArrRows, intLoop
dim idx, columnNames, params, paramVals, Separator, Field
dim i, j, k

idx = RequestCheckVar(request("idx"), 100)
paramVals = RequestCheckVar(request("paramVals"), 500)

''idx = "1"
''paramVals = "2019-04-01|2019-05-01|S|disney10x10"

'// 해킹대비
if Not IsNumeric(idx) then
	response.write "잘못된 접근입니다."
	dbget.close() : response.end
end if

'// 해킹대비
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
	response.write "잘못된 접근입니다."
	dbget.close() : response.end
end if

params = Split(params, "|")
paramVals = Split(paramVals, "|")

if UBound(params) <> UBound(paramVals) then
	response.write "파라미터가 동일하지 않습니다."
	dbget.close() : response.end
end if


for i = 0 to UBound(params)
	sqlStr = Replace(sqlStr, params(i), paramVals(i))
next

intLoop=0
Response.Buffer = true    '버퍼사용여부
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
            Response.Flush		' 버퍼리플래쉬
        end if
        Response.Write vbNewLine
		intLoop=intLoop+1
        rsget.MoveNext
	loop
END IF
rsget.close

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
