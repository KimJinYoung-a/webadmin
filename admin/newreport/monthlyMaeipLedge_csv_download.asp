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
' Description : ����ڻ�(����) FIX csv�ٿ�ε�
' Hieditor : �̻� ����
'			 2023.10.11 �ѿ�� ����(���ϸ� ����)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%


Const MaxPage   = 50
''Const PageSize = 2000  ''�Ǽ� ����..
dim PageSize : PageSize = 2000

dim yyyymm, placeGubun, PriceGbn
dim ver

yyyymm = request("yyyymm")
placeGubun = request("placeGubun")
PriceGbn = request("PriceGbn")
ver = request("ver")

if (ver = "") then
	ver = "V2"
end if

if (ver = "DW") then
    '// 5������ �������� ��.
	PageSize = 2500
end if

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

Dim FileName: FileName = "MonthlyMaeipLedge_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList, placeGubun)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
	if isarray(arrList) then
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "'" & arrList(1,intLoop) & "'"
		bufstr = bufstr & "," & trim(arrList(2,intLoop))
		bufstr = bufstr & "," & arrList(3,intLoop)
		bufstr = bufstr & "," & arrList(4,intLoop)
        ''�귣��
        bufstr = bufstr & "," & arrList(26,intLoop)
		''��ǰ����
		bufstr = bufstr & "," & arrList(5,intLoop)
		''��ǰ�ڵ�
		bufstr = bufstr & "," & arrList(6,intLoop)
		''�ɼ��ڵ�
		bufstr = bufstr & "," & "'" & arrList(7,intLoop) & "'"
        ''���ڵ�
		if (ver = "DW") then
            bufstr = bufstr & "," & "'" & arrList(44,intLoop) & "'"
        else
            bufstr = bufstr & "," & "'" & arrList(42,intLoop) & "'"
        end if

		'�ܰ�(���)
		bufstr = bufstr & "," & arrList(28,intLoop) ''arrList(30,intLoop)
		''�������(SYS)
		bufstr = bufstr & "," & arrList(8,intLoop)
		bufstr = bufstr & "," & arrList(9,intLoop)
		''�԰�
		bufstr = bufstr & "," & arrList(10,intLoop)
		bufstr = bufstr & "," & arrList(11,intLoop)
		''�̵�
		bufstr = bufstr & "," & arrList(12,intLoop)
		bufstr = bufstr & "," & arrList(13,intLoop)
		''�Ǹ�
		bufstr = bufstr & "," & arrList(14,intLoop)
		bufstr = bufstr & "," & arrList(15,intLoop)
        ''�������
		bufstr = bufstr & "," & arrList(16,intLoop)
		bufstr = bufstr & "," & arrList(17,intLoop)
		''��Ÿ���(��:�ν����)
		bufstr = bufstr & "," & arrList(20,intLoop)
		bufstr = bufstr & "," & arrList(21,intLoop)
		''CS���
		bufstr = bufstr & "," & arrList(22,intLoop)
		bufstr = bufstr & "," & arrList(23,intLoop)

		''����
		bufstr = bufstr & "," & (arrList(8,intLoop) + arrList(10,intLoop)+ arrList(12,intLoop)+ arrList(14,intLoop)+arrList(16,intLoop)+ arrList(18,intLoop)+arrList(20,intLoop) +arrList(22,intLoop)- arrList(24,intLoop))*-1
		bufstr = bufstr & "," & (arrList(9,intLoop) + arrList(11,intLoop)+ arrList(13,intLoop)+ arrList(15,intLoop)+arrList(17,intLoop)+ arrList(19,intLoop)+arrList(21,intLoop) +arrList(23,intLoop)- arrList(25,intLoop))*-1

		''�⸻���(�ý������)
		bufstr = bufstr & "," & arrList(24,intLoop)
		bufstr = bufstr & "," & arrList(25,intLoop)

		''�����԰��
		if placeGubun <> "S" then
			bufstr = bufstr & ",'" & arrList(29,intLoop) & "'"
		end if

		''�����԰��(���Ա��к�)
		if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then
			bufstr = bufstr & ",'" & arrList(30,intLoop) & "'"
		end if

		''��Ÿ���ó(�ν����ó)
		''bufstr = bufstr & "," & Replace(arrList(33,intLoop), ",", "__")

        ''����ڹ�ȣ
        bufstr = bufstr & "," & arrList(32,intLoop) ''arrList(28,intLoop)

		''�����
        bufstr = bufstr & "," & arrList(27,intLoop) ''arrList(29,intLoop)

        ''����ī�װ�
        bufstr = bufstr & "," & arrList(31,intLoop) ''arrList(27,intLoop)

        ''����ī��1  //2016/03/22
        bufstr = bufstr & "," & arrList(34,intLoop)
        ''����ī��2
        bufstr = bufstr & "," & arrList(35,intLoop)

		'// ��������
		bufstr = bufstr & "," & arrList(36,intLoop)
		bufstr = bufstr & "," & arrList(37,intLoop)
		bufstr = bufstr & "," & arrList(38,intLoop)

		'// �Һ��ڰ�, �����ǸŰ�, �����Ǹſ���
		bufstr = bufstr & "," & arrList(39,intLoop)
		bufstr = bufstr & "," & arrList(40,intLoop)
		bufstr = bufstr & "," & arrList(41,intLoop)

		if (ver = "DW") then
			bufstr = bufstr & "," & arrList(42,intLoop)	'��޾�(���ʽ��������밡)
			bufstr = bufstr & "," & arrList(43,intLoop)	'��ǰ��
			bufstr = bufstr & "," & arrList(47,intLoop)	'�ɼǸ�
			bufstr = bufstr & "," & arrList(45,intLoop)	'��ǰ��������
			bufstr = bufstr & "," & arrList(46,intLoop)	'�ɼǴ�������
		end if

        tFile.WriteLine bufstr
    Next
	end if
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage
dim otime
''otime=Timer()

if (ver = "V2") then
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count_V2] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
elseif (ver = "DW") then
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count_V2] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
else
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
end if
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly'', adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	FTotCnt = rsget(0)
END IF
rsget.close

''rw FormatNumber(Timer()-otime,4)
''response.write FTotCnt


Dim i, ArrRows
Dim headLine
IF FTotCnt > 0 THEN
	FTotPage =  CInt(FTotCnt\PageSize)
	If (FTotCnt\PageSize) <> (FTotCnt/PageSize) Then
		FTotPage = FTotPage + 1
	End If
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		If NOT fso.FolderExists(appPath) THEN
			fso.CreateFolder(appPath)
		END If
	Set tFile = fso.CreateTextFile(appPath & FileName )

	headLine = "YYYY-MM,�����ġ,�μ�,��������,�귣��,��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,�ܰ�(���),���ʼ���,���ʱݾ�,�԰����,�԰�ݾ�,�̵�����,�̵��ݾ�,�Ǹż���,�Ǹűݾ�,����������,�������ݾ�,��Ÿ������,��Ÿ���ݾ�,CS������,CS���ݾ�,��������,�����ݾ�,�⸻����,�⸻�ݾ�"

	''�����԰��
	if placeGubun <> "S" then
		headLine = headLine & ",�����԰��"
	end if
	''�����԰��(���Ա��к�)
	if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then
		headLine = headLine & ",�����԰��(���Ա��к�)"
	end if

	headLine = headLine & ",����ڹ�ȣ,�����,����ī�װ�"


	headLine = headLine & ",����ī��1,����ī��2"
	''headLine = ",,��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�"

	headLine = headLine & ",��������"
	headLine = headLine & ",���͸��Ա���"
	headLine = headLine & ",��ǰ���Ա���"
	headLine = headLine & ",�Һ��ڰ�"
	headLine = headLine & ",�����ǸŰ�"
	headLine = headLine & ",�����Ǹſ���"
	headLine = headLine & ",���ʽ��������밡"
	if (ver = "DW") then
		headLine = headLine & ",��ǰ��"
		headLine = headLine & ",�ɼǸ�"
		headLine = headLine & ",��ǰ��������"
		headLine = headLine & ",�ɼǴ�������"
	end if

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
        otime=Timer()

		if (ver = "V2") then
		    '' sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2 => sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1 ''�ӽú��� 2015/01/12
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"  ''��ġ���� 2015/04/13
		elseif (ver = "DW") then
			sqlStr ="exec [db_datamart].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_DW] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"
		else
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"
		end if

        ''response.write "1111111<br />"
        ''response.write sqlStr
        ''dbget.close : db3_dbget.close : response.end

		if (ver = "DW") then
    		db3_rsget.CursorLocation = adUseClient
			db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
			IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
				ArrRows = db3_rsget.getRows()
			END IF
			db3_rsget.close
		else
    		rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				ArrRows = rsget.getRows()
			END IF
			rsget.close
		end if



        CALL WriteMakeFile(tFile,ArrRows, placeGubun)

        ''rw FormatNumber(Timer()-otime,4)
    NExt
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

response.write FTotCnt&"�� ���� ["&FileName&"]"
response.redirect AdmPath&"/"&FileName


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
