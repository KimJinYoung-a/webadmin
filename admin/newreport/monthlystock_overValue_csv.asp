<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 30
%>
<%
'###########################################################
' Description : ������ csv�ٿ�ε�
' Hieditor : �̻� ����
'			 2023.10.11 �ѿ�� ����(���ϸ� ����)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 40
Const PageSize = 5000
Dim reqYYYYMM, reqStrplace, reqsysorreal, reqbPriceGbn, reqmygubun, reqYYYY, IsUsingV2, strNoType, strPriceType, strYearMonth
Dim AdmPath, appPath, sNow, sY, sM, sD, sH, sMi, sS, sDateName, FileName, fso, tFile, FTotCnt, FTotPage, FCurrPage, sqlStr
dim i, ArrRows, headLine
	reqYYYYMM = request("exYYYY")&"-"&request("exMM")
	reqStrplace = request("stplace")
	reqsysorreal = request("sysorreal")
	reqbPriceGbn = request("bPriceGbn")
	reqmygubun = request("mygubun")
	reqYYYY = request("exYYYY")
	IsUsingV2 = request("v2")

if (IsUsingV2 = "") then
	IsUsingV2 = "Y"
end if

AdmPath = "/admin/newreport/xldwn/"&request("exYYYY")&request("exMM")
appPath = server.mappath(AdmPath) + "/"
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

FileName = "MonthlyStock_"&sDateName&".csv"

if (IsUsingV2 = "Y") then
	sqlStr ="[db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_Count_V2] ('"&reqYYYYMM&"', '"&reqStrplace&"')"
else
	sqlStr ="[db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_Count] ('"&reqYYYYMM&"', '"&reqStrplace&"')"
end if

rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
IF Not (rsget.EOF OR rsget.BOF) THEN
	FTotCnt = rsget(0)
END IF
rsget.close

'response.write FTotCnt

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
	headLine = ""
	''If reqStrplace = "L" Then
	''	''headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,��������,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�,1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�,NULL,�հ�"
	''	if (reqmygubun = "Y") then
	''		headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,��������,�ý��ۼ���,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�," & (reqYYYY) & "," & (reqYYYY - 1) & "," & (reqYYYY - 2) & ",~ " & (reqYYYY - 3) & ",--,NULL,�հ�"
	''	else
	''		headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,��������,�ý��ۼ���,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�,1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�,NULL,�հ�"
	''	end if
	''ElseIf reqStrplace = "T" Then
	''	''headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,����,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�,1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�,NULL,�հ�"
	''	if (reqmygubun = "Y") then
	''		headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,�ý��ۼ���,����,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�," & (reqYYYY) & "," & (reqYYYY - 1) & "," & (reqYYYY - 2) & ",~ " & (reqYYYY - 3) & ",--,NULL,�հ�"
	''	else
	''		headLine = "�μ�,�ڵ屸��,���Ա���,�귣��,�ý��ۼ���,����,����,����,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��,�ɼǸ�,�����԰���,�ܰ�,1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�,NULL,�հ�"
	''	end if
	''End If

	strNoType		= "�ǻ�(+�ҷ�)"
	strPriceType	= "�ۼ��ø��԰�"
	strYearMonth	= "1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�"

	if (reqsysorreal = "sys") then
		strNoType = "�ý���"
	end if

	if (reqbPriceGbn = "V") then
		strPriceType = "��ո��԰�"
	end if

	if (reqmygubun = "Y") then
		strYearMonth = (reqYYYY) & "," & (reqYYYY - 1) & "," & (reqYYYY - 2) & ",~ " & (reqYYYY - 3)
	end if

	headLine = "�μ�,��������,���Ա���,�귣��,����,����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,��ǰ��,�ɼǸ�,�����԰���,����(�ý���),���ް�(" + CStr(strPriceType) + ")," + CStr(strYearMonth) + ",NULL,�հ�,����ī�װ�,����ī�װ�,����ī�װ�,����ī�װ�,�Һ��ڰ�,�����ǸŰ�,�����Ǹſ���,��������,���缾�͸��Ա���,������Ա���"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""

		if (IsUsingV2 = "Y") then
			sqlStr ="[db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_List_V2] ('"&reqYYYYMM&"','"&reqStrplace&"',"&i+1&","&PageSize&")"
		else
			sqlStr ="[db_summary].[dbo].[sp_Ten_monthlystock_overValue_MakeEXL_List] ('"&reqYYYYMM&"','"&reqStrplace&"',"&i+1&","&PageSize&")"
		end if

		'response.write sqlStr & "<br>"
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsget.EOF OR rsget.BOF) THEN
        	ArrRows = rsget.getRows()
        END IF
        rsget.close
       	CALL WriteMakeFile(tFile,ArrRows)
    NExt
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

response.write FTotCnt&"�� ���� ["&FileName&"]"
response.redirect AdmPath&"/"&FileName
response.end

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""
		bufstr = bufstr & arrList(1,intLoop)&","&arrList(2,intLoop)&","&arrList(3,intLoop)&","&arrList(4,intLoop)&","& trim(arrList(12,intLoop))&","&arrList(6,intLoop)&","&arrList(7,intLoop)&","&arrList(8,intLoop)&",'"&arrList(40,intLoop)&"',"&arrList(9,intLoop)&","&arrList(10,intLoop)&","&arrList(11,intLoop)

		bufstr = bufstr & ","&arrList(13,intLoop)

		if (reqbPriceGbn = "V") then
			bufstr = bufstr & ","&arrList(16,intLoop)
			tmpPrice = arrList(16,intLoop)
		else
			bufstr = bufstr & ","&arrList(15,intLoop)
			tmpPrice = arrList(15,intLoop)
		end if

		if (reqsysorreal = "sys") then
			if (reqmygubun = "Y") then
				bufstr = bufstr & ","&arrList(22,intLoop)*tmpPrice&","&arrList(23,intLoop)*tmpPrice&","&arrList(24,intLoop)*tmpPrice&","&arrList(25,intLoop)*tmpPrice
			else
				bufstr = bufstr & ","&arrList(17,intLoop)*tmpPrice&","&arrList(18,intLoop)*tmpPrice&","&arrList(19,intLoop)*tmpPrice&","&arrList(20,intLoop)*tmpPrice&","&arrList(21,intLoop)*tmpPrice
			end if

			bufstr = bufstr & ","&arrList(26,intLoop)*tmpPrice&","&arrList(13,intLoop)*tmpPrice
			bufstr = bufstr & ","&arrList(38,intLoop)
		else
			if (reqmygubun = "Y") then
				bufstr = bufstr & ","&arrList(22+10,intLoop)*tmpPrice&","&arrList(23+10,intLoop)*tmpPrice&","&arrList(24+10,intLoop)*tmpPrice&","&arrList(25+10,intLoop)*tmpPrice
			else
				bufstr = bufstr & ","&arrList(17+10,intLoop)*tmpPrice&","&arrList(18+10,intLoop)*tmpPrice&","&arrList(19+10,intLoop)*tmpPrice&","&arrList(20+10,intLoop)*tmpPrice&","&arrList(21+10,intLoop)*tmpPrice
			end if

			bufstr = bufstr & ","&arrList(26+10,intLoop)*tmpPrice&","&arrList(13+1,intLoop)*tmpPrice
			bufstr = bufstr & ","&arrList(38,intLoop)
		end if

		bufstr = bufstr & ","&arrList(41,intLoop)

		'// ����ī�װ�
		bufstr = bufstr & ","&arrList(42,intLoop)
		bufstr = bufstr & ","&arrList(43,intLoop)

		bufstr = bufstr & ","&arrList(44,intLoop)
		bufstr = bufstr & ","&arrList(45,intLoop)
		bufstr = bufstr & ","&arrList(46,intLoop)
		bufstr = bufstr & ","&arrList(47,intLoop)
		bufstr = bufstr & ","&arrList(48,intLoop)
		bufstr = bufstr & ","&arrList(49,intLoop)

        tFile.WriteLine bufstr
    Next
End function
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
