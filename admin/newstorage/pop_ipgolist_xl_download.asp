<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 30
%>
<%
'####################################################
' Description :  ���� �԰���Ʈ �����ٿ�
' History : 2007.01.01 �̻� ����
'			2018.10.11 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
Const MaxPage   = 40
Const PageSize = 5000

dim fromDate, toDate, yyyymm, designer, code, alinkcode, blinkcode, onoffgubun, divcode, rackipgoyn, vPurchaseType, sqlStr, FTotCnt, FTotPage, FCurrPage
dim searchType, searchText, minusjumun, ipgocheck, ExecuteDtStart, ExecuteDtEnd, CodeGubun, makerid, i, ArrRows, headLine, fso, tFile
	designer = request("designer")
	code = request("code")				' �԰� �ڵ�
	alinkcode = request("alinkcode")
	blinkcode = request("blinkcode")
	onoffgubun = request("onoffgubun")	' ��/���� ����
	divcode = request("divcode")		' ���� ����
	rackipgoyn = request("rackipgoyn")	'
	vPurchaseType = requestCheckVar(request("purchasetype"),2)
	searchType = request("searchType")
	searchText = request("searchText")
	minusjumun = request("minusjumun")

	'// �԰��� �˻��� �ʿ��� ���� ����
	ipgocheck = request("ipgocheck")
	fromDate = request("fromDate")
	toDate = request("toDate")
	yyyymm = Left(Now, 7)

if ipgocheck="on" then
	ExecuteDtStart = fromDate
	ExecuteDtEnd   = toDate
end if

if code="" then
	CodeGubun = "ST"  ''�԰�
	makerid = designer
else
	onoffgubun=""
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

Dim FileName: FileName = "IpgoList_"&sDateName&".csv"

sqlStr = " [db_summary].[dbo].[sp_Ten_IpgoList_MakeEXL_Count] ('"+CStr(ExecuteDtStart)+"', '"+CStr(ExecuteDtEnd)+"','"&CodeGubun&"','"&makerid&"','"&code&"','"&alinkcode&"','"&blinkcode&"','"&onoffgubun&"','"&divcode&"','"&rackipgoyn&"','"&vPurchaseType&"','"&searchType&"','"&trim(searchText)&"','"&minusjumun&"') "

'response.write sqlStr & "<br>"
'response.end
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	FTotCnt = rsget(0)
END IF
rsget.close

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

	headLine = "�԰��ڵ�,�ֹ��ڵ�,����IDX,��������,����óID,����ó,ó����,������,�԰���,�Һ��ڰ�,���԰�,����,������,����,��Ÿ����"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
		sqlStr ="[db_summary].[dbo].[sp_Ten_IpgoList_MakeEXL_List] ('"&ExecuteDtStart&"','"&ExecuteDtEnd&"','"&CodeGubun&"','"&makerid&"','"&code&"','"&alinkcode&"','"&blinkcode&"','"&onoffgubun&"','"&divcode&"','"&rackipgoyn&"','"&vPurchaseType&"','"&searchType&"','"&trim(searchText)&"','"&minusjumun&"',"&(i+1)&","&PageSize&")"

		'response.write sqlStr & "<br>"
		'response.end
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
dbget.close() : response.end

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
	dim FOneItem

	set FOneItem = new CIpCulmasterItem

    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		FOneItem.Fcode           = arrList(1,intLoop)
		FOneItem.Fblinkcode      = arrList(2,intLoop)
		FOneItem.FpurchaseType   = arrList(3,intLoop)
		FOneItem.Fsocid          = arrList(4,intLoop)
		FOneItem.Fdivcode        = arrList(11,intLoop)
		FOneItem.Fexecutedt      = arrList(8,intLoop)
		FOneItem.Fscheduledt     = arrList(7,intLoop)
		FOneItem.Ftotalsellcash  = arrList(9,intLoop)
		FOneItem.Ftotalsuplycash = arrList(10,intLoop)
		FOneItem.Fcomment        = Replace(db2html(arrList(12,intLoop)), vbCrLf, "")
		FOneItem.Fsocname        = db2html(arrList(5,intLoop))
		FOneItem.Fchargename     = db2html(arrList(6,intLoop))
		FOneItem.ftotalitemno     = arrList(13,intLoop) '��ǰ �� ����
		FOneItem.Fprizecnt     = arrList(14,intLoop) '��ǰ ������
		FOneItem.fpurchasetypename     = arrList(15,intLoop)
		FOneItem.fppMasterIdx           = arrList(16,intLoop)

		bufstr = FOneItem.Fcode
		bufstr = bufstr & "," & FOneItem.Fblinkcode
		bufstr = bufstr & "," & FOneItem.fppMasterIdx
		bufstr = bufstr & "," & FOneItem.fpurchasetypename
		bufstr = bufstr & "," & FOneItem.Fsocid
		bufstr = bufstr & "," & replace(FOneItem.Fsocname,",",".")
		bufstr = bufstr & "," & FOneItem.Fchargename
		bufstr = bufstr & "," & FOneItem.Fscheduledt
		bufstr = bufstr & "," & FOneItem.Fexecutedt
		bufstr = bufstr & "," & FOneItem.Ftotalsellcash
		bufstr = bufstr & "," & FOneItem.Ftotalsuplycash
		bufstr = bufstr & "," & FOneItem.ftotalitemno
		bufstr = bufstr & "," & FOneItem.Fprizecnt
		bufstr = bufstr & "," & FOneItem.GetDivCodeName
		bufstr = bufstr & "," & FOneItem.Fcomment

        tFile.WriteLine bufstr
    Next
End function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
