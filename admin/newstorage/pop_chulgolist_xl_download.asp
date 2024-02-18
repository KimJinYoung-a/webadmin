<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 30
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 5000

dim fromDate, toDate

fromDate = request("fromDate")
toDate = request("toDate")


dim yyyymm
yyyymm = Left(Now, 7)


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


Class CIpCulmasterItem
	public Fid
	public Fcode
	public Fsocid
	public Fdivcode
	public Fexecutedt
	public Fscheduledt
	public Ftotalsellcash
	public Ftotalsuplycash
	public Fvatcode
	public Fchargeid
	public Fcomment
	public Findt
	public Fupdt
	public Fdeldt
	public Ftotalbuycash
	public Fsocname
	public Fchargename
	public Frackipgoyn

	public FBrandMaeipdiv

	public Falinkcode
	public Fblinkcode

	public FpurchaseType

	' �������. ��񿡼� �ϰ��� �����ؼ� ���� ������.
	public function GetPurchaseTypeName()
		Select Case FpurchaseType
			Case "1"
				GetPurchaseTypeName = "�Ϲ�����"
			Case "4"
				GetPurchaseTypeName = "����"
			Case "5"
				GetPurchaseTypeName = "OFF����"
			Case "6"
				GetPurchaseTypeName = "����"
			Case "8"
				GetPurchaseTypeName = "����"
			Case Else
				GetPurchaseTypeName = FpurchaseType
		End Select
	end function

	public function GetMinusColor(icash)
		if (icash<0) then
			GetMinusColor = "#EE3333"
		else
			GetMinusColor = "#000000"
		end if
	end function

	public function GetDivCodeColor()
		if Fdivcode="002" then
			GetDivCodeColor = "#000000"
		elseif Fdivcode="001" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="801" then
			GetDivCodeColor = "#DD5555"
		elseif Fdivcode="802" then
			GetDivCodeColor = "#5555DD"
		end if
	end function

	public function GetDivCodeName()
		if Fdivcode="002" then
			GetDivCodeName = "��Ź"
		elseif Fdivcode="001" then
			GetDivCodeName = "����"
		elseif Fdivcode="003" then
			GetDivCodeName = "����"
		elseif Fdivcode="004" then
			GetDivCodeName = "�ܺ�"
		elseif Fdivcode="005" then
			GetDivCodeName = "����"
		elseif Fdivcode="006" then
			GetDivCodeName = "B2B"
		elseif Fdivcode="007" then
			GetDivCodeName = "��Ÿ"
		elseif Fdivcode="101" then
			GetDivCodeName = "��Ź���"
		elseif Fdivcode="801" then
			GetDivCodeName = "Off����"
		elseif Fdivcode="802" then
			GetDivCodeName = "Off��Ź"
		elseif Fdivcode="999" then
			GetDivCodeName = "��Ÿ(�������)"
		end if
	end function

	public function GetBrandMaeipDivCodeName()
		if FBrandMaeipdiv="W" then
			GetBrandMaeipDivCodeName = "��Ź"
		elseif FBrandMaeipdiv="M" then
			GetBrandMaeipDivCodeName = "����"
		elseif FBrandMaeipdiv="U" then
			GetBrandMaeipDivCodeName = "��ü"
		end if
	end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class


Dim FileName: FileName = "ChulgoList_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
	dim FOneItem

	set FOneItem = new CIpCulmasterItem

    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		FOneItem.Fcode           = arrList(1,intLoop)
		FOneItem.Falinkcode      = arrList(2,intLoop)
		FOneItem.Fsocid          = arrList(3,intLoop)
		FOneItem.Fsocname        = db2html(arrList(4,intLoop))
		FOneItem.Fdivcode        = arrList(11,intLoop)
		FOneItem.Fexecutedt      = arrList(7,intLoop)
		FOneItem.Fscheduledt     = arrList(6,intLoop)
		FOneItem.Ftotalsellcash  = arrList(8,intLoop)
		FOneItem.Ftotalsuplycash = arrList(9,intLoop)
		FOneItem.Ftotalbuycash 	 = arrList(10,intLoop)
		FOneItem.Fcomment        = Replace(db2html(arrList(12,intLoop)), vbCrLf, "")
		FOneItem.Fchargename     = db2html(arrList(5,intLoop))

		bufstr = FOneItem.Fcode
		bufstr = bufstr & "," & FOneItem.Falinkcode
		bufstr = bufstr & "," & FOneItem.Fsocid
		bufstr = bufstr & "," & FOneItem.Fsocname
		bufstr = bufstr & "," & FOneItem.Fchargename
		bufstr = bufstr & "," & FOneItem.Fscheduledt
		bufstr = bufstr & "," & FOneItem.Fexecutedt
		bufstr = bufstr & "," & FOneItem.Ftotalsellcash
		bufstr = bufstr & "," & FOneItem.Ftotalsuplycash
		bufstr = bufstr & "," & FOneItem.Ftotalbuycash
		bufstr = bufstr & "," & FOneItem.GetDivCodeName
		bufstr = bufstr & "," & FOneItem.Fcomment

        tFile.WriteLine bufstr
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr = " [db_summary].[dbo].[sp_Ten_ChulgoList_MakeEXL_Count] ('" + CStr(fromDate) + "', '" + CStr(toDate) + "') "
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	FTotCnt = rsget(0)
END IF
rsget.close

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

	headLine = "����ڵ�,�ֹ��ڵ�,���óID,���ó��,ó����,������,�����,�Һ��ڰ�,���,���԰�,����,��Ÿ����"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
		sqlStr ="[db_summary].[dbo].[sp_Ten_ChulgoList_MakeEXL_List] ('" & fromDate & "','" & toDate & "'," & (i+1) & "," & PageSize & ")"
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

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
