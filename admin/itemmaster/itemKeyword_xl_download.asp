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

dim makerid, sellyn, usingyn, itemid, itemname, keyword, itemidMn, itemidMx, searchKey

makerid = request("makerid")
sellyn = request("sellyn")
usingyn = request("usingyn")
itemid = request("itemid")
itemname = request("itemname")
keyword = request("keyword")
itemidMn = request("itemidMn")
itemidMx = request("itemidMx")
searchKey = request("searchKey")

if (makerid = "") and (itemidMn = "") and (itemidMx = "") and (searchKey = "") then
	response.write "잘못된 접속입니다."
	dbget.close
	response.end
end if




Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(Left(Now, 7), "-", "")
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

Dim FileName: FileName = "itemKeyword_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		headLine = "itemID,브랜드ID,상품명,키워드,판매여부,사용여부"

		bufstr = "" & arrList(1,intLoop) & ""
		bufstr = bufstr & "," & arrList(0,intLoop)
		bufstr = bufstr & ",""" & arrList(2,intLoop) & """"
		bufstr = bufstr & ",""" & arrList(3,intLoop) & """"
		bufstr = bufstr & "," & arrList(4,intLoop)
		bufstr = bufstr & "," & arrList(5,intLoop)

        tFile.WriteLine bufstr
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr = " db_item.dbo.usp_Ten_itemKeyword_MakeEXL_Count ('" + CStr(makerid) + "', '" + CStr(sellyn) + "', '" + CStr(usingyn) + "', " + CStr(itemid) + ", '" + CStr(itemname) + "', '" + CStr(keyword) + "', " + CStr(itemidMn) + ", " + CStr(itemidMx) + ", '" + CStr(searchKey) + "') "
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

	headLine = "itemID,브랜드ID,상품명,키워드,판매여부,사용여부"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""

		sqlStr ="db_item.dbo.usp_Ten_itemKeyword_MakeEXL_List ('" + CStr(makerid) + "', '" + CStr(sellyn) + "', '" + CStr(usingyn) + "', " + CStr(itemid) + ", '" + CStr(itemname) + "', '" + CStr(keyword) + "', " + CStr(itemidMn) + ", " + CStr(itemidMx) + ", '" + CStr(searchKey) + "'," & (i+1) & "," & PageSize & ")"
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

response.write FTotCnt&"건 생성 ["&FileName&"]"
response.redirect AdmPath&"/"&FileName
response.end

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
