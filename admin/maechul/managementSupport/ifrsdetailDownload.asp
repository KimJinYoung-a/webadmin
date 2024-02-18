<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  IFRS15-마일리지 출고일기준 안분 download
' History : 2020/05/06 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/combine_point_deposit_cls.asp" -->

<%
Server.ScriptTimeOut = 60*20		' 20분

Const MaxPage   = 50 
Const PageSize = 20000 

dim yyyymm : yyyymm=requestcheckvar(request("yyyymm"),7)
dim onoff : onoff=requestcheckvar(request("onoff"),10)


Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
Dim appPath : appPath = server.mappath(AdmPath) 
Dim pDownFile : pDownFile = appPath&"\ifrsmilePreDownFile.txt"

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

Dim FileName: FileName = "ifrs_mile_"&replace(yyyymm,"-","")&"_"&sDateName&".csv"
Dim fso, tFile

Dim sqlStr,ArrRows, iTotCnt, iTotPage

sqlStr ="exec [db_dataSummary].[dbo].[usp_TEN_IFRS15_Get_MileAnbunList_CNT] '"&yyyymm&"','"&onoff&"'"
db3_rsget.CursorLocation = adUseClient
db3_dbget.CommandTimeout = 60*10     ' 10분. 타임아웃 오류가 계속 나서 시간늘림
db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
	iTotCnt = db3_rsget(0)
END IF
db3_rsget.close

Dim bufstr1, i, bufline

' rw appPath & FileName
' rw appPath
' response.end

IF iTotCnt > 0 THEN
    iTotPage = CLNG(iTotCnt/PageSize)
    IF iTotPage<>(iTotCnt/PageSize) THEn iTotPage=iTotPage+1
    IF (iTotPage>MaxPage) THEn
		iTotPage=MaxPage
		iTotCnt=iTotPage*PageSize
	ENd IF

    ''기존파일 삭제.
    Call CheckPFileDelete(pDownFile)

    Set fso = CreateObject("Scripting.FileSystemObject")
        If NOT fso.FolderExists(appPath) THEN
			fso.CreateFolder(appPath)
		END If

	Set tFile = fso.CreateTextFile(appPath &"\"& FileName )


	bufstr1 = "구분"& "," &"매출처"& "," &"주문번호"& "," &"과세구분"& "," &"매입구분"& "," &"브랜드"& "," &"상품코드"& "," &"옵션코드"& "," &"수량"& "," &"매출총액"& "," &"업체정산액"& "," &"회계매출"& "," &"출고일"& "," &"사용마일리지"
	
	tFile.WriteLine bufstr1

    For i=0 to iTotPage-1
        ArrRows = ""
       
        sqlStr ="exec [db_dataSummary].[dbo].[usp_TEN_IFRS15_Get_MileAnbunList] "&i+1&","&PageSize&",'"&yyyymm&"','"&onoff&"'"
        db3_rsget.CursorLocation = adUseClient
        db3_dbget.CommandTimeout = 120 
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly

        IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
        	ArrRows = db3_rsget.getRows()
        END IF
        db3_rsget.close

        if isArray(ArrRows) then
            CALL WriteMileFile(tFile,ArrRows)
        end if

        
    NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing

    ''파일위치저장
    Call WriteOneLineToFile(pDownFile,appPath &"\"& FileName)

    response.write FormatNumber(iTotCnt,0)&"건 생성 ["&FileName&"]"
    response.write "<br><br><a href='"&AdmPath&"/"&FileName&"'>다운로드</a>"
    
else

    response.write "ERR: 건수 0"
END IF

Function WriteMileFile(tFile, arrList )
    Dim intLoop,iRow
    Dim bufstr
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
        bufstr = arrList(0,intLoop) & ","               ''구분
        bufstr = bufstr & arrList(1,intLoop) & ","      ''매출처
        bufstr = bufstr & arrList(2,intLoop) & ","      ''주문번호
        bufstr = bufstr & arrList(3,intLoop) & ","      ''과세구분
        bufstr = bufstr & arrList(4,intLoop) & ","      ''매입구분
        bufstr = bufstr & arrList(5,intLoop) & ","      ''브랜드
        bufstr = bufstr & arrList(6,intLoop) & ","  ''상품코드    
        bufstr = bufstr & "'"&arrList(7,intLoop) & ","      ''옵션코드
        bufstr = bufstr & arrList(8,intLoop) & "," 
        bufstr = bufstr & arrList(9,intLoop) & "," 
        bufstr = bufstr & arrList(10,intLoop) & "," 
        bufstr = bufstr & arrList(11,intLoop) & "," 
        bufstr = bufstr & arrList(12,intLoop) & "," 
        bufstr = bufstr & arrList(13,intLoop) 
        tFile.WriteLine bufstr
    Next

end function

function CheckPFileDelete(filePath)
    CheckPFileDelete = 0
    dim pfilePath : pfilePath=ReadOneLine(filePath)
    if (pfilePath<>"") then
        call DeleteExistFile(filePath)

        call DeleteExistFile(pfilePath)    
    end if

end function

function WriteOneLineToFile(filePath,oneline)
    Dim fso, tFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tFile = fso.CreateTextFile(filePath)
    tFile.WriteLine oneline
    tFile.Close
	Set tFile = Nothing
    Set fso = Nothing
end function 


function ReadOneLine(filePath)
  Dim fso, tFile
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filePath) Then
    set tFile = fso.OpenTextFile(filePath, 1, false)
    ReadOneLine = tFile.ReadLine
    tFile.Close
	Set tFile = Nothing
  end if
  Set fso = Nothing
end function


Function DeleteExistFile(filePath)
  Dim fso, result
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filePath) Then
    fso.DeleteFile(filePath) 
    result = 1
  Else
    result = 0
  End If
  Set fso = Nothing
  DeleteExistFile = result
End Function

%>

<!-- #include virtual="/lib/db/db3close.asp" -->
