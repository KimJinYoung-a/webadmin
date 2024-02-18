<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

'// 에러내자.
1

Dim iGLBSongjangDiv
iGLBSongjangDiv = CStr(getDefaultSongJangDiv(session("ssBctId")))

Function getDefaultSongJangDiv(iMakerid)
    dim sqlStr, ret
    ret = 0
    sqlstr = " select top 1 IsNULL(defaultsongjangdiv,0) as defaultsongjangdiv from db_partner.dbo.tbl_partner where id='"&iMakerid&"'"

    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF Not (rsget.EOF OR rsget.BOF) THEN
    	ret = rsget("defaultsongjangdiv")
    END IF
    rsget.Close
    getDefaultSongJangDiv = ret
end function

Class CSongJangItem
    public FDetailidx
    public FOrderserial
    public FSongjangDiv
    public FSongjangNo

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class

Function getFileExt(str)
	dim sp
	sp = split(str,".")
	getFileExt = sp(UBound(sp))
End Function

'//파일 확장자,크기 검사
function CheckCSVFiles(byval uprequest,ifile,imaxfilesize)
	dim file_name, file_size, file_mimetype, file_type

	file_name	= ifile.FileName '파일 이름 추출

	''file_size = ifile.FileLen '파일 사이즈 추출
	file_size = ifile.FileSize  '파일 사이즈 추출

	file_type = getFileExt(ifile.FileName)

	''file_mimetype = ifile.MimeType 		' 파일 mimetype 추출
	file_mimetype = ifile.ContentType  	' 파일 mimetype 추출

	'// 파일이 없을경우
	if (file_name="") then
		CheckCSVFiles=false
		exit function
	end if

	'//파일 사이즈가 크거나 작을 경우
	if ((file_size > imaxfilesize) or (file_size < 1))  then
    	response.write "<script type='text/javascript' language='javascript'>alert('파일사이즈 " + Formatnumber(imaxfilesize,0) + "Byte 보다 크거나 잘못된 파일 입니다.\n -" & file_size & "');</script>"
        CheckCSVFiles=false
        exit function
    end if

	'//마임타입이 틀릴경우
	'response.write file_mimetype
	If ((LCase(file_mimetype)<>"application/octet-stream") and (LCase(file_mimetype)<>"application/vnd.ms-excel")) then
    	response.write "<script type='text/javascript' language='javascript'>alert('CSV 화일만 지원됩니다.(2) \n- 올바른 파일이 아니거나 파일 형식이 다를 수 있습니다.\n - " & file_mimetype & "');</script>"
        CheckCSVFiles=false
        exit function
    end if

    IF (LCase(file_type)<>"csv") then
    	response.write "<script type='text/javascript' language='javascript'>alert('CSV 화일만 지원됩니다. \n- 올바른 파일이 아니거나 파일 형식이 다를 수 있습니다.\n - " & file_type & "');</script>"
        CheckCSVFiles=false
        exit function
    end if
    CheckCSVFiles=true

end function

function IsTopLine(ioneLine)
    IsTopLine = (Left(ioneLine,Len("일련번호"))="일련번호")
end function

' 정규식 함수
Function ReplaceText(str, patrn, repStr)
	Dim regEx
	Set regEx = New RegExp
	with regEx
		.Pattern = patrn
		.IgnoreCase = True
		.Global = True
	End with
	ReplaceText = regEx.Replace(str, repStr)
End Function

Function replaceDoublequatComma(oStr)
    dim OrgString : OrgString = oStr
    dim i, leng, pos1, pos2, RetStr
    leng = Len(OrgString)

    RetStr = OrgString
    for i=0 to leng-1
        pos1 = InStr(OrgString,chr(34))
        if (pos1>0) then
            pos2 = InStr(Mid(OrgString,pos1+1,1024),chr(34))

            if (pos2>0) then
                RetStr = Left(OrgString,pos1-1) + Replace(Mid(OrgString,pos1+1,pos2-1),",","") + Mid(OrgString,pos1+pos2+1,1024)
            end if
        end if
    next
    replaceDoublequatComma = RetStr
end function

Function replaceInComma(oStr)
    dim OrgString : OrgString = oStr
    dim i,maxloop , retStr
    maxloop = 8

    retStr = oStr
    For i=0 to maxloop-1
        if InStr(retStr,chr(34))>0 then
			retStr = replaceDoublequatComma(retStr)
        end if
    next
    replaceInComma = retStr
end function

''Data 유효성 검사
function IsValidLine(ioneLine,byref iValidArr)
    dim buf
    IsValidLine = False

    ''' "aaa,bb" 로 쌓여있는내역 "" 내부의 콤머를 지움
    ioneLine = replaceInComma(ioneLine)
	''rw ioneLine

    buf = split(ioneLine,",")

    if Not IsArray(buf) then Exit Function

    if UBound(buf)<>7 then Exit Function

    dim iDetailidx, iOrderserial, iSongjangDiv, iSongjangNo
    iDetailidx      = Trim(buf(0))
    iOrderserial    = Trim(buf(1))
    iSongjangDiv    = Trim(buf(6))
    iSongjangNo     = Trim(buf(7))
    iSongjangNo     = Replace(iSongjangNo,"-","")

    ''길이 Check
    if Len(iDetailidx)<7 then Exit Function
    if Len(iOrderserial)<>11 then Exit Function
    if Len(iSongjangDiv)<1 then Exit Function
    if Len(iSongjangNo)<7 or Len(iSongjangNo)>32 then Exit Function

    ''숫자 Check
    if Not IsNumeric(iDetailidx) then Exit Function
    if Not IsNumeric(iOrderserial) then Exit Function
    if Not IsNumeric(iSongjangDiv) then Exit Function
    if Not IsNumeric(iSongjangNo) then Exit Function

    ''ETC Check
    if (iSongjangDiv>99) or (iSongjangDiv<1) then Exit Function

    ''E+ 엑셀관련
    if InStr(iSongjangNo,"E+")>0 then Exit Function

    dim oArrLen
    if IsArray(iValidArr) then
        oArrLen = UBound(iValidArr)
    else
        oArrLen = -1
    end if

    if (oArrLen<0) then
        redim iValidArr(oArrLen+1)  '' ==0
    else
        redim preserve iValidArr(oArrLen+1)
    end if

    set iValidArr(oArrLen+1) = New CSongJangItem
    iValidArr(oArrLen+1).FDetailidx = iDetailidx
    iValidArr(oArrLen+1).FOrderserial = iOrderserial

    ''코코로박스 자꾸 이노지스택배로 등록함.. 강제로 기본 택배사로 변경
    if (LCASE(session("ssBctId"))="cocorobox") or (LCASE(session("ssBctId"))="cocoroboxdeco") or (LCASE(session("ssBctId"))="kamomekitchen") or (LCASE(session("ssBctId"))="emalia") then
        if (iGLBSongjangDiv<>"0") and (iGLBSongjangDiv<>"") then
            iSongjangDiv=iGLBSongjangDiv
        end if
    end if

    if (LCASE(session("ssBctId"))="loand") then
        if (iGLBSongjangDiv<>"0") and (iGLBSongjangDiv<>"") then
            iSongjangDiv=iGLBSongjangDiv
        end if
    end if

    if (LCASE(session("ssBctId"))="vintagevende") or (LCASE(session("ssBctId"))="vikinivender") then
        if (iGLBSongjangDiv<>"0") and (iGLBSongjangDiv<>"") then
            iSongjangDiv=iGLBSongjangDiv
        end if
    end if
    ''=============================================================================================

    iValidArr(oArrLen+1).FSongjangDiv = iSongjangDiv
    iValidArr(oArrLen+1).FSongjangNo = iSongjangNo

    IsValidLine = True
end function


Const CMAxLines=300
Dim DefaultPath
DefaultPath = server.MapPath("/designer/upcsvfile/")
dim Upload,i

'// 업로드 컨퍼넌트 선언 //
IF (application("Svr_Info")	= "Dev") then
	Set Upload = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
	Set Upload = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Upload.MaxBytesToAbort  = 5 * 1024 * 1024
Upload.Start DefaultPath '업로드경로

dim ret, uploadedFileName, saveFileName
ret = CheckCSVFiles(Upload, Upload("songjangfile"),1000*1024)

if (ret) then
    saveFileName = DefaultPath & "\" & "SJ" & Left(CStr(now()),10) & "_" & session.sessionid & "_" & session("ssBctID") & ".csv"
        uploadedFileName = Upload.Form("songjangfile").SaveAs(saveFileName, True)

    end if

    Set Upload = Nothing

    dim IsUploadErr

    IsUploadErr = Not ((ret) and (uploadedFileName<>""))


    dim iLines, iTotCnt, iSuccCnt, iFailCnt, TopLineExists, iFailLineStr, iBlankCount
    dim ValidArr

    iTotCnt     = 0
    iSuccCnt    = 0
    iFailCnt    = 0
		iBlankCount = 0
%>

<% if (IsUploadErr) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td align="center">
        업로드 오류
    </td>
</tr>
</table>
<% else %>
<%
    dim UploededStr, objFSO, objFile
    Set objFSO  = Server.CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(saveFileName,1)
    UploededStr = objFile.ReadAll()

    objFile.Close

    ''임시 파일 삭제
    ''objFSO.DeleteFile saveFileName

    Set objFile = Nothing
    Set objFSO  = Nothing

    iLines = split(UploededStr,vbCRLF)


    if IsArray(iLines) then
        iTotCnt = UBound(iLines)
        if (iTotCnt>CMAxLines) then iTotCnt=CMAxLines

        for i=0 to iTotCnt    ''빈줄도 검사
            if (i=0) then
                TopLineExists = IsTopLine(iLines(0))
            end if

            if ((i=0) and (TopLineExists)) then
                ''Skip
            elseif (iLines(i)="") then
                ''Skip
                iBlankCount = iBlankCount +1
            else
                if (IsValidLine(iLines(i),ValidArr)) then
                    iSuccCnt = iSuccCnt + 1
                else
                    iFailCnt = iFailCnt + 1
                    iFailLineStr = iFailLineStr + CStr(i+1) + ","
                end if
            end if
        Next

        ''' 탑라인은 뺌
        if (TopLineExists) then iTotCnt = iTotCnt - 1
        ''빈줄은 뺌
        iTotCnt = iTotCnt -iBlankCount
        iTotCnt = iTotCnt + 1

        ''response.write "iBlankCount="&iBlankCount
    end if

    if Right(iFailLineStr,1)="," then iFailLineStr=Left(iFailLineStr,Len(iFailLineStr)-1)

%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td>
    다음과 같이 송장 업로드 파일이 분석 되었습니다.<br>
    실제 업로드 된 건수가 아닌 CSV파일 포맷 검사입니다.<br>
    택배사코드나, 송장번호가 없는경우 실패 할 수 있습니다.<br>
    엑셀에서 <strong>송장번호가 <font color=red>6.8996E+12</font> 형태로 변환된 경우</strong> 수정해 주시기 바랍니다.<br>
    "송장입력진행" 버튼을 누르시면 송장 일괄 업로드가 실행 됩니다.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">

        총 업로드건수 : <%= iTotCnt %> <br>
        정상 건수 : <b><%= iSuccCnt %></b> <br>
        실패 건수 : <%= iFailCnt %> <br>
        <% if (iFailLineStr<>"") then %>
        실패 라인 : <%= iFailLineStr %> 번째 라인
        <% end if %>
    </td>
</tr>
<form name="frmSv" method="post" action="upchebeasong_Process.asp">
<input type="hidden" name="mode" value="SongjangInputCSV">
<% if IsArray(ValidArr) then %>
<% for i=0 to UBound(ValidArr) %>
    <input type="hidden" name="detailidxArr" value="<%= ValidArr(i).FDetailidx %>">
    <input type="hidden" name="orderserialArr" value="<%= ValidArr(i).FOrderserial %>">
    <input type="hidden" name="songjangdivArr" value="<%= ValidArr(i).FSongjangDiv %>">
    <input type="hidden" name="songjangnoArr" value="<%= ValidArr(i).FSongjangNo %>">
<% next %>
<% end if %>
</form>
<tr bgcolor="#FFFFFF">
    <td align="center"><input type="button" value="송장입력진행" onClick="regSongJangProc(frmSv)"></td>
</tr>
</table>
<% end if %>

<script language='javascript'>
function regSongJangProc(frm){
    var succCnt = <%= iSuccCnt %>;

    if (succCnt<1){
        alert('업로드할 정상 건수가 존재하지 않습니다. CSV포맷을 확인하세요.');
        return;
    }

    if (confirm('송장 일괄입력을 실행 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
