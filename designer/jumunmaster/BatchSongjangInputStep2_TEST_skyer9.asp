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

'// ��������.
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

'//���� Ȯ����,ũ�� �˻�
function CheckCSVFiles(byval uprequest,ifile,imaxfilesize)
	dim file_name, file_size, file_mimetype, file_type

	file_name	= ifile.FileName '���� �̸� ����

	''file_size = ifile.FileLen '���� ������ ����
	file_size = ifile.FileSize  '���� ������ ����

	file_type = getFileExt(ifile.FileName)

	''file_mimetype = ifile.MimeType 		' ���� mimetype ����
	file_mimetype = ifile.ContentType  	' ���� mimetype ����

	'// ������ �������
	if (file_name="") then
		CheckCSVFiles=false
		exit function
	end if

	'//���� ����� ũ�ų� ���� ���
	if ((file_size > imaxfilesize) or (file_size < 1))  then
    	response.write "<script type='text/javascript' language='javascript'>alert('���ϻ����� " + Formatnumber(imaxfilesize,0) + "Byte ���� ũ�ų� �߸��� ���� �Դϴ�.\n -" & file_size & "');</script>"
        CheckCSVFiles=false
        exit function
    end if

	'//����Ÿ���� Ʋ�����
	'response.write file_mimetype
	If ((LCase(file_mimetype)<>"application/octet-stream") and (LCase(file_mimetype)<>"application/vnd.ms-excel")) then
    	response.write "<script type='text/javascript' language='javascript'>alert('CSV ȭ�ϸ� �����˴ϴ�.(2) \n- �ùٸ� ������ �ƴϰų� ���� ������ �ٸ� �� �ֽ��ϴ�.\n - " & file_mimetype & "');</script>"
        CheckCSVFiles=false
        exit function
    end if

    IF (LCase(file_type)<>"csv") then
    	response.write "<script type='text/javascript' language='javascript'>alert('CSV ȭ�ϸ� �����˴ϴ�. \n- �ùٸ� ������ �ƴϰų� ���� ������ �ٸ� �� �ֽ��ϴ�.\n - " & file_type & "');</script>"
        CheckCSVFiles=false
        exit function
    end if
    CheckCSVFiles=true

end function

function IsTopLine(ioneLine)
    IsTopLine = (Left(ioneLine,Len("�Ϸù�ȣ"))="�Ϸù�ȣ")
end function

' ���Խ� �Լ�
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

''Data ��ȿ�� �˻�
function IsValidLine(ioneLine,byref iValidArr)
    dim buf
    IsValidLine = False

    ''' "aaa,bb" �� �׿��ִ³��� "" ������ �޸Ӹ� ����
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

    ''���� Check
    if Len(iDetailidx)<7 then Exit Function
    if Len(iOrderserial)<>11 then Exit Function
    if Len(iSongjangDiv)<1 then Exit Function
    if Len(iSongjangNo)<7 or Len(iSongjangNo)>32 then Exit Function

    ''���� Check
    if Not IsNumeric(iDetailidx) then Exit Function
    if Not IsNumeric(iOrderserial) then Exit Function
    if Not IsNumeric(iSongjangDiv) then Exit Function
    if Not IsNumeric(iSongjangNo) then Exit Function

    ''ETC Check
    if (iSongjangDiv>99) or (iSongjangDiv<1) then Exit Function

    ''E+ ��������
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

    ''���ڷιڽ� �ڲ� �̳������ù�� �����.. ������ �⺻ �ù��� ����
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

'// ���ε� ���۳�Ʈ ���� //
IF (application("Svr_Info")	= "Dev") then
	Set Upload = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
	Set Upload = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Upload.MaxBytesToAbort  = 5 * 1024 * 1024
Upload.Start DefaultPath '���ε���

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
        ���ε� ����
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

    ''�ӽ� ���� ����
    ''objFSO.DeleteFile saveFileName

    Set objFile = Nothing
    Set objFSO  = Nothing

    iLines = split(UploededStr,vbCRLF)


    if IsArray(iLines) then
        iTotCnt = UBound(iLines)
        if (iTotCnt>CMAxLines) then iTotCnt=CMAxLines

        for i=0 to iTotCnt    ''���ٵ� �˻�
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

        ''' ž������ ��
        if (TopLineExists) then iTotCnt = iTotCnt - 1
        ''������ ��
        iTotCnt = iTotCnt -iBlankCount
        iTotCnt = iTotCnt + 1

        ''response.write "iBlankCount="&iBlankCount
    end if

    if Right(iFailLineStr,1)="," then iFailLineStr=Left(iFailLineStr,Len(iFailLineStr)-1)

%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
    <td>
    ������ ���� ���� ���ε� ������ �м� �Ǿ����ϴ�.<br>
    ���� ���ε� �� �Ǽ��� �ƴ� CSV���� ���� �˻��Դϴ�.<br>
    �ù���ڵ峪, �����ȣ�� ���°�� ���� �� �� �ֽ��ϴ�.<br>
    �������� <strong>�����ȣ�� <font color=red>6.8996E+12</font> ���·� ��ȯ�� ���</strong> ������ �ֽñ� �ٶ��ϴ�.<br>
    "�����Է�����" ��ư�� �����ø� ���� �ϰ� ���ε尡 ���� �˴ϴ�.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">

        �� ���ε�Ǽ� : <%= iTotCnt %> <br>
        ���� �Ǽ� : <b><%= iSuccCnt %></b> <br>
        ���� �Ǽ� : <%= iFailCnt %> <br>
        <% if (iFailLineStr<>"") then %>
        ���� ���� : <%= iFailLineStr %> ��° ����
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
    <td align="center"><input type="button" value="�����Է�����" onClick="regSongJangProc(frmSv)"></td>
</tr>
</table>
<% end if %>

<script language='javascript'>
function regSongJangProc(frm){
    var succCnt = <%= iSuccCnt %>;

    if (succCnt<1){
        alert('���ε��� ���� �Ǽ��� �������� �ʽ��ϴ�. CSV������ Ȯ���ϼ���.');
        return;
    }

    if (confirm('���� �ϰ��Է��� ���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
