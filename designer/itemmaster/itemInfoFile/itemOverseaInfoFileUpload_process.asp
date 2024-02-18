<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ�ؿܹ������ �ϰ����� Excel ���ε�
' Hieditor : 2016.06.03 ������ ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<%
Dim uploadform, objfile, sDefaultPath
Dim iMaxLen,sFolderPath, monthFolder, sFile, sFilePath, orgFileName
dim xIfCd		'���ϳ� �ڵ��

iMaxLen	= 10 * 1024 * 1024	'���ε� ���� ���ѿ뷮(Byte)


'���ε� ���۳�Ʈ ����
IF (application("Svr_Info")	= "Dev") then
    Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

'���� ���۳�Ʈ ����
Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
sDefaultPath   = Server.MapPath("\designer\etc\infoUploadFiles\")

uploadform.Start sDefaultPath '���ε���

monthFolder = Replace(Left(CStr(now()),10),"-","")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'����üũ

    '���� ����
    sFolderPath = sDefaultPath&"/"&monthFolder&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    '��������
	sFile = fnMakeFileName(uploadform("sFile"))
	''sFilePath = sFolderPath&sFile
	sFilePath = sFolderPath&replace(session("ssBctId"),"&","")&"_"&sFile  ''�귣�屸�� ����.
	sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

	orgFileName = uploadform("sFile").FileName

END IF

Set objfile		= Nothing
Set uploadform = Nothing

Dim xlRowALL

'�������� ��¡
dim ret : ret = fnGetXLFileArray(xlRowALL, sFilePath, xIfCd)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('������ �ùٸ��� �ʰų� ������ �����ϴ�. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="�ܺ� ���̺� ������ �߸��Ǿ����ϴ�.") then
        response.write "<script>alert('�������� Save As Excel 97 -2003 ���չ��� ���·� ������ ����ϼ���.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

if Not(isArray(xIfCd)) then
    response.write "<script type='text/javascript'>alert('������ �ùٸ��� �ʽ��ϴ�. \n���ε� ����� �ٿ�޾� ������ �ùٸ��� �ۼ����ֽʽÿ�.');history.back();</script>"
    response.end
end if

''������ ó��.
dim lp, iLine, j, strSql, itemid, cntSc, msgFl, deliverOverseas,  itemweight, chkCnt, delSql
cntSc=0: msgFl=""
dim isBlankCont
dim isYnFieldErr,isNumErr

for lp=1 to ubound(xlRowALL)
	Set iLine = xlRowALL(lp)

	itemid =  iLine.FItemArray(0) 
	if isNumeric(itemid) then
		if ubound(iLine.FItemArray)=ubound(xIfCd) then
			'//�귣�� ��ǰ���� Ȯ��
			strSql = "Select count(itemid) from db_item.dbo.tbl_item where itemid='" & itemid &"' and makerid='" & session("ssBctId") & "'"  
			rsget.Open strSql,dbget,1
			if not rsget.eof then
				chkCnt = rsget(0)
			end if
			rsget.Close

			if chkCnt>0 then


				'On Error Resume Next
				strSql="": deliverOverseas="N":   itemweight=""
				isBlankCont  = false
				isYnFieldErr = false
				isNumErr =false
				
				for j=1 to ubound(xIfCd)
			 
					If xIfCd(j)="sYn" then	'//�ؿܹ�ۿ���
						deliverOverseas = html2db(trim(iLine.FItemArray(j)))
						deliverOverseas = replace(replace(replace(deliverOverseas,VbCRLF,""),VbCr,""),VbLf,"")
						deliverOverseas = UCASE(deliverOverseas)
						if Not(deliverOverseas="Y" or deliverOverseas="N") then isYnFieldErr = true
						
					elseIf xIfCd(j)="iW" then
						itemweight = Cint(trim(replace(replace(replace(iLine.FItemArray(j),"g",""),"(",""),")",""))) 
 
						if not isNumeric(itemweight) then 
							isNumErr = true
						end if
						if (deliverOverseas="Y") and (itemweight="" or itemweight="0") then
						    isBlankCont = true 
						end if
					end if
				next
	 
				if (isBlankCont) then
				    msgFl = msgFl & itemid & "(��ǰ���� ���Է¿���)\n" 
				elseif (isYnFieldErr) then
				    msgFl = msgFl &  itemid & "(�ؿܹ�ۿ��� Y/N�ʵ����)\n"
				elseif (isNumErr) then
					    msgFl = msgFl &  itemid & "(��ǰ���� ���ڿ��� ��������-���ڸ� �Է°���)\n"  
				else
    				strSql = "Update db_item.dbo.tbl_item  Set deliverOverseas='" & deliverOverseas & "' , itemweight= '" & itemweight & "'  where itemid='" & itemid & "'"
    	  	 	dbget.Execute strSql

    	            IF (application("Svr_Info")	= "Dev") then
    	                rw strSql
    	            end if

    				IF (ERR) then
    				    IF (application("Svr_Info")	= "Dev") then
    				        msgFl = msgFl &itemid & "["&replace(Err.Description,"'","")&"]"&"(�Է³���)\n"
    				    ELSE
        					msgFl = msgFl & itemid & "(�Է³���)\n"
        				END IF
    				else
    					cntSc = cntSc+1
    				End if
    		    end if
				'On Error Goto 0
			else
				msgFl = msgFl &  itemid & "(���»�ǰ)\n"
			end if
		else
			msgFl = msgFl &  itemid & "(�ʵ����)"''&ubound(iLine.FItemArray)&","&ubound(xIfCd)&VbLf
		end if
	else
		msgFl = msgFl &  itemid & "(�߸��� ��ǰ�ڵ�)\n"
	end if

	Set iLine = Nothing
next

IF msgFl<>"" then
    msgFl = msgFl & " ����.  > ������ ���� \n"
end if

IF (application("Svr_Info")	= "Dev") then
    rw msgFl
    ' response.write "<script type='text/javascript'>alert('"& replace(replace(replace(msgFl,VbCRLF,""),VbCr,""),VbLf,"") & " �� " & cntSc&"�� ���� ó���Ǿ����ϴ�.'); history.back();</script>"
ELSE
    response.write "<script type='text/javascript'>alert('"& replace(replace(replace(msgFl,VbCRLF,""),VbCr,""),VbLf,"") & " �� " & cntSc&"�� ���� ó���Ǿ����ϴ�.'); history.back();</script>"
END IF

'-- Functions --------------------------------------------------------------------------------
Class TXLRowObj
    public FItemArray

    public function setArrayLength(ln)
        Redim FItemArray(ln)
    end function
End Class

function IsSKipRow(ixlRow, skipCol0Str)
    if Not IsArray(ixlRow) then
        IsSKipRow = true
        Exit function
    end if

    if  LCASE(ixlRow(0))=LCASE(skipCol0Str) then
        IsSKipRow = true
        Exit function
    end if

    IsSKipRow = false
end function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, byref xIfCd)
    Dim conDB, Rs, strQry, iResult, i, j, iObj, iArrayLen
    Dim irowObj, strTable
    '' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.

    Set conDB = Server.CreateObject("ADODB.Connection")
    conDB.Provider = "Microsoft.Jet.oledb.4.0"
    'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
    ''conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"		'ù����� ������(HDR), �ʵ�Ӽ�����(IMEX;����/�ؽ�Ʈ)
    conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"	'' 1.17038e+006 ��ȯ�Ǵ� CASE ''2014/11/24 �����ʿ�
    
    'On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
			'/������ �˼� ���� ������ ������. "����ġ ���� ����. �ܺ� ��ü�� Ʈ�� ������ ����(C0000005)�� �߻��߽��ϴ�. ��ũ��Ʈ�� ��� ������ �� �����ϴ�"
			set conDB = nothing
            exit function
        End if
    'On Error Goto 0

    '' get First Sheet Name=============''��Ʈ�� �������ΰ�� ������ �� ����.
    Set Rs = conDB.OpenSchema(adSchemaTables)

    IF Not Rs.Eof Then
        strTable = Rs.Fields("table_name").Value
        ''rw "strTable="&strTable
    ENd IF
    Set Rs = Nothing
    ''==================================

    Set Rs = Server.CreateObject("ADODB.Recordset")

    ''strQry = "Select * From [sheet1$]"
    strQry = "Select * From ["&strTable&"]"

    ReDim xlRowALL(0)
    fnGetXLFileArray = true

	'On Error Resume Next
    Rs.Open strQry, conDB
        IF (ERR) then
            fnGetXLFileArray=false
            Rs.Close
            Set Rs = Nothing
            Set conDB = Nothing
            exit function
        End if

        j = 0
        If Not Rs.Eof Then
            Do Until Rs.Eof
                IF (ERR) then
                    fnGetXLFileArray=false
                    Rs.Close
                    Set Rs = Nothing
                    Set conDB = Nothing
                    exit function
                End if
                iArrayLen = rs.Fields.count-1

                set irowObj = new TXLRowObj
                irowObj.setArrayLength(iArrayLen)

				if j=1 then
					'# ǰ���׸��ڵ� ����
	                redim xIfCd(iArrayLen)
	                For i=0 to iArrayLen
	                    xIfCd(i) = cStr(null2blank(Rs(i)))
	                Next
				elseif j>=3 then
					'# ǰ�񳻿� ����
	                For i=0 to iArrayLen
						irowObj.FItemArray(i) = cStr(null2blank(Rs(i)))
	                    ''rw irowObj.FItemArray(i)
	                Next

	                IF (Not IsSKipRow(irowObj.FItemArray,"")) then
	                    ReDim Preserve xlRowALL(UBound(xlRowALL)+1)

	                    set xlRowALL(UBound(xlRowALL)) =  irowObj
	                    ''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow
	                END IF
				end if

                set irowObj = Nothing
                Rs.MoveNext
                j = j + 1
            Loop
       else
          fnGetXLFileArray=false
       end if

       ''''On Error Goto 0
        IF (ERR) then
            fnGetXLFileArray=false
        End if
    Rs.Close
	'On Error Goto 0

    Set Rs = Nothing
    Set conDB = Nothing

    if Ubound(xlRowALL)< 1 then fnGetXLFileArray=false

End Function
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->