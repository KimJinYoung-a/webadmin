<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품해외배송정보 일괄변경 Excel 업로드
' Hieditor : 2016.06.03 정윤정 생성
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
dim xIfCd		'파일내 코드들

iMaxLen	= 10 * 1024 * 1024	'업로드 파일 제한용량(Byte)


'업로드 컨퍼넌트 선언
IF (application("Svr_Info")	= "Dev") then
    Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

'파일 컨퍼넌트 선언
Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
sDefaultPath   = Server.MapPath("\designer\etc\infoUploadFiles\")

uploadform.Start sDefaultPath '업로드경로

monthFolder = Replace(Left(CStr(now()),10),"-","")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'파일체크

    '폴더 생성
    sFolderPath = sDefaultPath&"/"&monthFolder&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    '파일저장
	sFile = fnMakeFileName(uploadform("sFile"))
	''sFilePath = sFolderPath&sFile
	sFilePath = sFolderPath&replace(session("ssBctId"),"&","")&"_"&sFile  ''브랜드구분 쉽게.
	sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

	orgFileName = uploadform("sFile").FileName

END IF

Set objfile		= Nothing
Set uploadform = Nothing

Dim xlRowALL

'엑셀파일 파징
dim ret : ret = fnGetXLFileArray(xlRowALL, sFilePath, xIfCd)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="외부 테이블 형식이 잘못되었습니다.") then
        response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

if Not(isArray(xIfCd)) then
    response.write "<script type='text/javascript'>alert('파일이 올바르지 않습니다. \n업로드 양식을 다운받아 내용을 올바르게 작성해주십시요.');history.back();</script>"
    response.end
end if

''데이터 처리.
dim lp, iLine, j, strSql, itemid, cntSc, msgFl, deliverOverseas,  itemweight, chkCnt, delSql
cntSc=0: msgFl=""
dim isBlankCont
dim isYnFieldErr,isNumErr

for lp=1 to ubound(xlRowALL)
	Set iLine = xlRowALL(lp)

	itemid =  iLine.FItemArray(0) 
	if isNumeric(itemid) then
		if ubound(iLine.FItemArray)=ubound(xIfCd) then
			'//브랜드 상품인지 확인
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
			 
					If xIfCd(j)="sYn" then	'//해외배송여부
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
				    msgFl = msgFl & itemid & "(상품무게 미입력오류)\n" 
				elseif (isYnFieldErr) then
				    msgFl = msgFl &  itemid & "(해외배송여부 Y/N필드오류)\n"
				elseif (isNumErr) then
					    msgFl = msgFl &  itemid & "(상품무게 숫자외의 문자포함-숫자만 입력가능)\n"  
				else
    				strSql = "Update db_item.dbo.tbl_item  Set deliverOverseas='" & deliverOverseas & "' , itemweight= '" & itemweight & "'  where itemid='" & itemid & "'"
    	  	 	dbget.Execute strSql

    	            IF (application("Svr_Info")	= "Dev") then
    	                rw strSql
    	            end if

    				IF (ERR) then
    				    IF (application("Svr_Info")	= "Dev") then
    				        msgFl = msgFl &itemid & "["&replace(Err.Description,"'","")&"]"&"(입력내용)\n"
    				    ELSE
        					msgFl = msgFl & itemid & "(입력내용)\n"
        				END IF
    				else
    					cntSc = cntSc+1
    				End if
    		    end if
				'On Error Goto 0
			else
				msgFl = msgFl &  itemid & "(없는상품)\n"
			end if
		else
			msgFl = msgFl &  itemid & "(필드오류)"''&ubound(iLine.FItemArray)&","&ubound(xIfCd)&VbLf
		end if
	else
		msgFl = msgFl &  itemid & "(잘못된 상품코드)\n"
	end if

	Set iLine = Nothing
next

IF msgFl<>"" then
    msgFl = msgFl & " 오류.  > 오류건 제외 \n"
end if

IF (application("Svr_Info")	= "Dev") then
    rw msgFl
    ' response.write "<script type='text/javascript'>alert('"& replace(replace(replace(msgFl,VbCRLF,""),VbCr,""),VbLf,"") & " 총 " & cntSc&"건 정상 처리되었습니다.'); history.back();</script>"
ELSE
    response.write "<script type='text/javascript'>alert('"& replace(replace(replace(msgFl,VbCRLF,""),VbCr,""),VbLf,"") & " 총 " & cntSc&"건 정상 처리되었습니다.'); history.back();</script>"
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
    '' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.

    Set conDB = Server.CreateObject("ADODB.Connection")
    conDB.Provider = "Microsoft.Jet.oledb.4.0"
    'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
    ''conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"		'첫행까지 데이터(HDR), 필드속성무시(IMEX;숫자/텍스트)
    conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"	'' 1.17038e+006 반환되는 CASE ''2014/11/24 검토필요
    
    'On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
			'/이유를 알수 없는 서버단 에러남. "예기치 않은 오류. 외부 개체에 트랩 가능한 오류(C0000005)가 발생했습니다. 스크립트를 계속 실행할 수 없습니다"
			set conDB = nothing
            exit function
        End if
    'On Error Goto 0

    '' get First Sheet Name=============''시트가 여러개인경우 오류날 수 있음.
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
					'# 품목항목코드 접수
	                redim xIfCd(iArrayLen)
	                For i=0 to iArrayLen
	                    xIfCd(i) = cStr(null2blank(Rs(i)))
	                Next
				elseif j>=3 then
					'# 품목내용 접수
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