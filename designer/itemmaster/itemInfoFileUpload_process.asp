<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 품목정보 일괄변경 Excel 처리
' Hieditor : 2012.10.25 허진원 생성
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
Dim infoDiv
dim xIfDiv, xIfCd		'파일내 코드들

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
infoDiv = uploadform.form("infoDiv")
if infoDiv="" or isNull(infoDiv) then
    response.write "<script type='text/javascript'>alert('선택된 품목유형이 없습니다.\n품목유형을 선택하시고 다시 시도해주세요.');history.back();</script>"
    response.end
end if

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
dim ret : ret = fnGetXLFileArray(xlRowALL, sFilePath, xIfDiv, xIfCd)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="외부 테이블 형식이 잘못되었습니다.") then
        response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

if infoDiv<>trim(xIfDiv) then
    response.write "<script type='text/javascript'>alert('선택된 품목유형과 파일의 내용의 품목유형이 다릅니다.\n선택하신 품목유형과 업로드하신 파일을 확인해주세요.');history.back();</script>"
    response.end
end if

if Not(isArray(xIfCd)) then
    response.write "<script type='text/javascript'>alert('파일이 올바르지 않습니다. \n업로드 양식을 다운받아 내용을 올바르게 작성해주십시요.');history.back();</script>"
    response.end
end if

''데이터 처리.
dim lp, iLine, j, chkDiv, strSql, itemid, cntSc, msgFl, safetyYn, safetyDiv, safetyNum, chkCnt, delSql
cntSc=0: msgFl=""
dim isBlankCont
dim isYnFieldErr
dim isSafrtyErr

for lp=1 to ubound(xlRowALL)
	Set iLine = xlRowALL(lp)

	itemid =  iLine.FItemArray(0)
	if isNumeric(itemid) then
		if ubound(iLine.FItemArray)=ubound(xIfCd) then
			'//브랜드 상품인지 확인
			strSql = "Select count(itemid) from db_item.dbo.tbl_item where itemid='" & itemid & "' and makerid='" & session("ssBctId") & "'"
			rsget.Open strSql,dbget,1
				chkCnt = rsget(0)
			rsget.Close
			
			if chkCnt>0 then
				
				
				'On Error Resume Next
				strSql="": safetyYn="N": safetyDiv="": safetyNum=""
				isBlankCont  = false
				isYnFieldErr = false
				isSafrtyErr  = false
				for j=1 to ubound(xIfCd)
					if not(xIfCd(j)="code" or xIfCd(j)="sYn" or xIfCd(j)="sDiv" or xIfCd(j)="sNum") then
					    if (xIfCd(j)<>"") then  ''빈값이 있는경우 있음..
    						'// 상품품목고시 정보 처리
    						if left(xIfCd(j),1)="C" then
    							chkDiv = Trim(iLine.FItemArray(j))
    							chkDiv = UCASE(chkDiv)
    							j=j+1
    							
    							if (chkDiv<>"Y") and (chkDiv<>"N") then 
    							    chkDiv="N"
    							    isYnFieldErr=true    ''20121114 추가.
    							end if
    						else
    							chkDiv = "N"
    						end if
    		
    						strSql = strSql & "Insert into db_item.dbo.tbl_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
    						strSql = strSql & "('" & itemid & "'"
    						strSql = strSql & ",'" & xIfCd(j) & "'"
    						strSql = strSql & ",'" & chkDiv & "'"
    						strSql = strSql & ",'" & html2db(trim(iLine.FItemArray(j))) & "')" & vbCrLf
    						
    						if (trim(iLine.FItemArray(j))="") then isBlankCont=true
    				    end if
						'rw strSql
					elseIf xIfCd(j)="sYn" then	'//안정인증정보 처리
						safetyYn = html2db(trim(iLine.FItemArray(j)))
						safetyYn = replace(replace(replace(safetyYn,VbCRLF,""),VbCr,""),VbLf,"")
						safetyYn = UCASE(safetyYn)
					elseIf xIfCd(j)="sDiv" then
						safetyDiv = html2db(trim(iLine.FItemArray(j)))
						if (safetyYn="Y") and (safetyDiv<>"10") and (safetyDiv<>"20") and (safetyDiv<>"30") and (safetyDiv<>"40") and (safetyDiv<>"50") then
						    isSafrtyErr = true
						end if
						if (safetyYn="N") then safetyDiv="0"
					elseIf xIfCd(j)="sNum" then
						safetyNum = html2db(trim(iLine.FItemArray(j)))
						if (safetyYn="Y") and (safetyNum="") then
						    isSafrtyErr = true
						end if
					end if
				next
				
				if (isBlankCont) then
				    msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(항목누락)"
				elseif (isYnFieldErr) then
				    msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(Y/N필드오류)"
				elseif (isSafrtyErr) then
				    msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(안전인증필드오류)"
				else
				    '//상품품목정보 리셋 (정상적인경우만 리셋)
    				delSql = "Delete from  db_item.dbo.tbl_item_infoCont where itemid='" & itemid & "'"
    				dbget.Execute delSql
				
    				dbget.Execute strSql
    	            IF (application("Svr_Info")	= "Dev") then
    	                rw strSql
    	            end if
    	            
    				strSql = "Update db_item.dbo.tbl_item_Contents Set infoDiv='" & infoDiv & "', safetyYn='" & safetyYn & "', safetyDiv='" & safetyDiv & "', safetyNum=convert(varchar(30),'" & safetyNum & "') where itemid='" & itemid & "'"
    				dbget.Execute strSql
    	            
    	            IF (application("Svr_Info")	= "Dev") then
    	                rw strSql
    	            end if
    	            
    	            '''2012/11/09추가
    	            strSql = " update c"&VbCRLF
                    strSql = strSql&" set infoContent= '텐바이텐 고객행복센터 1644-6030'"&VbCRLF
                    strSql = strSql&" from db_item.dbo.tbl_item_infoCont c"&VbCRLF
                    strSql = strSql&" where convert(varchar(28),c.infoContent)= '텐바이텐 고객행복센터 1644-6'"&VbCRLF
                    strSql = strSql&" and c.infoContent<>'텐바이텐 고객행복센터 1644-6030'"&VbCRLF
                    strSql = strSql&" and c.itemid='"&itemid&"'"&VbCRLF
                    dbget.Execute strSql
                    
    				IF (ERR) then
    				    IF (application("Svr_Info")	= "Dev") then
    				        msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "["&replace(Err.Description,"'","")&"]"&"(입력내용)"
    				    ELSE
        					msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(입력내용)"
        				END IF
    				else
    					cntSc = cntSc+1
    				End if
    		    end if
				'On Error Goto 0
			else
				msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(없는상품)"
			end if
		else
			msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(항목누락)"
		end if
	else
		msgFl = msgFl & chkIIF(msgFl<>"",", ","") & itemid & "(잘못된 상품코드)"
	end if

	Set iLine = Nothing
next

IF msgFl<>"" then
    msgFl = msgFl & " 오류.\n\n오류건 제외 "
end if

IF (application("Svr_Info")	= "Dev") then
    rw msgFl
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

Function fnGetXLFileArray(byref xlRowALL, sFilePath, byref xIfDiv, byref xIfCd)
    Dim conDB, Rs, strQry, iResult, i, j, iObj, iArrayLen
    Dim irowObj, strTable
    '' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.

    Set conDB = Server.CreateObject("ADODB.Connection")
    conDB.Provider = "Microsoft.Jet.oledb.4.0"
    'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
    conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;HDR=NO;IMEX=1"		'첫행까지 데이터(HDR), 필드속성무시(IMEX;숫자/텍스트)

    'On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
			'/이유를 알수 없는 서버단 에러남. "예기치 않은 오류. 외부 개체에 트랩 가능한 오류(C0000005)가 발생했습니다. 스크립트를 계속 실행할 수 없습니다"
            Set conDB = Nothing
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

				if j=0 then
					'# 품목코드 값 접수
					xIfDiv = cStr(null2blank(Rs(0)))
				elseif j=2 then
					'# 품목항목코드 접수
	                redim xIfCd(iArrayLen)
	                For i=0 to iArrayLen
	                    xIfCd(i) = cStr(null2blank(Rs(i)))
	                Next
				elseif j>=4 then
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