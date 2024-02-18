<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim uploadform, objfile, sDefaultPath, sFolderPath ,monthFolder
Dim sFile, sFilePath, iMaxLen, orgFileName, maybeSheetName
Dim sCode, egCode, eCode
monthFolder = Replace(Left(CStr(now()),7),"-","")

'소숫점포함 숫자여부 체크
'-----------------------------------------
Function FIsNum(ByVal iValue)
    Dim iLength , i , retValue
    For i = 1 To Len(iValue)
     If Not (( asc(Mid(iValue,i,1)) > 47 And asc(Mid(iValue,i,1)) < 58 ) or asc(Mid(iValue,i,1))  = 46 ) Then
       FIsNum  = False
            Exit For
     else
        FIsNum = True
    end if
    Next


End Function
'-----------------------------------------
If (application("Svr_Info")	= "Dev") Then
    Set uploadform = Server.CreateObject("TABS.Upload")			'' - TEST : TABS.Upload
Else
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
End If

Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	sDefaultPath = Server.MapPath("/admin/shopmaster/sale/upFiles/")
	uploadform.Start sDefaultPath '업로드경로

	iMaxLen = uploadform.Form("iML")	'이미지파일크기
	sCode	= uploadform.Form("sC")		'세일 코드
	eCode	= uploadform.Form("eC")		'이벤트 코드
	egCode	= uploadform.Form("egC")	'이벤트 그룹 코드

	If (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) Then	'파일체크
	    sFolderPath = sDefaultPath&"/"
	    If NOT  objfile.FolderExists(sFolderPath) Then
	    	objfile.CreateFolder sFolderPath
	    End If

	    sFolderPath = sDefaultPath&"/"&monthFolder&"/"
	    If NOT  objfile.FolderExists(sFolderPath) Then
	    	objfile.CreateFolder sFolderPath
	    End If

		sFile = fnMakeFileName(uploadform("sFile"))
		sFilePath = sFolderPath&sFile
		sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

		orgFileName = uploadform("sFile").FileName
		maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")
	End If

Set objfile = Nothing
Set uploadform = Nothing

Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i
''''	0			1				2				3
''''상품코드, 	할인판매가, 	할인율,			할인마진율
xlPosArr = Array(0,1,2,3)
ArrayLen = UBound(xlPosArr)
skipString = "Sheet1"
afile = sFilePath
aSheetName = ""

Dim xlRowALL
Dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)
If (Not ret) or (Not IsArray(xlRowALL)) then
	response.write "<script>alert('파일이 올바르지 않거나 내용이 없습니다. "&Replace(Err.Description,"'","")&"');</script>"

	If (Err.Description="외부 테이블 형식이 잘못되었습니다.") Then
		response.write "<script>alert('엑셀에서 Save As Excel 97 -2003 통합문서 형태로 저장후 사용하세요.');</script>"
	End If
	response.write "<script>history.back();</script>"
	response.end
End If

''데이터 처리.
Dim iLine, iResult, strSql, addSqlDB, addSql, RetErr
Dim errCNT, objCmd, strStatus, errItem
Dim pcnt : pcnt = UBound(xlRowALL)
Dim itemArr, itemidarr, arrList, intLoop
Dim dSDate, dEDate, iSaleRate, iSaleMargin, iSaleMarginValue, saleStatus
errCNT = 0
addSqlDB = ""
addSql = ""
For i = 0 to pcnt
	If IsObject(xlRowALL(i)) Then
		Set iLine = xlRowALL(i)
			If Right(iLine.FItemArray(2), 1) = "%" Then
				iLine.FItemArray(2) = Left(iLine.FItemArray(2), Len(iLine.FItemArray(2)) - 1)
			End If

			If Right(iLine.FItemArray(3), 1) = "%" Then
				iLine.FItemArray(3) = Left(iLine.FItemArray(3), Len(iLine.FItemArray(3)) - 1)
			End If

			If iLine.FItemArray(1) = "" Then
				iLine.FItemArray(1) = 0
			ElseIf iLine.FItemArray(2) = "" Then
				iLine.FItemArray(2) = 0
			End If

			iLine.FItemArray(0) = replace(iLine.FItemArray(0) ,",","")
			iLine.FItemArray(1) = replace(iLine.FItemArray(1) ,",","")
			iLine.FItemArray(2) = replace(iLine.FItemArray(2) ,",","")
			iLine.FItemArray(3) = replace(iLine.FItemArray(3) ,",","")

			If (Not FIsNum(iLine.FItemArray(0))) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "상품코드("& iLine.FItemArray(0) &")는 숫자만 입력가능 합니다." & " \n"

			elseIf (Not FIsNum(iLine.FItemArray(1))) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "할인판매가("& iLine.FItemArray(1) &")는 숫자만 입력가능 합니다." & " \n"

			elseIf (Not FIsNum(iLine.FItemArray(2))) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "할인율("& iLine.FItemArray(2) &")은 숫자만 입력가능 합니다." & " \n"

			elseIf (Not FIsNum(iLine.FItemArray(3))) Then
				errCNT = errCNT + 1
				RetErr = RetErr & "할인마진율("& iLine.FItemArray(3) &")은 숫자만 입력가능 합니다." & " \n"

			else
				if 	iLine.FItemArray(2) >99 then
					errCNT = errCNT + 1
					RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & " 할인율은 100%를 넘길 수 없습니다." & " \n"
				end if
			    If iLine.FItemArray(3) < 0 Then
						errCNT = errCNT + 1
						RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & "역마진 상품입니다." & " \n"

			    End If
			     dim sMargin
			    strSql = "select sale_margin from db_Event.dbo.tbl_sale where sale_code = "&sCode
			     rsget.Open strSql,dbget,1
					 if not rsget.eof Then
					 	sMargin = rsget(0)
					end if
					rsget.close

			    If iLine.FItemArray(3) < 1 and sMargin <> 4 Then
						errCNT = errCNT + 1
						RetErr = RetErr & "상품코드 : " & iLine.FItemArray(0) & " 할인마진율이 1% 미만인 상품입니다. " & " \n"

			    End If
			End If

			itemArr				= itemArr & iLine.FItemArray(0) & ","			'상품코드
		Set iLine = xlRowALL(i)
	End If
Next

 If errCNT > 0 Then
 		response.write "<script>alert('"&errCNT&"건 오류.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
 		response.end
 end if

	If Right(itemArr,1) = "," Then
		itemidarr = Left(itemArr, Len(itemArr) - 1)
	End If

	'1.tmp테이블에 내용 있는 그대로 등록
	strSql = ""
	strSql = strSql & " DELETE FROM db_temp.[dbo].[tbl_saleItem_Upload] WHERE sale_code = '"&sCode&"' "
	dbget.execute strSql
	strSql = ""
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
				If Right(int(iLine.FItemArray(1)),1) <> "0" Then			'할인 판매가의 원단위 절삭위해 분기
					iLine.FItemArray(1) = Left(iLine.FItemArray(1),Len(iLine.FItemArray(1))-1)&"0"
				End If
				strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_saleItem_Upload] (sale_code, itemid, saleprice, saleper, salesupplycashPer) VALUES ('"&sCode&"', '"&iLine.FItemArray(0)&"', '"& int(iLine.FItemArray(1)) &"', '"&iLine.FItemArray(2)&"', '"&iLine.FItemArray(3)&"'); " & VBCRLF
			Set iLine = xlRowALL(i)
		End If
	Next
	dbget.execute strSql

'중복 상품 삭제
	 dim arritem, intI,errover
	 strSql =  " select itemid into #dblitem from db_temp.[dbo].[tbl_saleItem_Upload] where sale_code =  '"&sCode&"' group by itemid having count(itemid) > 1 "
	 dbget.execute strSql
	 strSql =  " select itemid from #dblitem order by itemid  "
	 rsget.Open strSql,dbget,1
	 if not rsget.eof Then
	 	arritem = rsget.getrows()
		end if
		rsget.close
		if isArray(arritem) Then
			for intI = 0 To UBOund(arritem)
			  if intI =0 Then
			  	errover = arritem(0,intI)
			  else
			  errover = errover &","&arritem(0,intI)
			end if
		  next

		  strSql = "delete from  db_temp.[dbo].[tbl_saleItem_Upload] where itemid in ( select itemid from #dblitem )"
		  dbget.execute strSql

		  strSql = "drop table #dblitem "
		  dbget.execute strSql

	  	strStatus = " 엑셀에 상품이 중복되어있습니다. 가격 확인 후 한개의 상품코드만 등록해주세요 "
			RetErr = RetErr & "상품코드 : [ " & errover & "]" & strStatus & " \n"
			errItem = errItem&","& errover
		end if
	''상품 가격및 마진율 확인
	dim arrList1
	strSql = ""
	strSql = strSql & " SELECT i.itemid, i.sailyn, i.sellcash, i.buycash, i.orgprice, i.orgsuplycash , u.saleprice, u.saleper, u.salesupplycashper, u.salesupplycash"
	strSql = strSql & " FROM db_item.dbo.tbl_item as i  "
	strSql = strSql & " JOIN db_temp.[dbo].[tbl_saleItem_Upload] as u on i.itemid = u.itemid "
	strSql = strSql & " WHERE sale_code = '"&sCode&"' "
	rsget.Open strSql,dbget,1
	If Not(rsget.EOF or rsget.BOF) Then
		arrList1 = rsget.getRows
	End If
	rsget.Close

'dim t_orgprice, t_saleprice, t_saleper
'dim errSalePer, errSupplyCashper
'	if isArray(arrList) then
'	errSalePer = ""
'		for intLoop = 0 To UBound(arrList,2)
'			t_orgprice = arrList(4,intLoop)
'			t_saleprice = arrList(6,intLoop)
'			t_saleper = arrList(7,intLoop)

			''할인율 88% 이상한 상품 확인
'			if t_saleprice = 0 then
'				if t_saleper>=88 then
'					if errSalePer ="" then
'						errSalePer = arrList(0,intLoop)
'					else
'						errSalePer = errSalePer & ", "&arrList(0,intLoop)
'					end if
'				end if
'			else
'				if round(t_orgprice - ((t_orgprice * t_saleper) / 100)) >= 88 then
'					if errSalePer ="" then
'					errSalePer = arrList(0,intLoop)
'				else
'					errSalePer = errSalePer & ", "&arrList(0,intLoop)
'				end if

'			end if
'		next

'		RetErr = RetErr & "상품코드 : [ " & errSalePer & "] 할인율 88% 이상인 상품입니다.  \n"

'	end if


	'3.이벤트코드가 있는 할인코드일 때 엑셀내의 상품코드가 해당 이벤트코드의 상품에 포함 되있는 지 판단
	If eCode <> "" Then

		strSql = ""
		strSql = strSql & " SELECT itemid from "
		strSql = strSql & " db_temp.[dbo].[tbl_saleItem_Upload]  "
		strSql = strSql & " WHERE itemid not in ( "
		strSql = strSql & " 	SELECT itemid FROM [db_event].[dbo].[tbl_eventitem] WHERE evt_code = '"&eCode&"' and evtgroup_code ='"&egCode&"' "
		strSql = strSql & " ) and sale_code = '"&sCode&"' "
		rsget.Open strSql,dbget,1
		If Not(rsget.EOF or rsget.BOF) Then
			i = 0
			Do until rsget.EOF
				errCNT = errCNT + 1
				strStatus = " 이벤트에 속한 상품이 아닙니다."
				RetErr = RetErr & "상품코드 : " & rsget("itemid") & strStatus & " \n"
					errItem = errItem&","& rsget("itemid")
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		'If errCNT > 0 Then
		'	response.write "<script>alert('"&errCNT&"건 오류.\n\n"&RetErr&"');opener.location.reload();self.close();</script>"
		'	response.end
		'End If
	End If
dim sale_type
	'4.tmp테이블에서 sailyn이 Y인 애들과 역마진인 애들이 있는 지 검사통과
	strSql = ""
	strSql = strSql & " SELECT convert(varchar(19),sale_startdate,121) as sale_startdate, convert(varchar(19),sale_enddate,121) as sale_enddate, sale_rate, sale_margin, sale_marginvalue, sale_status, sale_type FROM [db_event].[dbo].tbl_sale WHERE sale_code= "&sCode
	rsget.Open strSql,dbget
	If not rsget.EOF Then
		dSDate = rsget("sale_startdate")
		dEDate = rsget("sale_enddate")
		iSaleRate = rsget("sale_rate")
		iSaleMargin = rsget("sale_margin")
		iSaleMarginValue = rsget("sale_marginvalue")
		saleStatus	= rsget("sale_status")
		sale_type   = rsget("sale_type")
	End If
	rsget.Close

	'RetErr = ""
	'strStatus = ""
	strSql = ""
	strSql = strSql & " SELECT b.itemid, a.sale_code, a.sale_status "
	strSql = strSql & " FROM [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b "
	strSql =strSql&	"   WHERE  a.sale_code = b.sale_code "
	strSql =strSql&	"           and  ( "
	strSql =strSql&	"                    ( ( a.sale_type ='"&sale_type&"' and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"') "
	strSql =strSql&	"	                    and a.sale_using =1 and a.sale_status <> 8 and  b.saleitem_status <> 8 "
	strSql =strSql&	"                    ) "
	strSql =strSql&	"                    or "
	strSql =strSql&	"                    (a.sale_code = "&sCode&")"
	strSql =strSql&	"                 ) "
	strSql =strSql&	"            and b.itemid in ("&itemidarr&")"
	'strSql = strSql & " WHERE a.sale_code = b.sale_code and (( a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"
	'strSql = strSql & " and a.sale_using =1 and a.sale_status <> 8 and  b.saleitem_status <> 8 ) or (a.sale_code = "&sCode&")) and b.itemid in ("&itemidarr&")"

	rsget.Open strSql,dbget
	IF not rsget.EOF THEN
		arrList = rsget.getRows()
	End IF
	rsget.Close

	If isArray(arrList) Then
		For intLoop = 0 To UBound(arrList,2)
			errCNT = errCNT + 1
			Select Case arrList(2,intLoop)
				Case 6
					strStatus = "할인중"
				Case 7
					strStatus = "할인예정"
				Case 0
					strStatus = "등록대기"
			End Select
			RetErr = RetErr & "할인코드 : " & CStr(arrList(1,intLoop)) & " - 상품번호 : " & CStr(arrList(0,intLoop)) &" "& strStatus & " \n"
			errItem = errItem&","&arrList(0,intLoop)
		Next
	'	response.write "<script>alert('"&RetErr&"');opener.location.reload();self.close();</script>"
	'response.end
	End If

 if errItem <> "" then
 	erritem =    right(erritem, Len(erritem) - 1)
  strSql = "delete from db_temp.[dbo].[tbl_saleItem_Upload] where itemid in ("&errItem&") "
  dbget.execute strSql
 end if
'	If errCNT = 0 Then
		If eCode <> "" Then
			strSql = ""
			strSql = strSql & " INSERT INTO [db_event].[dbo].[tbl_saleItem]([sale_code], [itemid], [saleItem_status], [saleprice],[salesupplycash])"
			'strSql = strSql & " SELECT "&sCode&", i.itemid, 7, u .saleprice, u.salesupplycash"
			strSql = strSql & " SELECT "&sCode&", u.itemid, '7' "
			strSql = strSql & " , Case When u.saleprice = 0 Then round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) "
			strSql = strSql & " 	Else u.saleprice end saleprice "
			if sMargin ="4" then
			strSql = strSql & " , i.buycash  "
			Else
				strSql = strSql & " , Case When u.saleprice = 0 Then Round(round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) - ((round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) * u.salesupplycashper) / 100), 0) "
			strSql = strSql & " 	Else Round(u.saleprice - ((u.saleprice * u.salesupplycashper) / 100), 0) end  "
			end if
			strSql = strSql & "	FROM [db_item].[dbo].tbl_item i "
			strSql = strSql & "	JOIN [db_event].[dbo].[tbl_eventitem] c on i.itemid = c.itemid and c.evt_code = "&eCode&" and c.evtgroup_code ="&egCode
			strSql = strSql & " JOIN db_temp.[dbo].[tbl_saleItem_Upload] u on i.itemid = u.itemid and u.sale_code = '"&sCode&"'  "
			strSql = strSql & " WHERE i.itemid not in "
			strSql = strSql & " (select b.itemid from [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b"
			strSql = strSql&" 	where a.sale_code = b.sale_code and "
		strSql = strSql&"           ("
		strSql = strSql&"               ("
		strSql = strSql&"                   ( a.sale_type ="&sale_type&" and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"')  "
		strSql = strSql&"	   		         and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 "
		strSql = strSql&"               ) "
		strSql = strSql&"               or "
		strSql = strSql&"               (a.sale_code = "&sCode&")"
		strSql = strSql&"            ) "
		strSql = strSql&"  )"&addSql
		'	strSql = strSql & " 	where a.sale_code = b.sale_code and (( a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"'"
		'	strSql = strSql & "	 and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 ) or (a.sale_code = "&sCode&")) )"
		Else
			strSql = ""
			strSql = strSql & " INSERT INTO [db_event].[dbo].[tbl_saleItem]([sale_code], [itemid], [saleItem_status], [saleprice],[salesupplycash]) "
			'strSql = strSql & " SELECT sale_code, u.itemid, '7', u.saleprice, u.salesupplycash "
			strSql = strSql & " SELECT "&sCode&", u.itemid, '7' "
			strSql = strSql & " , Case When u.saleprice = 0 Then round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) "
			strSql = strSql & " 	Else u.saleprice end saleprice "
				if sMargin ="4" then
			strSql = strSql & " , i.buycash  "
			Else
			strSql = strSql & " , Case When u.saleprice = 0 Then Round(round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) - ((round(i.orgprice - (((i.orgprice * u.saleper) / 100)), -1, 1) * u.salesupplycashper) / 100), 0) "
			strSql = strSql & " 	Else Round(u.saleprice - ((u.saleprice * u.salesupplycashper) / 100), 0) end  "
end if
			strSql = strSql & " FROM db_temp.[dbo].[tbl_saleItem_Upload] as u "
			strSql = strSql & " JOIN [db_item].[dbo].tbl_item i on i.itemid = u.itemid "
			strSql = strSql & " WHERE u.itemid not in "
			strSql = strSql & " (SELECT b.itemid from [db_event].[dbo].tbl_sale a, [db_event].[dbo].[tbl_saleitem] b"
			strSql = strSql&" 	where a.sale_code = b.sale_code and "
		strSql = strSql&"           ("
		strSql = strSql&"               ("
		strSql = strSql&"                   (a.sale_type ="&sale_type&" and a.sale_startdate <= '"&dEDate&"' and a.sale_enddate >= '"&dSDate&"') "
		strSql = strSql&"	            and a.sale_using =1 and a.sale_status <> 8  and  b.saleitem_status <> 8 "
		strSql = strSql&"               ) "
		strSql = strSql&"               or "
		strSql = strSql&"               (a.sale_code = "&sCode&")"
		strSql = strSql&"            ) "
		strSql = strSql&"  )"&addSql
			strSql = strSql & " and sale_code = '"&sCode&"' "
		End If

		dbget.execute strSql

		If Err.Number <> 0 Then
			Alert_move "데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요","about:blank"
			dbget.close()	:	response.End
		End If

		If saleStatus = 6 Then
			Set objCmd = Server.CreateObject("ADODB.COMMAND")
				With objCmd
				.ActiveConnection = dbget
				.CommandType = adCmdText
				.CommandText = "{?= call [db_item].[dbo].[sp_Ten_item_SetPrice_RealTime] ("&sCode&",'I')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
				iResult = objCmd(0).Value
			Set objCmd = nothing
			If iResult <> 1 Then
				dbget.RollBackTrans
				Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
				dbget.close()	:	response.End
			End If
		End If
	'End If
	if RetErr <> "" then
		response.write "<script>alert('"&RetErr&"')</script>"
		end if
	response.write "<script>opener.location.reload();self.close();</script>"
'End If
'''====================================================================================
Class TXLRowObj
	Public FItemArray

	Public Function setArrayLength(ln)
		Redim FItemArray(ln)
	End Function
End Class

Function IsSKipRow(ixlRow, skipCol0Str)
	If Not IsArray(ixlRow) Then
		IsSKipRow = true
		Exit Function
	End if

	If  LCASE(ixlRow(0)) = LCASE(skipCol0Str) Then
		IsSKipRow = true
		Exit Function
	End If
	IsSKipRow = false
End Function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
	Dim conDB, Rs, strQry, iResult, i, J, iObj
	Dim irowObj, strTable
	'' on Error 구문 쓰면 안됨.. 서버 무한루프 도는듯.
	Set conDB = Server.CreateObject("ADODB.Connection")
		conDB.Provider = "Microsoft.Jet.oledb.4.0"
		'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 추가 2013/12/19
	'On Error Resume Next
		conDB.Open sFilePath
		If (ERR) Then
			fnGetXLFileArray=false
			'/이유를 알수 없는 서버단 에러남. "예기치 않은 오류. 외부 개체에 트랩 가능한 오류(C0000005)가 발생했습니다. 스크립트를 계속 실행할 수 없습니다"
			set conDB = nothing
			Exit Function
		End If
	'On Error Goto 0
		'' get First Sheet Name=============''시트가 여러개인경우 오류날 수 있음.
		Set Rs = conDB.OpenSchema(adSchemaTables)
			If Not Rs.Eof Then
				aSheetName = Rs.Fields("table_name").Value
				''rw "aSheetName="&aSheetName
			End If
		Set Rs = Nothing
		''==================================
		Set Rs = Server.CreateObject("ADODB.Recordset")
			''strQry = "Select * From [sheet1$]"
			strQry = "Select * From ["&aSheetName&"]"
			ReDim xlRowALL(0)
			fnGetXLFileArray = true

		'On Error Resume Next
			Rs.Open strQry, conDB
			IF (ERR) then
				fnGetXLFileArray=false
				Rs.Close
				Set Rs = Nothing
				Set conDB = Nothing
				Exit Function
			End if

			If Not Rs.Eof Then
				Do Until Rs.Eof
					If (ERR) Then
						fnGetXLFileArray=false
						Rs.Close
						Set Rs = Nothing
						Set conDB = Nothing
						Exit Function
					End if

					Set irowObj = new TXLRowObj
						irowObj.setArrayLength(iArrayLen)

						For i=0 to ArrayLen
							If (xlPosArr(i) < 0) Then
								irowObj.FItemArray(i) = ""
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
							''rw irowObj.FItemArray(i)
						Next

						IF (Not IsSKipRow(irowObj.FItemArray,skipString)) then
							ReDim Preserve xlRowALL(UBound(xlRowALL)+1)
							set xlRowALL(UBound(xlRowALL)) =  irowObj
							''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow
						END IF
					set irowObj = Nothing
					Rs.MoveNext
				Loop
			Else
				fnGetXLFileArray=false
			End If
			''''On Error Goto 0
			If (ERR) Then
				fnGetXLFileArray=false
			End If
			Rs.Close
		'On Error Goto 0
		Set Rs = Nothing
	Set conDB = Nothing
	If Ubound(xlRowALL) <  1 Then fnGetXLFileArray = false
End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->