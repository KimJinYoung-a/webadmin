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
Dim sCode, egCode, eCode,adminid
monthFolder = Replace(Left(CStr(now()),7),"-","")
adminid  = session("ssBctId")
'�Ҽ������� ���ڿ��� üũ 
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
	uploadform.Start sDefaultPath '���ε���

	iMaxLen = uploadform.Form("iML")	'�̹�������ũ��
	 

	If (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) Then	'����üũ
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
''''	0			1				2			 
''''��ǰ�ڵ�, 	�����ǸŰ�, 	���θ��԰� 
xlPosArr = Array(0,1,2,3)
ArrayLen = UBound(xlPosArr)
skipString = "Sheet1"
afile = sFilePath
aSheetName = ""

Dim xlRowALL
Dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)
If (Not ret) or (Not IsArray(xlRowALL)) then
	response.write "<script>alert('������ �ùٸ��� �ʰų� ������ �����ϴ�. "&Replace(Err.Description,"'","")&"');</script>"
	
	If (Err.Description="�ܺ� ���̺� ������ �߸��Ǿ����ϴ�.") Then
		response.write "<script>alert('�������� Save As Excel 97 -2003 ���չ��� ���·� ������ ����ϼ���.');</script>"
	End If
	response.write "<script>history.back();</script>"
	response.end
End If

''������ ó��.
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

			If iLine.FItemArray(1) = "" Then
				iLine.FItemArray(1) = 0
			ElseIf iLine.FItemArray(2) = "" Then
				iLine.FItemArray(2) = 0
			End If
			
			iLine.FItemArray(1) = replace(iLine.FItemArray(1) ,",","")
			iLine.FItemArray(2) = replace(iLine.FItemArray(2) ,",","")
			 
  
			If (Not FIsNum(iLine.FItemArray(0))) OR (Not FIsNum(iLine.FItemArray(1))) OR (Not FIsNum(iLine.FItemArray(2)))   Then
				errCNT = errCNT + 1
				RetErr = RetErr & "���� ������ �� ���ڿ��� ���ڰ� �ֽ��ϴ�." & " \n" 
				 
			else  
				iSaleMarginValue = round(((iLine.FItemArray(1) -iLine.FItemArray(2))/iLine.FItemArray(1))*100)
			    If iSaleMarginValue < 0 Then
						errCNT = errCNT + 1
						RetErr = RetErr & "��ǰ�ڵ� : " & iLine.FItemArray(0) & "������ ��ǰ�Դϴ�." & " \n"  
			    elseIf iSaleMarginValue < 1  then
						errCNT = errCNT + 1
						RetErr = RetErr & "��ǰ�ڵ� : " & iLine.FItemArray(0) & " ���θ������� 1% �̸��� ��ǰ�Դϴ�. " & " \n"
					 
			    End If
			End If 

			itemArr				= itemArr & iLine.FItemArray(0) & ","			'��ǰ�ڵ�
		Set iLine = xlRowALL(i)
	End If
Next
  
 
	If Right(itemArr,1) = "," Then
		itemidarr = Left(itemArr, Len(itemArr) - 1)
	End If

	'1.tmp���̺� ���� �ִ� �״�� ��� 
	strSql = "create table #tmpATSale( itemid int, sellprice money, buyprice money)"
	For i = 0 to pcnt
		If IsObject(xlRowALL(i)) Then
			Set iLine = xlRowALL(i)
				If Right(int(iLine.FItemArray(1)),1) <> "0" Then			'���� �ǸŰ��� ������ �������� �б�
					iLine.FItemArray(1) = Left(iLine.FItemArray(1),Len(iLine.FItemArray(1))-1)&"0"
				End If
				strSql = strSql & " INSERT INTO #tmpATSale (itemid, sellprice, buyprice ) VALUES ('"&iLine.FItemArray(0)&"', '"& int(iLine.FItemArray(1)) &"', '"&iLine.FItemArray(2)&"'); " & VBCRLF
			Set iLine = xlRowALL(i)
		End If
	Next
	dbget.execute strSql

	'' 2018/08/14 ���⼭ ������ �� ����.. (�̹����� �ȵǸ� �߰� �������� �ϳ� ����°� ���� ������..?)
	''CASE 1 ��ǰ�ڵ带 �߸�����ϴ� ���̽��� ����. => ���Ϻ귣�峢���� ���� �����ϰ� ����.
	dim MultipleMakerArr
	strSql = " select i.makerid,count(*) CNT from #tmpATSale T"
	strSql = strSql & " Join db_item.dbo.tbl_item i"
	strSql = strSql & " on T.itemid=i.itemid"
	strSql = strSql & " group by i.makerid"
	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		MultipleMakerArr = rsget.getRows()
	end if
	rsget.Close
	
	if IsArray(MultipleMakerArr) then
		if UBound(MultipleMakerArr,2)>0 then
			response.write "<script>alert('Oops! ���� �귣�带 ���ÿ� ó�� �� �� �����ϴ�.\r\n������ ��ҵǾ����ϴ�.')</script>"
			response.write "ERROR : ��ϵ��� �ʾҽ��ϴ�. - ���� �귣�带 ���ÿ� ó�� �� �� �����ϴ�"& "<br>"
			FOR intLoop = 0 TO UBound(MultipleMakerArr,2)
				response.write MultipleMakerArr(0,intLoop) & " : " & MultipleMakerArr(1,intLoop) & "�� <br>"
			Next
			strSql = "drop table #tmpATSale "
    	 	dbget.execute strSql

			dbget.close() : response.end
		end if
	else
		response.write "<script>alert('Oops! ��ϵ� ��ǰ�� �����ϴ�. ��ǰ�ڵ尡 �ùٸ��� �ʰų� ������ �����ϴ�.')</script>"
		strSql = "drop table #tmpATSale "
    	dbget.execute strSql

		dbget.close() : response.end
	end if

	''CASE 2 �Һ��ڰ����� ������ ������� 
	dim SalePriceHighArr
	strSql = " select T.itemid, T.sellprice, i.orgprice from #tmpATSale T"
	strSql = strSql & " Join db_item.dbo.tbl_item i"
	strSql = strSql & " on T.itemid=i.itemid"
	strSql = strSql & " where T.sellprice>i.orgprice"
	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		SalePriceHighArr = rsget.getRows()
	end if
	rsget.Close
	
	if IsArray(SalePriceHighArr) then
		response.write "<script>alert('Oops! ���ΰ��� �Һ��ڰ����� ���� ���� �� �� �����ϴ�.\r\n������ ��ҵǾ����ϴ�.')</script>"
		response.write "ERROR : ��ϵ��� �ʾҽ��ϴ�. - ���ΰ��� �Һ��ڰ����� ���� ���� �� �� �����ϴ�"& "<br>"
		response.write "��ǰ��ȣ" & " : " & "������ΰ�" & " / " & "����Һ��ڰ�" & "<br>"
		FOR intLoop = 0 TO UBound(SalePriceHighArr,2)
			response.write SalePriceHighArr(0,intLoop) & " : " & FormatNumber(SalePriceHighArr(1,intLoop),0) & " / " & FormatNumber(SalePriceHighArr(2,intLoop),0) & "<br>"
		Next
		strSql = "drop table #tmpATSale "
		dbget.execute strSql

		dbget.close() : response.end
	end if

	''CASE 3 �����ǸŰ����� ������ N% �̻� ���εɰ�� - ��Ƽ��
	dim SalePriceLowArr
	strSql = " select T.itemid, T.sellprice, i.sellcash from #tmpATSale T"
	strSql = strSql & " Join db_item.dbo.tbl_item i"
	strSql = strSql & " on T.itemid=i.itemid"
	strSql = strSql & " where i.sellcash<>0"
	strSql = strSql & " and (100-T.sellprice/sellcash*100>88)"
	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		SalePriceLowArr = rsget.getRows()
	end if
	rsget.Close
	
	if IsArray(SalePriceLowArr) then
		
		FOR intLoop = 0 TO UBound(SalePriceLowArr,2)
			RetErr = RetErr & "��ǰ��ȣ "& SalePriceLowArr(0,intLoop) &" ���� �ǸŰ����� �������� ����� �����ϴ�. " & FormatNumber(SalePriceLowArr(1,intLoop),0) & " / " & FormatNumber(SalePriceLowArr(2,intLoop),0) &"\n"
		Next

	end if
	'' CASE END



   ''�̺�Ʈ ���� ������� ����ó��  
   	strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	 	strSql = strSql & " (  select si.itemid, 2, s.sale_code, si.saleprice, si.salesupplycash,4,'��������','"&adminid&"'"
	 	strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
		strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "	  
		strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "         
		strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
		strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
		strSql = strSql & "              	and s.sale_using =1   "	           
		strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  <> 8  )" 
	 	dbget.execute strSql
	 		
   strSql = " update si  SET saleitem_status = 9 ,closedate=getdate(), lastupdate =getdate()"
   strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
	 strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "	  
	 strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "         
	 strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
	 strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
	 strSql = strSql & "              	and s.sale_using =1   "	           
	 strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  <> 8  " 
	 	dbget.execute strSql
	 	
	 	'just1day ��ǰ Ȯ��
	 	strSql ="select si.itemid  "
	 	strSql = strSql & " into #tmpJ1day "
	 	strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
	  strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "	  
	  strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "         
	  strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
	  strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
	  strSql = strSql & "              	and s.sale_using =1   "	           
	  strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  " 
	  	dbget.execute strSql
						
	 	strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	 	strSql = strSql & " (  select si.itemid, 2, s.sale_code, si.saleprice, si.salesupplycash,5,'����Ʈ������ ������-������ε�ϴ��ó��','"&adminid&"'"
	 	strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
		strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "	  
		strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "         
		strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
		strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
		strSql = strSql & "              	and s.sale_using =1   "	           
		strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  )" 
	 	dbget.execute strSql
	 		
   strSql = " update si  SET orgsailprice = t.sellprice,orgsailsuplycash = t.buyprice , orgsailyn='Y',lastupdate =getdate()"
   strSql = strSql & "    FROM 	db_event.dbo.tbl_saleitem as si   "
	 strSql = strSql & "              inner join db_event.dbo.tbl_sale as s   on s.sale_code = si.sale_code   "	  
	 strSql = strSql & "							inner join #tmpATSale as t on si.itemid = t.itemid "         
	 strSql = strSql & "              where (s.sale_status = 6  or s.sale_status = 7  ) "
	 strSql = strSql & "              	and ( si.saleItem_status = 6  or si.saleitem_status = 7 or si.saleitem_status = 1 ) "
	 strSql = strSql & "              	and s.sale_using =1   "	           
	 strSql = strSql & "              	and s.sale_enddate >=convert(varchar(10),getdate(),121) and s.availPayType  = 8  " 
	 	dbget.execute strSql
	 	  
		   ' ������� ó��	  	  		 
			 strSql = "update  i "
			 strSql = strSql & " set sellcash = t.sellprice, buycash = t.buyprice, sailprice =t.sellprice  , sailsuplycash =t.buyprice , sailyn ='Y'"
			 strSql = strSql & " , mileage=case when (1-(convert(float,t.sellprice)/ i.orgprice)) >= 0.4 then 0 else convert(int, t.sellprice*0.005) end, lastupdate =getdate()"
			  strSql = strSql & " from db_item.dbo.tbl_item as i "
			  strSql = strSql & " inner join #tmpATSale as t on i.itemid  = t.itemid "
			  strSql = strSql & " left outer join #tmpJ1day as j on t.itemid = j.itemid "
			 strSql = strSql & " where j.itemid is null "
			 dbget.execute strSql
			 
			 strSql = "	insert into db_log.dbo.tbl_saleitem_allLog(itemid,saletype, sale_code, sellcash, buycash, logtype, logconts, adminid ) "
	 		strSql = strSql &" select t.itemid ,1,0, t.sellprice,t.buyprice,1,'������ε��','"&adminid&"' "
	 		 strSql = strSql & " from #tmpATSale as t  "
			  strSql = strSql & " left outer join #tmpJ1day as j on t.itemid = j.itemid "
			 strSql = strSql & " where j.itemid is null "
	 		 dbget.execute strSql
		 
 
   '�ӽ����̺� ����
    strSql = "drop table #tmpJ1day"
    	 dbget.execute strSql
    strSql = "drop table #tmpATSale "
    	 dbget.execute strSql
	if RetErr <> "" then 
		response.write "<script>alert('"&RetErr&"')</script>"
		end if 
	response.write "<script>opener.location.reload();self.close();</script>"
 
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
	'' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.
	Set conDB = Server.CreateObject("ADODB.Connection")
		conDB.Provider = "Microsoft.Jet.oledb.4.0"
		'conDB.Provider = "Microsoft.ACE.OLEDB.12.0"
		conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;IMEX=1"  ''';IMEX=1 �߰� 2013/12/19
	'On Error Resume Next
		conDB.Open sFilePath
		If (ERR) Then
			fnGetXLFileArray=false
			'/������ �˼� ���� ������ ������. "����ġ ���� ����. �ܺ� ��ü�� Ʈ�� ������ ����(C0000005)�� �߻��߽��ϴ�. ��ũ��Ʈ�� ��� ������ �� �����ϴ�"
			set conDB = nothing
			Exit Function
		End If
	'On Error Goto 0
		'' get First Sheet Name=============''��Ʈ�� �������ΰ�� ������ �� ����.
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

						For i=0 to ArrayLen-1
							If (xlPosArr(i) < 0) Then
								irowObj.FItemArray(i) = ""
							Else
								irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
							End If
							'' rw irowObj.FItemArray(i) 
							
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