<%
'###########################################################
' Description : �������� ������ ���� �Լ�
' Hieditor : 2011.03.08 �ѿ�� ����
'###########################################################

CLASS CsActionMailCls
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Dim FAsID
	Dim FDivCD
	Dim FDivCDName
	Dim Forderno
	Dim FCustomerName	
	Dim FBuyHP
	Dim FBuyEmail
	Dim FWriteUser
	Dim FFinishUser
	Dim FTitle
	Dim FContents_jupsu
	Dim FContents_finish
	Dim FCurrstate
	Dim FCurrstateName
	Dim FRegDate
	Dim FFinishDate
	Dim FDeleteyn
	Dim FOpenTitle
	Dim FOpenContents
	Dim FSongjangDiv
	Dim FSongjangNo
	Dim FSongjangDivName
	Dim FRequireUpche
	Dim FMakerid
	Dim FAdd_upchejungsanDeliveryPay
	Dim FAdd_upchejungsanCause
	Dim FRefundRequire
	Dim FRefundResult
	Dim FReturnMethod
	Dim FRefundMileageSum
	Dim FRefundCouponSum
	Dim FAllatSubTractSum
	Dim FRefundItemCostSum
	Dim FRefundBeasongPay
	Dim FRefundDeliveryPay
	Dim FRefundAdjustPay
	Dim FCancelTotal
	Dim FReturnName
	Dim FReturnPhone
	Dim FReturnHP
	Dim FReturnZipCode
	Dim FReturnZipAddr
	Dim FReturnEtcAddr
	Dim FReBankName
	Dim FReBankAccount
	Dim FReBankOwnerName
	Dim FPayGateTid
	Dim FPayGateResultTid
	Dim FPayGateResultMsg
	Dim FReturnMethodName
	Dim FReqName
	Dim FReqPhone
	Dim FReqHP
	Dim FReqZipcode
	Dim FReqZipAddr
	Dim FReqEtcAddr
	Dim FReqEtcStr
    Dim FInfoHtml
    Dim FupcheReturnSongjangDivName
    Dim FupcheReturnSongjangDivTel
	Dim FSendDate
	Dim FResultCount	
    Dim FRectForceCurrState     ''���°� ���� ����.
    Dim FRectForceBuyEmail      ''�̸��� ��������.

 	public function GetAsDivCDName_off()
        GetAsDivCDName_off = db2html(FDivCDName)
	end function
	
	Public Sub GetOneCSASMaster_off(FRectCsAsID)
		dim tmpZipCode, tmpaddress1, tmpaddress2
			tmpZipCode="11154"
			tmpaddress1="��⵵ ��õ�� ������ ����������2�� 83"
			tmpaddress2="�ٹ����� ��������"

		dim strSQL
		strSQL =" SELECT TOP 1 " &_
				" 	A.masteridx ,A.DivCD ,A.orderno ,A.CustomerName ,A.WriteUser ,A.FinishUser " &_
				"	,A.Title ,A.Contents_Jupsu ,A.Contents_Finish ,A.CurrState ,A.RegDate ,A.FinishDate ,A.Deleteyn "&_
				"	,A.OpenTitle ,A.OpenContents ,A.RequireUpche ,A.Makerid ,A.SongjangDiv ,A.SongjangNo"&_
				"	,(SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=A.SongjangDiv) AS SongjangDivName " &_
				" 	,o.BuyHp,o.BuyEmail " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.divCD) as divcdname "

					IF (FRectForceCurrState<>"") then
					    strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd='"&FRectForceCurrState&"') as currstatename "
			        ELSE
			            strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.currstate) as currstatename "
			        END IF
			    				
				strSQL = strSQL & "	,IsNULL(D.ReqName,o.reqname) as ReqName ,IsNULL(D.ReqPhone,o.reqphone) as ReqPhone ,IsNULL(D.ReqHP,o.reqhp) as ReqHP " &_
				" 	,IsNULL(D.ReqZipcode,o.reqzipcode) as ReqZipcode ,IsNULL(D.ReqZipAddr,o.reqzipaddr) as ReqZipAddr ,IsNULL(D.ReqEtcAddr,o.reqaddress) as ReqEtcAddr ,IsNULL(D.ReqEtcStr,'') as ReqEtcStr " &_
				" 	,isNull(p.company_name,'(��)�ٹ�����') as ReturnName ,isNull(p.deliver_phone,'1644-6030') as ReturnPhone ,isNull(p.deliver_hp,'') as ReturnHP "&_
				" 	,isNull(p.return_zipcode,'"& tmpZipCode &"') as ReturnZipCode ,isNull(p.return_address,'"& tmpaddress1 &"') as ReturnZipAddr ,isNull(p.return_address2,'"& tmpaddress2 &"') as ReturnEtcAddr "&_
                " 	,isNull((SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivName "&_
                " 	,isNull((SELECT TOP 1 tel FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivTel "&_
                
				" FROM db_shop.dbo.tbl_shopbeasong_cs_master A " &_
				" LEFT JOIN db_shop.dbo.tbl_shopbeasong_order_master o " &_
				" 	on A.orgmasteridx=o.masteridx " &_
				" LEFT JOIN [db_shop].dbo.tbl_shopbeasong_cs_delivery d " &_
				" 	on A.masteridx = d.asid " &_
				" LEFT JOIN [db_partner].[dbo].tbl_partner p " &_
				" 	on A.makerid= p.id " &_
				" WHERE A.masteridx=" & CStr(FRectCsAsID)

			'response.write strSQL &"<br>"
			rsget.Open strSQL, dbget, 1

	        FResultCount = rsget.RecordCount

	        if  not rsget.EOF  then	 
	        	       	
				FAsID		= rsget("masteridx")
				FDivCD	= rsget("divCD")
				FDivCDName	= rsget("divcdname")							
				Forderno	= rsget("orderno")
				FCustomerName	= rsget("customername")				
				FWriteUser	= rsget("writeuser")
				FFinishUser	= rsget("finishuser")
				FBuyHP		= rsget("BuyHP")
				FBuyEmail	= rsget("BuyEmail")
				
				if (FRectForceBuyEmail<>"") then
				    FBuyEmail = FRectForceBuyEmail
				end if
				
				FTitle	= rsget("title")
				FContents_jupsu	= rsget("contents_jupsu")
				FContents_finish	= rsget("contents_finish")
				
				IF (FRectForceCurrState<>"") then  ''���°� ���� ���� (���� ��߼۽� ���.)
				    FCurrState = FRectForceCurrState
				ELSE
    				FCurrState	= rsget("currstate")
    			END IF
    			
				FCurrStateName	= db2html(rsget("currstatename"))
				FRegDate	= rsget("regdate")
				FFinishDate	= rsget("finishdate")
				FDeleteyn	= rsget("Deleteyn")				
				FOpenTitle	= rsget("OpenTitle")
				FOpenContents	= rsget("OpenContents")				
				FSongjangDiv	= rsget("SongjangDiv")
				FSongjangNo	= rsget("SongjangNo")
				FSongjangDivName = rsget("SongjangDivName")
				FRequireUpche	= rsget("RequireUpche")
				FMakerid	= rsget("Makerid")

				'//GetReturnAddress
				FReturnName	= rsget("ReturnName")
				FReturnPhone	= rsget("ReturnPhone")
				FReturnHP	= rsget("ReturnHP")
				FReturnZipCode	= rsget("ReturnZipCode")
				FReturnZipAddr	= rsget("ReturnZipAddr")
				FReturnEtcAddr	= rsget("ReturnEtcAddr")
				FReqName	= rsget("ReqName")
				FReqPhone	= rsget("ReqPhone")
				FReqHP		= rsget("ReqHP")
				FReqZipcode	= rsget("ReqZipcode")
				FReqZipAddr	= rsget("ReqZipAddr")
				FReqEtcAddr	= rsget("ReqEtcAddr")
				FReqEtcStr	= rsget("ReqEtcStr")
                
                FupcheReturnSongjangDivName = db2html(rsget("upcheReturnSongjangDivName"))
                FupcheReturnSongjangDivTel  = db2html(rsget("upcheReturnSongjangDivTel"))
			END IF
		rsget.close
		
		''��Ÿ �ȳ� ����
		if (FDivCD<>"") and ((FCurrState="B001") or (FCurrState="B007")) then
		    strSQL = " SELECT TOP 1 IsNULL(infoHtml,'') as infoHtml from db_cs.dbo.tbl_cs_comm_div_info"
		    strSQL = strSQL + " where div_comm_cd='" + FDivCD + "'"
		    strSQL = strSQL + " and state_comm_cd='" + FCurrState + "'"
		    
		    rsget.Open strSQL, dbget, 1
		    if  not rsget.EOF  then
		        FInfoHtml = db2Html(rsget("infoHtml"))
		    end if
		    rsget.Close
		end if
	End Sub

	''// ���� �⺻ ���� ��������
	Function getAsInfo_off()
		dim tmpHTML
		tmpHTML = ""

		tmpHTML=tmpHTML&"<!-- ���� �⺻ ���� ���� --> " & vbcrlf
		tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td colspan=""2"" class=""sky12pxb"" style=""padding: 10 0 5 0"">*��������</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" align=""center"" style=""padding-top:2px;"">�����ڵ�</td> " & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding-left:10px;padding-top:2px;"">"& FAsID &"</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">�ֹ���ȣ</td> " & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding-left:10px;padding-top:2px;"">"& forderno &"</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">�����Ͻ�</td> " & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding-left:10px;padding-top:2px;"">"& FRegDate &"</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">��������</td> " & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding-left:10px;padding-top:2px;"">"& FTitle &"</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		</table> " & vbcrlf

		tmpHTML=tmpHTML&"<!-- ���� �⺻ ���� �� --> " & vbcrlf

		getAsInfo_off =tmpHTML

	END Function

	''//���� ��ǰ ���� ��������
	Function getAsItemLIst_off()
		dim tmpHTML
		dim OCsDetail,i

		tmpHTML = ""

		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A008" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"<!-- ���� ��ǰ ���� ���� --> " & vbcrlf

			Set OCsDetail = New corder
			OCsDetail.FRectCsAsID = FAsID
			
			IF FResultCount>0 THEN
				OCsDetail.fGetCsDetailList
			END IF
			
			if (OCsDetail.FresultCount<1) then Exit function
			
				tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
				tmpHTML=tmpHTML&"		<tr> " & vbcrlf
				tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">������ǰ</td> " & vbcrlf
				tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding:5px 0px 5px 5px;""> " & vbcrlf
				tmpHTML=tmpHTML&"				<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" class=""a"" bgcolor=""#cccccc""> " & vbcrlf
				tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
				tmpHTML=tmpHTML&"					<td style=""width:50;"">��ǰ�ڵ�</td> " & vbcrlf
				tmpHTML=tmpHTML&"					<td>��ǰ��[�ɼ�]</td> " & vbcrlf
				tmpHTML=tmpHTML&"					<td style=""width:60px;"">�ǸŰ�</td> " & vbcrlf
				tmpHTML=tmpHTML&"					<td style=""width:30px;"">����</td> " & vbcrlf
				tmpHTML=tmpHTML&"				</tr> " & vbcrlf
												IF OCsDetail.FresultCount>0 Then
													FOR i=0 TO OCsDetail.FResultCount-1
													    IF (OCsDetail.FItemList(i).Fitemid<>0) or (OCsDetail.FItemList(i).fOrdersellprice<>0) then
				tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#FFFFFF"" > " & vbcrlf
				tmpHTML=tmpHTML&"					<td>"& OCsDetail.FItemList(i).fitemgubun & OCsDetail.FItemList(i).Fitemid & OCsDetail.FItemList(i).fitemoption & "</td> " & vbcrlf														
				tmpHTML=tmpHTML&"					<td>"& OCsDetail.FItemList(i).Fitemname &"</td> " & vbcrlf				
				tmpHTML=tmpHTML&"					<td>"& FormatNumber(OCsDetail.FItemList(i).fOrdersellprice,0) &"</td> " & vbcrlf
				tmpHTML=tmpHTML&"					<td>"& OCsDetail.FItemList(i).Fregitemno &"</td> " & vbcrlf
				tmpHTML=tmpHTML&"				</tr> " & vbcrlf
				                                        END IF
													NEXT
												END IF
				tmpHTML=tmpHTML&"				</table> " & vbcrlf
				tmpHTML=tmpHTML&"			</td> " & vbcrlf
				tmpHTML=tmpHTML&"		</tr> " & vbcrlf
				tmpHTML=tmpHTML&"		<tr> " & vbcrlf
				tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
				tmpHTML=tmpHTML&"		</tr> " & vbcrlf
				tmpHTML=tmpHTML&"		</table> " & vbcrlf
												Set OCsDetail= nothing
				tmpHTML=tmpHTML&"<!-- ���� ��ǰ ���� �� --> " & vbcrlf
		END IF
		getAsItemLIst_off = tmpHTML
	END Function

	''//���ּ� ��������
	Function getReqInfo_off()
		dim tmpHTML
		tmpHTML=""
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A010" THEN 'or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"<!-- ���ּ� ���� --> " & vbcrlf
			tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">���ּ�</td> " & vbcrlf
			tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding:5px 0px 5px 5px;""> " & vbcrlf
			tmpHTML=tmpHTML&"				<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" class=""a"" bgcolor=""#cccccc""> " & vbcrlf
			tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
			tmpHTML=tmpHTML&"					<td width=""50"" align=""center"" bgcolor=""#f7f7f7"">����</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td width=""80"" bgcolor=""#FFFFFF"">"& FReqName &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td width=""50"" align=""center"" bgcolor=""#f7f7f7"">����ó</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td bgcolor=""#FFFFFF"">"& FReqPhone &" / "& FReqHP &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"				</tr> " & vbcrlf
			tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
			tmpHTML=tmpHTML&"					<td bgcolor=""#f7f7f7"">�ּ�</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td colspan=""3"" bgcolor=""#FFFFFF"">["& FReqZipcode &"] "& FReqZipAddr &"&nbsp;"& FReqEtcAddr &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"				</tr> " & vbcrlf
			tmpHTML=tmpHTML&"				</table> " & vbcrlf
			tmpHTML=tmpHTML&"			</td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		</table> " & vbcrlf
			tmpHTML=tmpHTML&"<!-- ���ּ� �� --> " & vbcrlf
		END IF
		getReqInfo_off = tmpHTML
	END Function

	''//��ü �ּ� ��������
	Function getReturnInfo_off()
		dim tmpHTML
		tmpHTML=""
		IF FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"<!-- ��ü�ּ� ���� --> " & vbcrlf
			tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td width=""100"" height=""24"" align=""center"" bgcolor=""#f7f7f7"" class=""gray12px02b"" style=""padding-top:2px;"">��ǰȸ���ּ�</td> " & vbcrlf
			tmpHTML=tmpHTML&"			<td class=""gray12px02"" style=""padding:5px 0px 5px 5px;""> " & vbcrlf
			tmpHTML=tmpHTML&"				<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" class=""a"" bgcolor=""#cccccc""> " & vbcrlf
			tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
			tmpHTML=tmpHTML&"					<td width=""80"" bgcolor=""#f7f7f7"">��ü��</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td bgcolor=""#FFFFFF"">"& FReturnName &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td width=""80"" bgcolor=""#f7f7f7"">����ó</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td bgcolor=""#FFFFFF"">"& FReturnPhone &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"				</tr> " & vbcrlf
			tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
			tmpHTML=tmpHTML&"					<td bgcolor=""#f7f7f7"">�ּ�</td> " & vbcrlf
			tmpHTML=tmpHTML&"					<td colspan=""3"" bgcolor=""#FFFFFF"">["& FReturnZipCode &"] "& FReturnZipAddr &" &nbsp;"& FReturnEtcAddr &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"				</tr> " & vbcrlf
			
			if (FReturnName<>"(��)�ٹ�����") and (FupcheReturnSongjangDivName<>"") and (Left(FupcheReturnSongjangDivTel,1)="1" or Left(FupcheReturnSongjangDivTel,1)="0") then
			    tmpHTML=tmpHTML&"				<tr height=""22"" align=""center"" bgcolor=""#f7f7f7""> " & vbcrlf
    			tmpHTML=tmpHTML&"					<td width=""80"" bgcolor=""#f7f7f7"">�̿��ù��</td> " & vbcrlf
    			tmpHTML=tmpHTML&"					<td bgcolor=""#FFFFFF"">"& FupcheReturnSongjangDivName &"</td> " & vbcrlf
    			tmpHTML=tmpHTML&"					<td width=""80"" bgcolor=""#f7f7f7"">�ù�翬��ó</td> " & vbcrlf
    			tmpHTML=tmpHTML&"					<td bgcolor=""#FFFFFF"">"& FupcheReturnSongjangDivTel &"</td> " & vbcrlf
    			tmpHTML=tmpHTML&"				</tr> " & vbcrlf
			end if
			
			tmpHTML=tmpHTML&"				</table> " & vbcrlf
			tmpHTML=tmpHTML&"			</td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		</table> " & vbcrlf
			tmpHTML=tmpHTML&"<!-- ��ü�ּ� �� --> " & vbcrlf
		END IF

		getReturnInfo_off = tmpHTML
	END Function

	'//ó�� ��� ��������
	Function getFinishResult_off()
		dim tmpHTML
		tmpHTML=""

		IF FCurrState="B007" THEN
		    ''ó�� ������ ������..
		    if (FOpenContents="") then
		        if (FDivCD="A000") then
		            FOpenContents = "�±�ȯ��ǰ ���Ϸ�"
		        elseif (FDivCD="A001") then
		            FOpenContents = "������ǰ ���Ϸ�"
		        elseif (FDivCD="A002") then
		            FOpenContents = "��ǰ ���Ϸ�"
		        elseif (FDivCD="A003") then 
		        
		        elseif (FDivCD="A004") then   
		            FOpenContents = "��ǰ ��ǰ(ȸ��)�Ϸ�" '' / ȯ�ҵ��"
		            
		        elseif (FDivCD="A010") then      
		            FOpenContents = "��ǰ ȸ���Ϸ�" '' / ȯ�ҵ��"
		        elseif (FDivCD="A011") then      
		            FOpenContents = "�±�ȯ��ǰ ȸ���Ϸ�"
		        else
		            
		        end if
		    end if
		    
			tmpHTML=tmpHTML&"<!-- ó�� ��� ����--> " & vbcrlf
			tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td colspan=""2"" class=""sky12pxb"" style=""padding: 10 0 5 0;"">*ó�����</td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td width=""100"" height=""22"" align=""center"" bgcolor=""#f7f7f7"" class=""black12pxb"" style=""padding-top:2px;"">ó���Ϸ���</td> " & vbcrlf
			tmpHTML=tmpHTML&"			<td class=""black12px"" style=""padding-left:10px;padding-top:2px;"">"& FFinishDate &"</td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			
			IF (Trim(FOpenContents)<>"") then
    			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
    			tmpHTML=tmpHTML&"			<td height=""22"" align=""center"" bgcolor=""#f7f7f7"" class=""black12pxb"" style=""padding-top:2px;"">ó������</td> " & vbcrlf
    			tmpHTML=tmpHTML&"			<td class=""black12px"" style=""padding-left:10px;padding-top:2px;""> " & vbcrlf
    			tmpHTML=tmpHTML&"			"& nl2br(FOpenContents) &" " & vbcrlf
    			tmpHTML=tmpHTML&"			</td> " & vbcrlf
    			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			END IF
			
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		</table> " & vbcrlf
			tmpHTML=tmpHTML&"<!-- ó�� ��� ��--> " & vbcrlf
		END IF
		getFinishResult_off = tmpHTML
	END Function

	''// �ù� ���� ��������
	Function getDlvInfo_off()
		dim tmpHTML
		
		tmpHTML=""
        
        if (IsNULL(FSongjangNo)) or (FSongjangNo="") then Exit function 
        
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"<!-- �ù����� ���� --> " & vbcrlf
			tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td width=""100"" height=""22"" align=""center"" bgcolor=""#f7f7f7"" class=""black12pxb"" style=""padding-top:2px;"">�ù�����</td> " & vbcrlf
			tmpHTML=tmpHTML&"			<td class=""black12px"" style=""padding-left:10px;padding-top:2px;""> " & vbcrlf
			
			IF FSongjangNo<>"" then
				tmpHTML=tmpHTML& FSongjangDivName &" &nbsp;"& FSongjangNo &"&nbsp;"& vbcrlf
				tmpHTML=tmpHTML& "<a href="""& DeliverDivTrace(Trim(FSongjangDiv)) & FSongjangNo &""" target=""_blank"">>>�����ϱ�</a> " & vbcrlf
			ELSE
				IF FDivCD = "A004" THEN
					tmpHTML=tmpHTML&" 				�ù������� ��ϵ��� �ʾҽ��ϴ�.<!-- >>�ù�������� --> " & vbcrlf
				ELSE
					tmpHTML=tmpHTML&"				�ù������� ��ϵ��� �ʾҽ��ϴ�. " & vbcrlf
				END IF
			END IF
			
			tmpHTML=tmpHTML&"			</td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		<tr> " & vbcrlf
			tmpHTML=tmpHTML&"			<td height=""1"" colspan=""2"" bgcolor=""#cccccc""></td> " & vbcrlf
			tmpHTML=tmpHTML&"		</tr> " & vbcrlf
			tmpHTML=tmpHTML&"		</table> " & vbcrlf
			tmpHTML=tmpHTML&"<!-- �ù� ���� �� --> " & vbcrlf
		END IF

		getDlvInfo_off =  tmpHTML
	END Function

	'// ��Ÿ �ȳ�����
	Public Function getEtcNotice_off()
		dim tmpHTML
		
        getEtcNotice_off = ""
        
        if (Trim(FInfoHtml)="") then Exit function
        
		tmpHTML=tmpHTML&"<!-- ��Ÿ�ȳ����� START --> " & vbcrlf
		tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
		tmpHTML=tmpHTML&"		<tr>" & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""sky12pxb"" style=""padding:10 0 5 0;"">*��Ÿ�ȳ�����</td>" & vbcrlf
		tmpHTML=tmpHTML&"		</tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<tr>" & vbcrlf
		tmpHTML=tmpHTML&"			<td class=""black12px"" style=""padding:5px;"" bgcolor=""#99CCCC"">" & vbcrlf

		tmpHTML=tmpHTML&" 				"& FInfoHtml & vbcrlf
		
		tmpHTML=tmpHTML&"			</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		</table>" & vbcrlf
		tmpHTML=tmpHTML&"<!-- ��Ÿ�ȳ����� END --> " & vbcrlf

		
		getEtcNotice_off = tmpHTML
	End Function
	
	'// mail ������
	Function makeMailTemplate_off(id)
		dim tmpHTML

		Call GetOneCSASMaster_off(id) '// �� ����

		tmpHTML=tmpHTML&"<link href=""http://www.10x10.co.kr/lib/css/2008ten.css"" rel=""stylesheet"" type=""text/css""> " & vbcrlf
		tmpHTML=tmpHTML&"<table width=""600"" border=""0"" align=""center"" cellspacing=""0"" cellpadding=""0"" class=""a""> " & vbcrlf
		tmpHTML=tmpHTML&"<tr> " & vbcrlf
		tmpHTML=tmpHTML&"	<td><a href=""http://www.10x10.co.kr"" target=""_blank"" onFocus=""blur()""> " & vbcrlf
		tmpHTML=tmpHTML&"		<img src=""http://fiximage.10x10.co.kr/web2008/mail/mail_header.gif"" width=""600"" height=""60"" border=""0"" /></a> " & vbcrlf
		tmpHTML=tmpHTML&"	</td> " & vbcrlf
		tmpHTML=tmpHTML&"</tr> " & vbcrlf
		tmpHTML=tmpHTML&"<tr> " & vbcrlf
		tmpHTML=tmpHTML&"	<td style=""border:7px solid #eeeeee;""> " & vbcrlf
		tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""a""> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td>"& getMailHeadImage_off &"</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td height=""30"" style=""padding:0 15px 0 15px""> " & vbcrlf
		tmpHTML=tmpHTML&"				<!-- ���� / �ֹ���ȣ --> " & vbcrlf
		tmpHTML=tmpHTML&"				<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""a""> " & vbcrlf
		tmpHTML=tmpHTML&"				<tr> " & vbcrlf
		tmpHTML=tmpHTML&"					<td class=""black12px""> " & vbcrlf
		tmpHTML=tmpHTML&"						<strong>"& Fcustomername &"</strong>���� ��û�Ͻ� <span class=""sky12pxb"">["& GetAsDivCDName_off &"]</span>ó���� " & FCurrStateName & " �Ǿ����ϴ�. " & vbcrlf
		tmpHTML=tmpHTML&"					</td> " & vbcrlf
		tmpHTML=tmpHTML&"					<td align=""right"" class=""gray11px02"">�ֹ���ȣ : <span class=""sale11px01"">"& Forderno &"</span></td> " & vbcrlf
		tmpHTML=tmpHTML&"				</tr> " & vbcrlf
		tmpHTML=tmpHTML&"				<tr> " & vbcrlf
		tmpHTML=tmpHTML&"					<td height=""3"" colspan=""2"" class=""black12px"" style=""padding:5px;"" bgcolor=""#99CCCC""></td> " & vbcrlf
		tmpHTML=tmpHTML&"				</tr> " & vbcrlf
		tmpHTML=tmpHTML&"				</table> " & vbcrlf
		tmpHTML=tmpHTML&"			</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td style=""padding:5px 15px 20px 15px""> " & vbcrlf
		tmpHTML=tmpHTML&"				<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ���� �⺻ ���� ��������
										tmpHTML=tmpHTML& getAsInfo_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ���� ��ǰ ���� ��������
										tmpHTML=tmpHTML& getAsItemLIst_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ���ּ� ��������
										tmpHTML=tmpHTML& getReqInfo_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ��ü�ּ� ��������
										tmpHTML=tmpHTML& getReturnInfo_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ó����� ��������
										tmpHTML=tmpHTML& getFinishResult_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// �ù����� ��������
										tmpHTML=tmpHTML& getDlvInfo_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				<tr><td> " & vbcrlf
										''// ��Ÿ �ȳ�����
										tmpHTML=tmpHTML&  getEtcNotice_off()
		tmpHTML=tmpHTML&"				</td></tr> " & vbcrlf

		tmpHTML=tmpHTML&"				</table> " & vbcrlf
		tmpHTML=tmpHTML&"			</td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		</table> " & vbcrlf
		tmpHTML=tmpHTML&"	</td> " & vbcrlf
		tmpHTML=tmpHTML&"</tr> " & vbcrlf
		tmpHTML=tmpHTML&"<tr> " & vbcrlf
		tmpHTML=tmpHTML&"	<td><img src=""http://fiximage.10x10.co.kr/web2008/mail/mail_footer01.gif"" width=""600"" height=""30"" /></td> " & vbcrlf
		tmpHTML=tmpHTML&"</tr> " & vbcrlf
		tmpHTML=tmpHTML&"<tr> " & vbcrlf
		tmpHTML=tmpHTML&"	<td height=""51"" style=""border-bottom:1px solid #eaeaea;""> " & vbcrlf
		tmpHTML=tmpHTML&"		<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""> " & vbcrlf
		tmpHTML=tmpHTML&"		<tr> " & vbcrlf
		tmpHTML=tmpHTML&"			<td style=""padding-left:20px;""><img src=""http://fiximage.10x10.co.kr/web2008/mail/mail_footer02.gif"" width=""245"" height=""26"" /></td> " & vbcrlf
		tmpHTML=tmpHTML&"			<td width=""128""><a href=""http://www.10x10.co.kr/cscenter/csmain.asp"" onFocus=""blur()"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/mail/mail_btn_cs.gif"" width=""108"" height=""31"" border=""0"" /></a></td> " & vbcrlf
		tmpHTML=tmpHTML&"		</tr> " & vbcrlf
		tmpHTML=tmpHTML&"		</table> " & vbcrlf
		tmpHTML=tmpHTML&"	</td> " & vbcrlf
		tmpHTML=tmpHTML&"</tr> " & vbcrlf
		tmpHTML=tmpHTML&"<tr> " & vbcrlf
		tmpHTML=tmpHTML&"	<td style=""padding:10px 0 15px 0;line-height:17px;"" class=""gray11px02""> " & vbcrlf
		tmpHTML=tmpHTML&"	(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����<br> " & vbcrlf
		tmpHTML=tmpHTML&"	��ǥ�̻� : ������  &nbsp;����ڵ�Ϲ�ȣ : 211-87-00620  &nbsp;����Ǹž� �Ű��ȣ : �� 01-1968ȣ  &nbsp;�������� ��ȣ �� û�ҳ� ��ȣå���� : �̹���<br> " & vbcrlf
		tmpHTML=tmpHTML&"	<span class=""black11px"">���ູ����:TEL 1644-6030  &nbsp;E-mail:<a href=""mailto:customer@10x10.co.kr"" class=""link_black11pxb"">customer@10x10.co.kr</a> </span> " & vbcrlf
		tmpHTML=tmpHTML&"	</td> " & vbcrlf
		tmpHTML=tmpHTML&"</tr> " & vbcrlf
		tmpHTML=tmpHTML&"</table> " & vbcrlf
		tmpHTML=tmpHTML&"</body> " & vbcrlf
		tmpHTML=tmpHTML&"</html> " & vbcrlf

		makeMailTemplate_off = tmpHTML
	End Function

	''// ���� ��� �̹���
	Public Function getMailHeadImage_off()
		dim tmpImg
		IF FDivCD="A000" Then '// �±�ȯ���
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a000_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a000_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A001" Then '// ������߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a001_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a001_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A002" Then '// ���񽺹߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a002_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a002_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A003" Then '// ȯ�ҿ�û
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a003_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a003_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A004" Then '// ��ǰ����(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a004_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a004_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A007" Then '// �ſ�/��ü���
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a007_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a007_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A008" Then '// �ֹ����
			IF FCurrState="B001" Then
				'tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a008_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a008_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A010" Then '// ȸ����û(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a010_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a010_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A011" Then '// �±�ȯȸ��(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a011_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a011_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSEIF FDivCD="A900" Then '// �ֹ���������
			IF FCurrState="B001" Then
				'tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a011_01.gif"" width=""586"" height=""240"" border=""0"">"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src=""http://fiximage.10x10.co.kr/web2008/mail/csmail_top_a900_07.gif"" width=""586"" height=""240"" border=""0"">"
			End IF
		ELSE

		END IF
		getMailHeadImage_off = tmpImg
	End Function
End Class

function ReSendCsActionMail_off(id, iForceCurrState, iForceBuyEmail)
    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	if (iForceCurrState<>"") then
        oCsAction.FRectForceCurrState = iForceCurrState
    end if
    
    if (iForceBuyEmail<>"") then
        oCsAction.FRectForceBuyEmail = iForceBuyEmail
    end if
    
	strMailHTML = oCsAction.makeMailTemplate_off(id)
	strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "�Բ��� ��û�Ͻ� ["& oCsAction.GetAsDivCDName_off &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls

	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		oMail.MailerMailGubun = 1		' ���Ϸ� �ڵ����� ��ȣ
		oMail.Send_TMSMailer()		'TMS���Ϸ�
		'oMail.Send_Mailer()
	End IF

	SET oMail = nothing
   
    Set oCsAction = Nothing
end function

Function SendCsActionMail_off(id)
    dim oCsAction,strMailHTML,strMailTitle
    
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate_off(id)
	strMailTitle = "[�ٹ�����]"& oCsAction.FCustomerName & "�Բ��� ��û�Ͻ� ["& oCsAction.GetAsDivCDName_off &"] ó���� "& oCsAction.FCurrStateName &" �Ǿ����ϴ�."

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls
		
	IF oCsAction.FBuyEmail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		oMail.SenderMail	= "customer@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oCsAction.FCustomerName
		oMail.ReceiverMail	= oCsAction.FBuyEmail
		oMail.MailConts 	= strMailHTML
		
		oMail.Send_CDO()
	End IF

	SET oMail = nothing    
    Set oCsAction = Nothing    
End Function

%>
