<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/cscenter/lib/TaxSheetFunc.asp"-->
<%

	'// ���� ���� //
	dim taxIdx, busiIdx, mode, menupos
	dim page, searchDiv, searchKey, searchString, param
	dim SQL, msg, retURL

	dim oTax
	dim taxSheetIssueType

	'// �Ķ���� ���� //
	taxIdx = request("taxIdx")
	if taxIdx="" and request("chkSelect")<>"" then
		taxIdx = request("chkSelect")
	end if
	busiIdx = request("busiIdx")
	mode = request("mode")
	menupos = request("menupos")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	param = "&menupos=" & menupos & "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����


	'// ���۰� üũ �Լ� //
	Function chkTRStr(str)
		dim stmp
		'stmp = Replace(str,"&", "#amp;")
		'stmp = Replace(str,"=", "#equ;")
		stmp = Server.URLEncode(str)
		chkTRStr = stmp
	End Function


	'///// ��庰 �б� ó�� /////
	Select Case mode
		'����ڵ���� Ȯ�� ó��
		Case "BusiOk"
			SQL =	" Update db_order.[dbo].tbl_busiinfo Set " &_
					"	confirmYn='Y '" &_
					" Where busiIdx=" & busiIdx
			msg = "����ڵ������ Ȯ��ó���Ͽ����ϴ�."
			retURL = "tax_view.asp?taxIdx=" & taxIdx & param

		'��û�� ��� ó��
		Case "sheetOk"
			msg = fnTaxSheetPrint(taxIdx)
			if msg="OK" then
				SQL=""
				msg = "���ݰ�꼭�� �����Ͽ����ϴ�."
				retURL = "tax_view.asp?taxIdx=" & taxIdx & param
			else
			    ''response.write "msg:"&msg
				response.write	"<script language='javascript'>" &_
								"	alert('" & msg & "');" &_
								"	history.back();" &_
								"</script>"
				dbget.close()	:	response.End
			end if

		'��û�� �ϰ���� ó��
		Case "BatchOk"
			dim arrIdx, lp
			arrIdx = split(taxIdx,",")
			for lp=0 to ubound(arrIdx)
				msg = fnTaxSheetPrint(arrIdx(lp))
				if msg<>"OK" then
					response.write	"<script language='javascript'>" &_
									"	alert('" & msg & " [" & arrIdx(lp) & "]');" &_
									"	history.back();" &_
									"</script>"
					dbget.close()	:	response.End
				end if
			next

			SQL=""
			msg = "���ݰ�꼭 �� " & ubound(arrIdx)+1 & " �� �����Ͽ����ϴ�."
			retURL = "tax_list.asp?taxIdx=" & taxIdx & param

		'��û�� ����
		Case "sheetDel"
			SQL =	" Update db_order.[dbo].tbl_taxSheet Set " &_
					"	delYn = 'Y' " &_
					" Where taxIdx=" & taxIdx
			msg = "���ݰ�꼭 ��û ������ �����Ͽ����ϴ�."
			retURL = "tax_list.asp?taxIdx=" & param
		Case Else

	End Select

	'// DBó�� �� ������ �̵�

	'Ʈ������ ����
	if SQL<>"" then
		dbget.beginTrans
		dbget.Execute(SQL)

		if (mode = "sheetDel") then

			set oTax = new CTax
			oTax.FRecttaxIdx = taxIdx

			oTax.GetTaxRead

			taxSheetIssueType = GetTaxSheetIssueType(taxIdx)

			if (taxSheetIssueType = "orderserial") then
			    SQL = " update [db_order].[dbo].tbl_order_master" & VbCrlf
			    SQL = SQL & " set " & VbCrlf
			    SQL = SQL & " authcode = (case when accountdiv in ('7', '20') then NULL else authcode end) " + VbCrlf
			    SQL = SQL & " , cashreceiptreq = NULL " + VbCrlf
			    SQL = SQL & " where orderserial='" & oTax.FTaxList(0).Forderserial & "'"
			    dbget.Execute SQL

			elseif (taxSheetIssueType = "etcmeachul") then
				SQL = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
				SQL = SQL + " set issuestatecd = NULL " + vbCrlf
				SQL = SQL + " where idx=" & oTax.FTaxList(0).Forderidx
				rsget.Open SQL, dbget, 1

				if CStr(oTax.FTaxList(0).Forderidx) = "0" then
					SQL = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
					SQL = SQL + " set issuestatecd = NULL " + vbCrlf
					SQL = SQL + " where idx in (select matchlinkkey from db_order.dbo.tbl_taxSheet_Match where matchtype = 'E' and taxidx = " & oTax.FTaxList(0).Ftaxidx & ") "
					rsget.Open SQL, dbget, 1
				end if
			end if


			''if (Not IsNull(oTax.FTaxList(0).Forderserial)) and (oTax.FTaxList(0).Forderserial <> "") then '�� ���ݰ�꼭 ��û
			''	if (Left(oTax.FTaxList(0).Forderserial, 2) <> "SO") then
			''	end if
			''end if
		end if
	end if

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		if SQL<>"" then
			dbget.CommitTrans				'Ŀ��(����)
		end if

		response.write	"<script language='javascript'>" &_
						"	alert('" & msg & "');" &_
						"	self.location='" & retURL & "';" &_
						"</script>"
	Else
	    dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
						"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
						"	history.back();" &_
						"</script>"
	End If

	'### ���ݰ�꼭 �߱� �Լ� ###
	Function fnTaxSheetPrint(taxIdx)

		dim oTax, strSql, isueDate
	    dim preIssuedCnt
	    ''����೻�� �ִ��� üũ (�ߺ�üũ) : ���ݰ�꼭.
	    strSql = "select count(*) as cnt from db_order.[dbo].tbl_taxSheet"
        strSql = strSql&" where orderserial in (select orderserial from db_order.[dbo].tbl_taxSheet where taxidx="&taxIdx&")"
        strSql = strSql&" and isueYn='Y' and DelYn='N'"
        strSql = strSql&" and taxidx<>"&taxIdx&""

        rsget.Open strSql, dbget, 1
           preIssuedCnt  = rsget("cnt")
        rsget.Close

        if (preIssuedCnt>0) then
            errMsg = "���� ���� ������ ���� �մϴ�."
            fnTaxSheetPrint = errMsg
            exit function
        end if

        ''���ݿ����� ��û���� �ִ��� üũ
        strSql = 	" Select count(*) as cnt " &_
			" From [db_log].[dbo].tbl_cash_receipt " &_
			" Where orderserial in (select orderserial from db_order.[dbo].tbl_taxSheet where taxidx="&taxIdx&")" &_
			"	and cancelyn='N'" &_
			"	and resultcode='00'"

		rsget.Open strSql, dbget, 1
           preIssuedCnt  = rsget("cnt")
        rsget.Close

        if (preIssuedCnt>0) then
            errMsg = "���ݿ����� ���� ������ ���� �մϴ�."
            fnTaxSheetPrint = errMsg
            exit function
        end if


		Dim objXMLHTTP
		Dim reqParam 'ȣ���Ķ����
		Dim tmpArr, errMsg
		Dim retval  'ȣ����
		Dim tax_no  '���ݰ�꼭 ��ȣ
		Dim tel1, tel2, tel3, sIdx
		Dim cur_c_corp_no, cur_u_user_no
		errMsg = "OK"

		'# �⺻ ���� ����
		set oTax = new CTax
		oTax.FRecttaxIdx = taxIdx
		oTax.GetTaxRead

		'��ȭ��ȣ �и�
		if db2html(oTax.FTaxList(0).FrepTel)<>"" then
		    oTax.FTaxList(0).FrepTel = Replace(oTax.FTaxList(0).FrepTel,")","-")
			if instr(db2html(oTax.FTaxList(0).FrepTel),"-")<1 then
				tmpArr = array(left(db2html(oTax.FTaxList(0).FrepTel),3),mid(db2html(oTax.FTaxList(0).FrepTel),4,4),right(db2html(oTax.FTaxList(0).FrepTel),4))
			else
				tmpArr = split(db2html(oTax.FTaxList(0).FrepTel),"-")
			end if
			tel1 = tmpArr(0)
			tel2 = tmpArr(1)
			tel3 = tmpArr(2)
		end if
		sIdx = Replace(Date(),"-", "") & taxIdx

		'�ۼ��� ����(2009.01.06;������)
		if datediff("m",oTax.FTaxList(0).Fipkumdate,date())=0 then
			'�Ա����� ����ް� ������ �Ա��Ϸ�
			isueDate = dateSerial(Year(oTax.FTaxList(0).Fipkumdate),Month(oTax.FTaxList(0).Fipkumdate),Day(oTax.FTaxList(0).Fipkumdate))
		elseif datediff("m",oTax.FTaxList(0).Fipkumdate,date())=1 and datediff("d",date(),dateSerial(year(date),month(date),5))>=0 then
			'�Ա����� �������̸鼭 ��� 5�� �����̶�� �Ա��Ϸ�
			isueDate = dateSerial(Year(oTax.FTaxList(0).Fipkumdate),Month(oTax.FTaxList(0).Fipkumdate),Day(oTax.FTaxList(0).Fipkumdate))
		else
			'�׷��� ������ �ݿ� 1�Ϸ� ����
			isueDate = dateSerial(Year(date),Month(date),1)
		end if

		'�׽�Ʈ/�Ǽ��� ����
		if (application("Svr_Info")="Dev") then
    		cur_c_corp_no = "10001568"		'(��� ��ȣ - test:10001568)
    		cur_u_user_no = "1000794"		'(����� ��ȣ - test:1000794)
    	else
    		cur_c_corp_no = "57911"			'(��� ��ȣ - Real:57911)
    		cur_u_user_no = "261746"		'(����� ��ȣ - Real:261746)
    	end if

	    '���� �Ķ���� ����
	    '"&item_nm=" & chkTRStr(LeftB(db2html(oTax.FTaxList(0).Fitemname),40)) &_ (��ǰ�� ����;2006-06-26;������)
		reqParam =	"uniq_id=" & sIdx &_
	    			"&biz_no=" & Replace(oTax.FTaxList(0).FbusiNo,"-", "") &_
	    			"&corp_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiName)) &_
	    			"&ceo_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiCEOName)) &_
	    			"&biz_status=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiType)) &_
	    			"&biz_type=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiItem)) &_
	    			"&addr=" & chkTRStr(db2html(oTax.FTaxList(0).FbusiAddr)) &_
	    			"&dam_nm=" & chkTRStr(db2html(oTax.FTaxList(0).FrepName)) &_
	    			"&email=" & chkTRStr(db2html(oTax.FTaxList(0).FrepEmail)) &_
	    			"&hp_no1=" & tel1 &_
	    			"&hp_no2=" & tel2 &_
	    			"&hp_no3=" & tel3 &_
	    			"&serial_no=" & sIdx &_
	    			"&item_count=1" &_
	    			"&item_nm=" & chkTRStr("�����λ繫��ǰ") &_
	    			"&item_qty=1" &_
	    			"&item_price=" & oTax.FTaxList(0).FtotalPrice-oTax.FTaxList(0).FtotalTax &_
	    			"&item_vat=" & oTax.FTaxList(0).FtotalTax &_
	    			"&cur_c_corp_no=" & cur_c_corp_no &_
	    			"&cur_u_user_no=" & cur_u_user_no &_
					"&api_no=1" &_
					"&enc_yn=N" &_
					"&cur_biz_no=2118700620" &_
					"&cur_corp_nm=" & chkTRStr("(��)�ٹ�����") &_
					"&cur_ceo_nm=" & chkTRStr("��â��") &_
					"&cur_biz_status=" & chkTRStr("����") &_
					"&cur_biz_type=" & chkTRStr("���ڻ�ŷ�") &_
					"&cur_addr=" & chkTRStr("����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����") &_
					"&cur_dam_nm=" & chkTRStr("�ٹ�����") &_
					"&cur_email=" & chkTRStr("customer@10x10.co.kr") &_
					"&cur_hp_no1=02" &_
					"&cur_hp_no2=1644" &_
					"&cur_hp_no3=6030" &_
					"&write_date=" & replace(isueDate,"-","") &_
					"&sb_type=01" &_
					"&pc_gbn=C" &_
					"&tax_type=01" &_
					"&bill_type=01" &_
					"&approve_type=11" &_
					"&final_status=12" &_
					"&Cash_amt=" & oTax.FTaxList(0).FtotalPrice
		'' "&credit_amt=" & oTax.FTaxList(0).FtotalPrice ������ �߰� : �ܻ�̼������� ǥ��
		'response.Write reqParam
		'dbget.close()	:	response.End

		'XML��� ���۳�Ʈ ����
		Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
		if (application("Svr_Info")="Dev") then
			objXMLHTTP.Open "POST",	"http://ifs.neoport.net:8383/tx_create.req", False
		else
			objXMLHTTP.Open "POST",	"http://web1.neoport.net:8383/tx_create.req", False
		end if
		objXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		objXMLHTTP.Send reqParam
'response.write "retval="&retval
		retval = trim(objXMLHTTP.responseText)
		retval = replace(retval,Vbcrlf,"")
		retval = replace(retval,Vbcr,"")
		retval = replace(retval,Vblf,"")
		Set objXMLHTTP = Nothing


	    If retval <> "" Then
	        tmpArr = Split(retval, "|")
	        tax_no = Clng(tmpArr(0))

	        if tax_no<0 then
	        	errMsg = retval
	        else
	        	errMsg = "OK"
	        end if
	    else
			errMsg = "����� ������ �߻��߽��ϴ�."
	    End If

		if errMsg="OK" then
			'������ ó��
			strSql =	" Update db_order.[dbo].tbl_busiinfo Set " &_
					"	confirmYn='Y '" &_
					" Where busiIdx=" & oTax.FTaxList(0).FbusiIdx & ";" & chr(13)&chr(10)

			strSql = strSql & " Update db_order.[dbo].tbl_taxSheet Set " &_
						"	neoTaxNo = '" & tax_no & "' " &_
						"	,curUserId = '" & Session("ssBctId") & "' " &_
						"	,printDate = getdate() " &_
						"	,isueYn = 'Y' " &_
						"	,isueDate = '" & isueDate & "' " &_
						" Where taxIdx=" & taxIdx
			dbget.Execute(strSql)
		end if

		set oTax = Nothing

		'// ��� �ݳ�
		fnTaxSheetPrint = errMsg
	End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

