<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim i, vQuery, vTemp, vUserID, vCatdID, vCardOpt, vOrderID, vCardPrice, vUserCell, vMMSTitle, vMMSMessage, vIsReg, vMMSOrderID, vMMSDB, vErrQu
	Dim vGiftProcQuery
	vTemp = Trim(Request("userid"))
	vTemp = Replace(vTemp," ","")
	vCatdID = Request("iid")
	vCardOpt = Request("opt")
	vMMSTitle = Request("mmstitle")
	vMMSMessage = Request("mmsmessage")
	
	IF application("Svr_Info") = "Dev" THEN
    	vMMSDB = "[ACADEMYDB].[db_LgSMS].[dbo].[mms_msg]"
    else
    	vMMSDB = "[LOGISTICSDB].[db_LgSMS].[dbo].[mms_msg]"
    end if
	
	'############################################## [1] ī�������ޱ� ##################################################
	vQuery = "Select top 1 cardSellCash From [db_item].[dbo].[tbl_giftcard_option] " &_
			" Where cardItemid = '" & vCatdID & "' and cardOption = '" & vCardOpt & "' and optSellYn = 'Y' and optIsUsing = 'Y' "
	rsget.Open vQuery, dbget, 1
	if Not(rsget.EOF or rsget.BOF) then
		vCardPrice = rsget(0)
	elseif vCardOpt = "0000" then		'�����Է�
		vCardPrice = 0
	end if
	rsget.Close
	'############################################## [1] ī�������ޱ� ##################################################
	
	
	For i = LBound(Split(vTemp,",")) To UBound(Split(vTemp,","))
	
		vUserID = Split(vTemp,",")(i)
		
		'############################################## [2] ȸ�������ޱ� ����ȣ ##################################################
		vQuery = "Select usercell From [db_user].[dbo].[tbl_user_n] Where userid = '" & vUserID & "' "
		rsget.Open vQuery, dbget, 1
		if Not rsget.EOF then
			vUserCell = rsget(0)
			rsget.Close
		else
			rsget.Close
			Response.Write "<script>alert('"&vUserID&" ���� ID�Դϴ�. Ȯ���غ�����.');</script>"
			dbget.close()
			Exit For
			Response.End
		end if
		'############################################## [2] ȸ�������ޱ� ����ȣ ##################################################
		
		
		
		'############################################## [3] �ֹ���ȣ �޾ƿ��� ##################################################
		vOrderID = fnGiftCardReg(vCatdID, vCardOpt, vCardPrice)
		
		If vOrderID = "x" Then
			Response.Write "<script>alert('�Է��� �ȵǾ����ϴ�. Ȯ���غ�����.');</script>"
			dbget.close()
			Response.End
		End If
		
		vMMSOrderID = vMMSOrderID & "'" & vOrderID & "',"
		'############################################## [3] �ֹ���ȣ �޾ƿ��� ##################################################
		
		
		
		'############################################## [4] ������ �޴��� �� ����, LMS ���� ����(��ϿϷ�� ����) ##################################################
		vQuery = "UPDATE [db_order].[dbo].[tbl_giftcard_order] SET " & "<br>" & _
				 "	jumunDiv = '7', sendhp = '1644-6030', reqhp = '" & vUserCell & "', MMSTitle = '" & vMMSTitle & "', " & "<br>" & _
				 "	MMSContent = '" & vMMSMessage & "' " & "<br>" & _
				 " WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [4] ������ �޴��� �� ����, LMS ���� ����(��ϿϷ�� ����) ##################################################
		
		
		
		'############################################## [5] ���ó�� ##################################################
		vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_regList](giftOrderSerial, masterCardCode, cardItemid, cardOption, cardPrice, buyDate, cardExpire, userid, cardStatus) " & "<br>" & _
				 "	SELECT giftOrderSerial, masterCardCode, cardItemid, cardOption, totalsum, regdate, dateadd(year,5,regdate), '" & vUserID & "', '1' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [5] ���ó�� ##################################################
		
		
		
		'############################################## [6] �α� �߰� ##################################################
		vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_log](userid, useCash, jukyocd, jukyo, orderserial, reguserid, siteDiv) " & "<br>" & _
				 "	SELECT '" & vUserID & "', totalsum, 100, 'GIFTī�� ���', giftOrderSerial, '" & session("ssBctId") & "', 'T' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		'############################################## [6] �α� �߰� ##################################################



		'############################################## [7] ����Ȳ Ȯ�� �� �߰� �� ���� ##################################################
		vQuery = "Select distinct userid From [db_user].[dbo].[tbl_giftcard_current] Where userid = '" & vUserID & "' "
		rsget.Open vQuery, dbget, 1
		if Not rsget.EOF then
			vIsReg = "o"
		else
			vIsReg = "x"
		end if
		rsget.close

		If vIsReg = "o" Then	'### ������ UPDATE
			vQuery = "UPDATE [db_user].[dbo].[tbl_giftcard_current] SET " & "<br>" & _
					 "	currentCash = (currentCash + " & vCardPrice & "), gainCash = (gainCash + " & vCardPrice & "), lastUpdate = getdate() " & "<br>" & _
					 " WHERE userid = '" & vUserID & "'"
			'dbget.execute vQuery
			vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		ElseIf vIsReg = "x" Then	'### ������ INSERT
			vQuery = "INSERT INTO [db_user].[dbo].[tbl_giftcard_current](userid, currentCash, gainCash, lastupdate) " & "<br>" & _
					 "	SELECT '" & vUserID & "', totalsum, totalsum, getdate() " & "<br>" & _
					 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
			'dbget.execute vQuery
			vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br>"
		End If
		vErrQu = vErrQu & vQuery & "<br>"
		'############################################## [7] ����Ȳ Ȯ�� �� �߰� �� ���� ##################################################
		
		
		
		'############################################## [8] MMS ������ ##################################################
		vQuery = "INSERT INTO " & vMMSDB & "(SUBJECT,PHONE,CALLBACK,STATUS,REQDATE,MSG,FILE_CNT, EXPIRETIME) " & "<br>" & _
				 "	SELECT '" & vMMSTitle & "', reqhp, '1644-6030','0',getdate() , '" & vMMSMessage & "','0','43200' " & "<br>" & _
				 "	FROM [db_order].[dbo].[tbl_giftcard_order] WHERE giftOrderSerial = '" & vOrderID & "'"
		'dbget.execute vQuery
		vErrQu = vErrQu & vQuery & "<br>"
		vGiftProcQuery = vGiftProcQuery & vQuery & "<br><br><br><br><br>"
		'############################################## [8] MMS ������ ##################################################
		
		vUserCell = ""
		vIsReg = ""
	Next
	
	vMMSOrderID = Left(vMMSOrderID,Len(vMMSOrderID)-1)
	
	'response.write vErrQu & "<br>"
	'Response.Write "<strong>�߱� �Ϸ� �Ǿ����ϴ�.</strong>"
	Response.write vGiftProcQuery

'===========================================================================================================================================
'####### �ʿ� �Լ��� #######
	'### ī���ֹ���ȣ�ޱ��Լ�
	Function fnGiftCardReg(giftItemid, giftOption, giftcardPrice)
		Dim strSql, rndjumunno, ordUserid, ordUserNm, tmpOrdSn, tmpMstCd, ordIdx
		'### �ֹ���
		ordUserid = "system"
		'ordUserid = "10x10phone"
		ordUserNm = "�ٹ�����"
		
			tmpOrdSn = "": tmpMstCd = ""
		    '�ӽ��ֹ���ȣ ����
		    Randomize
			rndjumunno = CLng(Rnd * 100000) + 1
			rndjumunno = CStr(rndjumunno)
	
			'@�ֹ��� ���� (GiftCardGbn:0, ���� 1���� ����;POS������)
			strSql = "Insert Into [db_order].[dbo].tbl_giftcard_order "
			strSql = strSql & " (giftOrderSerial,cardItemid,cardOption,masterCardCode,userid,buyname,totalsum,jumundiv,accountdiv,ipkumdiv,ipkumdate "
			strSql = strSql & " ,discountrate,subtotalprice,miletotalprice,tencardspend,referip,userlevel,sumPaymentEtc,designId,resendCnt,GiftCardGbn,notRegSpendSum) "
			strSql = strSql & " Values "
			strSql = strSql & " ('" & rndjumunno & "'," & giftItemid & ",'" & giftOption & "','','" & ordUserid & "','" & ordUserNm & "'," & giftcardPrice
			strSql = strSql & " ,'5','10','8',getdate(),1," & giftcardPrice & ",0,0,'" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "'"
			strSql = strSql & " ,7,0,'101',0,0,0)"
			dbget.Execute strSql
	
			'@IDX����
			strSql = "Select IDENT_CURRENT('[db_order].[dbo].tbl_giftcard_order') as maxitemid "
			rsget.Open strSql,dbget,1
				ordIdx = rsget("maxitemid")
			rsget.close
	
			'## �� �ֹ���ȣ/ī���ڵ� Setting
			if (Not IsNull(ordIdx)) and (ordIdx<>"") then
				dim sh: sh = 0
				tmpOrdSn = "G" & Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),4,256)
				tmpOrdSn = tmpOrdSn & Format00(5,Right(CStr(ordIdx),5))
				tmpMstCd = getMasterCode(ordIdx,16,sh)
	
				strSql = " update [db_order].[dbo].tbl_giftcard_order" + vbCrlf
				strSql = strSql + " set giftOrderSerial = '" + tmpOrdSn + "'" + vbCrlf
				strSql = strSql + " ,masterCardCode = '" + tmpMstCd + "'" + vbCrlf
				strSql = strSql + " where idx = " + CStr(ordIdx) + vbCrlf
	
				dbget.Execute strSql
	
				'# ����Ʈī�� ������ȣ �߱� �α� ����
				Call putGiftCardMasterCDLog(tmpOrdSn,tmpMstCd,sh-1)
				
				fnGiftCardReg = tmpOrdSn
			else
				fnGiftCardReg = "x"
		    end if
	End Function


    '// ���ڵ�����(+�ߺ���ϰ˻�)
	function getMasterCode(no,sz,byRef sh)
		dim strSql, blChk, bufCode
		blChk = false
		if sh="" then sh=0
		do Until blChk
			if (sz-sh-1)<=0 then blChk=true
			bufCode = makeMasterCode(no,sz,sh)
			strSql = "Select count(idx) from db_order.dbo.tbl_giftcard_cdLog Where masterCardCode='" & bufCode & "'"
			rsget.Open strSql, dbget, 1
				if rsget(0)<=0 then
					IF Not(Left(bufCode,4)="1010" or Left(bufCode,4)="6979") THEN ''preFix �� �ߺ��ȵǰ�. (1010: Point1010ȸ��ī��, 6979: �ǹ�ī��)
					    blChk=true
					    getMasterCode = bufCode
					END IF
				end if
			rsget.Close
			sh = sh +1
		loop
	end function
	
	
	'// �ڵ����(�Ϸù�ȣ, �ڵ����, �ߺ�����Ʈ / MD5�ʿ�)
	function makeMasterCode(no,sz,sh)
		dim tmpMD, tmpNo, tmpChk, i

		'���� �˻�
		if (sz>32) or ((31-sz)<sh) then
			makeMasterCode = string(sz,"0")
			exit Function
		end if

		'�����ڵ� ����
		tmpMD = MD5(no)
		for i=1 to Len(tmpMD)
			if mid(tmpMD,i,1)>="0" and mid(tmpMD,i,1)<="9" then
				tmpNo = tmpNo & mid(tmpMD,i,1)
			else
				tmpNo = tmpNo & ASC(mid(tmpMD,i,1)) mod 10
			end if
		next

		tmpNo = left(right(tmpNo,len(tmpNo)-sh),sz-1)
		
		'�����ڵ� ����
		for i=1 to Len(tmpNo)
			tmpChk = tmpChk + (mid(tmpNo,i,1) * i)
		next
		tmpChk = right(tmpChk\Len(tmpNo),1)
		
		'��� ��ȯ
		makeMasterCode = tmpNo & tmpChk
	end function
	
	
	'// ����Ʈī�� ������ȣ �߱� �α� ����
	sub putGiftCardMasterCDLog(osn,mcd,sh)
		dim strSql
		strSql = "Insert into db_order.dbo.tbl_giftcard_cdLog " &_
				"(giftOrderSerial, masterCardCode, shiftNum) values " &_
				"('" & osn & "', '" & mcd & "'," & sh & ")"
		dbget.Execute strSql
	end sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->