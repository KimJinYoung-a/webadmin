<%

'����Ʈ�� �ִ� Ŭ������ �״�� �����´�.
'�ٸ�, CS������ ������Ƚ���� �������̴�.
'procGiftCardReg ���� Ʈ������ ���ش�.

	'//����Ʈī�� ��Ȳ ������
	Class myGiftCarditem
		public FgiftOrderSerial
		public FmasterCardCode
		public FbuyDate
		public FregDate
		public FcardExpire
		public FsmallImage
		public FcardPrice
		public FcardStatus
		public FcardItemid

		Private Sub Class_initialize()
		End Sub

		Private Sub Class_Terminate()
		End Sub
	End Class

	'//����Ʈī�� ��Ȳ ������
	Class myGiftCard
		public FItemList()
		public FTotalCount
		public FResultCount
		public FCurrPage
		public FTotalPage
		public FPageSize
		public FScrollCount
		public FRectUserid

		'# ����Ʈī�� �ܾ� Ȯ��
		Public Function myGiftCardCurrentCash()
			Dim strSql
			strSQL = "exec [db_user].[dbo].sp_Ten_giftCardCurrentCash '" & CStr(FRectUserid) & "'"
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSQL, dbget, 1
				if not rsget.EOF then
					myGiftCardCurrentCash = rsget(0)
				else
					myGiftCardCurrentCash = 0
				end if
			rsget.Close
		end Function

		'# ��� ����Ʈī�� ���
		Public Sub myGiftCardRegList()
			Dim i, strSql

			'ī��Ʈ
			strSql = "exec [db_user].[dbo].sp_Ten_giftCardRegListCnt '" & CStr(FRectUserid) & "'," & FPageSize
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSQL, dbget, 1
				FTotalCount = rsget("cnt")
				FTotalPage = rsget("totPg")
			rsget.Close

			'������������ ��ü ���������� Ŭ �� �Լ�����
			if Cint(FCurrPage)>Cint(FTotalPage) then
				FResultCount = 0
				exit sub
			end if

			'���� ����
			strSql = "exec [db_user].[dbo].sp_Ten_giftCardRegList '" & CStr(FRectUserid) & "'," & FPageSize & "," & FCurrPage
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.pagesize = FPageSize
			rsget.Open strSQL, dbget, 1

			if (FCurrPage * FPageSize < FTotalCount) then
				FResultCount = FPageSize
			else
				FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
			end if

			redim preserve FItemList(FResultCount)

			i=0
			if Not(rsget.EOF or rsget.BOF) then
				rsget.absolutepage = FCurrPage
				do until rsget.EOF
					set FItemList(i) = new myGiftCarditem

					FItemList(i).FgiftOrderSerial = rsget("giftOrderSerial")
					FItemList(i).FmasterCardCode = rsget("masterCardCode")
					FItemList(i).FbuyDate = rsget("buyDate")
					FItemList(i).FregDate = rsget("regDate")
					FItemList(i).FcardExpire = rsget("cardExpire")
					FItemList(i).FcardItemid = rsget("cardItemid")
					FItemList(i).FsmallImage = webImgUrl & "/giftcard/small/" & GetImageSubFolderByItemid(rsget("cardItemid")) & "/" & rsget("smallImage")
					FItemList(i).FcardPrice = rsget("cardPrice")
					FItemList(i).FcardStatus = rsget("cardStatus")

					rsget.movenext
					i=i+1
				Loop
			end if

			rsget.Close

		End Sub

		Private Sub Class_initialize()
		End Sub

		Private Sub Class_Terminate()
		End Sub
	End Class

	'// ī�� ��� ó��
	Function procGiftCardReg(mcd)
		dim strSql, strChk
		dim giftOrderSerial, cardItemid, cardOption, cardPrice, buyDate

		'����ڵ� Ȯ�� (���ۿϷ�� �ڵ常)
		strSql = "Select giftOrderSerial, cardItemid, cardOption, totalsum, ipkumdate, jumunDiv " &_
			" From db_order.dbo.tbl_giftcard_order " &_
			" Where masterCardCode='" & mcd & "'" &_
			"	and ipkumDiv>='4' " &_
			"	and cancelYn='N' "
		rsget.Open strSql, dbget, 1

		if rsget.EOF or rsget.BOF then
			procGiftCardReg = "W"			'����ī�� ��ȣ
			rsget.close: exit Function
		else
			if rsget("jumunDiv")="7" then
				procGiftCardReg = "R"		'��ϵ� ī��
				rsget.close: exit Function
			elseif rsget("jumunDiv")="9" then
				procGiftCardReg = "C"		'��ҵ� ī��
				rsget.close: exit Function
			elseif datediff("d",rsget("ipkumdate"),date()) > (365*5) then
				procGiftCardReg = "L"		'��ȿ�Ⱓ ����
				rsget.close: exit Function
			else
				giftOrderSerial = rsget("giftOrderSerial")
				cardItemid = rsget("cardItemid")
				cardOption = rsget("cardOption")
				cardPrice = rsget("totalsum")
				buyDate = rsget("ipkumdate")
			end if
		end if

		rsget.Close

		'���ó��
		strSql = "Insert into db_user.dbo.tbl_giftcard_regList (giftOrderSerial, masterCardCode, cardItemid, cardOption, cardPrice, buyDate, cardExpire, userid, cardStatus)" &_
				" Values " &_
				" ('" & giftOrderSerial & "'" &_
				" ,'" & mcd & "'" &_
				" ,'" & cardItemid & "'" &_
				" ,'" & cardOption & "'" &_
				" ,'" & cardPrice & "'" &_
				" ,'" & formatdatetime(buyDate,2) & " " & formatdatetime(buyDate,4) & "'" &_
				" ,'" & formatdatetime(dateadd("yyyy",5,buyDate),2) & " " & formatdatetime(dateadd("yyyy",5,buyDate),4) & "'" &_
				" ,'" & GetLoginUserID & "'" &_
				" ,'1')"
		dbget.execute(strSql)

		'�ֹ����� ���� (���ó��)
		strSql = "Update db_order.dbo.tbl_giftcard_order " &_
				" Set jumunDiv='7' " &_
				" where giftOrderSerial='" & giftOrderSerial & "'"
		dbget.execute(strSql)

		'��� �α� �߰�
		strSql = "Insert into db_user.dbo.tbl_giftcard_log (userid, useCash, jukyocd, jukyo, orderserial, reguserid, siteDiv)" &_
				" Values " &_
				" ('" & GetLoginUserID & "'" &_
				" ," & cardPrice &_
				" ,100,'GIFTī�� ���'" &_
				" ,'" & giftOrderSerial & "'" &_
				" ,'" & GetLoginUserID & "'" &_
				" ,'T')"
		dbget.execute(strSql)


		'����Ȳ �߰�
		strSql = "select count(*) from db_user.dbo.tbl_giftcard_current where userid='" & GetLoginUserID & "'"
		rsget.Open strSql, dbget, 1
			strChk = rsget(0)
		rsget.Close

		if strChk>0 then
			strSql = "Update db_user.dbo.tbl_giftcard_current Set " &_
					"	currentCash = (currentCash + " & cardPrice & ") " &_
					"	,gainCash = (gainCash + " & cardPrice & ") " &_
					"	,lastUpdate = getdate() " &_
					" Where userid='" & GetLoginUserID & "'"
			dbget.execute(strSql)
		else
			strSql = "Insert Into db_user.dbo.tbl_giftcard_current (userid, currentCash, gainCash, lastupdate) values " &_
					" ('" & GetLoginUserID & "'" &_
					" ," & cardPrice &_
					" ," & cardPrice & ",getdate())"
			dbget.execute(strSql)
		end if

        IF (Err) then
		    procGiftCardReg = "E"			'ó���� �����߻�
		    Exit Function
		ELSE
		    procGiftCardReg = cardPrice		'ó�� �Ϸ�(ī��ݾ� ��ȯ)
		end if
	end Function


	'// ������ȣ Ȯ��(�ֹ���ȣ ���)
	Function getGiftCardMasterCD(osn, byRef resendCnt, byRef oIdx)
		dim strSql, strChk

		'����ڵ� Ȯ��
		strSql = "Select masterCardCode, ipkumdiv, jumunDiv, ipkumdate, resendCnt, cancelyn, idx " &_
			" From db_order.dbo.tbl_giftcard_order " &_
			" Where giftOrderSerial='" & osn & "'" &_
			"	and cancelYn='N' "
		rsget.Open strSql, dbget, 1

		if rsget.EOF or rsget.BOF then
			getGiftCardMasterCD = "W"			'����ī�� ��ȣ
			rsget.close: exit Function
		else
			if rsget("jumunDiv")="1" or rsget("ipkumdiv")<"3" then
				getGiftCardMasterCD = "A"		'������ �ֹ�
				rsget.close: exit Function
			end if

			if rsget("jumunDiv")="7" then
				getGiftCardMasterCD = "R"		'��ϵ� ī��
				rsget.close: exit Function
			end if

			if rsget("jumunDiv")="9" or rsget("ipkumdiv")="9" or rsget("cancelyn")="Y" then
				getGiftCardMasterCD = "C"		'��� �ֹ�
				rsget.close: exit Function
			end if

			if datediff("d",rsget("ipkumdate"),date()) > (365*5) then
				getGiftCardMasterCD = "E"		'��ȿ�Ⱓ ����
				rsget.close: exit Function
			end if
		end if

		oIdx = rsget("idx")									'// �ֹ� �Ϸù�ȣ ��ȯ
		resendCnt = rsget("resendCnt")						'// ������ Ƚ�� ��ȯ
		getGiftCardMasterCD = rsget("masterCardCode")		'// ������ȣ ��ȯ

		rsget.Close

	end Function

	'// ����Ʈī�� ������ȣ �߱� �α� ����
	sub putGiftCardMasterCDLog(osn,mcd,sh)
		dim strSql
		strSql = "Insert into db_order.dbo.tbl_giftcard_cdLog " &_
				"(giftOrderSerial, masterCardCode, shiftNum) values " &_
				"('" & osn & "', '" & mcd & "'," & sh & ")"
		dbget.Execute strSql
	end sub

	'// ��߼� ���� ����(������ȣ ����)
	sub chgOrderInfoResendMasterCD(osn,mcd)
		dim strSql
		strSql = "Update db_order.dbo.tbl_giftcard_order Set " &_
				"	masterCardCode='" & mcd & "' " &_
				"	,jumunDiv='5' " &_
				"	,resendCnt=resendCnt+1 " &_
				"Where giftOrderSerial='" & osn & "'"
		dbget.Execute strSql
	end sub

	'// ���ڵ�����(+�ߺ���ϰ˻�)
	function getMasterCode(no,sz,byRef sh)
		dim strSql, blChk
		blChk = false
		if sh="" then sh=0
		do Until blChk
			if (sz-sh-1)<=0 then blChk=true
			strSql = "Select count(idx) from db_order.dbo.tbl_giftcard_cdLog Where masterCardCode='" & makeMasterCode(no,sz,sh) & "'"
			rsget.Open strSql, dbget, 1
				if rsget(0)<=0 then
					blChk=true
					getMasterCode = makeMasterCode(no,sz,sh)
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

	'// �ڵ� ��ȿ�� �˻�
	function chkMasterCode(cd)
		dim tmpChk, i

		if cd="" or len(cd)<=1 then
			chkMasterCode=false
			exit function
		end if

		for i=1 to Len(cd)-1
			tmpChk = tmpChk + (mid(cd,i,1) * i)
		next
		tmpChk = right(tmpChk\(Len(cd)-1),1)

		if tmpChk=right(cd,1) then
			chkMasterCode = true
		else
			chkMasterCode = false
		end if
	end function

%>