<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim eappidx, reqTitle, reqContent, userid, opt, sugiPrice, MMSTitle, MMSContent, makecnt, mode, menupos
eappidx			= request("eappidx")
reqTitle		= html2db(request("reqTitle"))
reqContent		= html2db(request("reqContent"))
userid			= request("userid")
opt				= request("opt")
sugiPrice		= request("sugiPrice")
MMSTitle		= html2db(request("MMSTitle"))
MMSContent		= html2db(request("MMSContent"))
makecnt			= request("makecnt")
mode			= request("mode")
menupos			= request("menupos")

If LenB(MMSTitle) >= 60 Then
	Response.Write "<script>alert('MMS ������ 60Byte �̸����� �Է��ϼž� �մϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
	Response.End
End If

Dim strRst, strSql, lp, tmpOrdSn, tmpMstCd, ordIdx
Dim sh
If userid <> "" then
	Dim useridCnt, iA2, arrTemp2, arruserid
	userid = replace(userid,",",chr(10))
	userid = replace(userid,chr(13),"")
	arrTemp2 = Split(userid,chr(10))
	iA2 = 0
	useridCnt = 0
	Do While iA2 <= ubound(arrTemp2)
		If Trim(arrTemp2(iA2))<>"" then
			arruserid = arruserid & trim(arrTemp2(iA2)) & ","
			useridCnt = useridCnt + 1
		End If
		iA2 = iA2 + 1
	Loop
	arruserid = left(arruserid,len(arruserid)-1)
End If

If mode = "I" or mode = "U" Then
	If Clng(makecnt) <> Clng(useridCnt) Then
		response.write "�߱��� ������ ���̵� ���ڰ� �ٸ�"
		response.end
	End If
End If

Dim vIdx : vIdx = request("idx")
If mode = "I" Then					'Giftī�� �� ���
	For lp = 0 to makeCnt - 1
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CNT " & vbcrlf
		strSql = strSql & " FROM db_user.dbo.tbl_user_n " & vbcrlf
		strSql = strSql & " WHERE userid = '"&Split(arruserid, ",")(lp)&"' "
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("CNT") <> 1 Then
			Response.Write "<script>alert('"& Split(arruserid, ",")(lp) &"�� ���� ID �Դϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
			Response.End
		End If
		rsget.close
	Next

	strSql = ""
	strSql = strSql & " INSERT INTO db_cs.dbo.tbl_giftcard_master " & vbcrlf
	strSql = strSql & " ([eappIdx], [reqTitle], [reqContent], [makeCnt], [opt], [sugiPrice], [mmsTitle], [mmsContent], [regdate], [isSend], regUserId) VALUES " & vbcrlf
	strSql = strSql & " ('"&eappIdx&"', '"&reqTitle&"', '"&reqContent&"', '"&makeCnt&"', '"&opt&"', '"&sugiPrice&"', '"&mmsTitle&"', '"&mmsContent&"', getdate(),'N', '"& session("ssBctID") &"')"
	dbget.execute strSql

	'@IDX����
	strSql = "Select SCOPE_IDENTITY() as maxitemid "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		ordIdx = rsget("maxitemid")
	rsget.close

	For lp = 0 to makeCnt - 1
		strSql = ""
		strSql = strSql & " INSERT INTO db_cs.dbo.tbl_giftcard_detail" & vbcrlf
		strSql = strSql & " ([midx], [userid]) VALUES " & vbcrlf
		strSql = strSql & " ('"&ordIdx&"', '"&Split(arruserid, ",")(lp)&"')"
		dbget.execute strSql
	Next

	Response.Write "<script>alert('���� �Ͽ����ϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
	Response.End
ElseIf mode = "U" Then				'Giftī�� �� ����
	For lp = 0 to makeCnt - 1
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CNT " & vbcrlf
		strSql = strSql & " FROM db_user.dbo.tbl_user_n " & vbcrlf
		strSql = strSql & " WHERE userid = '"&Split(arruserid, ",")(lp)&"' "
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("CNT") <> 1 Then
			Response.Write "<script>alert('"& Split(arruserid, ",")(lp) &"�� ���� ID �Դϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
			Response.End
		End If
		rsget.close
	Next

	strSql = ""
	strSql = strSql & " UPDATE db_cs.dbo.tbl_giftcard_master " & vbcrlf
	strSql = strSql & " SET eappIdx = '"& eappIdx &"' " & vbcrlf
	strSql = strSql & " , reqTitle = '"& reqTitle &"'" & vbcrlf
	strSql = strSql & " , reqContent = '"& reqContent &"'" & vbcrlf
	strSql = strSql & " , makeCnt = '"& makeCnt &"'" & vbcrlf
	strSql = strSql & " , opt = '"& opt &"'" & vbcrlf
	strSql = strSql & " , sugiPrice = '"& sugiPrice &"'" & vbCrlf
	strSql = strSql & " , mmsTitle = '"& mmsTitle &"'" & vbcrlf
	strSql = strSql & " , mmsContent = '"& mmsContent &"'" & vbcrlf
	strSql = strSql & " WHERE idx = '"& vIdx &"'  "
	dbget.execute strSql

	strSql = ""
	strSql = strSql & " DELETE FROM db_cs.dbo.tbl_giftcard_detail" & vbcrlf
	strSql = strSql & " WHERE midx = '"& vIdx &"'  "
	dbget.execute strSql

	For lp = 0 to makeCnt - 1
		strSql = ""
		strSql = strSql & " INSERT INTO db_cs.dbo.tbl_giftcard_detail" & vbcrlf
		strSql = strSql & " ([midx], [userid]) VALUES " & vbcrlf
		strSql = strSql & " ('"&vIdx&"', '"&Split(arruserid, ",")(lp)&"')"
		dbget.execute strSql
	Next
	Response.Write "<script>alert('���� �Ͽ����ϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
	Response.End
ElseIf mode = "S" Then				'Giftī�� ���� �߱�
	'0. ���������� ���� ���� Ȯ��
	If vIdx = "" OR opt = "" Then
		Response.Write "<script>alert('���� ��ΰ� �ƴմϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
		Response.End
	End If

	'1. �̹� �Խñ��� �߼��ߴ� �� üũ
	Dim alreadySend : alreadySend = "N"
	Dim alreadyOrderReg : alreadyOrderReg = "N"
	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt FROM db_cs.dbo.tbl_giftcard_master WHERE idx = '"& vIdx &"' and isSend = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		alreadySend = "Y"
	End If
	rsget.close

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt FROM db_cs.dbo.tbl_giftcard_detail WHERE midx = '"& vIdx &"' and isnull(giftOrderSerial, '') <> '' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") > 0 Then
		alreadyOrderReg = "Y"
	End If
	rsget.close

	'1-1. �߼��ߴٸ� ���޼��� �� ����Ʈ �̵�
	If alreadySend = "Y" OR alreadyOrderReg = "Y" Then
		Response.Write "<script>alert('�̹� �߼��Ͽ����ϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
		Response.End
	End If

	'2. �Խñ��� �߼� ���� �ʾҴٸ�
	If alreadySend <> "Y" Then
		On Error Resume Next
		dbget.beginTrans

		'3. detailIdx, useridList array�� ����
		Dim detailIdxList(), useridList()
		strSql = ""
		strSql = strSql & " SELECT idx, userid FROM db_cs.dbo.tbl_giftcard_detail WHERE midx = '"& vIdx &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		makecnt = rsget.RecordCount
		Redim preserve detailIdxList(makecnt)
		Redim preserve useridList(makecnt)
		lp = 1
		If not rsget.EOF Then
			Do until rsget.EOF
				detailIdxList(lp)	= rsget("idx")
				useridList(lp)		= rsget("userid")
				lp = lp + 1
				rsget.moveNext
			Loop
		End If
		rsget.close

		'4. �� ID���� giftOrderSerial, masterCardCode �Ҵ� ��Ŵ
		'Q1) giftOrderSerial unique ���? -> ���⼱ ����Ʈ ���� �״��, �ձ��ڸ� G => J�� ������ => �ٽ� G�� ������
		'Q2) masterCardCode  unique ���? -> ���⼱ ����Ʈ ���� �״��
		For lp=1 to makecnt
			tmpOrdSn = "G" & Mid(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),4,256)
			tmpOrdSn = tmpOrdSn & Format00(5,Right(detailIdxList(lp),5))
			tmpMstCd = getMasterCode(detailIdxList(lp),16,sh)

			strSql = ""
			strSql = strSql & " UPDATE db_cs.dbo.tbl_giftcard_detail "
			strSql = strSql & " SET giftOrderSerial = '"& tmpOrdSn &"' "
			strSql = strSql & " ,masterCardCode = '"& tmpMstCd &"' "
			strSql = strSql & " WHERE midx = '"& vIdx &"' "
			strSql = strSql & " and idx = '"& detailIdxList(lp) &"'"
			strSql = strSql & " and giftOrderSerial is null  "
			dbget.Execute strSql
		Next

		'4-1. �ɼǿ� �ش��ϴ� ���� ��� ������ ����
		Dim giftcardPrice
		giftcardPrice = fntGiftCardPrice(vIdx, opt)

		'5. tmpGiftTBL �� �ӽ����̺� ����
		strSql = ""
		strSql = strSql & " SELECT * "
		strSql = strSql & " INTO #tmpGiftTBL "
		strSql = strSql & " FROM [db_order].[dbo].tbl_giftcard_order "
		strSql = strSql & " WHERE 1=2 "
		dbget.execute strSql

		'6. tmpGiftTBL ������ �Է�
		strSql = ""
		strSql = strSql & " INSERT INTO #tmpGiftTBL "
		strSql = strSql & " (giftOrderSerial,cardItemid,cardOption,masterCardCode,userid,buyname,totalsum,jumundiv,accountdiv,ipkumdiv,ipkumdate "
		strSql = strSql & " ,discountrate,subtotalprice,miletotalprice,tencardspend,referip,userlevel,sumPaymentEtc,designId,resendCnt,GiftCardGbn,notRegSpendSum "
		strSql = strSql & " , regdate, cancelyn, sendDiv, bookingYn, sendhp, reqhp, MMSTitle, MMSContent) "
		strSql = strSql & " SELECT d.giftOrderSerial, '101', m.opt, d.masterCardCode, 'system', '�ٹ�����' "
		strSql = strSql & " , CASE WHEN m.opt = '0000' THEN m.sugiPrice ELSE '"& giftcardPrice &"' END, '7' as jumundiv,'10','8',getdate() "
		strSql = strSql & " , 1, CASE WHEN m.opt = '0000' THEN m.sugiPrice ELSE '"& giftcardPrice &"' END, 0,0,'"& Left(request.ServerVariables("REMOTE_ADDR"),32) &"', 7, 0, '101', 0, 0, 0 "
		strSql = strSql & " , getdate(), 'N', 'S', 'N', '1644-6030', n.usercell, m.mmsTitle, m.mmsContent "
		strSql = strSql & " FROM db_cs.dbo.tbl_giftcard_master as m "
		strSql = strSql & " JOIN db_cs.dbo.tbl_giftcard_detail as d on m.idx = d.midx "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_n as n on d.userid = n.userid "
		strSql = strSql & " WHERE m.idx = '"& vIdx &"'  "
		dbget.execute strSql

		'7.���� ������ �ʿ��ϴٸ�  db_order.dbo.tbl_giftcard_order ���⿡ �Է����� �����ϱ�
		'������� userid�� ���� �����ϴ� id�����..���

		'8. 7���� ��������  db_order.dbo.tbl_giftcard_order�� �Է��ϱ�
		strSql = ""
		strSql = strSql & " INSERT INTO [db_order].[dbo].tbl_giftcard_order (giftOrderSerial,cardItemid,cardOption,masterCardCode,userid,buyname,totalsum,jumundiv,accountdiv,ipkumdiv,ipkumdate "
		strSql = strSql & " ,discountrate,subtotalprice,miletotalprice,tencardspend,referip,userlevel,sumPaymentEtc,designId,resendCnt,GiftCardGbn,notRegSpendSum "
		strSql = strSql & " , regdate, cancelyn, sendDiv, bookingYn, sendhp, reqhp, MMSTitle, MMSContent) "
		strSql = strSql & " SELECT giftOrderSerial,cardItemid,cardOption,masterCardCode,userid,buyname,totalsum,jumundiv,accountdiv,ipkumdiv,ipkumdate "
		strSql = strSql & " ,discountrate,subtotalprice,miletotalprice,tencardspend,referip,userlevel,sumPaymentEtc,designId,resendCnt,GiftCardGbn,notRegSpendSum "
		strSql = strSql & " , regdate, cancelyn, sendDiv, bookingYn, sendhp, reqhp, MMSTitle, MMSContent "
		strSql = strSql & " FROM #tmpGiftTBL "
		dbget.execute strSql

		'9. ����Ʈī�� ������ȣ �߱� �α� ����..shiftNum�� 0�̿��� ��?
		strSql = ""
		strSql = strSql & " INSERT INTO db_order.dbo.tbl_giftcard_cdLog "
		strSql = strSql & " (giftOrderSerial, masterCardCode, shiftNum) "
		strSql = strSql & " SELECT d.giftOrderSerial, d.masterCardCode, 0 "
		strSql = strSql & " FROM db_cs.dbo.tbl_giftcard_detail as d "
		strSql = strSql & " JOIN db_order.dbo.tbl_giftcard_order as o on d.giftOrderSerial = o.giftOrderSerial "
		strSql = strSql & " WHERE d.midx = '"& vIdx &"'  "
		dbget.Execute strSql

		'10. ���ó��
		strSql = ""
		strSql = strSql & " INSERT INTO db_user.dbo.tbl_giftcard_regList "
		strSql = strSql & " (giftOrderSerial, masterCardCode, cardItemid, cardOption, cardPrice, buyDate, cardExpire, userid, cardStatus) "
		strSql = strSql & " SELECT d.giftOrderSerial, d.masterCardCode, o.cardItemid, o.cardOption, o.totalsum, o.regdate, dateadd(year,5,o.regdate), d.userid, '1'  "
		strSql = strSql & " FROM db_cs.dbo.tbl_giftcard_detail as d "
		strSql = strSql & " JOIN db_order.dbo.tbl_giftcard_order as o on d.giftOrderSerial = o.giftOrderSerial "
		strSql = strSql & " WHERE d.midx = '"& vIdx &"'  "
		dbget.execute strSql

		'11. �α� �߰�
		strSql = ""
		strSql = strSql & " INSERT INTO db_user.dbo.tbl_giftcard_log "
		strSql = strSql & " (userid, useCash, jukyocd, jukyo, orderserial, reguserid, siteDiv) "
		strSql = strSql & " SELECT d.userid, o.totalsum, 100, 'GIFTī�� ���', o.giftOrderSerial, '"& session("ssBctID") &"', 'T' "
		strSql = strSql & " FROM db_cs.dbo.tbl_giftcard_detail as d "
		strSql = strSql & " JOIN db_order.dbo.tbl_giftcard_order as o on d.giftOrderSerial = o.giftOrderSerial "
		strSql = strSql & " WHERE d.midx = '"& vIdx &"'  "
		dbget.execute strSql

		'12. ����Ȳ �߰�
		For lp=1 to makecnt
			strSql = ""
			strSql = strSql & " IF EXISTS(SELECT userid FROM db_user.dbo.tbl_giftcard_current WHERE userid = '"& useridList(lp) &"') "
			strSql = strSql & " 	BEGIN "
			strSql = strSql & " 		UPDATE db_user.dbo.tbl_giftcard_current "
			strSql = strSql & " 		SET currentCash = (currentCash + "& giftcardPrice &")  "
			strSql = strSql & " 		,gainCash = (gainCash + "& giftcardPrice &")  "
			strSql = strSql & " 		,lastUpdate = getdate()  "
			strSql = strSql & " 		WHERE userid = '"& useridList(lp) &"' "
			strSql = strSql & " 	END "
			strSql = strSql & " ELSE "
			strSql = strSql & " 	BEGIN "
			strSql = strSql & " 		INSERT INTO db_user.dbo.tbl_giftcard_current (userid, currentCash, gainCash, lastupdate) "
			strSql = strSql & " 		SELECT TOP 1 d.userid, o.totalsum, o.totalsum, getdate() "
			strSql = strSql & " 		FROM db_cs.dbo.tbl_giftcard_detail as d  "
			strSql = strSql & " 		JOIN db_order.dbo.tbl_giftcard_order as o on d.giftOrderSerial = o.giftOrderSerial  "
			strSql = strSql & " 		WHERE d.midx = '"& vIdx &"' and d.userid = '"& useridList(lp) &"' "
			strSql = strSql & " 	END "
			dbget.execute strSql
		Next

		If (Err) then
			rw Err.Description
			dbget.RollbackTrans
			response.end
		Else
			dbget.CommitTrans
		End if
		On error Goto 0

		'13. �߼ۿϷ� LMS ����
		strSql = ""
		strSql = strSql & " INSERT INTO [SMSDB].db_LgSMS.dbo.MMS_MSG (SUBJECT, PHONE, CALLBACK, STATUS, REQDATE, MSG, FILE_CNT, EXPIRETIME) "
		strSql = strSql & " SELECT o.MMSTitle, o.reqhp, '1644-6030','0',getdate() , o.MMSContent, '0','43200' "
		strSql = strSql & " FROM db_cs.dbo.tbl_giftcard_detail as d   "
		strSql = strSql & " JOIN db_order.dbo.tbl_giftcard_order as o on d.giftOrderSerial = o.giftOrderSerial "
		strSql = strSql & " WHERE d.midx = '"& vIdx &"' "
		dbget.execute strSql

		'14. �ش� �� �߼� ó��
		strSql = ""
		strSql = strSql & " UPDATE db_cs.dbo.tbl_giftcard_master " & vbcrlf
		strSql = strSql & " SET isSend = 'Y', isSendDate = getdate() "
		strSql = strSql & " WHERE idx = '"& vIdx &"'  "
		dbget.execute strSql
		Response.Write "<script>alert('�߼� �Ͽ����ϴ�.');location.replace('/admin/giftcard/list.asp?menupos="&menupos&"')</script>"
		Response.End
	End If
Else
	rw "�ý��� ����"
	response.end
End If

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

Function fntGiftCardPrice(iidx, iopt)
	Dim strSql
	If iopt = "0000" Then
		strSql = ""
		strSql = strSql & " SELECT TOP 1 sugiPrice FROM db_cs.dbo.tbl_giftcard_master "
		strSql = strSql & " WHERE idx = '"& iidx &"' "
	Else
		strSql = ""
		strSql = strSql & " SELECT TOP 1 cardSellCash FROM db_item.dbo.tbl_giftcard_option "
		strSql = strSql & " WHERE cardItemid='101' "
		strSql = strSql & "	and cardOption='" & iopt & "'"
		strSql = strSql & "	and optSellYn='Y' and optIsUsing='Y' "
	End If
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not(rsget.EOF or rsget.BOF) then
		fntGiftCardPrice = rsget(0)
	end if
	rsget.Close
End Function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->