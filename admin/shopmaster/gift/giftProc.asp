<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ǰ db ó��
' History : 2008.04.02 ������ ����
'			2020.03.27 �ѿ�� ����(����ǰ���� üũ �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i
Dim sMode, strSql,strSqlAdd, itemoption, itemgubun
Dim eCode,gCode
Dim sTitle, dSDay, dEDay, iGiftScope, sBrand, iGroupCode, iGiftType, iGiftRange1, iGiftRange2, iGiftKindCnt, iGiftKindType, iGiftLimit
Dim dSDayTime, dEDayTime
Dim sGiftKindName, itemid, sGiftKindImg, iGiftKindCode, sGiftDelivery, iGiftStatus, sGiftUsing,igStatus,sOpenDate,sCloseDate
Dim iSerachType,sSearchTxt, sDate,sSdate,sEdate,iCurrpage,sgStatus
Dim iSiteScope, sPartnerID, barcode, stockitemexists, sqlStr
Dim strParm, giftkind_name
Dim s120Img, s401Img, s402Img, s403Img, s404Img, s405Img
Dim giftkind_linkGbn, bcouponidx
Dim givecnt, vChangeContents, vSCMChangeSQL
Dim prd_itemgubun, prd_itemid, prd_itemoption, gift_delivery, gift_code, makerid
dim gift_isusing, gift_text1, gift_img1, gift_text2, gift_img2, gift_text3, gift_img3, gift_infotext
	gift_delivery = requestCheckVar(request("gift_delivery"),1)
	gift_code = requestCheckVar(getNumeric(request("gift_code")),10)
	sMode = requestCheckVar(Request.Form("sM"),32)
	giftkind_linkGbn = requestCheckVar(Request.Form("giftkind_linkGbn"),1)
	bcouponidx       = requestCheckVar(Request.Form("bcouponidx"),10)
	gift_isusing = requestCheckVar(Request.Form("gift_isusing"),1)
	gift_text1 = requestCheckVar(Request.Form("gift_text1"),256)
	gift_img1 = requestCheckVar(Request.Form("gift_img1"),128)
	gift_text2 = requestCheckVar(Request.Form("gift_text2"),256)
	gift_img2 = requestCheckVar(Request.Form("gift_img2"),128)
	gift_text3 = requestCheckVar(Request.Form("gift_text3"),256)
	gift_img3 = requestCheckVar(Request.Form("gift_img3"),128)
	gift_infotext = requestCheckVar(Request.Form("gift_infotext"),1)
	makerid = requestCheckVar(Request.Form("makerid"),32)
	giftkind_name = requestCheckVar(Request.Form("giftkind_name"),60)

if gift_infotext="" then gift_infotext="N"
if gift_text1 <> "" then
	if checkNotValidHTML(gift_text1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if gift_img1 <> "" then
	if checkNotValidHTML(gift_img1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if gift_text2 <> "" then
	if checkNotValidHTML(gift_text2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if gift_img2 <> "" then
	if checkNotValidHTML(gift_img2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if gift_text3 <> "" then
	if checkNotValidHTML(gift_text3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if gift_img3 <> "" then
	if checkNotValidHTML(gift_img3) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
''rw 	bcouponidx
''rw giftkind_linkGbn
'===========================================================================
'���������� ���� ��� Ȯ��
	Function fnChkDelivery(ByVal iGiftScope, ByVal sGiftDelivery, ByVal eCode, ByVal Brand, ByVal egCode, ByVal gCode)
		IF sGiftDelivery ="Y" THEN '��ü����� ���
			strSqlAdd = " and deliverytype not in (2,5,7,9)"
		ELSE
			strSqlAdd = " and deliverytype not in (1,4)"
		END IF

		IF 	iGiftScope = 1 THEN '��� ������ ������ ���
			IF sGiftDelivery ="Y" THEN
				Alert_return("��������� ����ǰ�� ��쿡�� �ٹ����ٹ�۸� ���� �մϴ�. ������ �ٽ� �������ּ��� ")
	      	 dbget.close()	:	response.End
			END IF

			''2011-10 �߰�
			IF (sGiftDelivery ="C") THEN  ''�����ΰ��.
			    if (Len(requestCheckVar(Request.Form("bcouponidx"),9))<1) then
			        Alert_return("���� ��ȣ ����.... ")
			        response.End
			    end if

			    strSql = " select top 1 C.* "
                strSql = strSql & "	from db_user.dbo.tbl_user_coupon_master C "
                strSql = strSql & " where idx="&requestCheckVar(Request.Form("bcouponidx"),9)
                strSql = strSql & " and isopenlistcoupon='Y'"
                strSql = strSql & " and startdate>=(select evt_startdate from db_event.dbo.tbl_event where evt_code="&eCode&")"
                strSql = strSql & " and expiredate>(select evt_enddate from db_event.dbo.tbl_event where evt_code="&eCode&")"
''rw strSql
    			rsget.Open strSql, dbget
    			IF not (rsget.EOF OR rsget.BOF) THEN
    			    ''
    			ELSE
    			    rsget.close
    			    Alert_return("��ϵ� ���� �ڵ尡 ���� ���� �ʰų�, ��¥ ���� �Ǵ� ���ð�(������) ����Ÿ���� �ƴմϴ�. ")
    	      	    dbget.close()	:	response.End
    			END IF
    			rsget.close
			END IF
		ELSEIF 	iGiftScope = 2 THEN '�̺�Ʈ��ϻ�ǰ ������ ���
			IF eCode ="" OR eCode = "0" THEN
				Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")
	       dbget.close()	:	response.End
	    	END IF

			strSql = " SELECT deliverytype FROM [db_event].[dbo].[tbl_eventitem] AS A INNER JOIN [db_item].[dbo].[tbl_Item] AS B ON A.itemid = B.itemid  "&_
					"  WHERE  evt_code = "&eCode& strSqlAdd
			rsget.Open strSql, dbget
			IF not (rsget.EOF OR rsget.BOF) THEN
				Alert_return("������ �̺�Ʈ��ϻ�ǰ�� ����ǰ���Ÿ�԰� �ٸ� ��ǰ�� �����մϴ�. ������ �ٽ� �������ּ��� ")
	      	 dbget.close()	:	response.End
			END IF
			rsget.close
		ELSEIF 	iGiftScope = 3 THEN '���ú귣�� ������ ���
			strSql = " SELECT deliverytype FROM  [db_item].[dbo].[tbl_Item] where makerid = '"&sBrand&"' "& strSqlAdd
			rsget.Open strSql, dbget
			IF not (rsget.EOF OR rsget.BOF) THEN
			%>
			<script language="javascript">
			<!--
			if(confirm("������ �귣���ǰ�� ����ǰ���Ÿ�԰� �ٸ� ��ǰ�� �����մϴ�. \n �� ��ǰ�� ���ؼ��� ����ǰ�� �߼۵��� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?")){
				return;
			}else{
				history.back();
			}
			//-->
			</script>
			<%
			END IF
			rsget.close
		ELSEIF 	iGiftScope = 4 THEN '���ñ׷��ǰ  ������ ���
			strSql = " SELECT deliverytype FROM [db_event].[dbo].[tbl_eventitem] AS A INNER JOIN [db_item].[dbo].[tbl_Item] AS B ON A.itemid = B.itemid  "&_
					"  WHERE  evt_code = "&eCode& " and evtgroup_code ="&egCode&strSqlAdd
			rsget.Open strSql, dbget
			IF not (rsget.EOF OR rsget.BOF) THEN
				Alert_return("������ �׷��ǰ��  ����ǰ���Ÿ�԰� �ٸ� ��ǰ�� �����մϴ�. ������ �ٽ� �������ּ��� ")
	      	 dbget.close()	:	response.End
			END IF
			rsget.close
		ELSEIF 	iGiftScope = 5 THEN '���û�ǰ  ������ ���
			strSql = " SELECT deliverytype FROM [db_event].[dbo].[tbl_giftitem] AS A INNER JOIN [db_item].[dbo].[tbl_Item] AS B ON A.itemid = B.itemid  "&_
					"  WHERE  gift_code = "&gCode&strSqlAdd
			rsget.Open strSql, dbget
			IF not (rsget.EOF OR rsget.BOF) THEN
				Alert_return("���û�ǰ��  ����ǰŸ�԰� �ٸ� ��ǰ�� �����մϴ�. ������ �ٽ� �������ּ��� ")
	      	 dbget.close()	:	response.End
			END IF
			rsget.close
		END IF
	End Function
'===========================================================================
SELECT CASE sMode
Case "I"	'//����ǰ ���
	eCode			= requestCheckVar(Request.Form("eC"),10)
	IF eCode ="" THEN eCode = 0
	sTitle			= html2db(requestCheckVar(Request.Form("sGN"),64))
	dSDay 			= requestCheckVar(Request.Form("sSD"),10)
	dEDay			= requestCheckVar(Request.Form("sED"),10)
	dSDayTime		= request("sSDTime")
	dEDayTime		= request("sEDTime")

	iGiftScope		= requestCheckVar(Request.Form("giftscope"),4)
	sBrand			= requestCheckVar(Request.Form("ebrand"),32)
	iGroupCode		= requestCheckVar(Request.Form("selG"),10)
	iGiftType		= requestCheckVar(Request.Form("gifttype"),10)
	iGiftRange1		= requestCheckVar(Request.Form("sGR1"),10)
	iGiftRange2		= requestCheckVar(Request.Form("sGR2"),10)
	iGiftKindCnt	= requestCheckVar(Request.Form("iGKC"),10)
	iGiftKindType	= requestCheckVar(Request.Form("chkKT"),10)
	iGiftLimit		= requestCheckVar(Request.Form("iL"),10)
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)
	sGiftDelivery	= requestCheckVar(Request.Form("selD"),1)
	iGiftStatus		= requestCheckVar(Request.Form("giftstatus"),10)
	sOpenDate		= requestCheckVar(Request.Form("sOD"),30)
	sCloseDate		= requestCheckVar(Request.Form("sCD"),30)
	iSiteScope		= requestCheckVar(Request.Form("eventscope"),4)

	IF CStr(iSiteScope) = "3" THEN 
		sPartnerID 	= requestCheckVar(Request.Form("selP"),32)
		If Len(dSDayTime) <> "8" Then 	dSDayTime = "00:00:00"
		If Len(dEDayTime) <> "8" Then 	dSDayTime = "23:59:00"
		If dSDayTime <> "" Then
			dSDay = dSDay & " " & dSDayTime
		End If

		If dEDayTime <> "" Then
			dEDay = dEDay & " " & dEDayTime
		End If
	End If

	IF iGiftStatus = "7" THEN
		if sOpenDate = "" then
			 sOpenDate = "getdate()"
		else
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		end if
	ELSEIF 	iGiftStatus = "9" THEN
		if sCloseDate = "" then
			 sCloseDate = "getdate()"
		else
			sCloseDate = " convert(nvarchar(10),'"&sCloseDate&"',21)"&"+' "&formatdatetime(sCloseDate,4)&"'"
		end if
	ELSE
		IF sOpenDate = "" THEN
			sOpenDate = "null"
		ELSE
			sOpenDate = " convert(nvarchar(10),'"&sOpenDate&"',21)"&"+' "&formatdatetime(sOpenDate,4)&"'"
		END IF

		IF sCloseDate = "" THEN
			sCloseDate = "null"
		ELSE
			sCloseDate = " convert(nvarchar(10),'"&sCloseDate&"',21)"&"+' "&formatdatetime(sCloseDate,4)&"'"
		END IF
	END IF

	IF iGiftKindType = "" THEN iGiftKindType = 1
	IF iGiftLimit ="" THEN iGiftLimit = 0
	IF iGiftType = "" THEN iGiftType =0
	IF iGiftRange1 = "" THEN iGiftRange1 = 0
	IF iGiftRange2 = "" THEN iGiftRange2 = 0
	IF iGroupCode = "" THEN iGroupCode = 0
	' ���������� �����϶� 1+1 , 1:1 �ʱ�ȭ
	if iGiftType<>"" and not(isnull(iGiftType)) then
		if iGiftType=2 then
			iGiftKindType=0
		end if
	end if

	'//���������� ���� ��� Ȯ��
	CALL fnChkDelivery(iGiftScope,sGiftDelivery,eCode, sBrand,iGroupCode, 0)

	On Error Resume Next

	'//������ ���
	strSql = "INSERT INTO [db_event].[dbo].[tbl_gift] ( [gift_name], [gift_scope], [evt_code], [evtgroup_code], [makerid], [gift_type], [gift_range1], [gift_range2]"&_
			", [giftkind_code], [giftkind_type], [giftkind_cnt], [giftkind_limit], [gift_startdate], [gift_enddate],[gift_status],[gift_delivery],[adminid],opendate,lastupdate"&_
			", site_scope, partner_id)"&_
			" VALUES ('"&sTitle&"','"&iGiftScope&"','"&eCode&"','"&iGroupCode&"','"&sBrand&"','"&iGiftType&"','"&iGiftRange1&"','"&iGiftRange2&"' "&_
			",'"&iGiftKindCode&"','"&iGiftKindType&"','"&iGiftKindCnt&"','"&iGiftLimit&"','"&dSDay&"','"&dEDay&"','"&iGiftStatus&"','"&sGiftDelivery&"','"&session("ssBctId")&"',"&sOpenDate&",getdate()"&_
			", '"&iSiteScope&"','"&sPartnerID&"') " + VbCRLF
	dbget.execute strSql

	strSql = "select SCOPE_IDENTITY()"
	rsget.Open strSql, dbget, 0
	gCode = rsget(0)
	rsget.Close

	strSql = " update [db_event].[dbo].[tbl_giftkind] "
	strSql = strSql + " set org_gift_code = " + CStr(gCode) + " "
	strSql = strSql + " where giftkind_code = " + CStr(iGiftKindCode) + " and org_gift_code is NULL "
	dbget.execute strSql

	IF Err.Number <> 0 THEN
		response.Write strSql
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")
		dbget.close()	:	response.End
	END IF

	if gift_text1<>"" or gift_img1<>"" then
		strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_md_theme where evt_code=" & eCode & ")"&vbCrlf 
		strSql = strSql & "begin"&vbCrlf 
		strSql = strSql & " UPDATE db_event.dbo.tbl_event_md_theme"&vbCrlf
		strSql = strSql & " SET gift_text1 = '"& html2db(gift_text1) &"'"&vbCrlf
		strSql = strSql & " , gift_img1 = '"& gift_img1 &"'"&vbCrlf
		strSql = strSql & " , gift_text2 = '"& html2db(gift_text2) &"'"&vbCrlf
		strSql = strSql & " , gift_img2 = '"& gift_img2 &"'"&vbCrlf
		strSql = strSql & " , gift_text3 = '"& html2db(gift_text3) &"'"&vbCrlf
		strSql = strSql & " , gift_img3 = '"& gift_img3 &"'"&vbCrlf
		strSql = strSql & " , gift_isusing = '" & gift_isusing & "'"&vbCrlf
		strSql = strSql & " , contentsAlign = '" & gift_infotext & "'"&vbCrlf
		strSql = strSql & " WHERE  evt_code = "& eCode &vbCrlf
		strSql = strSql & "end"&vbCrlf 
		strSql = strSql & " ELSE "&vbCrlf
		strSql = strSql & "begin"&vbCrlf 
		strSql = strSql & " INSERT INTO db_event.dbo.tbl_event_md_theme (evt_code, gift_isusing, gift_img1, gift_text1, gift_img2, gift_text2, gift_img3, gift_text3, contentsAlign)"&vbCrlf 
		strSql = strSql & " VALUES("&eCode&",'" & gift_isusing & "', '"& gift_img1 &"' ,'"& gift_text1 & "','" & gift_img2 &"' ,'"& gift_text2 & "','" & gift_img3 &"' ,'"& gift_text3 &"' ,'"& gift_infotext &"')"&vbCrlf 
		strSql = strSql & "end"
		dbget.execute strSql
	END IF

	'#################################### ����ǰ �α� ���� #########################################################################
	vChangeContents = vChangeContents & "����ǰ �α� " & vbCrLf
	vChangeContents = vChangeContents & "- ���� : gift_name = " & sTitle & vbCrLf
	vChangeContents = vChangeContents & "- ���� �̺�Ʈ�ڵ� : evt_code = " & eCode & vbCrLf
	vChangeContents = vChangeContents & "- ������� / ���� : gift_scope = " & iGiftScope & ", gift_type = " & iGiftType & vbCrLf
	vChangeContents = vChangeContents & "- �Ⱓ : gift_startdate = " & dSDay & " ~ gift_enddate = " & dEDay & vbCrLf
	vChangeContents = vChangeContents & "- �귣�� : makerid = " & sBrand & vbCrLf
	vChangeContents = vChangeContents & "- �������� : gift_range1 = " & iGiftRange1 & ", gift_range2 = " & iGiftRange2 & vbCrLf
	vChangeContents = vChangeContents & "- ����ǰ���� : giftkind_code = " & iGiftKindCode & ", giftkind_type = " & iGiftKindType & vbCrLf
	vChangeContents = vChangeContents & "- ����ǰ���� / ���� : giftkind_cnt = " & iGiftKindCnt & ", giftkind_limit = " & iGiftLimit & vbCrLf
	vChangeContents = vChangeContents & "- ��۹�� : gift_delivery = " & sGiftDelivery & vbCrLf
	vChangeContents = vChangeContents & "- ���� : gift_status = " & iGiftStatus & vbCrLf
	vChangeContents = vChangeContents & "- ���� / ������ : opendate = " & sOpenDate & ", closedate = " & sCloseDate & vbCrLf
	vChangeContents = vChangeContents & "- ���� : site_scope = " & iSiteScope & vbCrLf
	'### ���� �α� ����(event)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'gift', '" & gCode & "', '" & menupos & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

	on error Goto 0

	IF eCode = 0 THEN eCode = ""
	response.redirect("giftList.asp?menupos="&menupos&"&eC="&eCode)
dbget.close()	:	response.End
Case "U"	'//����ǰ����
	Dim strAdd : strAdd = ""

	'�˻��� üũ--------------------------------------------------------------
	 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	 sSearchTxt     = requestCheckVar(Request("sTxt"),30)		'�˻���
	 sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
	 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	 sEdate     	= requestCheckVar(Request("iED"),10)		'������
	 sgStatus	    = requestCheckVar(Request("gstatus"),4)	'����ǰ ����

	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
 	strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&sgStatus
 	'--------------------------------------------------------------

	gCode			= requestCheckVar(Request.Form("gC"),10)
	eCode			= requestCheckVar(Request.Form("eC"),10)
	IF eCode ="" THEN eCode = 0
	sTitle			= html2db(requestCheckVar(Request.Form("sGN"),64))
	dSDay 			= requestCheckVar(Request.Form("sSD"),10)
	dEDay			= requestCheckVar(Request.Form("sED"),10)
	iGiftScope		= requestCheckVar(Request.Form("giftscope"),4)
	sBrand			= requestCheckVar(Request.Form("ebrand"),32)
	iGroupCode		= requestCheckVar(Request.Form("selG"),10)
	iGiftType		= requestCheckVar(Request.Form("gifttype"),10)
	iGiftRange1		= requestCheckVar(Request.Form("sGR1"),10)
	iGiftRange2		= requestCheckVar(Request.Form("sGR2"),10)
	iGiftKindCnt	= requestCheckVar(Request.Form("iGKC"),10)
	iGiftKindType	= requestCheckVar(Request.Form("chkKT"),10)
	iGiftLimit		= requestCheckVar(Request.Form("iL"),10)
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)
	iGiftStatus		= requestCheckVar(Request.Form("giftstatus"),10)
	sGiftUsing		= requestCheckVar(Request.Form("sGU"),1)
	sOpenDate		= requestCheckVar(Request.Form("sOD"),30)
	sCloseDate		= requestCheckVar(Request.Form("sCD"),30)
	sGiftDelivery	= requestCheckVar(Request.Form("selD"),1)
	iSiteScope		= requestCheckVar(Request.Form("eventscope"),4)

	givecnt		= requestCheckVar(Request.Form("givecnt"),5) '����ǰ ���� ���� ����
	dSDayTime		= request("sSDTime")
	dEDayTime		= request("sEDTime")

	IF CStr(iSiteScope) = "3" THEN	sPartnerID 		= requestCheckVar(Request.Form("selP"),32)

	If CStr(iSiteScope) = "3" THEN
		If Len(dSDayTime) <> "8" Then 	dSDayTime = "00:00:00"
		If Len(dEDayTime) <> "8" Then 	dSDayTime = "23:59:00"
		If dSDayTime <> "" Then
			dSDay = dSDay & " " & dSDayTime
		End If

		If dEDayTime <> "" Then
			dEDay = dEDay & " " & dEDayTime
		End If
	End If

	'DB�� �ܰ� Ȯ��(�ߺ�ó��üũ)
	IF iGiftStatus="0" THEN
		strSql = "SELECT COUNT(*) FROM db_event.dbo.tbl_gift WHERE gift_code=" & gCode & " AND gift_status in (6,7,9)"
		rsget.Open strSql, dbget
		IF rsget(0)>0 THEN
			rsget.Close
			Alert_return("����ǰ�� ���µǾ��ֽ��ϴ�.\n�ٽ� Ȯ�����ּ���.")
			dbget.close():	response.End
		End IF
		rsget.Close
	END IF

	IF iGiftStatus ="7" AND sOpenDate="" THEN
		strAdd = " , [opendate] = getdate()"
	ELSEIF (iGiftStatus = "9" and sCloseDate ="" ) THEN
		strAdd = ", [closedate] = getdate() "	'����ó���� ����
	END IF

	'������ ������ ����� ������ ���� ��¥�� ����
	IF iGiftStatus = 9 and  datediff("d",dEDay,date()) <0 THEN
			dEDay = date()
	END IF

	IF iGiftKindType = "" THEN iGiftKindType = 1
	IF iGiftLimit ="" THEN iGiftLimit = 0
 	IF iGiftType = "" THEN iGiftType =0
 	IF iGiftRange1 = "" THEN iGiftRange1 = 0
	IF iGiftRange2 = "" THEN iGiftRange2 = 0
	IF iGroupCode = "" THEN iGroupCode = 0
	' ���������� �����϶� 1+1 , 1:1 �ʱ�ȭ
	if iGiftType<>"" and not(isnull(iGiftType)) then
		if iGiftType=2 then
			iGiftKindType=0
		end if
	end if

 	'//���������� ���� ��� Ȯ��
 	CALL fnChkDelivery(iGiftScope,sGiftDelivery,eCode, sBrand,iGroupCode, gCode)

 	'//������ ����
	strSql = " UPDATE [db_event].[dbo].[tbl_gift] SET  [gift_name] = '"&sTitle&"', [gift_scope]="&iGiftScope&", [evtgroup_code] ="&iGroupCode&_
			" , [makerid]='"&sBrand&"', [gift_type]="&iGiftType&", [gift_range1]="&iGiftRange1&", [gift_range2]= "&iGiftRange2&_
			", [giftkind_code]= "&iGiftKindCode&", [giftkind_type] ="&iGiftKindType&" , [giftkind_cnt]= "&iGiftKindCnt&", [giftkind_limit]="&iGiftLimit&_
			", [gift_startdate]= '"&dSDay&"', [gift_enddate]='"&dEDay&"', [gift_status] = "&iGiftStatus&", [gift_using] = '"&sGiftUsing&"'"&_
			" , gift_delivery = '"&sGiftDelivery&"'"&_
			",[adminid]= '"&session("ssBctId")&"', [lastupdate] = getdate(), site_scope="&iSiteScope&", partner_id ='"&sPartnerID&"' , giftkind_givecnt = '"& givecnt &"' "&strAdd&_
			" WHERE gift_code = "&gCode

	dbget.execute strSql

	IF Err.Number <> 0 THEN
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���1")
       dbget.close()	:	response.End
	END IF

	if gift_text1<>"" or gift_img1<>"" then
		strSql = "IF EXISTS(SELECT evt_code FROM db_event.dbo.tbl_event_md_theme where evt_code=" & eCode & ")"&vbCrlf 
		strSql = strSql & "begin"&vbCrlf 
		strSql = strSql & " UPDATE db_event.dbo.tbl_event_md_theme"&vbCrlf
		strSql = strSql & " SET gift_text1 = '"& html2db(gift_text1) &"'"&vbCrlf
		strSql = strSql & " , gift_img1 = '"& gift_img1 &"'"&vbCrlf
		strSql = strSql & " , gift_text2 = '"& html2db(gift_text2) &"'"&vbCrlf
		strSql = strSql & " , gift_img2 = '"& gift_img2 &"'"&vbCrlf
		strSql = strSql & " , gift_text3 = '"& html2db(gift_text3) &"'"&vbCrlf
		strSql = strSql & " , gift_img3 = '"& gift_img3 &"'"&vbCrlf
		strSql = strSql & " , gift_isusing = '" & gift_isusing & "'"&vbCrlf
		strSql = strSql & " , contentsAlign = '" & gift_infotext & "'"&vbCrlf
		strSql = strSql & " WHERE  evt_code = "& eCode &vbCrlf
		strSql = strSql & "end"&vbCrlf 
		strSql = strSql & " ELSE "&vbCrlf
		strSql = strSql & "begin"&vbCrlf 
		strSql = strSql & " INSERT INTO db_event.dbo.tbl_event_md_theme (evt_code, gift_isusing, gift_img1, gift_text1, gift_img2, gift_text2, gift_img3, gift_text3, contentsAlign)"&vbCrlf 
		strSql = strSql & " VALUES("&eCode&",'" & gift_isusing & "', '"& gift_img1 &"' ,'"& gift_text1 & "','" & gift_img2 &"' ,'"& gift_text2 & "','" & gift_img3 &"' ,'"& gift_text3 &"' ,'"& gift_infotext &"')"&vbCrlf 
		strSql = strSql & "end"
		dbget.execute strSql
	END IF
	'#################################### ����ǰ �α� ���� #########################################################################
	vChangeContents = vChangeContents & "����ǰ �α� " & vbCrLf
	vChangeContents = vChangeContents & "- ���� : gift_name = " & sTitle & vbCrLf
	vChangeContents = vChangeContents & "- ���� �̺�Ʈ�ڵ� : evt_code = " & eCode & vbCrLf
	vChangeContents = vChangeContents & "- ������� / ���� : gift_scope = " & iGiftScope & ", gift_type = " & iGiftType & vbCrLf
	vChangeContents = vChangeContents & "- �Ⱓ : gift_startdate = " & dSDay & " ~ gift_enddate = " & dEDay & vbCrLf
	vChangeContents = vChangeContents & "- �귣�� : makerid = " & sBrand & vbCrLf
	vChangeContents = vChangeContents & "- �������� : gift_range1 = " & iGiftRange1 & ", gift_range2 = " & iGiftRange2 & vbCrLf
	vChangeContents = vChangeContents & "- ����ǰ���� : giftkind_code = " & iGiftKindCode & ", giftkind_type = " & iGiftKindType & vbCrLf
	vChangeContents = vChangeContents & "- ����ǰ���� / ���� : giftkind_cnt = " & iGiftKindCnt & ", giftkind_limit = " & iGiftLimit & vbCrLf
	vChangeContents = vChangeContents & "- ��۹�� : gift_delivery = " & sGiftDelivery & vbCrLf
	vChangeContents = vChangeContents & "- ���� : gift_status = " & iGiftStatus & vbCrLf
	vChangeContents = vChangeContents & "- ���� / ������ : opendate = " & sOpenDate & ", closedate = " & sCloseDate & vbCrLf
	vChangeContents = vChangeContents & "- ���� : site_scope = " & iSiteScope & vbCrLf
	'### ���� �α� ����(event)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'gift', '" & gCode & "', '" & menupos & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)

	IF eCode = 0 THEN eCode = ""
	response.redirect("giftList.asp?menupos="&menupos&"&"&strParm)
dbget.close()	:	response.End
Case "KI"  '//����ǰ ���� ���
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))
	itemid			= requestCheckVar(Request.Form("itemid"),10)
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100)
	prd_itemgubun	= requestCheckVar(Request.Form("prd_itemgubun"),32)
	prd_itemid		= requestCheckVar(Request.Form("prd_itemid"),32)
	prd_itemoption	= requestCheckVar(Request.Form("prd_itemoption"),32)

	if (prd_itemid <> "") then
		'// �����ڵ� üũ
		strSql = " select top 1 shopitemid from db_shop.dbo.tbl_shop_item "
		strSql = strSql & " where itemgubun = '" + CStr(prd_itemgubun) + "' and shopitemid = " + CStr(prd_itemid) + " and itemoption = '" + CStr(prd_itemoption) + "' "

		rsget.Open strSql, dbget
		IF rsget.EOF OR rsget.BOF THEN
			rsget.Close
			Alert_return("�߸��� �����ڵ��Դϴ�.")
			dbget.close()	:	response.End
		End IF
		rsget.Close
	end if

	IF itemid = "" THEN itemid =0
	IF bcouponidx = "" THEN bcouponidx=0

	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_item].[dbo].[tbl_item] where itemid = "&itemid
	rsget.Open strSql, dbget
	IF rsget.EOF OR rsget.BOF THEN
		rsget.Close
		Alert_return("�������� �ʴ� ��ǰ��ȣ�Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
       dbget.close()	:	response.End
	End IF
	rsget.Close
	END IF

	' ��۹�� üũ
	IF (giftkind_linkGbn="I") Then
		if gift_delivery="N" then
			if prd_itemgubun="" or isnull(prd_itemgubun) or prd_itemid="" or isnull(prd_itemid) or prd_itemoption="" or isnull(prd_itemoption) then
				Alert_return("����ǰ ������ ��ǰ�� ���� �ϼ̽��ϴ�. �����ڵ带 �Է��� �ּ���.")
				dbget.close() : response.End
			end if
		end if

	elseIF (giftkind_linkGbn="B") Then
	    if bcouponidx=0 then
	        Alert_return("�������� �ʴ� ���ʽ� ���� ��ȣ�Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
	        response.End
	    end if

	    strSql = "select idx from db_user.dbo.tbl_user_coupon_master C "
        strSql = strSql & " where idx="&bcouponidx&VbCRLF
        strSql = strSql & " and isopenlistcoupon='Y'"

    	rsget.Open strSql, dbget
    	IF rsget.EOF OR rsget.BOF THEN
    		rsget.Close
    		Alert_return("�������� �ʴ� ���ʽ� ���� ��ȣ �Ǵ� ���ð�(������) ����Ÿ���� �ƴմϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
            dbget.close()	:	response.End
    	End IF
    	rsget.Close
	End IF

''���� 2013/09/25 SELECT SCOPE_IDENTITY ������, IDENT_CURRENT Ʈ����� �������
    strSql = "select * from [db_event].[dbo].[tbl_giftkind] where 1=0"
	rsget.Open strSql,dbget,1,3
	rsget.AddNew
		rsget("giftkind_name")      = sGiftKindName
		rsget("giftkind_img")       = sGiftKindImg
		rsget("itemid")             = CHKIIF(giftkind_linkGbn="B",0,itemid)
		rsget("giftkind_linkGbn")   = giftkind_linkGbn
		rsget("bcouponidx")         = CHKIIF(giftkind_linkGbn="B",bcouponidx,0)
		rsget("reguserid")          = CStr(session("ssBctId"))
		if (prd_itemid <> "") then
    		rsget("prd_itemgubun")      = prd_itemgubun
    		rsget("prd_itemid")         = prd_itemid
    		rsget("prd_itemoption")     = prd_itemoption
    	end if
	rsget.update
		iGiftKindCode = rsget("giftkind_code")
	rsget.close

'response.redirect("popgiftkindReg.asp?sGKN="&sGiftKindName)
%>
	<script language="javascript">
	<!--
		var strImg = "<%=sGiftKindImg%>";
		opener.document.all.iGK.value = "<%=iGiftKindCode%>";
		opener.document.all.sGKN.value= "<%=sGiftKindName%>";

		var gKLGbn = "<%= giftkind_linkGbn %>";
		if (opener.document.all.giftkind_linkGbn){
		    opener.document.all.giftkind_linkGbn.value= gKLGbn;
		}
		if (gKLGbn=='B'){
		    if (opener.document.all.bcouponidx){
		        opener.document.all.bcouponidx.value= "<%= bcouponidx %>";
		    }
		}

		if(strImg !=""){
		opener.document.all.spanImg.innerHTML = "<a href=javascript:jsImgView('"+strImg+"')><img src='"+strImg+"' border=0></a>";
		}
		window.close();
	//-->
	</script>
<%
dbget.close()	:	response.End
Case "KU"  '//����ǰ ���� ����
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))
	itemid			= requestCheckVar(Request.Form("itemid"),10)
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100)
	prd_itemgubun	= trim(requestCheckVar(getNumeric(Request.Form("prd_itemgubun")),32))
	prd_itemid		= trim(requestCheckVar(getNumeric(Request.Form("prd_itemid")),32))
	prd_itemoption	= trim(requestCheckVar(Request.Form("prd_itemoption"),32))

	IF itemid = "" THEN itemid =0
	IF bcouponidx = "" THEN bcouponidx=0

	if (prd_itemid <> "") then
		'// �����ڵ� üũ
		strSql = " select top 1 shopitemid from db_shop.dbo.tbl_shop_item "
		strSql = strSql & " where itemgubun = '" + CStr(prd_itemgubun) + "' and shopitemid = " + CStr(prd_itemid) + " and itemoption = '" + CStr(prd_itemoption) + "' "

		rsget.Open strSql, dbget
		IF rsget.EOF OR rsget.BOF THEN
			rsget.Close
			Alert_return("�߸��� �����ڵ��Դϴ�.")
			dbget.close()	:	response.End
		End IF
		rsget.Close
	end if

	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_item].[dbo].[tbl_item] where itemid = "&itemid
	rsget.Open strSql, dbget
	IF rsget.EOF OR rsget.BOF THEN
		rsget.Close
		Alert_return("�������� �ʴ� ��ǰ��ȣ�Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
       dbget.close()	:	response.End
	End IF
	rsget.Close
	END IF

	' ��۹�� üũ
	IF (giftkind_linkGbn="I") Then
		if gift_delivery="N" then
			if prd_itemgubun="" or isnull(prd_itemgubun) or prd_itemid="" or isnull(prd_itemid) or prd_itemoption="" or isnull(prd_itemoption) then
				Alert_return("����ǰ ������ ��ǰ�� ���� �ϼ̽��ϴ�. �����ڵ带 �Է��� �ּ���.")
				dbget.close() : response.End
			end if
		end if

	elseIF (giftkind_linkGbn="B") Then
	    if bcouponidx=0 then
	        Alert_return("�������� �ʴ� ���ʽ� ���� ��ȣ�Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
	        response.End
	    end if

	    strSql = "select idx from db_user.dbo.tbl_user_coupon_master C "
        strSql = strSql & " where idx="&bcouponidx&VbCRLF
        strSql = strSql & " and isopenlistcoupon='Y'"

    	rsget.Open strSql, dbget
    	IF rsget.EOF OR rsget.BOF THEN
    		rsget.Close
    		Alert_return("�������� �ʴ� ���ʽ� ���� ��ȣ �Ǵ� ���ð�(������) ����Ÿ���� �ƴմϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
            dbget.close()	:	response.End
    	End IF
    	rsget.Close
	End IF

	strSql = " UPDATE [db_event].[dbo].[tbl_giftkind]" & VbCRLF
	strSql = strSql & " set [giftkind_name] ='"&sGiftKindName&"'" & VbCRLF
	strSql = strSql & " , [giftkind_img] ='"&sGiftKindImg&"'" & VbCRLF
	strSql = strSql & " , [itemid] ="&CHKIIF(giftkind_linkGbn="B",0,itemid) & VbCRLF
	strSql = strSql & " , [giftkind_linkGbn]='"&giftkind_linkGbn&"'" & VbCRLF
	strSql = strSql & " , [bcouponidx]="&CHKIIF(giftkind_linkGbn="B",bcouponidx,0) & VbCRLF
	if (prd_itemid <> "") then
		strSql = strSql & " , [prd_itemgubun]='"&prd_itemgubun&"'" & VbCRLF
		strSql = strSql & " , [prd_itemid]='"&prd_itemid&"'" & VbCRLF
		strSql = strSql & " , [prd_itemoption]='"&prd_itemoption&"'" & VbCRLF
	end if
	strSql = strSql & " WHERE giftkind_code = "&iGiftKindCode & VbCRLF

	dbget.execute strSql

	IF Err.Number <> 0 THEN
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")
       dbget.close()	:	response.End
	END IF

response.redirect("popgiftkindReg.asp?sGKN="&sGiftKindName)
dbget.close()	:	response.End

Case "KM"  '//����ǰ ���� ����	2010 �߰�
	iGiftKindCode	= requestCheckVar(Request.Form("iGK"),10)
	sGiftKindName 	= html2db(requestCheckVar(Request.Form("sGKN"),60))
	itemid			= requestCheckVar(Request.Form("itemid"),10)
	sGiftKindImg	= requestCheckVar(Request.Form("sGKImg"),100)
	s120Img	        = requestCheckVar(Request.Form("S120"),100)
	s401Img         = requestCheckVar(Request.Form("S401"),100)
	s402Img         = requestCheckVar(Request.Form("S402"),100)
	s403Img         = requestCheckVar(Request.Form("S403"),100)
	s404Img         = requestCheckVar(Request.Form("S404"),100)
	s405Img         = requestCheckVar(Request.Form("S405"),100)

	IF itemid = "" THEN itemid =0

	IF itemid > 0 THEN
	strSql = "SELECT itemid FROM [db_item].[dbo].[tbl_item] where itemid = "&itemid
	rsget.Open strSql, dbget
	IF rsget.EOF OR rsget.BOF THEN
		rsget.Close
		Alert_return("�������� �ʴ� ��ǰ��ȣ�Դϴ�. Ȯ�� �� �ٽ� �Է����ּ���")
       dbget.close()	:	response.End
	End IF
	rsget.Close
	END IF
	strSql = " UPDATE [db_event].[dbo].[tbl_giftkind] " & VbCRLF
	strSql = strSql & " set [giftkind_name] ='"&sGiftKindName&"'" & VbCRLF
	strSql = strSql & " , [giftkind_img] ='"&sGiftKindImg&"'" & VbCRLF
	strSql = strSql & " , [itemid] ="&itemid & VbCRLF
	strSql = strSql & " , image120 ='"&s120Img &"'"& VbCRLF
	strSql = strSql & " WHERE giftkind_code = "&iGiftKindCode
	dbget.execute strSql

	strSql = " Delete from db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	strSql = strSql & " WHERE gift_kind_code = "&iGiftKindCode
	dbget.execute strSql

	if (s401Img<>"") then
	    strSql = " Insert Into  db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode
	    strSql = strSql & " ,1"
	    strSql = strSql & " ,'"& s401Img& "')"
	    dbget.execute strSql
	end if

	if (s402Img<>"") then
	    strSql = " Insert Into  db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode
	    strSql = strSql & " ,2"
	    strSql = strSql & " ,'"& s402Img& "')"
	    dbget.execute strSql
	end if

	if (s403Img<>"") then
	    strSql = " Insert Into  db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode
	    strSql = strSql & " ,3"
	    strSql = strSql & " ,'"& s403Img& "')"
	    dbget.execute strSql
	end if

	if (s404Img<>"") then
	    strSql = " Insert Into  db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode
	    strSql = strSql & " ,4"
	    strSql = strSql & " ,'"& s404Img& "')"
	    dbget.execute strSql
	end if

	if (s405Img<>"") then
	    strSql = " Insert Into  db_event.dbo.tbl_giftkind_AddImage " & VbCRLF
	    strSql = strSql & " (gift_kind_code, addnum, gift_kind_addimage) "
	    strSql = strSql & " values(" & iGiftKindCode
	    strSql = strSql & " ,5"
	    strSql = strSql & " ,'"& s405Img& "')"
	    dbget.execute strSql
	end if

	''�ɼ�
	Dim optCnt , gift_kind_option, gift_kind_optionName, gift_kind_Limit, gift_kind_LimitSold, gift_kind_LimitYN
	gift_kind_option = Split(request("gift_kind_option"),",")
	gift_kind_optionName = Split(request("gift_kind_optionName"),",")
	gift_kind_Limit = Split(request("gift_kind_Limit"),",")
	gift_kind_LimitSold = Split(request("gift_kind_LimitSold"),",")
	gift_kind_LimitYN = Split(request("gift_kind_LimitYN"),",")

	prd_itemgubun = Split(request("prd_itemgubun"),",")
	prd_itemid = Split(request("prd_itemid"),",")
	prd_itemoption = Split(request("prd_itemoption"),",")

	'rw request("gift_kind_option")
	'rw request("gift_kind_optionName")

	if IsArray(gift_kind_option) then
	    for i=LBound(gift_kind_option) to UBound(gift_kind_option)
	        if (Trim(gift_kind_option(i))<>"") then
	            strSql = "IF Exists(select * from db_event.dbo.tbl_giftkind_Option where gift_kind_code="& iGiftKindCode &" and  gift_kind_option='"&Trim(gift_kind_option(i))&"' )"
	            strSql = strSql & " BEGIN"
	            strSql = strSql & " update db_event.dbo.tbl_giftkind_Option " & VbCRLF
	            strSql = strSql & " set gift_kind_optionName='" & Trim(gift_kind_optionName(i)) & "'"  & VbCRLF
	            strSql = strSql & " ,gift_kind_Limit=" & Trim(gift_kind_Limit(i)) & ""  & VbCRLF
	            strSql = strSql & " ,gift_kind_LimitSold=" & Trim(gift_kind_LimitSold(i)) & ""  & VbCRLF
	            strSql = strSql & " ,gift_kind_optionUsing='" & Trim(request("gift_kind_optionUsing_"&Trim(gift_kind_option(i)))) & "'"  & VbCRLF
	            strSql = strSql & " ,gift_kind_LimitYN='" & Trim(gift_kind_LimitYN(i)) & "'"  & VbCRLF
				if (Trim(prd_itemid(i)) <> "") then
					strSql = strSql & " ,prd_itemgubun='" & Trim(prd_itemgubun(i)) & "'"  & VbCRLF
					strSql = strSql & " ,prd_itemid='" & Trim(prd_itemid(i)) & "'"  & VbCRLF
					strSql = strSql & " ,prd_itemoption='" & Trim(prd_itemoption(i)) & "'"  & VbCRLF
				end if
	            strSql = strSql & " where gift_kind_code="& iGiftKindCode & VbCRLF
	            strSql = strSql & " and gift_kind_option='"&Trim(gift_kind_option(i))&"'" & VbCRLF
	            strSql = strSql & " END"
	            strSql = strSql & " ELSE"
	            strSql = strSql & " BEGIN"
	            strSql = strSql & " Insert Into  db_event.dbo.tbl_giftkind_Option " & VbCRLF
	            strSql = strSql & " (gift_kind_code, gift_kind_option, gift_kind_optionName, gift_kind_Limit, gift_kind_LimitSold, gift_kind_optionUsing, gift_kind_LimitYN, prd_itemgubun, prd_itemid, prd_itemoption)"
	            strSql = strSql & " values("
	            strSql = strSql & " "& iGiftKindCode & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_option(i))&"'" & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_optionName(i))&"'" & VbCRLF
	            strSql = strSql & " ,"&Trim(gift_kind_Limit(i))&"" & VbCRLF
	            strSql = strSql & " ,"&Trim(gift_kind_LimitSold(i))&"" & VbCRLF
	            strSql = strSql & " ,'"&Trim(request("gift_kind_optionUsing_"&Trim(gift_kind_option(i))))&"'" & VbCRLF
	            strSql = strSql & " ,'"&Trim(gift_kind_LimitYN(i))&"'" & VbCRLF
				if (Trim(prd_itemid(i)) <> "") then
					strSql = strSql & " ,'"&Trim(prd_itemgubun(i))&"'" & VbCRLF
					strSql = strSql & " ,'"&Trim(prd_itemid(i))&"'" & VbCRLF
					strSql = strSql & " ,'"&Trim(prd_itemoption(i))&"'" & VbCRLF
				else
					strSql = strSql & " , NULL " & VbCRLF
					strSql = strSql & " , NULL " & VbCRLF
					strSql = strSql & " , NULL " & VbCRLF
				end if
	            strSql = strSql & " )"
	            strSql = strSql & " END"
				''response.write strSql
	            dbget.execute strSql
	        end if
	    next
    end if

	IF Err.Number <> 0 THEN
		Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���")
       dbget.close()	:	response.End
	else
		response.write "<script>alert('����Ǿ����ϴ�.');</script>"
		response.write "<script>history.back();</script>"
		response.write "<script>location.reload();</script>"
	END IF
''response.end
''response.redirect("popgiftkindManage.asp?iGK="&iGiftKindCode)
dbget.close()	:	response.End

Case "regautogiftitem"	'//����ǰ����
	if makerid="" or isnull(makerid) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�귣�尡 �����ϴ�.');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	makerid = trim(makerid)
	if giftkind_name="" or isnull(giftkind_name) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('����ǰ�������� �����ϴ�.');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	giftkind_name = trim(giftkind_name)

	itemgubun="85"
	itemoption="0000"
	sqlStr = " select top 1 shopitemid" & vbcrlf
	sqlStr = sqlStr & " from [db_shop].[dbo].tbl_shop_item with (readuncommitted)" & vbcrlf
	sqlStr = sqlStr & " where itemgubun='"& itemgubun &"'" & vbcrlf
	sqlStr = sqlStr & " order by shopitemid desc" & vbcrlf

	'response.write sqlStr & "<Br>"
	rsget.Open sqlStr,dbget,1
		if not rsget.Eof then
			itemid = rsget("shopitemid")+1
		else
			itemid = 1
		end if
	rsget.close

	sqlStr = "insert into [db_shop].[dbo].tbl_shop_item (" & vbCrlf
	sqlStr = sqlStr & " itemgubun,shopitemid,itemoption,makerid,shopitemname,shopitemoptionname,orgsellprice,shopitemprice" & vbCrlf
	sqlStr = sqlStr & " , discountsellprice,shopsuplycash,shopbuyprice,centermwdiv,vatinclude)" & vbCrlf
	sqlStr = sqlStr & " 	select" & vbCrlf
	sqlStr = sqlStr & "		N'"& itemgubun &"',"& itemid &",N'"& itemoption &"',N'"& makerid &"',N'"& giftkind_name &"',NULL,0,0" & vbCrlf
	'sqlStr = sqlStr & " 	,0,0,0,c.maeipdiv,c.vatinclude" & vbCrlf
	sqlStr = sqlStr & " 	,0,0,0,'W' as maeipdiv,c.vatinclude" & vbCrlf	' �űԵ�Ͻÿ��� ������ ��Ź���� ����.	2023.06.23 �̹����̻�� ��û
	sqlStr = sqlStr & "		from db_user.dbo.tbl_user_c c with (readuncommitted)" & vbCrlf
	sqlStr = sqlStr & "		where c.userid='"& makerid &"'" & vbCrlf

	'response.write sqlStr & "<Br>"
	dbget.execute sqlStr

	sqlStr = " exec db_shop.dbo.sp_ten_shop_tnbarcode_update N'"& itemgubun &"',"& itemid &",N'"& itemoption &"'"

	'response.write sqlStr & "<Br>"
	dbget.Execute sqlStr

	barcode = trim(itemgubun)&trim(Format00(6,itemid))&itemoption

	sqlStr = "update [db_shop].[dbo].tbl_shop_item" & vbCrlf
	sqlStr = sqlStr & " set extbarcode=N'"& barcode &"'" & vbCrlf
	sqlStr = sqlStr & " where itemgubun='"& itemgubun &"'" & vbCrlf
	sqlStr = sqlStr & " and shopitemid="& itemid &"" & vbCrlf
	sqlStr = sqlStr & " and itemoption='"& itemoption &"'" & vbCrlf

	'response.write sqlStr
	dbget.execute sqlStr

	sqlStr = " select top 1 * from [db_item].[dbo].tbl_item_option_stock" & vbCrlf
	sqlStr = sqlStr & " where itemgubun=N'"& itemgubun &"'" & vbCrlf
	sqlStr = sqlStr & " and itemid="& itemid &"" & vbCrlf
	sqlStr = sqlStr & " and itemoption=N'"& itemoption &"'" & vbCrlf

	'response.write sqlStr
	rsget.Open sqlStr,dbget,1
		stockitemexists = (not rsget.Eof)
	rsget.close

	if (stockitemexists) then
		sqlStr = " update [db_item].[dbo].tbl_item_option_stock" & vbCrlf
		sqlStr = sqlStr & " set barcode=N'"& barcode &"'" & vbCrlf
		sqlStr = sqlStr & " where itemgubun='"& itemgubun &"'" & vbCrlf
		sqlStr = sqlStr & " and itemid="& itemid &"" & vbCrlf
		sqlStr = sqlStr & " and itemoption='"& itemoption &"'" & vbCrlf

		'response.write sqlStr
		dbget.execute sqlStr
	else
		sqlStr = " insert into [db_item].[dbo].tbl_item_option_stock" & vbCrlf
		sqlStr = sqlStr & " (itemgubun,itemid,itemoption,barcode)" & vbCrlf
		sqlStr = sqlStr & " values("
		sqlStr = sqlStr & " N'"& itemgubun &"'" & vbCrlf
		sqlStr = sqlStr & " ,"& itemid &"" & vbCrlf
		sqlStr = sqlStr & " ,N'"& itemoption &"'" & vbCrlf
		sqlStr = sqlStr & " ,'"& barcode &"'" & vbCrlf
		sqlStr = sqlStr & " )" & vbCrlf

		'response.write sqlStr
		'response.end
		dbget.execute sqlStr
	end if

	response.write "<script type='text/javascript'>"
	response.write "	parent.ReActWithThis('"&itemgubun&"', '"&itemid&"', '"&itemoption&"');"
	response.write "	alert('85�ڵ� ����ǰ�� ��� �Ǿ����ϴ�.');"
	response.write "</script>"
	dbget.close()	:	response.End

CASE Else
	Alert_return("������ ó���� ������ �߻��Ͽ����ϴ�.�����ڿ��� ������ �ּ���2")
       dbget.close()	:	response.End
END SELECT
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->