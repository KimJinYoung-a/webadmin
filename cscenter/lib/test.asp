<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->

<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<%
dim id : id = request("id")
	if id = "" then id = "2351713"
'dim detailidx : detailidx = request("detailidx")
'	if detailidx = "" then detailidx = "42757801"

    dim oCsAction,strMailHTML,strMailTitle
	Set oCsAction = New CsActionMailCls
	strMailHTML = oCsAction.makeMailTemplate_GiftCard(id)

'dim tmp_sendmailmsg
'dim MisendReason : MisendReason = "03"
'if (MisendReason <> "05") then
'	tmp_sendmailmsg = GetMichulgoMailString(MisendReason)
'	tmp_sendmailmsg = Replace(tmp_sendmailmsg, "\n", "<br>")
'end if

'Call SendMiChulgoMailWithMessage(detailidx, tmp_sendmailmsg)
response.write strMailHTML

CLASS CsActionMailCls

	''// ���� ��� �̹���
	Public Function getMailHeadImage()
		dim tmpImg
		IF FDivCD="A000" Then '// �±�ȯ���
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A001" Then '// ������߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A002" Then '// ���񽺹߼�
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A003" Then '// ȯ�ҿ�û
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A004" Then '// ��ǰ����(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A007" Then '// �ſ�/��ü���
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A008" Then '// �ֹ����
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A010" Then '// ȸ����û(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A011" Then '// �±�ȯȸ��(��)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A900" Then '// �ֹ���������
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order_comp.png' alt='CS�ȳ�����' style='vertical-align:top;' />"
			End IF
		ELSE

		END IF
		getMailHeadImage = tmpImg
	End Function

	''// ���� ����	'2017.12.19 �ѿ�� ����
	Public Function getMailHeadtitle()
		dim tmptitle
		IF FDivCD="A000" Then '// �±�ȯ���
			IF FCurrState="B001" Then
				tmptitle = "��ȯ��� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ȯ��� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A001" Then '// ������߼�
			IF FCurrState="B001" Then
				tmptitle = "������߼� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "������߼� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A002" Then '// ���񽺹߼�
			IF FCurrState="B001" Then
				tmptitle = "���� �߼� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "���� �߼� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A003" Then '// ȯ�ҿ�û
			IF FCurrState="B001" Then
				tmptitle = "ȯ�� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "ȯ�� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A004" Then '// ��ǰ����(��)
			IF FCurrState="B001" Then
				tmptitle = "��ǰ ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ǰ �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A007" Then '// �ſ�/��ü���
			IF FCurrState="B001" Then
				tmptitle = "������� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "������� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A008" Then '// �ֹ����
			IF FCurrState="B001" Then
				'tmptitle = "�ֹ���� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "�ֹ���� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A010" Then '// ȸ����û(��)
			IF FCurrState="B001" Then
				tmptitle = "��ǰȸ�� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ǰȸ�� �Ϸ� �ȳ�����"
			End IF
		ELSEIF FDivCD="A011" Then '// �±�ȯȸ��(��)
			IF FCurrState="B001" Then
				tmptitle = "��ȯ��ǰ ȸ�� ���� ����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "��ȯ��ǰ ȸ�� �Ϸ� ����"
			End IF
		ELSEIF FDivCD="A900" Then '// �ֹ���������
			IF FCurrState="B001" Then
				'tmptitle = "�ֹ��������� ���� �ȳ�����"
			ELSEIF FCurrState="B007" Then
				tmptitle = "�ֹ��������� �Ϸ� �ȳ�����"
			End IF
		ELSE

		END IF
		getMailHeadtitle = tmptitle
	End Function

	''//���� ��ǰ ���� ��������		'/2017.12.19 �ѿ��
	Function getAsItemLIst()
		dim tmpHTML
		dim OCsDetail,i

		tmpHTML = ""

		'A001(������߼�), A008(�ֹ����), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�), A004(��ǰ����(��))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A008" or FDivCD="A010" or FDivCD="A011" THEN
			Set OCsDetail = New CCSASList
			OCsDetail.FRectCsAsID = FAsID
			IF FResultCount>0 THEN
				OCsDetail.GetCsDetailList
			END IF

			if (OCsDetail.FresultCount<1) then Exit function

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left; color:#000;'>���� ��ǰ ����</th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:12px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""����"", sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;'>��ǰ</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>��ǰ�ڵ�</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:295px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>��ǰ��[�ɼ�]</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:right; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>�ǸŰ���</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:25px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>&nbsp;</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>����</th>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF OCsDetail.FresultCount>0 Then
				FOR i=0 TO OCsDetail.FResultCount-1
				    IF (OCsDetail.FItemList(i).Fitemid<>0) or (OCsDetail.FItemList(i).Fitemcost<>0) then
						tmpHTML=tmpHTML&"							<tr>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;'>" & vbcrlf
						tmpHTML=tmpHTML&"									<img src='"& OCsDetail.FItemList(i).FSmallImage &"' alt='' />" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
						tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemid &"" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).Fitemid=0) Then
							tmpHTML=tmpHTML&"									��ۺ�" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemname &"" & vbcrlf
						END IF
						if ( OCsDetail.FItemList(i).Fitemoptionname <>"") then
							tmpHTML=tmpHTML&"									["& OCsDetail.FItemList(i).Fitemoptionname &"]" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf

						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, ""����"", sans-serif;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).FdiscountAssingedCost<>0) and (OCsDetail.FItemList(i).Fitemcost>OCsDetail.FItemList(i).FdiscountAssingedCost) then
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:6px 0; font-size:11px; line-height:16px; color:#707070; font-family:dotum, ""����"", sans-serif; text-decoration:line-through; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) & "��</span><br />" & vbcrlf
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""����"", sans-serif; text-align:right;'>" & FormatNumber(OCsDetail.FItemList(i).FdiscountAssingedCost,0) &"��</span>" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""����"", sans-serif; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) &"��</span>" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, ""����"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>"& OCsDetail.FItemList(i).Fregitemno &"</td>" & vbcrlf
						tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			        END IF
				NEXT
			END IF

			tmpHTML=tmpHTML&"						</table>" & vbcrlf
			tmpHTML=tmpHTML&"					</td>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"			</table>" & vbcrlf
			tmpHTML=tmpHTML&"		</td>" & vbcrlf
			tmpHTML=tmpHTML&"	</tr>" & vbcrlf

			Set OCsDetail= nothing
		END IF
		getAsItemLIst = tmpHTML
	END Function

	''//���ּ� ��������		'/2017.12.19 �ѿ��
	Function getReqInfo()
		dim tmpHTML
		tmpHTML=""

		'A001(������߼�), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�)
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A010" THEN 'or FDivCD="A011"
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>���ּ�</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& AstarUserName(trim(FReqName)) &" ���� &nbsp; &nbsp;"& AstarPhoneNumber(trim(FReqPhone)) &" / "& AstarPhoneNumber(trim(FReqHP)) &" <br />["& printUserId(trim(FReqZipcode), 2, "*") &"] "& printUserId(trim(FReqZipAddr), 2, "*") &"&nbsp;(���ϻ���)" & vbcrlf		' FReqEtcAddr
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF
		getReqInfo = tmpHTML
	END Function

	''//��ü �ּ� ��������		'/2017.12.19 �ѿ��
	Function getReturnInfo()
		dim tmpHTML
		tmpHTML=""

		' A011(�±�ȯȸ��(��)), A010(ȸ����û(��)), A004(��ǰ����(��))
		IF FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��ǰȸ���ּ�</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnName &" &nbsp; &nbsp;"& FReturnPhone &"<br />["& FReturnZipCode &"] "& FReturnZipAddr &"&nbsp;"& FReturnEtcAddr &"" & vbcrlf
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			if (FReturnName<>"(��)�ٹ�����") and (FupcheReturnSongjangDivName<>"") and (Left(FupcheReturnSongjangDivTel,1)="1" or Left(FupcheReturnSongjangDivTel,1)="0") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�̿��ù��</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
				tmpHTML=tmpHTML&"									"& FupcheReturnSongjangDivName &"<br />�ù�翬��ó : "& FupcheReturnSongjangDivTel &"" & vbcrlf
				tmpHTML=tmpHTML&"								</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			END IF
		END IF

		getReturnInfo = tmpHTML
	END Function

	''//ȯ������ ��������		'/2017.12.19 �ѿ��
	Function getRefundInfo()
		dim tmpHTML
		tmpHTML=""

		' A008(�ֹ����), A010(ȸ����û(��)), A007(�ſ�/��ü���), A003(ȯ�ҿ�û), A004(��ǰ����(��))
		IF FDivCD="A003" or FDivCD="A004" or FDivCD="A007" or FDivCD="A008" or FDivCD="A010" THEN
		    ''ȯ�Ҿ�0�̸� return
		    if (FRefundRequire=0) then Exit function

		    ''����Ȯ�� ȯ�� ���� ����
		    if (FReturnMethod="R007") then
		        if (Len(Replace(FReBankAccount,"-",""))<7) then
    		        FReBankName = ""
    		        FReBankAccount = "����Ȯ�ο��"
    		        FReBankOwnerName =""
    		    else
    		        FReBankAccount = Left(FReBankAccount,Len(Trim(FReBankAccount))-3) + "***"
    		    end if
		    end if

			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ȯ�ҿ�����</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FormatNumber(FRefundRequire,0) &" ��" & vbcrlf

			'��ۺ����� �ȳ��� ���ظ� ������ ǥ�þ��ϰ� ����
			'if (FRefundDeliveryPay<>0) then
			'			tmpHTML=tmpHTML&"									(��ۺ����� : " & FormatNumber(FRefundDeliveryPay+Frefundbeasongpay,0) &")" & vbcrlf
			'end if

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ȯ������(����)</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnMethodName &"&nbsp;&nbsp;" & vbcrlf
	
			IF (FReturnMethod="R007") THEN
				tmpHTML=tmpHTML&"									"& FReBankName &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& FReBankAccount &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& AstarUserName(FReBankOwnerName) &" " & vbcrlf
			ELSEIF (FReturnMethod="R900") THEN
				tmpHTML=tmpHTML&"									(�������̵� : "& FUserID &") " & vbcrlf
			ELSEIF (FReturnMethod="R100") or (FReturnMethod="R550") or (FReturnMethod="R560") or (FReturnMethod="R120") or (FReturnMethod="R020") or (FReturnMethod="R080") THEN
				if (Left(FPayGateTid,6)="IniTec") and (FCurrState="B007") and (FReturnMethod<>"R120") then
					tmpHTML=tmpHTML&"									<a target='_blank' href=https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="& FPayGateTid &"&noMethod=1>[������ǥ���]</a> " & vbcrlf
				end if
				if (FReturnMethod = "R550") or (FReturnMethod = "R560") then
					tmpHTML=tmpHTML&"									������/����Ƽ�� �� ����� ���������� �߱޵Ǹ�, ��ǰ���Žÿ��� ����Ұ��Դϴ�." & vbcrlf
				end if
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		END IF
		getRefundInfo = tmpHTML
	END Function

	'// ó�� ��� ��������		'/2017.12.19 �ѿ��
	Function getFinishResult()
		dim tmpHTML
		tmpHTML=""

		'A001(������߼�), A011(�±�ȯȸ��(��)), A010(ȸ����û(��)), A003(ȯ�ҿ�û), A002(���񽺹߼�), A004(��ǰ����(��)), A000(�±�ȯ���)
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

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; text-align:left; font-family:dotum, ""����"", sans-serif; text-align:left;'>ó����� ����<th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ó���Ϸ���</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FFinishDate &"</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF (Trim(FOpenContents)<>"") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>ó������</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:16px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& nl2br(FOpenContents) &"</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			end IF

			''// �ù����� ��������
			tmpHTML=tmpHTML& getDlvInfo()

			tmpHTML=tmpHTML&"						</table>" & vbcrlf
			tmpHTML=tmpHTML&"					</td>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"			</table>" & vbcrlf
			tmpHTML=tmpHTML&"		</td>" & vbcrlf
			tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		END IF
		getFinishResult = tmpHTML
	END Function

	''// �ù� ���� ��������		'/2017.12.19 �ѿ��
	Function getDlvInfo()
		dim tmpHTML
		tmpHTML=""

        if (IsNULL(FSongjangNo)) or (FSongjangNo="") then Exit function

		'A001(������߼�), A011(�±�ȯȸ��(��)), A000(�±�ȯ���), A010(ȸ����û(��)), A002(���񽺹߼�), A004(��ǰ����(��))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�ù�����</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

			IF FSongjangNo<>"" then
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>"& FSongjangDivName &"</span>"& vbcrlf
				tmpHTML=tmpHTML&"									&nbsp;&nbsp;<a href='"& DeliverDivTrace(Trim(FSongjangDiv)) & FSongjangNo &"' target='_blank' style='margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; color:#0066cc; text-align:left;'>"& FSongjangNo &"</a>" & vbcrlf
			ELSE
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070;'>�ù������� ��ϵ��� �ʾҽ��ϴ�.</span>" & vbcrlf
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF

		getDlvInfo =  tmpHTML
	END Function

	'// ��Ÿ �ȳ�����		'/2017.12.19 �ѿ��
	Public Function getEtcNotice()
		dim tmpHTML

        getEtcNotice = ""

        if (Trim(FInfoHtml)="") then Exit function

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left;'>��Ÿ �ȳ� ����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FInfoHtml &"" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		getEtcNotice = tmpHTML
	End Function

	''// ���� �⺻ ���� ��������		'/2017.12.19 �ѿ��
	Function getAsInfo()
		dim tmpHTML

		tmpHTML = ""
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FRegDate &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�����ڵ�</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FAsID &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>�ֹ���ȣ</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& FOrderSerial &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FTitle &" <a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp?orderserial=" & FOrderSerial & "' target='_blank'>[�󼼳�������]</a>" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>��������</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:left;'>"& GetCauseDetailString &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
				
		getAsInfo =tmpHTML
	END Function

	'/ �̸��� ���ø� �����ͼ� ����°ɷ� ����.	2017.12.18 �ѿ�� ����
	Function makeMailTemplate_GiftCard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster_GiftCard(id) '// �� ����

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' ���

		tmpHTML=mailheader

		'A001(������߼�), A008(�ֹ����), A000(�±�ȯ���)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			������ "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			������ ��û�Ͻ� "& GetAsDivCDName &"�� ���������� ó���� "& FCurrStateName &" �Ǿ����ϴ�.<br />�ٹ������� �̿����ּż� �����մϴ�." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; text-align:right;'>�ֹ���ȣ :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""����"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ó����� ��������
		tmpHTML=tmpHTML& getFinishResult()

		''// ���� ��ǰ ���� ��������
		'tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""����"", sans-serif; text-align:left; color:#000;'>���� ����</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""����"", sans-serif; color:#707070;'>" & vbcrlf

		''// ���� �⺻ ���� ��������
		tmpHTML=tmpHTML& getAsInfo()

		''// ���ּ� ��������
		'tmpHTML=tmpHTML& getReqInfo()

		''// ��ü�ּ� ��������
		'tmpHTML=tmpHTML& getReturnInfo()

		''// ȯ������ ��������
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// ��Ÿ �ȳ�����
		'tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='���� ���� �󼼺���' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='�ٹ����� �����ϱ�' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""����"", sans-serif; color:#707070; text-align:center;'>������ ��� ���� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' Ǫ��

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' �̸�������

		makeMailTemplate_GiftCard = tmpHTML
	End Function


	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

	Dim FAsID
	Dim FDivCD
	Dim FGubun01
	Dim FGubun02

	Dim FDivCDName
	Dim FGubun01Name
	Dim FGubun02Name

	Dim FOrderSerial
	Dim FCustomerName
	Dim FUserid
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
	Dim FExtSiteName

	Dim FOpenTitle
	Dim FOpenContents

	Dim FSiteGubun

	Dim FSongjangDiv
	Dim FSongjangNo
	Dim FSongjangDivName

	Dim FRequireUpche
	Dim FMakerid

	Dim FAdd_upchejungsanDeliveryPay
	Dim FAdd_upchejungsanCause

	Dim FOrgSubTotalPrice
	Dim FOrgItemCostSum
	Dim FOrgBeasongPay
	Dim FOrgMileageSum
	Dim FOrgCouponSum
	Dim FOrgAllatDiscountSum

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

	Dim Fencmethod
	Dim FdecAccount

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
    
    Dim Faccountdiv      ''2016/08/05 �߰�
    Dim Fpggubun         ''2016/08/05 �߰�
    
 	public function GetAsDivCDName()
        GetAsDivCDName = db2html(FDivCDName)
	end function

	public function GetCauseDetailString()
        GetCauseDetailString = Fgubun02Name
    end function

	Public Sub GetOneCSASMaster(FRectCsAsID)

		dim strSQL
		strSQL =" SELECT TOP 1 " &_
				" 	A.ID ,A.DivCD ,A.Gubun01 ,A.Gubun02 ,A.OrderSerial ,A.CustomerName ,A.Userid ,A.WriteUser ,A.FinishUser " &_
				"	,A.Title ,A.Contents_Jupsu ,A.Contents_Finish ,A.CurrState ,A.RegDate ,A.FinishDate ,A.Deleteyn ,A.ExtSiteName "&_
				"	,A.OpenTitle ,A.OpenContents ,A.RequireUpche ,A.Makerid ,A.SongjangDiv ,A.SongjangNo ,A.SiteGubun "&_
				"	,(SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=A.SongjangDiv) AS SongjangDivName " &_
				" 	,o.BuyHp,o.BuyEmail " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.divCD) as divcdname " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun01) as gubun01name " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun02) as gubun02name "
		IF (FRectForceCurrState<>"") then
		    strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd='"&FRectForceCurrState&"') as currstatename "
        ELSE
            strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.currstate) as currstatename "
        END IF

		strSQL = strSQL & " 	,IsNULL(J.add_upchejungsandeliverypay,0) as add_upchejungsandeliverypay , J.add_upchejungsancause " &_

				" 	,r.OrgSubTotalPrice,r.OrgItemCostSum,r.OrgBeasongPay,r.OrgMileageSum,r.OrgCouponSum,r.OrgAllatDiscountSum "&_
				" 	,IsNULL(r.RefundRequire,0) as RefundRequire ,isNULL(r.RefundResult,0) as RefundResult "&_
				"	,r.ReturnMethod,r.RefundMileageSum,r.RefundCouponSum,r.AllatSubTractSum "&_
				"	,r.RefundItemCostSum,r.RefundBeasongPay,r.RefundDeliveryPay,r.RefundAdjustPay,r.CancelTotal "&_
				" 	,r.RebankName ,r.RebankAccount ,r.RebankOwnerName ,r.PayGateTid " &_
				"   ,r.encmethod " &_
				"   , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount " &_
				" 	,r.paygateresultTid,r.PayGateResultMsg " &_
				" 	,(SELECT top 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=r.returnmethod and comm_group='Z090') as ReturnMethodName " &_

				" 	,IsNULL(D.ReqName,o.reqname) as ReqName ,IsNULL(D.ReqPhone,o.reqphone) as ReqPhone ,IsNULL(D.ReqHP,o.reqhp) as ReqHP " &_
				" 	,IsNULL(D.ReqZipcode,o.reqzipcode) as ReqZipcode ,IsNULL(D.ReqZipAddr,o.reqzipaddr) as ReqZipAddr ,IsNULL(D.ReqEtcAddr,o.reqaddress) as ReqEtcAddr ,IsNULL(D.ReqEtcStr,'') as ReqEtcStr " &_
				" 	,isNull(p.company_name,'(��)�ٹ�����') as ReturnName ,isNull(p.deliver_phone,'1644-6030') as ReturnPhone ,isNull(p.deliver_hp,'') as ReturnHP "&_
				" 	,isNull(p.return_zipcode,'132-010') as ReturnZipCode ,isNull(p.return_address,'���� ������') as ReturnZipAddr ,isNull(p.return_address2,'������ 63���� ���δ��� 3��') as ReturnEtcAddr "&_
                " 	,isNull((SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivName "&_
                " 	,isNull((SELECT TOP 1 tel FROM db_order.dbo.tbl_songjang_div WHERE divcd=p.defaultsongjangdiv),'') as upcheReturnSongjangDivTel "&_
                "   ,isNULL(o.accountdiv,'') as accountdiv, isNULL(o.pggubun,'') as pggubun"&_
                
				" FROM [db_cs].[dbo].tbl_new_as_list A " &_
				" LEFT JOIN db_order.dbo.tbl_order_master o " &_
				" 	on A.orderserial=o.orderserial " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_upcheAddjungsan J " &_
				" 	on A.id=J.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_refund_info r " &_
				" 	on A.id=r.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery d " &_
				" 	on A.id = d.asid " &_
				" LEFT JOIN [db_partner].[dbo].tbl_partner p " &_
				" 	on A.makerid= p.id " &_
				" WHERE A.id=" & CStr(FRectCsAsID)

			rsget.Open strSQL, dbget, 1

	        FResultCount = rsget.RecordCount

	        if  not rsget.EOF  then
	        	'//GetOneCSASMaster
				FAsID		= rsget("ID")
				FDivCD	= rsget("divCD")
				FGubun01	= rsget("gubun01")
				FGubun02	= rsget("gubun02")

				FDivCDName	= rsget("divcdname")
				FGubun01Name	= rsget("gubun01name")
				FGubun02Name	= rsget("gubun02name")

				FOrderSerial	= rsget("orderserial")
				FCustomerName	= rsget("customername")
				FUserid	= rsget("userid")
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
				FExtSiteName	= rsget("ExtSiteName")

				FOpenTitle	= rsget("OpenTitle")
				FOpenContents	= rsget("OpenContents")

				FSiteGubun	= rsget("SiteGubun")

				FSongjangDiv	= rsget("SongjangDiv")
				FSongjangNo	= rsget("SongjangNo")
				FSongjangDivName = rsget("SongjangDivName")
				FRequireUpche	= rsget("RequireUpche")
				FMakerid	= rsget("Makerid")

				FAdd_upchejungsanDeliveryPay	= rsget("Add_upchejungsanDeliveryPay")
				FAdd_upchejungsanCause	= rsget("Add_upchejungsanCause")

				'//GetOneRefundInfo
				FOrgSubTotalPrice	= rsget("OrgSubTotalPrice")
				FOrgItemCostSum	= rsget("OrgItemCostSum")
				FOrgBeasongPay	= rsget("OrgBeasongPay")
				FOrgMileageSum	= rsget("OrgMileageSum")
				FOrgCouponSum	= rsget("OrgCouponSum")
				FOrgAllatDiscountSum	= rsget("OrgAllatDiscountSum")
				FRefundRequire	= rsget("RefundRequire")
				FRefundResult	= rsget("RefundResult")
				FReturnMethod	= rsget("ReturnMethod")
				FRefundMileageSum	= rsget("RefundMileageSum")
				FRefundCouponSum	= rsget("RefundCouponSum")
				FRefundItemCostSum	= rsget("RefundItemCostSum")
				FRefundBeasongPay	= rsget("RefundBeasongPay")
				FRefundDeliveryPay	= rsget("RefundDeliveryPay")
				FRefundAdjustPay	= rsget("RefundAdjustPay")

				FAllatSubTractSum	= rsget("AllatSubTractSum")
				FCancelTotal	= rsget("CancelTotal")

				FReBankName	= rsget("ReBankName")
				FReBankAccount	= rsget("ReBankAccount")
				Fencmethod      = rsget("encmethod")
				FdecAccount      = rsget("decAccount")
				IF (Fencmethod="PH1") then FReBankAccount=FdecAccount
				IF (Fencmethod="AE2") then FReBankAccount=FdecAccount

				FReBankOwnerName	= rsget("ReBankOwnerName")
				FPayGateTid	= rsget("PayGateTid")

				FPayGateResultTid	= rsget("PayGateResultTid")
				FPayGateResultMsg	= rsget("PayGateResultMsg")

				FReturnMethodName	= rsget("ReturnMethodName")

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
                
                Faccountdiv = rsget("accountdiv")
                Fpggubun    = rsget("pggubun")
                
                if (Fpggubun="NP") then
                    if (FReturnMethod="R100" or FReturnMethod="R120" or FReturnMethod="R020" or FReturnMethod="R022") then
                        FReturnMethodName = "���̹��������"
                    end if
                end if
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

	Public Sub GetOneCSASMaster_GiftCard(FRectCsAsID)

		dim strSQL
		strSQL =" SELECT TOP 1 " &_
				" 	A.ID ,A.DivCD ,A.Gubun01 ,A.Gubun02 ,A.OrderSerial ,A.CustomerName ,A.Userid ,A.WriteUser ,A.FinishUser " &_
				"	,A.Title ,A.Contents_Jupsu ,A.Contents_Finish ,A.CurrState ,A.RegDate ,A.FinishDate ,A.Deleteyn ,A.ExtSiteName "&_
				"	,A.OpenTitle ,A.OpenContents ,A.RequireUpche ,A.Makerid ,A.SongjangDiv ,A.SongjangNo ,A.SiteGubun "&_
				"	,(SELECT TOP 1 divname FROM db_order.dbo.tbl_songjang_div WHERE divcd=A.SongjangDiv) AS SongjangDivName " &_
				" 	,o.BuyHp,o.BuyEmail " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.divCD) as divcdname " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun01) as gubun01name " &_
				" 	,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.gubun02) as gubun02name "
		IF (FRectForceCurrState<>"") then
		    strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd='"&FRectForceCurrState&"') as currstatename "
        ELSE
            strSQL = strSQL & "  ,(SELECT TOP 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=A.currstate) as currstatename "
        END IF


		strSQL = strSQL & " 	,r.OrgSubTotalPrice,r.OrgItemCostSum,r.OrgBeasongPay,r.OrgMileageSum,r.OrgCouponSum,r.OrgAllatDiscountSum "&_
				" 	,IsNULL(r.RefundRequire,0) as RefundRequire ,isNULL(r.RefundResult,0) as RefundResult "&_
				"	,r.ReturnMethod,r.RefundMileageSum,r.RefundCouponSum,r.AllatSubTractSum "&_
				"	,r.RefundItemCostSum,r.RefundBeasongPay,r.RefundDeliveryPay,r.RefundAdjustPay,r.CancelTotal "&_
				" 	,r.RebankName ,r.RebankAccount ,r.RebankOwnerName ,r.PayGateTid " &_
				"   ,r.encmethod " &_
				"   , (CASE WHEN r.encmethod='PH1' THEN IsNull(db_cs.dbo.uf_DecAcctPH1(r.encaccount), '') WHEN r.encmethod='AE2' THEN IsNull(db_cs.dbo.uf_DecAcctAES256(r.encaccount), '') ELSE '' END) as decaccount " &_
				" 	,r.paygateresultTid,r.PayGateResultMsg " &_
				" 	,(SELECT top 1 comm_name FROM db_cs.dbo.tbl_cs_comm_code WHERE comm_cd=r.returnmethod and comm_group='Z090') as ReturnMethodName " &_

				" FROM [db_cs].[dbo].tbl_new_as_list A " &_
				" LEFT JOIN db_order.dbo.tbl_giftcard_order o " &_
				" 	on A.orderserial=o.giftorderserial " &_
				" LEFT JOIN [db_cs].[dbo].tbl_as_refund_info r " &_
				" 	on A.id=r.asid " &_
				" LEFT JOIN [db_cs].[dbo].tbl_new_as_delivery d " &_
				" 	on A.id = d.asid " &_
				" WHERE A.id=" & CStr(FRectCsAsID)

			rsget.Open strSQL, dbget, 1

	        FResultCount = rsget.RecordCount

	        if  not rsget.EOF  then
	        	'//GetOneCSASMaster
				FAsID		= rsget("ID")
				FDivCD	= rsget("divCD")
				FGubun01	= rsget("gubun01")
				FGubun02	= rsget("gubun02")

				FDivCDName	= rsget("divcdname")
				FGubun01Name	= rsget("gubun01name")
				FGubun02Name	= rsget("gubun02name")

				FOrderSerial	= rsget("orderserial")
				FCustomerName	= rsget("customername")
				FUserid	= rsget("userid")
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
				FExtSiteName	= rsget("ExtSiteName")

				FOpenTitle	= rsget("OpenTitle")
				FOpenContents	= rsget("OpenContents")

				FSiteGubun	= rsget("SiteGubun")

				FSongjangDiv	= rsget("SongjangDiv")
				FSongjangNo	= rsget("SongjangNo")
				FSongjangDivName = rsget("SongjangDivName")
				FRequireUpche	= rsget("RequireUpche")
				FMakerid	= rsget("Makerid")

				'FAdd_upchejungsanDeliveryPay	= rsget("Add_upchejungsanDeliveryPay")
				'FAdd_upchejungsanCause	= rsget("Add_upchejungsanCause")

				'//GetOneRefundInfo
				FOrgSubTotalPrice	= rsget("OrgSubTotalPrice")
				FOrgItemCostSum	= rsget("OrgItemCostSum")
				FOrgBeasongPay	= rsget("OrgBeasongPay")
				FOrgMileageSum	= rsget("OrgMileageSum")
				FOrgCouponSum	= rsget("OrgCouponSum")
				FOrgAllatDiscountSum	= rsget("OrgAllatDiscountSum")
				FRefundRequire	= rsget("RefundRequire")
				FRefundResult	= rsget("RefundResult")
				FReturnMethod	= rsget("ReturnMethod")
				FRefundMileageSum	= rsget("RefundMileageSum")
				FRefundCouponSum	= rsget("RefundCouponSum")
				FRefundItemCostSum	= rsget("RefundItemCostSum")
				FRefundBeasongPay	= rsget("RefundBeasongPay")
				FRefundDeliveryPay	= rsget("RefundDeliveryPay")
				FRefundAdjustPay	= rsget("RefundAdjustPay")

				FAllatSubTractSum	= rsget("AllatSubTractSum")
				FCancelTotal	= rsget("CancelTotal")

				FReBankName	= rsget("ReBankName")
				FReBankAccount	= rsget("ReBankAccount")
				Fencmethod      = rsget("encmethod")
				FdecAccount      = rsget("decAccount")
				IF (Fencmethod="PH1") then FReBankAccount=FdecAccount
				IF (Fencmethod="AE2") then FReBankAccount=FdecAccount

				FReBankOwnerName	= rsget("ReBankOwnerName")
				FPayGateTid	= rsget("PayGateTid")

				FPayGateResultTid	= rsget("PayGateResultTid")
				FPayGateResultMsg	= rsget("PayGateResultMsg")

				FReturnMethodName	= rsget("ReturnMethodName")

				'//GetReturnAddress
				'FReturnName	= rsget("ReturnName")
				'FReturnPhone	= rsget("ReturnPhone")
				'FReturnHP	= rsget("ReturnHP")
				'FReturnZipCode	= rsget("ReturnZipCode")
				'FReturnZipAddr	= rsget("ReturnZipAddr")
				'FReturnEtcAddr	= rsget("ReturnEtcAddr")

				'FReqName	= rsget("ReqName")
				'FReqPhone	= rsget("ReqPhone")
				'FReqHP		= rsget("ReqHP")
				'FReqZipcode	= rsget("ReqZipcode")
				'FReqZipAddr	= rsget("ReqZipAddr")
				'FReqEtcAddr	= rsget("ReqEtcAddr")
				'FReqEtcStr	= rsget("ReqEtcStr")

                'FupcheReturnSongjangDivName = db2html(rsget("upcheReturnSongjangDivName"))
                'FupcheReturnSongjangDivTel  = db2html(rsget("upcheReturnSongjangDivTel"))
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
	
	
End Class

Class CCSASDetailItem
    ''tbl_as_detail's
    public Fid
    public Fmasterid
    public Fgubun01
    public Fgubun02
    public Fgubun01name
    public Fgubun02name
    public Fregdetailstate
    public Fregitemno
    public Fconfirmitemno
    public Fcausediv
    public Fcausedetail
    public Fcausecontent

    ''tbl_order_detail's
    public Forderdetailidx
    public Forderserial
    public Fitemid
    public Fitemoption
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    public Fitemcost
    public Fbuycash
    public Fitemno
    public Forderitemno
    public Fisupchebeasong
    public Fcancelyn

    public Foitemdiv
    public FodlvType
    public Fissailitem
    public Fitemcouponidx
    public Fbonuscouponidx

    public ForderDetailcurrstate
    public FdiscountAssingedCost    '' �ֹ��� ���εȰ��� ( ALL@ / %���α� �ݿ�)

    public Forgitemcost					'�Һ��ڰ�
    public FitemcostCouponNotApplied	'�ǸŰ�(���ΰ�)
    public FplusSaleDiscount			'�÷����������ξ�
    public FspecialshopDiscount			'��������ξ�

    public Forgprice					'����Һ��ڰ�(+�ɼǰ�)

	public Fprevcsreturnfinishno		'���� CS��ǰ����(�����̻�)

	public Freforderdetailidx

	Public Fsongjangdiv
	Public Fsongjangno

    ''public FAllAtDiscountedPrice

    ''tbl_item's
    public FSmallImage

    ''��ü ������� ��ǰ ��ۺ� ���� ����
    public function IsUpcheParticleDeliverPayCodeItem
        IsUpcheParticleDeliverPayCodeItem = (Fitemid=0) and (Left(Fitemoption,2)="90")
    end function

    ''��ü ������� ��ǰ���� ����
    public function IsUpcheParticleDeliverItem
        IsUpcheParticleDeliverItem = (FodlvType=9)
    end function

    ''��ǰ�� ����ϴ� ��ǰ����(All@ ���ΰ�, %���� ���ΰ� �ݿ�)
    public function GetOrgPayedItemPrice()
        GetOrgPayedItemPrice = Fitemcost

        if (FdiscountAssingedCost=0) then
            ''�������
            GetOrgPayedItemPrice = Fitemcost-getAllAtDiscountedPrice
        else
            if (FdiscountAssingedCost<>Fitemcost) then
                GetOrgPayedItemPrice = FdiscountAssingedCost
            end if
        end if
    end function

    ''All@ ���εȰ���
    public function getAllAtDiscountedPrice()
        getAllAtDiscountedPrice =0
        ''���� ��ǰ���� ���εǴ°�� �߰����ξ���.
        ''���ϸ���SHOP ��ǰ �߰� ���� ����.
	    ''���ϻ�ǰ �߰����� ����
	    '' 20070901�߰� : �������� ���ʽ��������� �߰����� ����.

'	    if (FdiscountAssingedCost=0) then
'	        ''�������
'            if (Fitemcouponidx<>0) or (IsMileShopSangpum) or (Fissailitem="Y") then
'    			getAllAtDiscountedPrice = 0
'    		else
'    			getAllAtDiscountedPrice = round(((1-0.94) * FItemCost / 100) * 100 ) * FItemNo
'    		end if
'    	else
    	    if (IsNULL(Fbonuscouponidx) or (Fbonuscouponidx=0)) and (Fitemcost>FdiscountAssingedCost) then
    	            getAllAtDiscountedPrice = Fitemcost-FdiscountAssingedCost
    	    else
    	        getAllAtDiscountedPrice = 0
    	    end if
'    	end if
    end function

    '' %���α� ���αݾ� or ī�� ���αݾ�
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0
'        if (Fitemcost>FdiscountAssingedCost) then
'                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
'        end if

		if (Fitemid = 0) and (Fitemcost > FdiscountAssingedCost) and not IsNull(Fbonuscouponidx) then
			'// ��ۺ� ����
			getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
        ''elseif (FdiscountAssingedCost=0) then
	        ''�������
	    ''    ''getPercentBonusCouponDiscountedPrice = Fitemcost*
		else
			'// ���� �������� ����(2014-06-23, skyer9)
            if (Fbonuscouponidx<>0)  and (Fitemcost>FdiscountAssingedCost) then
                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
            end if
        end if
    end function

    ''���ϸ����� ��ǰ
    public function IsMileShopSangpum()
		IsMileShopSangpum = false

		if Foitemdiv="82" then
			IsMileShopSangpum = true
		end if
	end function

    public function GetDefaultRegNo(IsRegState)
        if (IsRegState) then
            GetDefaultRegNo = Fitemno
        else
            GetDefaultRegNo = Fregitemno
        end if
    end function

    ''CsAction ������ ��ǰ ���� ���� ���ɿ���
    public function IsItemNoEditEnabled(byval idivcd)
        IsItemNoEditEnabled = false

        if (Fcancelyn="Y") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsItemNoEditEnabled = true

            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ ����
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true

        elseif (fnIsServiceDeliverProcess(idivcd)) or (fnIsServiceRecvProcess(idivcd)) then
            '���� - �׻� ���� ���� ����
            if (idivcd = "A002") or (idivcd = "A200") then
            	IsItemNoEditEnabled=true

            elseif (ForderDetailcurrstate>=7) then
            	IsItemNoEditEnabled=true

            end if
        end if
    end function


    ''CsAction ������ ��ǰ�� üũ ���ɿ���
    public function IsCheckAvailItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd)
        IsCheckAvailItem = false

        if (Fcancelyn="Y") then Exit function
        if (iMasterCancelYn<>"N") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsCheckAvailItem = true
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ ����
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true

            if (FItemId=0) then IsCheckAvailItem=true
        elseif (idivcd="A006") then
            ''���� ���ǻ���
            IsCheckAvailItem=true

            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false
        elseif (idivcd="A009") then
            ''��Ÿ����(�޸�) - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd="A700") then
            ''��Ÿ���� - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd = "A002") or (idivcd = "A200") then
        	'���� - �׻� üũ����
            if Fitemid=0 then
                IsCheckAvailItem=false
            else
                IsCheckAvailItem=true
            end if
        elseif (idivcd = "A001") then
            ''����
            if (ForderDetailcurrstate>=7) or ((Fcancelyn="A") and (iIpkumdiv>=7)) then IsCheckAvailItem=true
        elseif (idivcd = "A000") then
            ''�±�ȯ
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true
        else

        end if
    end function

    ''CsAction ������ ��ǰ�� ����Ʈ üũ��
    public function IsDefaultCheckedItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd, byval ckAll)
        IsDefaultCheckedItem =false

        if (Not IsCheckAvailItem(iIpkumdiv,iMasterCancelYn,idivcd)) then Exit function

        if (fnIsCancelProcess(idivcd)) then
            if (ckAll<>"") then
                IsDefaultCheckedItem = true
            else
                IsDefaultCheckedItem = false
            end if

            if (Fcancelyn="Y") or (iMasterCancelYn<>"N") then IsDefaultCheckedItem=false

            if (ForderDetailcurrstate>=3) then IsDefaultCheckedItem=false
        elseif (fnIsReturnProcess(idivcd)) then
            ''��ǰ�����ΰ�� - No action
        elseif (idivcd="A006") then
            ''���� ���ǻ��� - No action
        elseif (idivcd="A009") then
            ''��Ÿ����(�޸�) - No action
        else

        end if
    end function

	'==========================================================================
    '���ʽ����� ���� �ֹ����� üũ
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

	'��ǰ���� ���� �ֹ����� üũ
    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
        end if
    end function

    '��������� ���� �ֹ����� üũ
    public function IsSpecialShopDiscountItem()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (Not IsItemCouponDiscountItem) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : �Һ��ڰ�����, �ɼǰ������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetItemCouponDiscountPrice = (Forgprice - Fitemcost) = 0
        		exit function
        	end if

        	GetItemCouponDiscountPrice = false
        	exit function
        end if

		if (FspecialshopDiscount > 0) then
			IsSpecialShopDiscountItem = true
		else
			IsSpecialShopDiscountItem = false
		end if
    end function

	'��ǰ�������ξ�
    public function GetItemCouponDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (IsItemCouponDiscountItem = true) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : �Һ��ڰ�����, �ɼǰ�����, ����������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetItemCouponDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetItemCouponDiscountPrice = 0
        	exit function
        end if

        GetItemCouponDiscountPrice = FitemcostCouponNotApplied - Fitemcost
    end function

	'���ʽ��������ξ�
    public function GetBonusCouponDiscountPrice()
        GetBonusCouponDiscountPrice = Fitemcost - FdiscountAssingedCost
    end function

	'��ǰ���ξ�
    public function GetSaleDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'���ŵ���Ÿ
        	if (Not IsBonusCouponDiscountItem) and (Not IsItemCouponDiscountItem) and (Fissailitem = "Y") then
        		'TODO : �Һ��ڰ�����, �ɼǰ�����, ����������� �ִ°�� ����Ȯ�� ���� �ȴ�.
        		GetSaleDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetSaleDiscountPrice = 0
        	exit function
        end if

        GetSaleDiscountPrice = (Forgitemcost - (FitemcostCouponNotApplied + FplusSaleDiscount + FspecialshopDiscount))
    end function

    public function IsOldJumun()
    	'2011�� 4�� 1�� ���� �ֹ� �Ǵ� �� �ֹ��� ���� ���̳ʽ��ֹ�
    	IsOldJumun = (Forgitemcost = 0)
    end function

	public function GetOrgItemCostColor()
		if IsOldJumun then
			GetOrgItemCostColor = "gray"
		else
			GetOrgItemCostColor = "black"
		end if
	end function

	public function GetOrgItemCostPrice()
		if IsOldJumun then
			GetOrgItemCostPrice = Forgprice
		else
			GetOrgItemCostPrice = Forgitemcost
		end if
	end function

	public function GetSaleColor()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				GetSaleColor = "red"
			else
				GetSaleColor = "black"
			end if
		end if
	end function

	public function GetSalePrice()
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				GetSalePrice = Fitemcost
			else
				GetSalePrice = Forgprice
			end if
		else
			GetSalePrice = FitemcostCouponNotApplied
		end if
	end function

	public function GetSaleText()
		dim result

		result = ""
		if IsOldJumun then
			if (Fissailitem = "Y") or (Fissailitem = "P") or ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
				if (Fissailitem = "Y") then
					if (Forgprice <= Fitemcost) then
						result = result + "���λ�ǰ + �Һ��ڰ� ����" + vbCrLf
					else
						result = result + "���λ�ǰ" + vbCrLf
					end if
				end if
				if (Fissailitem = "P") then
					result = result + "�÷�������" + vbCrLf
				end if
				if ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
					result = result + "��������� �Ǵ� �Һ��ڰ�/�ɼǰ� ����" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				if (Fissailitem = "Y") then
					result = result + "���λ�ǰ : " + CStr(GetSaleDiscountPrice) + "��" + vbCrLf
				end if
				if (FplusSaleDiscount > 0) then
					result = result + "�÷������� : " + CStr(FplusSaleDiscount) + "��" + vbCrLf
				end if
				if (FspecialshopDiscount > 0) then
					result = result + "���ȸ������ : " + CStr(FspecialshopDiscount) + "��" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		end if

		GetSaleText = result
	end function

	public function GetItemCouponColor()
		if (IsItemCouponDiscountItem = true) then
			GetItemCouponColor = "green"
		else
			GetItemCouponColor = "black"
		end if
	end function

	public function GetItemCouponPrice()
		GetItemCouponPrice = Fitemcost
	end function

	public function GetItemCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsItemCouponDiscountItem = true) then
				if (GetSalePrice <> GetItemCouponPrice) then
					result = result + "��ǰ���������ǰ" + vbCrLf
				else
					result = result + "��ۺ����������ǰ" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		else
			if (IsItemCouponDiscountItem = true) then
				if (GetItemCouponDiscountPrice = 0) then
					result = result + "��ۺ����������ǰ" + vbCrLf
				else
					result = result + "��ǰ���� : " + CStr(GetItemCouponDiscountPrice) + "��" + vbCrLf
				end if
			else
				result = "���󰡰�"
			end if
		end if

		GetItemCouponText = result
	end function

	public function GetBonusCouponColor()
		if (IsBonusCouponDiscountItem = true) then
			GetBonusCouponColor = "purple"
		else
			GetBonusCouponColor = "black"
		end if
	end function

	public function GetBonusCouponPrice()
		GetBonusCouponPrice = FdiscountAssingedCost
	end function

	public function GetBonusCouponText()
		dim result

		result = ""
		if IsOldJumun then
			if (IsBonusCouponDiscountItem = true) then
				result = result + "���ʽ�����" + vbCrLf
			else
				result = "���󰡰�"
			end if
		else
			if (IsBonusCouponDiscountItem = true) then
				result = result + "���ʽ����� : " + CStr(GetBonusCouponDiscountPrice) + "��" + vbCrLf
			else
				result = "���󰡰�"
			end if
		end if

		GetBonusCouponText = result
	end function

	'==========================================================================
    public function CancelStateStr()
		CancelStateStr = "����"

		if Fcancelyn="Y" then
			CancelStateStr ="���"
		elseif Fcancelyn="D" then
			CancelStateStr ="����"
		elseif Fcancelyn="A" then
			CancelStateStr ="�߰�"
		end if
	end function

	public function CancelStateColor()
		CancelStateColor = "#000000"

		if Fcancelyn="Y" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="D" then
			CancelStateColor ="#FF0000"
		elseif Fcancelyn="A" then
			CancelStateColor ="#0000FF"
		end if
	end function

	''order Detail's State Name : ������
	Public function GetStateName()
        if ForderDetailcurrstate="2" then
            if (Fisupchebeasong="Y") then
		        GetStateName = "��ü�뺸"
		    else
		        GetStateName = "�����뺸"
		    end if
	    elseif ForderDetailcurrstate="3" then
		    GetStateName = "��ǰ�غ�"
	    elseif ForderDetailcurrstate="7" then
		    GetStateName = "���Ϸ�"
	    else
		    GetStateName = ForderDetailcurrstate
	    end if
	end Function

	'' ��Ͻ� ����..
	Public function GetRegDetailStateName()
        if (Fregdetailstate="2") then
            if (Fisupchebeasong="Y") then
		        GetRegDetailStateName = "��ü�뺸"
		    else
		        GetRegDetailStateName = "�����뺸"
		    end if
	    elseif Fregdetailstate="3" then
		    GetRegDetailStateName = "��ǰ�غ�"
	    elseif Fregdetailstate="7" then
		    GetRegDetailStateName = "���Ϸ�"
	    else
		    GetRegDetailStateName = "----"
	    end if
	end Function

	''order Detail's State color
	public function GetStateColor()
	    if ForderDetailcurrstate="2" then
			GetStateColor="#000000"
		elseif ForderDetailcurrstate="3" then
			GetStateColor="#CC9933"
		elseif ForderDetailcurrstate="7" then
			GetStateColor="#FF0000"
		else
			GetStateColor="#000000"
		end if
	end function

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CCSASList
    public FItemList()
    public FOneItem

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectUserID
    public FRectUserName
    public FRectOrderSerial
    public FRectStartDate
    public FRectEndDate
    public FRectSearchType
    public FRectIdx
    public FRectMakerid

    public FRectDivcd
    public FRectCurrstate

    public FRectCsAsID
    public FRectCsRefAsID
    public FRectNotCsID
    ''
    public FDeliverPay
    public IsUpchebeasongExists
    public IsTenbeasongExists

    public FRectOldOrder

    ''��ü���
    public FRectOnlyJupsu
	public FRectOnlyCustomerJupsu
	public FRectOnlyCSServiceRefund
    public FRectShowAX12
    public FRectReceiveYN
    public FRectExcludeB006YN
    public FRectExcludeA004YN
    public FRectExcludeOLDCSYN


	Public FRectDeleteYN	' �������ܿ���
	Public FRectWriteUser	' �����ھ��̵� �˻�

    Public FRectExtSitename

    Public FRectItemID

	public FRectDateType

    public Sub GetCsDetailList()
        dim SqlStr, i

		sqlStr = "select c.*"
		sqlStr = sqlStr + " ,IsNull(d.currstate, '2') as orderdetailcurrstate"
		sqlStr = sqlStr + " ,IsNull(d.reducedprice, 0) as discountAssingedCost, IsNull(d.oitemdiv, i.itemdiv) as oitemdiv, IsNull(d.odlvType, i.deliveryType) as odlvType, d.issailitem, d.itemcouponidx, d.bonuscouponidx"
		sqlStr = sqlStr + " ,IsNULL(d.itemcost,0) as OrderItemcost"
		sqlStr = sqlStr + " ,C2.comm_name as gubun01name, C3.comm_name as gubun02name"
		sqlStr = sqlStr + " ,i.smallimage "

		sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list m "
		sqlStr = sqlStr + " join [db_cs].[dbo].tbl_new_as_detail c "
		sqlStr = sqlStr + " on m.id = c.masterid "
		if (FRectOldOrder="on") then
		    sqlStr = sqlStr + " left join [db_log].[dbo].tbl_old_order_detail_2003 d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d.idx"
		else
		    sqlStr = sqlStr + " left join [db_order].[dbo].tbl_order_detail d"
		    sqlStr = sqlStr + "  on c.orderdetailidx=d.idx"
		end if

		sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item i "
		sqlStr = sqlStr + "  on c.itemid=i.itemid"
		sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C2"
        sqlStr = sqlStr + "  on c.gubun01=C2.comm_cd"
        sqlStr = sqlStr + " Left Join [db_cs].[dbo].tbl_cs_comm_code C3"
        sqlStr = sqlStr + "  on c.gubun02=C3.comm_cd"

		if (FRectCsRefAsID <> "") then
			sqlStr = sqlStr + " where m.refasid=" + CStr(FRectCsRefAsID) + ""
		else
			sqlStr = sqlStr + " where c.masterid=" + CStr(FRectCsAsID) + ""
		end if

        sqlStr = sqlStr + " order by c.isupchebeasong, c.makerid, c.itemid, c.itemoption"
		'response.write sqlStr

		rsget.Open sqlStr,dbget,1

		FTotalCount = rsget.RecordCount
		FResultCount = FTotalCount

		redim preserve FItemList(FResultCount)

		i=0
		do until rsget.eof
			set FItemList(i) = new CCSASDetailItem

            FItemList(i).Fid              = rsget("id")
            FItemList(i).Fmasterid        = rsget("masterid")
            FItemList(i).Fgubun01         = rsget("gubun01")
            FItemList(i).Fgubun02         = rsget("gubun02")
            FItemList(i).Fregitemno       = rsget("regitemno")
            FItemList(i).Fconfirmitemno   = rsget("confirmitemno")

            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''���� ��� ���� ����
            FItemList(i).Forderdetailidx  = rsget("orderdetailidx")
            FItemList(i).Forderserial     = rsget("orderserial")
            FItemList(i).Fitemid          = rsget("itemid")
            FItemList(i).Fitemoption      = rsget("itemoption")
            FItemList(i).Fmakerid         = rsget("makerid")
            FItemList(i).Fitemname        = db2html(rsget("itemname"))
            FItemList(i).Fitemoptionname  = db2html(rsget("itemoptionname"))
            FItemList(i).Fitemcost        = rsget("itemcost")
            FItemList(i).Fbuycash         = rsget("buycash")
            FItemList(i).Fitemno          = rsget("confirmitemno")
            FItemList(i).Forderitemno     = rsget("orderitemno")
            FItemList(i).Fisupchebeasong  = rsget("isupchebeasong")

            FItemList(i).FdiscountAssingedCost = rsget("discountAssingedCost")

            FItemList(i).Foitemdiv        = rsget("oitemdiv")
            FItemList(i).FodlvType        = rsget("odlvType")
            FItemList(i).Fissailitem      = rsget("issailitem")
            FItemList(i).Fitemcouponidx   = rsget("itemcouponidx")
            FItemList(i).Fbonuscouponidx  = rsget("bonuscouponidx")


            FItemList(i).Forderdetailcurrstate  = rsget("orderdetailcurrstate")

            FItemList(i).FSmallImage      = webImgUrl + "/image/small/" + GetImageSubFolderByItemID(FItemList(i).Fitemid) + "/" + rsget("smallimage")

            if (FItemList(i).Fitemid=0) then
                FDeliverPay          = FItemList(i).Fitemcost
            else
                IsUpchebeasongExists = IsUpchebeasongExists or (FItemList(i).Fisupchebeasong="Y")
                IsTenbeasongExists   = IsTenbeasongExists or (FItemList(i).Fisupchebeasong<>"Y")
            end if

            FItemList(i).Fgubun01name   = rsget("gubun01name")
            FItemList(i).Fgubun02name   = rsget("gubun02name")

            if (FItemList(i).Fitemcost=0) then
                FItemList(i).Fitemcost = rsget("OrderItemcost")
            end if

			rsget.movenext
			i=i+1
		loop
		rsget.close

    end Sub

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 10
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub

    Private Sub Class_Terminate()

    End Sub

    public Function HasPreScroll()
            HasPreScroll = StarScrollPage > 1
    end Function

    public Function HasNextScroll()
            HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
    end Function

    public Function StarScrollPage()
            StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function

end Class


'/�������� : [CS]��۰���>>�������Ʈ NEW, ��ü���� : �ֹ� ���� > ����� ����Ʈ	' 2017.12.20 �ѿ��
function SendMiChulgoMailWithMessage(idx, mailmessage)
    ''require /lib/classes/cscenter/oldmisendcls.asp
	''require /lib/classes/order/upchebeasongcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[�ٹ�����] ��� ���� �ȳ������Դϴ�."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
	oneMisend.FRectForMail = "Y"
    oneMisend.getOneOldMisendItem

	'//=======  ���� �߼� =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "�ٹ�����"
		oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "������� �ȳ�����")		' �̸�������

		strMailHTML = replace(strMailHTML,":ORDERSERIAL:",oneMisend.FOneItem.Forderserial)
		strMailHTML = replace(strMailHTML,":ITEMIMAGEURL:",oneMisend.FOneItem.Fsmallimage)
		strMailHTML = replace(strMailHTML,":ITEMID:",oneMisend.FOneItem.Fitemid)
		strMailHTML = replace(strMailHTML,":ITEMNAME:",oneMisend.FOneItem.Fitemname)

		if oneMisend.FOneItem.Fitemoptionname<>"" then
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","["&oneMisend.FOneItem.Fitemoptionname&"]")
		else
			strMailHTML = replace(strMailHTML,":ITEMOPTIONNAME:","")
		end if

		strMailHTML = replace(strMailHTML,":ITEMCNT:",oneMisend.FOneItem.Fitemcnt)
		strMailHTML = replace(strMailHTML,":COMPANYNAME:",oneMisend.FOneItem.getDlvCompanyName)
		strMailHTML = replace(strMailHTML,":MAYSENDDATE:",oneMisend.FOneItem.FMisendipgodate)

		if (oneMisend.FOneItem.FIsUpcheBeasong="Y") then
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�� ������ �ش� �Ǹ��ڰ� ���Բ� �����帮�� �����Դϴ�.<br>*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		else
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*�߼� �����Ϸ� ���� 1-2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.")
		end if

		if (oneMisend.FOneItem.FMisendipgodate<>"") then
			contentsHtml = nl2br(mailmessage)
			contentsHtml = Replace(contentsHtml, "\n", "<br>")

			if (GetMichulgoMailTitleString(oneMisend.FOneItem.FMisendReason) <> "") then
				oMail.MailTitles = GetMichulgoMailTitleString(oneMisend.FOneItem.FMisendReason)
			end if
		end if
		strMailHTML = replace(strMailHTML,":CONTENTSHTML:",contentsHtml)

		oMail.MailConts 	= strMailHTML
		response.write strMailHTML
		'oMail.Send_Mailer()
'		oMail.Send_CDO
	End IF

    ''�޸� ����.
    'contentsHtml = replace(contentsHtml,"�߼ۿ�����","�߼ۿ�����("&oneMisend.FOneItem.FMisendipgodate&")")
	'Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

CLASS MailCls

	dim MailTitles		'���� ����
	dim MailConts		'���� ���� 			(text/html)
	dim SenderMail		'���� �߼��� �ּ� 	(customer@10x10.co.kr,mailzine@10x10.co.kr)
	dim SenderNm		'���� �߼����̸� 	(�ٹ�����)

	dim MailType		'���ø� ��ȣ 		([4],5,6,7,8,9)

	dim ReceiverNm		'���� ������ �̸� 	($1)
	dim ReceiverMail	'���� ������ �ּ� 	(xxxx@aaa.com..)


	dim AddrType				'���ϼ��� ��� (event,userid)
	dim arrUserId 				'AddrType ="userid" �ϰ�� ���

	dim AddrString				'�����ּ� ������ ���� ����
	dim EvtCode,EvtGroupCode 	'AddrType ="event" �ϰ�� ���


	dim strQuery 		'�̸��� ���� ���� ����
	dim EmailDataType	'�̸��� ���� ���� ��� (Enum : string - ���� �Է�,sql - ���� �̿�)
	Dim DB_ID 			'�������� ��񿬰� ��ȣ - ���� (�Ǽ���- 4 ; �׽�Ʈ- 5)


	Private Sub Class_Initialize()
		EvtCode =0
		EvtGroupCode =0
		EmailDataType = "sql"
		MailType = 5

		IF application("Svr_Info")="Dev" THEN
			DB_ID = "5" '//(�Ǽ���- 4 ; �׽�Ʈ- 5)
		ELSE
			DB_ID = "4"
		END IF
		SenderMail	= "mailzine@10x10.co.kr"
		SenderNm	= "�ٹ�����"

	End Sub

	Private Sub Class_Terminate()

	End Sub

	'//+++	���� ���ø� �ҷ����� 	+++//	' 2017.12.20 �ѿ��
	Public Function getMailTemplate()
		dim mFileNm, dfPath, fso,ffso,fnHTML
		dim mailheader, mailfooter

		'/* ���� ���� */
		'// MailType - 5 �̻� ���� ��� (�����ڿ� ����/���� ����! ��.�Ѥ� )
		IF MailType ="5" Then '// �Ļ�������� ��� ����
			mFileNm =""
		ELSEIF MailType="6" Then 		'// �ֹ�����
			mFileNm ="mail_a01.htm"
		ELSEIF MailType ="7" Then '// ����Ȯ��
			mFileNm ="mail_a02.htm"
		ELSEIF MailType ="8" Then '// ������
			'mFileNm = "mail_delivery2011.htm"
			mFileNm ="mail_delivery2017.html"
		ELSEIF MailType ="9" Then '// �������ڵ���Ҿȳ�
			mFileNm ="mail_a04.htm"

		ELSEIF MailType ="10" Then '// ��ŸCS���߼�
			mFileNm ="mail_b01.htm"
		ELSEIF MailType ="11" Then '// �ֹ����(ȯ�Ҿȳ�)
			mFileNm ="mail_b02.htm"
		ELSEIF MailType ="12" Then '// ��ǰ����
			mFileNm ="mail_b03.htm"
		ELSEIF MailType ="13" Then '// ��ǰ�Ϸ�(ȯ�Ҿȳ�)
			mFileNm ="mail_b04.htm"
		ELSEIF MailType ="14" Then '// ȯ��/ī����ҿϷ�
			mFileNm ="mail_b05.htm"

		ELSEIF MailType ="15" Then '// 1:1��� �亯
			'mFileNm ="mail_c01.htm"
			mFileNm ="mail_c01_new.html"
		ELSEIF MailType ="16" Then '// ��ǰQ&A �亯
			mFileNm ="mail_c02.htm"
		ELSEIF MailType ="17" Then '// �Ϲ� ���� ����
			mFileNm ="mail_d01.htm"
		ELSEIF MailType ="18" Then '// ��ǰ���ۼ��ȳ�
			mFileNm ="mail_d02.htm"
		ELSEIF MailType ="19" Then '// ȸ����ް���
			mFileNm ="mail_d03.htm"
		ELSEIF MailType ="20" Then '// �̺�Ʈ��÷����
			mFileNm ="mail_d06.htm"
		ELSEIF MailType ="21" Then '// ��й�ȣ��߼۸���
			mFileNm ="mail_d07.htm"
		ELSEIF MailType ="22" Then '// �����������
			'mFileNm ="mail_misend.htm"
			mFileNm ="email_misend.html"
		End IF

		IF MailType<>"5" and mFileNm="" Then
			response.write "���ø� �ҷ����� ����"
			Exit Function
		End IF

		'//�Ǽ�,�׼�����
		IF application("Svr_Info")="Dev" THEN
			'dfPath = "C:\testweb\admin2009scm\lib\email\mailtemplate" 		'// �׼�(scm)
			dfPath = Server.MapPath("\lib\email\mailtemplate")
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")
			''dfPath = "E:\home\cube1010\admin2009scm\lib\email\mailtemplate" 	'// �Ǽ�(scm)
		END IF

		'/* ���� �ҷ����� */
		IF mFileNm<>"" Then
			Set fso = server.CreateObject("Scripting.FileSystemObject")
				IF fso.FileExists(dfPath & "\" & mFileNm) then
					set ffso = fso.OpenTextFile(dfPath & "\" & mFileNm,1)
					fnHTML = ffso.ReadAll
					ffso.close
					set ffso = nothing
				ELSE
					fnHTML = ""
				End IF
			Set fso = nothing
		End IF

		'/�ű� ������ �������� ���ø� ����� Ǫ�Ϳ� ������ �и���. ���� �ٸ��ǵ鵵 ������� ���� �и��ϰ�, �� �Ϸ� �Ǹ� �б�ó�� ����.
		IF MailType ="22" Then '// �����������
	        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("\lib\email")

	        mFileNm = dfPath&"\\email_header_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailheader = ffso.readall	' ���
		
	        ' ������ �ҷ��ͼ� ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("/lib/email")

	        mFileNm = dfPath&"\\email_footer_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailfooter = ffso.readall	' Ǫ��

			fnHTML = mailheader & fnHTML & mailfooter
		End IF

		getMailTemplate = fnHTML
	End Function

End CLASS

%>