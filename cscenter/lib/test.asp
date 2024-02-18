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

	''// 메일 헤더 이미지
	Public Function getMailHeadImage()
		dim tmpImg
		IF FDivCD="A000" Then '// 맞교환출고
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_exchange_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A001" Then '// 누락재발송
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_resend_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A002" Then '// 서비스발송
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_send_service_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A003" Then '// 환불요청
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_refund_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A004" Then '// 반품접수(업)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_return_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A007" Then '// 신용/이체취소
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_payment_cancel_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A008" Then '// 주문취소
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_order_cancel_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A010" Then '// 회수신청(텐)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_prd_recovery_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A011" Then '// 맞교환회수(텐)
			IF FCurrState="B001" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_cancel_prd_recovery_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSEIF FDivCD="A900" Then '// 주문내역변경
			IF FCurrState="B001" Then
				'tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order.png' alt='CS안내메일' style='vertical-align:top;' />"
			ELSEIF FCurrState="B007" Then
				tmpImg = "<img src='http://mailzine.10x10.co.kr/2017/txt_noti_change_order_comp.png' alt='CS안내메일' style='vertical-align:top;' />"
			End IF
		ELSE

		END IF
		getMailHeadImage = tmpImg
	End Function

	''// 메일 제목	'2017.12.19 한용민 생성
	Public Function getMailHeadtitle()
		dim tmptitle
		IF FDivCD="A000" Then '// 맞교환출고
			IF FCurrState="B001" Then
				tmptitle = "교환출고 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "교환출고 완료 안내메일"
			End IF
		ELSEIF FDivCD="A001" Then '// 누락재발송
			IF FCurrState="B001" Then
				tmptitle = "누락재발송 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "누락재발송 완료 안내메일"
			End IF
		ELSEIF FDivCD="A002" Then '// 서비스발송
			IF FCurrState="B001" Then
				tmptitle = "서비스 발송 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "서비스 발송 완료 안내메일"
			End IF
		ELSEIF FDivCD="A003" Then '// 환불요청
			IF FCurrState="B001" Then
				tmptitle = "환불 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "환불 완료 안내메일"
			End IF
		ELSEIF FDivCD="A004" Then '// 반품접수(업)
			IF FCurrState="B001" Then
				tmptitle = "반품 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "반품 완료 안내메일"
			End IF
		ELSEIF FDivCD="A007" Then '// 신용/이체취소
			IF FCurrState="B001" Then
				tmptitle = "결제취소 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "결제취소 완료 안내메일"
			End IF
		ELSEIF FDivCD="A008" Then '// 주문취소
			IF FCurrState="B001" Then
				'tmptitle = "주문취소 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "주문취소 완료 안내메일"
			End IF
		ELSEIF FDivCD="A010" Then '// 회수신청(텐)
			IF FCurrState="B001" Then
				tmptitle = "상품회수 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "상품회수 완료 안내메일"
			End IF
		ELSEIF FDivCD="A011" Then '// 맞교환회수(텐)
			IF FCurrState="B001" Then
				tmptitle = "교환상품 회수 접수 메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "교환상품 회수 완료 메일"
			End IF
		ELSEIF FDivCD="A900" Then '// 주문내역변경
			IF FCurrState="B001" Then
				'tmptitle = "주문내역변경 접수 안내메일"
			ELSEIF FCurrState="B007" Then
				tmptitle = "주문내역변경 완료 안내메일"
			End IF
		ELSE

		END IF
		getMailHeadtitle = tmptitle
	End Function

	''//접수 상품 정보 가져오기		'/2017.12.19 한용민
	Function getAsItemLIst()
		dim tmpHTML
		dim OCsDetail,i

		tmpHTML = ""

		'A001(누락재발송), A008(주문취소), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송), A004(반품접수(업))
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
			tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left; color:#000;'>접수 상품 정보</th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:50px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""돋움"", sans-serif; text-align:center; color:#707070; font-size:12px; line-height:12px;'>상품</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:100px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>상품코드</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:295px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>상품명[옵션]</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:right; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>판매가격</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:25px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>&nbsp;</th>" & vbcrlf
			tmpHTML=tmpHTML&"								<th style='width:85px; height:44px; margin:0; padding:0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; text-align:center; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>수량</th>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF OCsDetail.FresultCount>0 Then
				FOR i=0 TO OCsDetail.FResultCount-1
				    IF (OCsDetail.FItemList(i).Fitemid<>0) or (OCsDetail.FItemList(i).Fitemcost<>0) then
						tmpHTML=tmpHTML&"							<tr>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:50px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea;'>" & vbcrlf
						tmpHTML=tmpHTML&"									<img src='"& OCsDetail.FItemList(i).FSmallImage &"' alt='' />" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:100px;margin:0;  padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-size:11px; line-height:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
						tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemid &"" & vbcrlf
						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:295px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:left; font-size:11px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).Fitemid=0) Then
							tmpHTML=tmpHTML&"									배송비" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									"& OCsDetail.FItemList(i).Fitemname &"" & vbcrlf
						END IF
						if ( OCsDetail.FItemList(i).Fitemoptionname <>"") then
							tmpHTML=tmpHTML&"									["& OCsDetail.FItemList(i).Fitemoptionname &"]" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf

						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; font-size:12px; text-align:right; font-family:dotum, ""돋움"", sans-serif;'>" & vbcrlf

						IF (OCsDetail.FItemList(i).FdiscountAssingedCost<>0) and (OCsDetail.FItemList(i).Fitemcost>OCsDetail.FItemList(i).FdiscountAssingedCost) then
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:6px 0; font-size:11px; line-height:16px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-decoration:line-through; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) & "원</span><br />" & vbcrlf
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>" & FormatNumber(OCsDetail.FItemList(i).FdiscountAssingedCost,0) &"원</span>" & vbcrlf
						ELSE
							tmpHTML=tmpHTML&"									<span style='margin:0; padding:0; font-weight:bold; font-size:12px; line-height:17px; color:#707070; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>"& FormatNumber(OCsDetail.FItemList(i).Fitemcost,0) &"원</span>" & vbcrlf
						END IF

						tmpHTML=tmpHTML&"								</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:25px; padding:6px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
						tmpHTML=tmpHTML&"								<td style='width:85px; margin:0; padding:6px 0; border-bottom:solid 1px #eaeaea; text-align:center; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; color:#707070; font-size:12px; line-height:12px;'>"& OCsDetail.FItemList(i).Fregitemno &"</td>" & vbcrlf
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

	''//고객주소 가져오기		'/2017.12.19 한용민
	Function getReqInfo()
		dim tmpHTML
		tmpHTML=""

		'A001(누락재발송), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송)
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A010" THEN 'or FDivCD="A011"
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>고객주소</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& AstarUserName(trim(FReqName)) &" 고객님 &nbsp; &nbsp;"& AstarPhoneNumber(trim(FReqPhone)) &" / "& AstarPhoneNumber(trim(FReqHP)) &" <br />["& printUserId(trim(FReqZipcode), 2, "*") &"] "& printUserId(trim(FReqZipAddr), 2, "*") &"&nbsp;(이하생략)" & vbcrlf		' FReqEtcAddr
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF
		getReqInfo = tmpHTML
	END Function

	''//업체 주소 가져오기		'/2017.12.19 한용민
	Function getReturnInfo()
		dim tmpHTML
		tmpHTML=""

		' A011(맞교환회수(텐)), A010(회수신청(텐)), A004(반품접수(업))
		IF FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>반품회수주소</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnName &" &nbsp; &nbsp;"& FReturnPhone &"<br />["& FReturnZipCode &"] "& FReturnZipAddr &"&nbsp;"& FReturnEtcAddr &"" & vbcrlf
			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			if (FReturnName<>"(주)텐바이텐") and (FupcheReturnSongjangDivName<>"") and (Left(FupcheReturnSongjangDivTel,1)="1" or Left(FupcheReturnSongjangDivTel,1)="0") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>이용택배사</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
				tmpHTML=tmpHTML&"									"& FupcheReturnSongjangDivName &"<br />택배사연락처 : "& FupcheReturnSongjangDivTel &"" & vbcrlf
				tmpHTML=tmpHTML&"								</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			END IF
		END IF

		getReturnInfo = tmpHTML
	END Function

	''//환불정보 가져오기		'/2017.12.19 한용민
	Function getRefundInfo()
		dim tmpHTML
		tmpHTML=""

		' A008(주문취소), A010(회수신청(텐)), A007(신용/이체취소), A003(환불요청), A004(반품접수(업))
		IF FDivCD="A003" or FDivCD="A004" or FDivCD="A007" or FDivCD="A008" or FDivCD="A010" THEN
		    ''환불액0이면 return
		    if (FRefundRequire=0) then Exit function

		    ''부정확한 환불 정보 제거
		    if (FReturnMethod="R007") then
		        if (Len(Replace(FReBankAccount,"-",""))<7) then
    		        FReBankName = ""
    		        FReBankAccount = "계좌확인요망"
    		        FReBankOwnerName =""
    		    else
    		        FReBankAccount = Left(FReBankAccount,Len(Trim(FReBankAccount))-3) + "***"
    		    end if
		    end if

			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>환불예정액</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FormatNumber(FRefundRequire,0) &" 원" & vbcrlf

			'배송비차감 안내가 오해를 일으켜 표시안하게 수정
			'if (FRefundDeliveryPay<>0) then
			'			tmpHTML=tmpHTML&"									(배송비차감 : " & FormatNumber(FRefundDeliveryPay+Frefundbeasongpay,0) &")" & vbcrlf
			'end if

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>환불정보(계좌)</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
			tmpHTML=tmpHTML&"									"& FReturnMethodName &"&nbsp;&nbsp;" & vbcrlf
	
			IF (FReturnMethod="R007") THEN
				tmpHTML=tmpHTML&"									"& FReBankName &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& FReBankAccount &"&nbsp;&nbsp; " & vbcrlf
				tmpHTML=tmpHTML&"									"& AstarUserName(FReBankOwnerName) &" " & vbcrlf
			ELSEIF (FReturnMethod="R900") THEN
				tmpHTML=tmpHTML&"									(적립아이디 : "& FUserID &") " & vbcrlf
			ELSEIF (FReturnMethod="R100") or (FReturnMethod="R550") or (FReturnMethod="R560") or (FReturnMethod="R120") or (FReturnMethod="R020") or (FReturnMethod="R080") THEN
				if (Left(FPayGateTid,6)="IniTec") and (FCurrState="B007") and (FReturnMethod<>"R120") then
					tmpHTML=tmpHTML&"									<a target='_blank' href=https://iniweb.inicis.com/DefaultWebApp/mall/cr/cm/mCmReceipt_head.jsp?noTid="& FPayGateTid &"&noMethod=1>[매출전표출력]</a> " & vbcrlf
				end if
				if (FReturnMethod = "R550") or (FReturnMethod = "R560") then
					tmpHTML=tmpHTML&"									기프팅/기프티콘 은 발행시 증빙서류가 발급되며, 상품구매시에는 발행불가입니다." & vbcrlf
				end if
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		END IF
		getRefundInfo = tmpHTML
	END Function

	'// 처리 결과 가져오기		'/2017.12.19 한용민
	Function getFinishResult()
		dim tmpHTML
		tmpHTML=""

		'A001(누락재발송), A011(맞교환회수(텐)), A010(회수신청(텐)), A003(환불요청), A002(서비스발송), A004(반품접수(업)), A000(맞교환출고)
		IF FCurrState="B007" THEN
		    ''처리 내역이 없을때..
		    if (FOpenContents="") then
		        if (FDivCD="A000") then
		            FOpenContents = "맞교환상품 출고완료"
		        elseif (FDivCD="A001") then
		            FOpenContents = "누락상품 출고완료"
		        elseif (FDivCD="A002") then
		            FOpenContents = "상품 출고완료"
		        elseif (FDivCD="A003") then

		        elseif (FDivCD="A004") then
		            FOpenContents = "상품 반품(회수)완료" '' / 환불등록"

		        elseif (FDivCD="A010") then
		            FOpenContents = "상품 회수완료" '' / 환불등록"
		        elseif (FDivCD="A011") then
		            FOpenContents = "맞교환상품 회수완료"
		        else

		        end if
		    end if

			tmpHTML=tmpHTML&"	<tr>" & vbcrlf
			tmpHTML=tmpHTML&"		<td style='padding:45px 29px; margin:0;'>" & vbcrlf
			tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<th style='margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; text-align:left; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>처리결과 정보<th>" & vbcrlf
			tmpHTML=tmpHTML&"				</tr>" & vbcrlf
			tmpHTML=tmpHTML&"				<tr>" & vbcrlf
			tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
			tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>처리완료일</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FFinishDate &"</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf

			IF (Trim(FOpenContents)<>"") then
				tmpHTML=tmpHTML&"							<tr>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>처리내용</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
				tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:16px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& nl2br(FOpenContents) &"</td>" & vbcrlf
				tmpHTML=tmpHTML&"							</tr>" & vbcrlf
			end IF

			''// 택배정보 가져오기
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

	''// 택배 정보 가져오기		'/2017.12.19 한용민
	Function getDlvInfo()
		dim tmpHTML
		tmpHTML=""

        if (IsNULL(FSongjangNo)) or (FSongjangNo="") then Exit function

		'A001(누락재발송), A011(맞교환회수(텐)), A000(맞교환출고), A010(회수신청(텐)), A002(서비스발송), A004(반품접수(업))
		IF FDivCD="A000" or FDivCD="A001" or FDivCD="A002" or FDivCD="A004" or FDivCD="A010" or FDivCD="A011" THEN
			tmpHTML=tmpHTML&"							<tr>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>택배정보</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
			tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

			IF FSongjangNo<>"" then
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>"& FSongjangDivName &"</span>"& vbcrlf
				tmpHTML=tmpHTML&"									&nbsp;&nbsp;<a href='"& DeliverDivTrace(Trim(FSongjangDiv)) & FSongjangNo &"' target='_blank' style='margin:0; padding:0; font-size:12px; color:#dd5555; font-size:11px; line-height:18px; color:#0066cc; text-align:left;'>"& FSongjangNo &"</a>" & vbcrlf
			ELSE
				tmpHTML=tmpHTML&"									<span style='margin:0; text-align:left; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>택배정보가 등록되지 않았습니다.</span>" & vbcrlf
			END IF

			tmpHTML=tmpHTML&"								</td>" & vbcrlf
			tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		END IF

		getDlvInfo =  tmpHTML
	END Function

	'// 기타 안내사항		'/2017.12.19 한용민
	Public Function getEtcNotice()
		dim tmpHTML

        getEtcNotice = ""

        if (Trim(FInfoHtml)="") then Exit function

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left;'>기타 안내 정보</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:10px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; vertical-align:top; text-align:left;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:630px; margin:0; padding-top:14px; font-size:12px; line-height:19px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
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

	''// 접수 기본 내용 가져오기		'/2017.12.19 한용민
	Function getAsInfo()
		dim tmpHTML

		tmpHTML = ""
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수일</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FRegDate &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>서비스코드</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FAsID &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>주문번호</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& FOrderSerial &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수내용</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>" & vbcrlf
		tmpHTML=tmpHTML&"									"& FTitle &" <a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp?orderserial=" & FOrderSerial & "' target='_blank'>[상세내역보기]</a>" & vbcrlf
		tmpHTML=tmpHTML&"								</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
		tmpHTML=tmpHTML&"							<tr>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:110px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; background:#f8f8f8; font-weight:bold; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>접수사유</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:30px; padding:11px 0; border-bottom:solid 1px #eaeaea;'>&nbsp;</td>" & vbcrlf
		tmpHTML=tmpHTML&"								<td style='width:470px; margin:0; padding:11px 0; border-bottom:solid 1px #eaeaea; font-size:12px; line-height:20px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:left;'>"& GetCauseDetailString &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"							</tr>" & vbcrlf
				
		getAsInfo =tmpHTML
	END Function

	'/ 이메일 탬플릿 가져와서 만드는걸로 생성.	2017.12.18 한용민 생성
	Function makeMailTemplate_GiftCard(id)
		dim tmpHTML, fs, dirPath, fileName, objFile, mailheader, mailfooter

		Call GetOneCSASMaster_GiftCard(id) '// 값 세팅

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_header_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailheader = objFile.readall	' 헤더

		tmpHTML=mailheader

		'A001(누락재발송), A008(주문취소), A000(맞교환출고)
		tmpHTML=tmpHTML&" <table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='height:253px; text-align:center;'>"& getMailHeadImage &"</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:20px 0 45px; font-size:12px; line-height:22px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>" & vbcrlf

		if (FDivCD = "A008") then
			tmpHTML=tmpHTML&"			고객님의 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		elseif (FDivCD = "A000") then
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"가 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		else
			tmpHTML=tmpHTML&"			고객님이 요청하신 "& GetAsDivCDName &"이 정상적으로 처리가 "& FCurrStateName &" 되었습니다.<br />텐바이텐을 이용해주셔서 감사합니다." & vbcrlf		'Fcustomername
		end if

		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:0 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%; background:#f8f8f8; text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:297px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; text-align:right;'>주문번호 :</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:7px; padding:34px 0;'></td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='width:331px; margin:0; padding:34px 0; font-size:22px; line-height:22px; font-weight:bold; font-family:dotum, ""돋움"", sans-serif; color:#dd5555; text-align:left; letter-spacing:-1px;'>" & vbcrlf
		tmpHTML=tmpHTML&"						"& FOrderSerial &"" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 처리결과 가져오기
		tmpHTML=tmpHTML& getFinishResult()

		''// 접수 상품 정보 가져오기
		'tmpHTML=tmpHTML& getAsItemLIst()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45 29px; margin:0;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<th style='width:100%; margin:0; padding:0 0 15px 3px; font-size:17px; line-height:17px; font-family:dotum, ""돋움"", sans-serif; text-align:left; color:#000;'>접수 정보</th>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td style='border-top:solid 2px #000;'>" & vbcrlf
		tmpHTML=tmpHTML&"						<table border='0' cellpadding='0' cellspacing='0' style='width:100%; font-size:11px; font-family:dotum, ""돋움"", sans-serif; color:#707070;'>" & vbcrlf

		''// 접수 기본 내용 가져오기
		tmpHTML=tmpHTML& getAsInfo()

		''// 고객주소 가져오기
		'tmpHTML=tmpHTML& getReqInfo()

		''// 업체주소 가져오기
		'tmpHTML=tmpHTML& getReturnInfo()

		''// 환불정보 가져오기
		tmpHTML=tmpHTML& getRefundInfo()

		tmpHTML=tmpHTML&"						</table>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf

		''// 기타 안내사항
		'tmpHTML=tmpHTML& getEtcNotice()

		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='padding:45px 104px; margin:0;text-align:center;'>" & vbcrlf
		tmpHTML=tmpHTML&"			<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbcrlf
		tmpHTML=tmpHTML&"				<tr>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/my10x10/order/order_cslist.asp' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_receiption_info.png' alt='접수 정보 상세보기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"					<td>" & vbcrlf
		tmpHTML=tmpHTML&"						<a href='http://www.10x10.co.kr/' target='_blank'><img src='http://mailzine.10x10.co.kr/2017/btn_go_shopping.png' alt='텐바이텐 쇼핑하기' style='vertical-align:top; border:0;' /></a>" & vbcrlf
		tmpHTML=tmpHTML&"					</td>" & vbcrlf
		tmpHTML=tmpHTML&"				</tr>" & vbcrlf
		tmpHTML=tmpHTML&"			</table>" & vbcrlf
		tmpHTML=tmpHTML&"		</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"	<tr>" & vbcrlf
		tmpHTML=tmpHTML&"		<td style='margin:0; padding:25px 0; border-top:solid 1px #eaeaea; font-size:12px; line-height:12px; font-family:dotum, ""돋움"", sans-serif; color:#707070; text-align:center;'>끝까지 기분 좋은 쇼핑이 될 수 있도록 최선을 다하겠습니다.</td>" & vbcrlf
		tmpHTML=tmpHTML&"	</tr>" & vbcrlf
		tmpHTML=tmpHTML&"</table>" & vbcrlf

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/lib/email")

        fileName = dirPath&"\\email_footer_1.html"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailfooter = objFile.readall	' 푸터

		tmpHTML=tmpHTML&mailfooter

		tmpHTML = replace(tmpHTML,":mailtitle:", getMailHeadtitle())		' 이메일제목

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

    Dim FRectForceCurrState     ''상태값 강제 지정.
    Dim FRectForceBuyEmail      ''이메일 강제지정.
    
    Dim Faccountdiv      ''2016/08/05 추가
    Dim Fpggubun         ''2016/08/05 추가
    
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
				" 	,isNull(p.company_name,'(주)텐바이텐') as ReturnName ,isNull(p.deliver_phone,'1644-6030') as ReturnPhone ,isNull(p.deliver_hp,'') as ReturnHP "&_
				" 	,isNull(p.return_zipcode,'132-010') as ReturnZipCode ,isNull(p.return_address,'서울 도봉구') as ReturnZipAddr ,isNull(p.return_address2,'도봉동 63번지 여인닷컴 3층') as ReturnEtcAddr "&_
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

				IF (FRectForceCurrState<>"") then  ''상태값 강제 지정 (메일 재발송시 사용.)
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
                        FReturnMethodName = "네이버페이취소"
                    end if
                end if
			END IF
		rsget.close

		''기타 안내 사항
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

				IF (FRectForceCurrState<>"") then  ''상태값 강제 지정 (메일 재발송시 사용.)
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

		''기타 안내 사항
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
    public FdiscountAssingedCost    '' 주문시 할인된가격 ( ALL@ / %할인권 반영)

    public Forgitemcost					'소비자가
    public FitemcostCouponNotApplied	'판매가(할인가)
    public FplusSaleDiscount			'플러스세일할인액
    public FspecialshopDiscount			'우수고객할인액

    public Forgprice					'현재소비자가(+옵션가)

	public Fprevcsreturnfinishno		'이전 CS반품수량(접수이상)

	public Freforderdetailidx

	Public Fsongjangdiv
	Public Fsongjangno

    ''public FAllAtDiscountedPrice

    ''tbl_item's
    public FSmallImage

    ''업체 개별배송 상품 배송비 인지 여부
    public function IsUpcheParticleDeliverPayCodeItem
        IsUpcheParticleDeliverPayCodeItem = (Fitemid=0) and (Left(Fitemoption,2)="90")
    end function

    ''업체 개별배송 상품인지 여부
    public function IsUpcheParticleDeliverItem
        IsUpcheParticleDeliverItem = (FodlvType=9)
    end function

    ''반품시 사용하는 상품가격(All@ 할인값, %쿠폰 할인값 반영)
    public function GetOrgPayedItemPrice()
        GetOrgPayedItemPrice = Fitemcost

        if (FdiscountAssingedCost=0) then
            ''기존방식
            GetOrgPayedItemPrice = Fitemcost-getAllAtDiscountedPrice
        else
            if (FdiscountAssingedCost<>Fitemcost) then
                GetOrgPayedItemPrice = FdiscountAssingedCost
            end if
        end if
    end function

    ''All@ 할인된가격
    public function getAllAtDiscountedPrice()
        getAllAtDiscountedPrice =0
        ''기존 상품쿠폰 할인되는경우 추가할인없음.
        ''마일리지SHOP 상품 추가 할인 없음.
	    ''세일상품 추가할인 없음
	    '' 20070901추가 : 정율할인 보너스쿠폰사용시 추가할인 없음.

'	    if (FdiscountAssingedCost=0) then
'	        ''기존방식
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

    '' %할인권 할인금액 or 카드 할인금액
    public function getPercentBonusCouponDiscountedPrice()
        getPercentBonusCouponDiscountedPrice = 0
'        if (Fitemcost>FdiscountAssingedCost) then
'                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
'        end if

		if (Fitemid = 0) and (Fitemcost > FdiscountAssingedCost) and not IsNull(Fbonuscouponidx) then
			'// 배송비 쿠폰
			getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
        ''elseif (FdiscountAssingedCost=0) then
	        ''기존방식
	    ''    ''getPercentBonusCouponDiscountedPrice = Fitemcost*
		else
			'// 전액 할인쿠폰 생김(2014-06-23, skyer9)
            if (Fbonuscouponidx<>0)  and (Fitemcost>FdiscountAssingedCost) then
                getPercentBonusCouponDiscountedPrice = Fitemcost-FdiscountAssingedCost
            end if
        end if
    end function

    ''마일리지샵 상품
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

    ''CsAction 접수시 상품 갯수 수정 가능여부
    public function IsItemNoEditEnabled(byval idivcd)
        IsItemNoEditEnabled = false

        if (Fcancelyn="Y") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsItemNoEditEnabled = true

            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''반품 접수
            if (ForderDetailcurrstate>=7) then IsItemNoEditEnabled=true

        elseif (fnIsServiceDeliverProcess(idivcd)) or (fnIsServiceRecvProcess(idivcd)) then
            '서비스 - 항상 갯수 수정 가능
            if (idivcd = "A002") or (idivcd = "A200") then
            	IsItemNoEditEnabled=true

            elseif (ForderDetailcurrstate>=7) then
            	IsItemNoEditEnabled=true

            end if
        end if
    end function


    ''CsAction 접수시 상품별 체크 가능여부
    public function IsCheckAvailItem(byval iIpkumdiv, byval iMasterCancelYn, byval idivcd)
        IsCheckAvailItem = false

        if (Fcancelyn="Y") then Exit function
        if (iMasterCancelYn<>"N") then Exit function

        if (fnIsCancelProcess(idivcd)) then
            IsCheckAvailItem = true
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false

        elseif (fnIsReturnProcess(idivcd)) then
            ''반품 접수
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true

            if (FItemId=0) then IsCheckAvailItem=true
        elseif (idivcd="A006") then
            ''출고시 유의사항
            IsCheckAvailItem=true

            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=false
        elseif (idivcd="A009") then
            ''기타사항(메모) - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd="A700") then
            ''기타정산 - All case Avail
            IsCheckAvailItem=true
        elseif (idivcd = "A002") or (idivcd = "A200") then
        	'서비스 - 항상 체크가능
            if Fitemid=0 then
                IsCheckAvailItem=false
            else
                IsCheckAvailItem=true
            end if
        elseif (idivcd = "A001") then
            ''누락
            if (ForderDetailcurrstate>=7) or ((Fcancelyn="A") and (iIpkumdiv>=7)) then IsCheckAvailItem=true
        elseif (idivcd = "A000") then
            ''맞교환
            if (ForderDetailcurrstate>=7) then IsCheckAvailItem=true
        else

        end if
    end function

    ''CsAction 접수시 상품별 디폴트 체크드
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
            ''반품접수인경우 - No action
        elseif (idivcd="A006") then
            ''출고시 유의사항 - No action
        elseif (idivcd="A009") then
            ''기타사항(메모) - No action
        else

        end if
    end function

	'==========================================================================
    '보너스쿠폰 적용 주문인지 체크
    public function IsBonusCouponDiscountItem()
        IsBonusCouponDiscountItem = false
        if (Not IsNull(Fbonuscouponidx) and (Fbonuscouponidx<>0))  then
            IsBonusCouponDiscountItem = true
        end if
    end function

	'상품쿠폰 적용 주문인지 체크
    public function IsItemCouponDiscountItem()
        IsItemCouponDiscountItem = false
        if (Not IsNull(Fitemcouponidx) and (Fitemcouponidx<>0)) then
            IsItemCouponDiscountItem = true
        end if
    end function

    '우수고객할인 적용 주문인지 체크
    public function IsSpecialShopDiscountItem()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (Not IsItemCouponDiscountItem) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : 소비자가변경, 옵션가변경이 있는경우 부정확한 값이 된다.
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

	'상품쿠폰할인액
    public function GetItemCouponDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (IsItemCouponDiscountItem = true) and (Not IsBonusCouponDiscountItem) and (Fissailitem = "N") then
        		'TODO : 소비자가변경, 옵션가변경, 우수고객할인이 있는경우 부정확한 값이 된다.
        		GetItemCouponDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetItemCouponDiscountPrice = 0
        	exit function
        end if

        GetItemCouponDiscountPrice = FitemcostCouponNotApplied - Fitemcost
    end function

	'보너스쿠폰할인액
    public function GetBonusCouponDiscountPrice()
        GetBonusCouponDiscountPrice = Fitemcost - FdiscountAssingedCost
    end function

	'상품할인액
    public function GetSaleDiscountPrice()
        if (FitemcostCouponNotApplied = 0) then
        	'과거데이타
        	if (Not IsBonusCouponDiscountItem) and (Not IsItemCouponDiscountItem) and (Fissailitem = "Y") then
        		'TODO : 소비자가변경, 옵션가변경, 우수고객할인이 있는경우 부정확한 값이 된다.
        		GetSaleDiscountPrice = Forgprice - Fitemcost
        		exit function
        	end if

        	GetSaleDiscountPrice = 0
        	exit function
        end if

        GetSaleDiscountPrice = (Forgitemcost - (FitemcostCouponNotApplied + FplusSaleDiscount + FspecialshopDiscount))
    end function

    public function IsOldJumun()
    	'2011년 4월 1일 이전 주문 또는 그 주문에 대한 마이너스주문
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
						result = result + "할인상품 + 소비자가 인하" + vbCrLf
					else
						result = result + "할인상품" + vbCrLf
					end if
				end if
				if (Fissailitem = "P") then
					result = result + "플러스할인" + vbCrLf
				end if
				if ((Fissailitem = "N") and (Not IsItemCouponDiscountItem) and (Forgprice <> Fitemcost)) then
					result = result + "우수고객할인 또는 소비자가/옵션가 변동" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		else
			if (Forgitemcost <> FitemcostCouponNotApplied) then
				if (Fissailitem = "Y") then
					result = result + "할인상품 : " + CStr(GetSaleDiscountPrice) + "원" + vbCrLf
				end if
				if (FplusSaleDiscount > 0) then
					result = result + "플러스할인 : " + CStr(FplusSaleDiscount) + "원" + vbCrLf
				end if
				if (FspecialshopDiscount > 0) then
					result = result + "우수회원할인 : " + CStr(FspecialshopDiscount) + "원" + vbCrLf
				end if
			else
				result = "정상가격"
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
					result = result + "상품쿠폰적용상품" + vbCrLf
				else
					result = result + "배송비쿠폰적용상품" + vbCrLf
				end if
			else
				result = "정상가격"
			end if
		else
			if (IsItemCouponDiscountItem = true) then
				if (GetItemCouponDiscountPrice = 0) then
					result = result + "배송비쿠폰적용상품" + vbCrLf
				else
					result = result + "상품쿠폰 : " + CStr(GetItemCouponDiscountPrice) + "원" + vbCrLf
				end if
			else
				result = "정상가격"
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
				result = result + "보너스쿠폰" + vbCrLf
			else
				result = "정상가격"
			end if
		else
			if (IsBonusCouponDiscountItem = true) then
				result = result + "보너스쿠폰 : " + CStr(GetBonusCouponDiscountPrice) + "원" + vbCrLf
			else
				result = "정상가격"
			end if
		end if

		GetBonusCouponText = result
	end function

	'==========================================================================
    public function CancelStateStr()
		CancelStateStr = "정상"

		if Fcancelyn="Y" then
			CancelStateStr ="취소"
		elseif Fcancelyn="D" then
			CancelStateStr ="삭제"
		elseif Fcancelyn="A" then
			CancelStateStr ="추가"
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

	''order Detail's State Name : 현상태
	Public function GetStateName()
        if ForderDetailcurrstate="2" then
            if (Fisupchebeasong="Y") then
		        GetStateName = "업체통보"
		    else
		        GetStateName = "물류통보"
		    end if
	    elseif ForderDetailcurrstate="3" then
		    GetStateName = "상품준비"
	    elseif ForderDetailcurrstate="7" then
		    GetStateName = "출고완료"
	    else
		    GetStateName = ForderDetailcurrstate
	    end if
	end Function

	'' 등록시 상태..
	Public function GetRegDetailStateName()
        if (Fregdetailstate="2") then
            if (Fisupchebeasong="Y") then
		        GetRegDetailStateName = "업체통보"
		    else
		        GetRegDetailStateName = "물류통보"
		    end if
	    elseif Fregdetailstate="3" then
		    GetRegDetailStateName = "상품준비"
	    elseif Fregdetailstate="7" then
		    GetRegDetailStateName = "출고완료"
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

    ''업체사용
    public FRectOnlyJupsu
	public FRectOnlyCustomerJupsu
	public FRectOnlyCSServiceRefund
    public FRectShowAX12
    public FRectReceiveYN
    public FRectExcludeB006YN
    public FRectExcludeA004YN
    public FRectExcludeOLDCSYN


	Public FRectDeleteYN	' 삭제제외여부
	Public FRectWriteUser	' 접수자아이디 검색

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

            FItemList(i).Fregdetailstate  = rsget("regdetailstate")   ''접수 당시 진행 상태
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


'/직원어드민 : [CS]배송관리>>미출고리스트 NEW, 업체어드민 : 주문 관리 > 미출고 리스트	' 2017.12.20 한용민
function SendMiChulgoMailWithMessage(idx, mailmessage)
    ''require /lib/classes/cscenter/oldmisendcls.asp
	''require /lib/classes/order/upchebeasongcls.asp
    dim oneMisend
    dim strMailHTML,strMailTitle, contentsHtml
	strMailHTML = ""
	strMailTitle = "[텐바이텐] 출고 지연 안내메일입니다."

    set oneMisend = new COldMiSend
    oneMisend.FRectDetailIDx = idx
	oneMisend.FRectForMail = "Y"
    oneMisend.getOneOldMisendItem

	'//=======  메일 발송 =========/
	dim oMail
	dim MailHTML

	set oMail = New MailCls         '' mailLib2

	IF oneMisend.FOneItem.Fbuyemail<>"" THEN

		oMail.MailTitles	= strMailTitle
		oMail.SenderNm		= "텐바이텐"
		oMail.SenderMail	= "mailzine@10x10.co.kr"
		oMail.AddrType		= "string"
		oMail.ReceiverNm	= oneMisend.FOneItem.FBuyname
		oMail.ReceiverMail	= oneMisend.FOneItem.FBuyEmail
		oMail.MailType = "22"
		strMailHTML = oMail.getMailTemplate
		''parsing
		strMailHTML = replace(strMailHTML,":mailtitle:", "출고지연 안내메일")		' 이메일제목

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
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*본 메일은 해당 판매자가 고객님께 보내드리는 메일입니다.<br>*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
		else
	        strMailHTML = replace(strMailHTML,":BOTTOMMSG:","*발송 예정일로 부터 1-2일 후에 상품을 받아보실 수 있습니다.")
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

    ''메모에 저장.
    'contentsHtml = replace(contentsHtml,"발송예정일","발송예정일("&oneMisend.FOneItem.FMisendipgodate&")")
	'Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[Mail]" + strMailTitle + VbCrlf + contentsHtml)

	SET oMail = nothing
	set oneMisend = Nothing
end function

CLASS MailCls

	dim MailTitles		'메일 제목
	dim MailConts		'메일 내용 			(text/html)
	dim SenderMail		'메일 발송자 주소 	(customer@10x10.co.kr,mailzine@10x10.co.kr)
	dim SenderNm		'메일 발송자이름 	(텐바이텐)

	dim MailType		'템플릿 번호 		([4],5,6,7,8,9)

	dim ReceiverNm		'메일 수신자 이름 	($1)
	dim ReceiverMail	'메일 수신자 주소 	(xxxx@aaa.com..)


	dim AddrType				'메일수집 방식 (event,userid)
	dim arrUserId 				'AddrType ="userid" 일경우 사용

	dim AddrString				'메일주소 수집에 쓰일 정보
	dim EvtCode,EvtGroupCode 	'AddrType ="event" 일경우 사용


	dim strQuery 		'이메일 정보 수집 쿼리
	dim EmailDataType	'이메일 정보 수집 방식 (Enum : string - 직접 입력,sql - 쿼리 이용)
	Dim DB_ID 			'선더메일 디비연결 번호 - 고정 (실서버- 4 ; 테스트- 5)


	Private Sub Class_Initialize()
		EvtCode =0
		EvtGroupCode =0
		EmailDataType = "sql"
		MailType = 5

		IF application("Svr_Info")="Dev" THEN
			DB_ID = "5" '//(실서버- 4 ; 테스트- 5)
		ELSE
			DB_ID = "4"
		END IF
		SenderMail	= "mailzine@10x10.co.kr"
		SenderNm	= "텐바이텐"

	End Sub

	Private Sub Class_Terminate()

	End Sub

	'//+++	메일 템플릿 불러오기 	+++//	' 2017.12.20 한용민
	Public Function getMailTemplate()
		dim mFileNm, dfPath, fso,ffso,fnHTML
		dim mailheader, mailfooter

		'/* 파일 선택 */
		'// MailType - 5 이상 실제 사용 (관계자외 접근/수정 금지! ㅡ.ㅡㅋ )
		IF MailType ="5" Then '// 식사용자정의 양식 메일
			mFileNm =""
		ELSEIF MailType="6" Then 		'// 주문접수
			mFileNm ="mail_a01.htm"
		ELSEIF MailType ="7" Then '// 결제확인
			mFileNm ="mail_a02.htm"
		ELSEIF MailType ="8" Then '// 출고메일
			'mFileNm = "mail_delivery2011.htm"
			mFileNm ="mail_delivery2017.html"
		ELSEIF MailType ="9" Then '// 무통장자동취소안내
			mFileNm ="mail_a04.htm"

		ELSEIF MailType ="10" Then '// 기타CS출고발송
			mFileNm ="mail_b01.htm"
		ELSEIF MailType ="11" Then '// 주문취소(환불안내)
			mFileNm ="mail_b02.htm"
		ELSEIF MailType ="12" Then '// 반품접수
			mFileNm ="mail_b03.htm"
		ELSEIF MailType ="13" Then '// 반품완료(환불안내)
			mFileNm ="mail_b04.htm"
		ELSEIF MailType ="14" Then '// 환불/카드취소완료
			mFileNm ="mail_b05.htm"

		ELSEIF MailType ="15" Then '// 1:1상담 답변
			'mFileNm ="mail_c01.htm"
			mFileNm ="mail_c01_new.html"
		ELSEIF MailType ="16" Then '// 상품Q&A 답변
			mFileNm ="mail_c02.htm"
		ELSEIF MailType ="17" Then '// 일반 공지 메일
			mFileNm ="mail_d01.htm"
		ELSEIF MailType ="18" Then '// 상품평작성안내
			mFileNm ="mail_d02.htm"
		ELSEIF MailType ="19" Then '// 회원등급공지
			mFileNm ="mail_d03.htm"
		ELSEIF MailType ="20" Then '// 이벤트당첨공지
			mFileNm ="mail_d06.htm"
		ELSEIF MailType ="21" Then '// 비밀번호재발송메일
			mFileNm ="mail_d07.htm"
		ELSEIF MailType ="22" Then '// 출고지연메일
			'mFileNm ="mail_misend.htm"
			mFileNm ="email_misend.html"
		End IF

		IF MailType<>"5" and mFileNm="" Then
			response.write "템플릿 불러오기 실패"
			Exit Function
		End IF

		'//실섭,테섭구분
		IF application("Svr_Info")="Dev" THEN
			'dfPath = "C:\testweb\admin2009scm\lib\email\mailtemplate" 		'// 테섭(scm)
			dfPath = Server.MapPath("\lib\email\mailtemplate")
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")
			''dfPath = "E:\home\cube1010\admin2009scm\lib\email\mailtemplate" 	'// 실섭(scm)
		END IF

		'/* 파일 불러오기 */
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

		'/신규 리뉴얼 버전부터 탬플릿 헤더와 푸터와 내용을 분리함. 차후 다른건들도 리뉴얼시 전부 분리하고, 다 완료 되면 분기처리 뺄껏.
		IF MailType ="22" Then '// 출고지연메일
	        ' 파일을 불러와서 ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("\lib\email")

	        mFileNm = dfPath&"\\email_header_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailheader = ffso.readall	' 헤더
		
	        ' 파일을 불러와서 ---------------------------------------------------------------------------
	        Set fso = Server.CreateObject("Scripting.FileSystemObject")
	        dfPath = server.mappath("/lib/email")

	        mFileNm = dfPath&"\\email_footer_1.html"

	        Set ffso = fso.OpenTextFile(mFileNm,1)
	        mailfooter = ffso.readall	' 푸터

			fnHTML = mailheader & fnHTML & mailfooter
		End IF

		getMailTemplate = fnHTML
	End Function

End CLASS

%>