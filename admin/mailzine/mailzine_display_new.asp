<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V3.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/classes/event/eventColorCodeCls.asp"-->
<%
response.write "������"
response.end
dim headerHTML, tailHTML, tmpHTML, salePer, saleCPer, title, coupontitle
dim headerDB, tailDB, htmlForDB, sqlStr
dim weekendHTML, maineventHTML, eventList8, eventList4, eventList, mdpickHTML, just1dayHTML, tentenclassHTML
dim yyyymmdd, fromyyyymmdd, toyyyymmdd, datecount, maxpercentage
dim evtList, prevEvenOdd, currEvenOdd, imgURL
dim idx, i, j, k, member, typeGubun, tempeventHTML
	idx = requestCheckVar(request("idx"),32)
	member = requestCheckVar(request("member"),32)
	typeGubun = requestCheckVar(request("type"),32)

dim yyyymmddStr : yyyymmddStr = Year(Now) & "��" & month(Now) & "��" & day(Now) & "��"
dim rejectURL : rejectURL = "http://www.10x10.co.kr/member/mailzine/reject_mailzine.asp?M_ID=${TMS_M_EMAIL}"
if (member <> "member") then
	rejectURL = "http://www.10x10.co.kr/member/mailzine/notmember_del.asp?usermail=[$email]&site=10x10"
end if
dim arrSmallBig(), currState
dim omail, cEvtCont
set omail = new CMailzineList
	omail.frectidx = idx
	omail.frectmailergubun = "EMS"
	omail.MailzineDetail()

weekendHTML = ""
yyyymmdd = omail.FOneItem.Fregdate

' �ָ�Ư��
if (omail.FOneItem.Fregtype = "2") then
	if (Not IsNumeric(omail.FOneItem.Fevt_code)) then
		Call PrintErrorAndStop("�߸��� �ָ�Ư�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'")
	end if

	set cEvtCont = new ClsEvent
		cEvtCont.FECode = omail.FOneItem.Fevt_code
		cEvtCont.fnGetEventCont
		cEvtCont.fnGetEventDisplay

	if (cEvtCont.FEName = "") then
		Call PrintErrorAndStop("�߸��� �ָ�Ư�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'")
	end if

	if (DateDiff("d", cEvtCont.FESDay, cEvtCont.FEEDay) < 2) then
		Call PrintErrorAndStop("�߸��� �ָ�Ư�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />�ָ�Ư�� �̺�Ʈ�Ⱓ�� 3�� �̸��Դϴ�.")
	end if

	fromyyyymmdd = cEvtCont.FESDay
	toyyyymmdd = cEvtCont.FEEDay
	fromyyyymmdd = replace(fromyyyymmdd, "-", ".")
	toyyyymmdd = replace(toyyyymmdd, "-", ".")
	if (Left(cEvtCont.FESDay,7) = Left(cEvtCont.FEEDay,7)) then
		toyyyymmdd = Right(toyyyymmdd,2)
	end if

	datecount = DateDiff("d", cEvtCont.FESDay, cEvtCont.FEEDay) + 1
	if (cEvtCont.FESale = True) then
		maxpercentage = cEvtCont.FsalePer
	end if
	if (maxpercentage = "") then
		Call PrintErrorAndStop("�߸��� �ָ�Ư�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ǰ �ִ� ���ΰ� �Է¾ȵ�.")
	end if

	weekendHTML = "							<tr>" & vbCrLf
	weekendHTML = weekendHTML + "								<td background=""http://mailzine.10x10.co.kr/2018/common/bg_weekend_sale2.png"" bgcolor=""#ffffff"" width=""700"" height=""504"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_weekend_sale2.png); background-repeat:no-repeat; background-position:50% 0; background-size:100%; vertical-align:top;"">" & vbCrLf
	weekendHTML = weekendHTML + "								<!--[if gte mso 9]>" & vbCrLf
	weekendHTML = weekendHTML + "								<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; height:504px;"">" & vbCrLf
	weekendHTML = weekendHTML + "									<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_weekend_sale2.png"" color=""#ffffff"" />" & vbCrLf
	weekendHTML = weekendHTML + "									<v:textbox inset=""0,0,0,0"">" & vbCrLf
	weekendHTML = weekendHTML + "								<![endif]-->" & vbCrLf
	weekendHTML = weekendHTML + "								<div>" & vbCrLf
	weekendHTML = weekendHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px"" width=""700"">" & vbCrLf
	weekendHTML = weekendHTML + "										<tr>" & vbCrLf
	weekendHTML = weekendHTML + "											<td style=""text-align:left; vertical-align:top;"">" & vbCrLf
	weekendHTML = weekendHTML + "												<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:450px"" width=""450"">" & vbCrLf
	weekendHTML = weekendHTML + "													<tr>" & vbCrLf
	weekendHTML = weekendHTML + "														<td style=""padding:60px 30px 0 30px; font-size:16px; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;"">${EMS_M_NAME}���� ���� �غ��� Ư���� ����������!</td>" & vbCrLf
	weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	if (DateDiff("d", cEvtCont.FESDay, cEvtCont.FEEDay) = 2) then
		weekendHTML = weekendHTML + "													<tr>" & vbCrLf
		weekendHTML = weekendHTML + "														<td style=""padding:40px 30px 0 30px; font-size:16px; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=[eventcode]&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_weekend_sale.png"" alt=""SPECIAL WEEKEND SALE"" style=""border:0;"" /></a></td>" & vbCrLf
		weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	else
		weekendHTML = weekendHTML + "													<tr>" & vbCrLf
		weekendHTML = weekendHTML + "														<td style=""padding:40px 30px 0 30px; font-size:16px; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=[eventcode]&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_holiday_sale.png"" alt=""SPECIAL HOLIDAY SALE"" style=""border:0;"" /></a></td>" & vbCrLf
		weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	end if
	weekendHTML = weekendHTML + "													<tr>" & vbCrLf
	weekendHTML = weekendHTML + "														<td style=""padding:20px 30px 60px 30px; color:#ff3131; font-family:verdana, sans-serif; font-weight:bold; text-align:left;"">" & vbCrLf
	weekendHTML = weekendHTML + "															<span style=""font-size:22px; text-decoration:underline; line-height:22px; vertical-align:middle;"">MAX</span> <span style=""font-size:48px; line-height:48px; vertical-align:middle;"">[maxpercentage]%</span>" & vbCrLf
	weekendHTML = weekendHTML + "														</td>" & vbCrLf
	weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	weekendHTML = weekendHTML + "												</table>" & vbCrLf
	weekendHTML = weekendHTML + "											</td>" & vbCrLf
	weekendHTML = weekendHTML + "											<td style=""text-align:right; vertical-align:top;"">" & vbCrLf
	weekendHTML = weekendHTML + "												<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:250px"" width=""250"">" & vbCrLf
	weekendHTML = weekendHTML + "													<tr>" & vbCrLf
	weekendHTML = weekendHTML + "														<td style=""padding:60px 30px; font-size:12px; color:#666; font-family:verdana, sans-serif; text-align:right; line-height:2;"">" & vbCrLf
	weekendHTML = weekendHTML + "															<p style=""padding:0; margin:0;"">[fromyyyymmdd] - [toyyyymmdd]</p>" & vbCrLf
	weekendHTML = weekendHTML + "															<p style=""padding:0; margin:0; font-weight:bold;"">ONLY [datecount]Days</p>" & vbCrLf
	weekendHTML = weekendHTML + "														</td>" & vbCrLf
	weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	weekendHTML = weekendHTML + "													<tr>" & vbCrLf
	weekendHTML = weekendHTML + "														<td style=""padding:80px 45px 0 0; text-align:right;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=[eventcode]&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/btn_event_link.png"" alt=""����Ư�� ��������"" style=""border:0;"" /></a></td>" & vbCrLf
	weekendHTML = weekendHTML + "													</tr>" & vbCrLf
	weekendHTML = weekendHTML + "												</table>" & vbCrLf
	weekendHTML = weekendHTML + "											</td>" & vbCrLf
	weekendHTML = weekendHTML + "										</tr>" & vbCrLf
	weekendHTML = weekendHTML + "									</table>" & vbCrLf
	weekendHTML = weekendHTML + "								</div>" & vbCrLf
	weekendHTML = weekendHTML + "								<!--[if gte mso 9]>" & vbCrLf
	weekendHTML = weekendHTML + "									</v:textbox>" & vbCrLf
	weekendHTML = weekendHTML + "								</v:rect>" & vbCrLf
	weekendHTML = weekendHTML + "								<![endif]-->" & vbCrLf
	weekendHTML = weekendHTML + "								</td>" & vbCrLf
	weekendHTML = weekendHTML + "							</tr>" & vbCrLf

	weekendHTML = Replace(weekendHTML, "[fromyyyymmdd]", fromyyyymmdd)
	weekendHTML = Replace(weekendHTML, "[toyyyymmdd]", toyyyymmdd)
	weekendHTML = Replace(weekendHTML, "[datecount]", datecount)
	weekendHTML = Replace(weekendHTML, "[maxpercentage]", maxpercentage)
	weekendHTML = Replace(weekendHTML, "[eventcode]", omail.FOneItem.Fevt_code)

end if

if (omail.FOneItem.Fregtype = "3") or (omail.FOneItem.Fregtype = "4") then
	'// ���� ��ȹ��

	if (Not IsNumeric(omail.FOneItem.Fevt_code)) then
		Call PrintErrorAndStop("�߸��� ���� ��ȹ�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'")
	end if

	set cEvtCont = new ClsEvent
	cEvtCont.FECode = omail.FOneItem.Fevt_code
	cEvtCont.fnGetEventCont
	cEvtCont.fnGetEventDisplay

	if (cEvtCont.FEName = "") then
		Call PrintErrorAndStop("�߸��� ���� ��ȹ�� �̺�Ʈ�ڵ��Դϴ�. : '" & omail.FOneItem.Fevt_code & "'")
	end if

	salePer = ""
	if (cEvtCont.FESale = True) then
		salePer = cEvtCont.FsalePer
	end if

	coupontitle = ""
	if (cEvtCont.FECoupon = True) then
		saleCPer = cEvtCont.FsaleCPer
		coupontitle = "<strong style=""display:inline-block; font-size:16px; line-height:1.5; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>"
	end if

	title = CHKIIF(InStr(cEvtCont.FEName, "|") > 0, Mid(cEvtCont.FEName, 1, InStr(cEvtCont.FEName, "|")), cEvtCont.FEName)
	title = Replace(title, "|", "")
	if (salePer = "") then
		salePer = CHKIIF(InStr(cEvtCont.FEName, "|") > 0, Mid(cEvtCont.FEName, InStr(cEvtCont.FEName, "|")+1, 1000), "")
		salePer = Replace(salePer, "~", "")
		salePer = Replace(salePer, "%", "")
	end if

	maineventHTML = "							<tr>" & vbCrLf
	maineventHTML = maineventHTML + "								<td style=""vertical-align:top;"">" & vbCrLf
	maineventHTML = maineventHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf
	maineventHTML = maineventHTML + "										<tr>" & vbCrLf
	maineventHTML = maineventHTML + "											<td>" & vbCrLf
	maineventHTML = maineventHTML + "												<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
	maineventHTML = maineventHTML + "													<tr>" & vbCrLf
	maineventHTML = maineventHTML + "														<td style=""width:700px; height:378px;"">" & vbCrLf
	maineventHTML = maineventHTML + "															<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & omail.FOneItem.Fevt_code & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & cEvtCont.FEBImgMoListBanner & """ alt="""" width=""700"" height=""378"" style=""width:700px; height:378px;"" /></a>" & vbCrLf
	maineventHTML = maineventHTML + "														</td>" & vbCrLf
	maineventHTML = maineventHTML + "													</tr>" & vbCrLf
	maineventHTML = maineventHTML + "													<tr>" & vbCrLf
	maineventHTML = maineventHTML + "														<td style=""padding:20px 0 20px 0; vertical-align:top;"">" & vbCrLf
	maineventHTML = maineventHTML + "															<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
	maineventHTML = maineventHTML + "																<tr>" & vbCrLf
	maineventHTML = maineventHTML + "																	<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
	maineventHTML = maineventHTML + "																		<table width=""600"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:600px;"">" & vbCrLf
	maineventHTML = maineventHTML + "																			<tr>" & vbCrLf
	maineventHTML = maineventHTML + "																				<td style=""padding:0; font-size:32px; font-weight:bold; line-height:1.31; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & omail.FOneItem.Fevt_code & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none; font-size:32px; line-height:1.31;"" target=""_blank"">" & title & "</a></td>" & vbCrLf
	maineventHTML = maineventHTML + "																			</tr>" & vbCrLf
	maineventHTML = maineventHTML + "																			<tr>" & vbCrLf
	maineventHTML = maineventHTML + "																				<td style=""padding:20px 0 0 0; font-size:16px; line-height:1.5; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & omail.FOneItem.Fevt_code & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:16px; line-height:1.5;"" target=""_blank"">" & coupontitle & cEvtCont.FsubcopyK & "</a></td>" & vbCrLf
	maineventHTML = maineventHTML + "																			</tr>" & vbCrLf
	maineventHTML = maineventHTML + "																		</table>" & vbCrLf
	maineventHTML = maineventHTML + "																	</td>" & vbCrLf
	if salePer<>"" then
	maineventHTML = maineventHTML + "																	<td width=""80"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
	maineventHTML = maineventHTML + "																		<table width=""80"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:80px;"">" & vbCrLf
	maineventHTML = maineventHTML + "																			<tr>" & vbCrLf
	maineventHTML = maineventHTML + "																				<td style=""width:80px; height:80px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:16px; font-weight:bold; text-align:center; text-decoration:none;"">~" & salePer & "%</td>" & vbCrLf
	maineventHTML = maineventHTML + "																			</tr>" & vbCrLf
	maineventHTML = maineventHTML + "																		</table>" & vbCrLf
	maineventHTML = maineventHTML + "																	</td>" & vbCrLf
	end if
	maineventHTML = maineventHTML + "																</tr>" & vbCrLf
	maineventHTML = maineventHTML + "															</table>" & vbCrLf
	maineventHTML = maineventHTML + "														</td>" & vbCrLf
	maineventHTML = maineventHTML + "													</tr>" & vbCrLf
	maineventHTML = maineventHTML + "												</table>" & vbCrLf
	maineventHTML = maineventHTML + "											</td>" & vbCrLf
	maineventHTML = maineventHTML + "										</tr>" & vbCrLf
	maineventHTML = maineventHTML + "										<tr>" & vbCrLf
	maineventHTML = maineventHTML + "											<td style=""padding:20px 0;""><img src=""http://mailzine.10x10.co.kr/2018/common/deco_line.png"" alt="""" style=""vertical-align:top; border:0;"" /></td>" & vbCrLf
	maineventHTML = maineventHTML + "										</tr>" & vbCrLf
	maineventHTML = maineventHTML + "									</table>" & vbCrLf
	maineventHTML = maineventHTML + "								</td>" & vbCrLf
	maineventHTML = maineventHTML + "							</tr>" & vbCrLf

end if

if (omail.FOneItem.Fregtype = "2") or (omail.FOneItem.Fregtype = "3") or (omail.FOneItem.Fregtype = "4") then
	'��ȹ�� 8��
	set cEvtCont = new ClsEvent
	cEvtCont.FECodeArr = omail.FOneItem.Fimgmap1
	evtList = ""
	if (omail.FOneItem.Fimgmap1 <> "") then
		evtList = cEvtCont.fnGetMailzineEventListData
	end if
	if Not IsArray(evtList) then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ��� ����.")
	end if

	if (UBound(evtList, 2) - LBound(evtList, 2)) < 7 then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ����� 8�� �̸�.")
	end if

	eventList8 = "							<tr>" & vbCrLf
	eventList8 = eventList8 + "								<td style=""padding-top:20px;"">" & vbCrLf
	eventList8 = eventList8 + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf

	redim arrSmallBig(8)
	currState = "B"
	if (omail.FOneItem.Fregtype = "2") or (omail.FOneItem.Fregtype = "3") then
		arrSmallBig(0) = "S1"
		arrSmallBig(1) = "S2"
		arrSmallBig(2) = "B"
		arrSmallBig(3) = "B"
		arrSmallBig(4) = "S1"
		arrSmallBig(5) = "S2"
		arrSmallBig(6) = "B"
		arrSmallBig(7) = "B"
	else
		arrSmallBig(0) = "S1"
		arrSmallBig(1) = "S2"
		arrSmallBig(2) = "B"
		arrSmallBig(3) = "S1"
		arrSmallBig(4) = "S2"
		arrSmallBig(5) = "S1"
		arrSmallBig(6) = "S2"
		arrSmallBig(7) = "B"
	end if

	prevEvenOdd = 1
	for i = LBound(evtList, 2) to UBound(evtList, 2)
		'evtList(0, i)

		if i = 2 and Replace(yyyymmdd, ".", "-") = "2018-08-02" then
			'// ȸ����� ����ȳ� ����(1ȸ)
			eventList8 = eventList8 + "										<tr>" & vbCrLf
			eventList8 = eventList8 + "											<td style=""padding:15px 0 15px 0;""><a href=""http://www.10x10.co.kr/my10x10/special_info.asp"" target=""_blank"" style=""text-decoration:none; display:block;""><img src=""http://mailzine.10x10.co.kr/2018/common/grade_noti_20180802.jpg"" alt="""" width=""700"" height=""210"" style=""width:700px; height:210px;"" /></a></td>" & vbCrLf
			eventList8 = eventList8 + "										</tr>" & vbCrLf
		end if

		salePer = trim(evtList(5, i))
		saleCPer = evtList(6, i)
		title = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), 1, InStr(evtList(3, i), "|")), evtList(3, i))
		title = Replace(title, "|", "")

		''response.write "aaaa" & title & "aaa" & evtList(3, i) & "aaaa" & InStr(evtList(3, i), "|") & "<br />"
		if (salePer = "") then
			salePer = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), InStr(evtList(3, i), "|")+1, 1000), "")
			salePer = Replace(salePer, "~", "")
			salePer = Replace(salePer, "%", "")
		end if

		if arrSmallBig(i) = "S1" then
			eventList8 = eventList8 + "										<tr>" & vbCrLf
			eventList8 = eventList8 + "											<td>" & vbCrLf
			eventList8 = eventList8 + "												<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf
			eventList8 = eventList8 + "													<tr>" & vbCrLf
			eventList8 = eventList8 + "														<!-- set(small-left) -->" & vbCrLf
			eventList8 = eventList8 + "														<td style=""padding:20px 10px 20px 0; vertical-align:top;"">" & vbCrLf
			eventList8 = eventList8 + "<!-- small left item table -->"
			eventList8 = eventList8 + "														</td>" & vbCrLf
			eventList8 = eventList8 + "														<!--// set(small-left) -->" & vbCrLf
			eventList8 = eventList8 + "														<!-- set(small-right) -->" & vbCrLf
			eventList8 = eventList8 + "														<td style=""padding:20px 0 20px 10px; vertical-align:top;"">" & vbCrLf
			eventList8 = eventList8 + "<!-- small right item table -->"
			eventList8 = eventList8 + "														</td>" & vbCrLf
			eventList8 = eventList8 + "														<!--// set(small-right) -->" & vbCrLf
			eventList8 = eventList8 + "													</tr>" & vbCrLf
			eventList8 = eventList8 + "												</table>" & vbCrLf
			eventList8 = eventList8 + "											</td>" & vbCrLf
			eventList8 = eventList8 + "										</tr>" & vbCrLf
		end if

		coupontitle = ""
		if arrSmallBig(i) = "S1" then
			'// ���� �̹���
			if (saleCPer <> "") then
				coupontitle = "<strong style=""display:inline-block; font-size:14px; line-height:1.43; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>"
			end if

			tmpHTML = "															<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""width:340px; height:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & evtList(1, i) & """ alt="""" width=""340"" height=""340"" style=""width:340px; height:340px;"" /></a>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""padding:15px 0 0 0; vertical-align:top;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
			tmpHTML = tmpHTML + "																					<table width=""" & CHKIIF((salePer <> ""), "260", "335") & """ border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:" & CHKIIF((salePer <> ""), "260", "335") & "px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:0; font-size:20px; font-weight:bold; line-height:1.5; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;	text-decoration:none; font-size:20px; line-height:1.5; letter-spacing:-0.5px;"" target=""_blank"">" & title & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:7px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:14px; line-height:1.43; letter-spacing:-0.5px;"" target=""_blank"">" & coupontitle & evtList(4, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					</table>" & vbCrLf
			tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			if (salePer <> "") then
				tmpHTML = tmpHTML + "																				<td width=""64"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
				tmpHTML = tmpHTML + "																					<table width=""64"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:64px;"">" & vbCrLf
				tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
				tmpHTML = tmpHTML + "																							<td style=""width:64px; height:64px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:14px; font-weight:bold; text-align:center; text-decoration:none;"">~" & salePer & "%</td>" & vbCrLf
				tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
				tmpHTML = tmpHTML + "																					</table>" & vbCrLf
				tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			else
				tmpHTML = tmpHTML + "																				<td width=""5"" style=""vertical-align:top; text-align:right;"">&nbsp;</td>" & vbCrLf
			end if
			tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																		</table>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "															</table>" & vbCrLf

			eventList8 = replace(eventList8, "<!-- small left item table -->", tmpHTML)
		elseif arrSmallBig(i) = "S2" then
			'// ������ �̹���
			if (saleCPer <> "") then
				coupontitle = "<strong style=""display:inline-block; font-size:14px; line-height:1.43; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>"
			end if

			tmpHTML = "															<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""width:340px; height:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & evtList(1, i) & """ alt="""" width=""340"" height=""340"" style=""width:340px; height:340px;"" /></a>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""padding:15px 0 0 0; vertical-align:top;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
			tmpHTML = tmpHTML + "																					<table width=""" & CHKIIF((salePer <> ""), "260", "335") & """ border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:" & CHKIIF((salePer <> ""), "260", "335") & "px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:0; font-size:20px; font-weight:bold; line-height:1.5; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;	text-decoration:none; font-size:20px; line-height:1.5; letter-spacing:-0.5px;"" target=""_blank"">" & evtList(3, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:7px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:14px; line-height:1.43; letter-spacing:-0.5px;"" target=""_blank"">" & coupontitle & evtList(4, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					</table>" & vbCrLf
			tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			if (salePer <> "") then
				tmpHTML = tmpHTML + "																				<td width=""64"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
				tmpHTML = tmpHTML + "																					<table width=""64"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:64px;"">" & vbCrLf
				tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
				tmpHTML = tmpHTML + "																							<td style=""width:64px; height:64px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:14px; font-weight:bold; text-align:center; text-decoration:none;"">~" & evtList(5, i) & "%</td>" & vbCrLf
				tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
				tmpHTML = tmpHTML + "																					</table>" & vbCrLf
				tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			else
				tmpHTML = tmpHTML + "																				<td width=""5"" style=""vertical-align:top; text-align:right;"">&nbsp;</td>" & vbCrLf
			end if
			tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																		</table>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "															</table>" & vbCrLf

			eventList8 = replace(eventList8, "<!-- small right item table -->", tmpHTML)
		else
			if (saleCPer <> "") then
				coupontitle = "<strong style=""display:inline-block; font-size:14px; line-height:1.43; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>"
			end if

			'// ū �̹���
			eventList8 = eventList8 + "										<tr>" & vbCrLf
			eventList8 = eventList8 + "											<td>" & vbCrLf
			eventList8 = eventList8 + "												<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf
			eventList8 = eventList8 + "<!-- big first item -->"
			''eventList8 = eventList8 + "<!-- big second item -->"
			eventList8 = eventList8 + "												</table>" & vbCrLf
			eventList8 = eventList8 + "											</td>" & vbCrLf
			eventList8 = eventList8 + "										</tr>" & vbCrLf

			tmpHTML = "													<!-- set(big) -->" & vbCrLf
			tmpHTML = tmpHTML + "													<tr>" & vbCrLf
			tmpHTML = tmpHTML + "														<td style=""padding:20px 0; vertical-align:top;"">" & vbCrLf
			tmpHTML = tmpHTML + "															<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""width:700px; height:378px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & evtList(2, i) & """ alt="""" width=""700"" height=""378"" style=""width:700px; height:378px;"" /></a>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""padding:15px 0 0 0; vertical-align:top;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
			tmpHTML = tmpHTML + "																					<table width=""" & CHKIIF((salePer <> ""), "620", "695") & """ border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:" & CHKIIF((salePer <> ""), "620", "695") & "px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:0; font-size:20px; font-weight:bold; line-height:1.5; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none; font-size:20px; line-height:1.5; letter-spacing:-0.5px;"" target=""_blank"">" & evtList(3, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""padding:7px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:14px; line-height:1.43; letter-spacing:-0.5px;"" target=""_blank"">" & coupontitle & evtList(4, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					</table>" & vbCrLf
			tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			if (salePer <> "") then
				tmpHTML = tmpHTML + "																				<td width=""64"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
				tmpHTML = tmpHTML + "																					<table width=""64"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:64px;"">" & vbCrLf
				tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
				tmpHTML = tmpHTML + "																							<td style=""width:64px; height:64px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:14px; font-weight:bold; text-align:center; text-decoration:none;"">~" & evtList(5, i) & "%</td>" & vbCrLf
				tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
				tmpHTML = tmpHTML + "																					</table>" & vbCrLf
				tmpHTML = tmpHTML + "																				</td>" & vbCrLf
			else
				tmpHTML = tmpHTML + "																				<td width=""5"" style=""vertical-align:top; text-align:right;"">&nbsp;</td>" & vbCrLf
			end if

			tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																		</table>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "															</table>" & vbCrLf
			tmpHTML = tmpHTML + "														</td>" & vbCrLf
			tmpHTML = tmpHTML + "													</tr>" & vbCrLf
			tmpHTML = tmpHTML + "													<!--// set(big) -->" & vbCrLf

			eventList8 = replace(eventList8, "<!-- big first item -->", tmpHTML)
		end if

		if (i >= 7) then
			exit for
		end if
	next

	eventList8 = eventList8 + "									</table>" & vbCrLf
	eventList8 = eventList8 + "								</td>" & vbCrLf
	eventList8 = eventList8 + "							</tr>" & vbCrLf
end if

if (omail.FOneItem.Fregtype = "3") then
	'// ��ȹ�� 4��
	set cEvtCont = new ClsEvent
	cEvtCont.FECodeArr = omail.FOneItem.Fimgmap1
	evtList = ""
	if (omail.FOneItem.Fimgmap1 <> "") then
		evtList = cEvtCont.fnGetMailzineEventListData
	end if
	if Not IsArray(evtList) then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ��� ����.")
	end if

	if (UBound(evtList, 2) - LBound(evtList, 2)) < 11 then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ����� 4�� �̸�.")
	end if

	eventList4 = "							<tr>" & vbCrLf
	eventList4 = eventList4 + "								<td style=""padding:10px 0 20px 0;"">" & vbCrLf
	eventList4 = eventList4 + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf

	for i = LBound(evtList, 2) to UBound(evtList, 2)
		if (i <= 7) then
			'// �տ� 8�� ��ŵ
		else
			if (i mod 2) = 0 then
				eventList4 = eventList4 + "										<tr>" & vbCrLf
				eventList4 = eventList4 + "											<td style=""padding:20px 10px 20px 0; vertical-align:top;"">" & vbCrLf
				eventList4 = eventList4 + "<!-- item " & i & " -->"
				eventList4 = eventList4 + "											</td>" & vbCrLf
				eventList4 = eventList4 + "											<td style=""padding:20px 0 20px 10px; vertical-align:top;"">" & vbCrLf
				eventList4 = eventList4 + "<!-- item " & (i+1) & " -->"
				eventList4 = eventList4 + "											</td>" & vbCrLf
				eventList4 = eventList4 + "										</tr>" & vbCrLf
			end if

			salePer = trim(evtList(5, i))
			saleCPer = evtList(6, i)
			title = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), 1, InStr(evtList(3, i), "|")), evtList(3, i))
			title = Replace(title, "|", "")

			''response.write "aaaa" & title & "aaa" & evtList(3, i) & "aaaa" & InStr(evtList(3, i), "|") & "<br />"
			if (salePer = "") then
				salePer = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), InStr(evtList(3, i), "|")+1, 1000), "")
				salePer = Replace(salePer, "~", "")
				salePer = Replace(salePer, "%", "")
			end if

			coupontitle = ""
			if (saleCPer <> "") then
				coupontitle = "<strong style=""display:inline-block; font-size:14px; line-height:1.43; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>"
			end if
			tmpHTML = "												<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "													<tr>" & vbCrLf
			tmpHTML = tmpHTML + "														<td style=""width:340px; height:184px;"">" & vbCrLf
			tmpHTML = tmpHTML + "															<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & evtList(2, i) & """ alt="""" width=""340"" height=""184"" style=""width:340px; height:184px;"" /></a>" & vbCrLf
			tmpHTML = tmpHTML + "														</td>" & vbCrLf
			tmpHTML = tmpHTML + "													</tr>" & vbCrLf
			tmpHTML = tmpHTML + "													<tr>" & vbCrLf
			tmpHTML = tmpHTML + "														<td style=""padding:15px 0 0 0; vertical-align:top;"">" & vbCrLf
			tmpHTML = tmpHTML + "															<table width=""340"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:340px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
			tmpHTML = tmpHTML + "																		<table width=""270"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:270px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<td style=""padding:0; font-size:20px; font-weight:bold; line-height:1.5; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none; font-size:20px; line-height:1.5; letter-spacing:-0.5px;"" target=""_blank"">" & title & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<td style=""padding:7px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:14px; line-height:1.43; letter-spacing:-0.5px;"" target=""_blank"">" & coupontitle & evtList(4, i) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																		</table>" & vbCrLf
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																	<td width=""56"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
			if (salePer <> "") then
				tmpHTML = tmpHTML + "																		<table width=""56"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:56px;"">" & vbCrLf
				tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
				tmpHTML = tmpHTML + "																				<td style=""width:56px; height:56px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:12px; font-weight:bold; text-align:center; text-decoration:none;"">~" & salePer & "%</td>" & vbCrLf
				tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
				tmpHTML = tmpHTML + "																		</table>" & vbCrLf
			end if
			tmpHTML = tmpHTML + "																	</td>" & vbCrLf
			tmpHTML = tmpHTML + "																</tr>" & vbCrLf
			tmpHTML = tmpHTML + "															</table>" & vbCrLf
			tmpHTML = tmpHTML + "														</td>" & vbCrLf
			tmpHTML = tmpHTML + "													</tr>" & vbCrLf
			tmpHTML = tmpHTML + "												</table>" & vbCrLf

			eventList4 = replace(eventList4, "<!-- item " & i & " -->", tmpHTML)
		end if
	next

	eventList4 = eventList4 + "									</table>" & vbCrLf
	eventList4 = eventList4 + "								</td>" & vbCrLf
	eventList4 = eventList4 + "							</tr>" & vbCrLf
end if

' ���̾���丮
if (omail.FOneItem.Fregtype = "5") then
	set cEvtCont = new ClsEvent
	cEvtCont.FECodeArr = omail.FOneItem.Fimgmap1
	evtList = ""
	if (omail.FOneItem.Fimgmap1 <> "") then
		evtList = cEvtCont.fnGetMailzineEventListData
	end if
	if Not IsArray(evtList) then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�[0]. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ��� ����.")
	end if

	if (UBound(evtList, 2) - LBound(evtList, 2)) < 1 then
		Call PrintErrorAndStop("�߸��� ��ȹ�� ����Դϴ�[1]. : '" & omail.FOneItem.Fevt_code & "'" & "<br />��ȹ�� ����� 1�� �̸�.")
	end if

	maineventHTML = "							<tr>" & vbCrLf
	maineventHTML = maineventHTML + "				<td style='vertical-align:top;'><a href=""http://www.10x10.co.kr/diarystory2019/?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src='http://mailzine.10x10.co.kr/2018/common/@temp_img_diary.jpg' alt='2019 DIARY STORY' /></a></td>" & vbCrLf
	maineventHTML = maineventHTML + "			</tr>" & vbCrLf

	eventList = "							<tr>" & vbCrLf
	eventList = eventList + "					<td style='padding:20px 0; vertical-align:top;'>" & vbCrLf
    eventList = eventList + "						<table border='0' cellpadding='0' cellspacing='0' style='width:100%;'>" & vbCrLf

	for i = LBound(evtList, 2) to UBound(evtList, 2)
		salePer = trim(evtList(5, i))
		saleCPer = evtList(6, i)
		title = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), 1, InStr(evtList(3, i), "|")), evtList(3, i))
		title = Replace(title, "|", "")

		''response.write "aaaa" & title & "aaa" & evtList(3, i) & "aaaa" & InStr(evtList(3, i), "|") & "<br />"
		if (salePer = "") then
			salePer = CHKIIF(InStr(evtList(3, i), "|") > 0, Mid(evtList(3, i), InStr(evtList(3, i), "|")+1, 1000), "")
			salePer = Replace(salePer, "~", "")
			salePer = Replace(salePer, "%", "")
		end if

		coupontitle = ""
		if (saleCPer <> "") then
			coupontitle = "<strong style=""display:inline-block; font-size:14px; line-height:1.43; color:#00b160; font-family:verdana, 'MalgunGothic', '�������', sans-serif;"">���� ~" & saleCPer & "%<img src=""http://mailzine.10x10.co.kr/2018/common/img_sep.png"" alt="""" style=""margin:0 8px;"" /></strong>" & vbCrLf
		end if

		tmpHTML = tmpHTML + "													<tr>" & vbCrLf
		tmpHTML = tmpHTML + "														<td style=""padding:20px 0; vertical-align:top;"">" & vbCrLf
		tmpHTML = tmpHTML + "															<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
		tmpHTML = tmpHTML + "																<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																	<td style=""width:700px; height:378px;"">" & vbCrLf
		tmpHTML = tmpHTML + "																		<a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""text-decoration:none; display:block;"" target=""_blank""><img src=""" & evtList(2, i) & """ alt="""" width=""700"" height=""378"" style=""width:700px; height:378px;"" /></a>" & vbCrLf
		tmpHTML = tmpHTML + "																	</td>" & vbCrLf
		tmpHTML = tmpHTML + "																</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																	<td style=""padding:15px 0 0 0; vertical-align:top;"">" & vbCrLf
		tmpHTML = tmpHTML + "																		<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
		tmpHTML = tmpHTML + "																			<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																				<td style=""vertical-align:top; text-align:left;"">" & vbCrLf
		tmpHTML = tmpHTML + "																					<table width=""" & CHKIIF((salePer <> ""), "620", "695") & """ border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:" & CHKIIF((salePer <> ""), "620", "695") & "px;"">" & vbCrLf
		tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																							<td style=""padding:0; font-size:20px; font-weight:bold; line-height:1.5; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none; font-size:20px; line-height:1.5; letter-spacing:-0.5px;"" target=""_blank"">" & title & "</a></td>" & vbCrLf
		tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																							<td style=""padding:7px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; text-align:left; color:#000; font-family:'MalgunGothic', '�������', sans-serif; -webkit-text-size-adjust:none;""><a href=""http://www.10x10.co.kr/event/eventmain.asp?eventid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000; text-decoration:none; font-size:14px; line-height:1.43; letter-spacing:-0.5px;"" target=""_blank"">" & coupontitle & evtList(4, i) & "</a></td>" & vbCrLf
		tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																					</table>" & vbCrLf
		tmpHTML = tmpHTML + "																				</td>" & vbCrLf
		if (salePer <> "") then
			tmpHTML = tmpHTML + "																				<td width=""64"" style=""vertical-align:top; text-align:right;"">" & vbCrLf
			tmpHTML = tmpHTML + "																					<table width=""64"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:64px;"">" & vbCrLf
			tmpHTML = tmpHTML + "																						<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																							<td style=""width:64px; height:64px; background-color:#ff3131; color:#fff; font-family:verdana, sans-serif; font-size:14px; font-weight:bold; text-align:center; text-decoration:none;"">~" & salePer & "%</td>" & vbCrLf
			tmpHTML = tmpHTML + "																						</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					</table>" & vbCrLf
			tmpHTML = tmpHTML + "																				</td>" & vbCrLf
		else
			tmpHTML = tmpHTML + "																				<td width=""5"" style=""vertical-align:top; text-align:right;"">&nbsp;</td>" & vbCrLf
		end if

		tmpHTML = tmpHTML + "																			</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																		</table>" & vbCrLf
		tmpHTML = tmpHTML + "																	</td>" & vbCrLf
		tmpHTML = tmpHTML + "																</tr>" & vbCrLf
		tmpHTML = tmpHTML + "															</table>" & vbCrLf
		tmpHTML = tmpHTML + "														</td>" & vbCrLf
		tmpHTML = tmpHTML + "													</tr>" & vbCrLf

		eventList = tmpHTML
	next

	eventList = eventList + "									</table>" & vbCrLf
	eventList = eventList + "								</td>" & vbCrLf
	eventList = eventList + "							</tr>" & vbCrLf
end if

if (omail.FOneItem.Fregtype = "2") or (omail.FOneItem.Fregtype = "4") or (omail.FOneItem.Fregtype = "5") then
	'// �ָ�Ư�� ������ 12��
	'// ��ȹ��+������ 6��

	if (omail.FOneItem.Fimgmap2 = "") then
		Call PrintErrorAndStop("�߸��� ������ ����Դϴ�[0]. : '" & omail.FOneItem.Fimgmap2 & "'" & "<br />������ ��� ����.")
	end if

	set cEvtCont = new ClsEvent
	cEvtCont.FRectItemidArr = omail.FOneItem.Fimgmap2
	cEvtCont.FESDay = omail.FOneItem.Fregdate

	' ���̾���丮
	if omail.FOneItem.Fregtype = "5" then
		evtList = cEvtCont.fnGetMailzinediaryData
	else
		evtList = cEvtCont.fnGetMailzineMDPickData
	end if

	if Not IsArray(evtList) then
		Call PrintErrorAndStop("�߸��� ������ ����Դϴ�[1]. : '" & omail.FOneItem.Fimgmap2 & "'" & "<br />������ ��� ����.")
	end if

	if (omail.FOneItem.Fregtype = "2") and (UBound(evtList, 2) - LBound(evtList, 2)) < 11 then
		Call PrintErrorAndStop("�߸��� ������ ����Դϴ�[2]. : '" & omail.FOneItem.Fimgmap2 & "'" & "<br />������ ����� 12�� �̸�.")
	end if

	if (omail.FOneItem.Fregtype = "4") and (UBound(evtList, 2) - LBound(evtList, 2)) < 5 then
		Call PrintErrorAndStop("�߸��� ������ ����Դϴ�[3]. : '" & omail.FOneItem.Fimgmap2 & "'" & "<br />������ ����� 6�� �̸�.")
	end if

	mdpickHTML = "							<tr>" & vbCrLf
	mdpickHTML = mdpickHTML + "								<td style=""padding:50px 0 0 0;"">" & vbCrLf
	mdpickHTML = mdpickHTML + "									<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; margin:0 auto;"">" & vbCrLf
	mdpickHTML = mdpickHTML + "										<tr>" & vbCrLf

	' ���̾���丮
	if omail.FOneItem.Fregtype = "5" then
		mdpickHTML = mdpickHTML + "											<td><img src=""http://mailzine.10x10.co.kr/2018/common/tit_recommend_diary.png"" alt=""��õ ���̾"" style=""vertical-align:top;"" /></td>" & vbCrLf
	else
		mdpickHTML = mdpickHTML + "											<td><img src=""http://mailzine.10x10.co.kr/2018/common/tit_mdpick.png"" alt=""MD's PICK"" style=""vertical-align:top;"" /></td>" & vbCrLf
	end if

	mdpickHTML = mdpickHTML + "										</tr>" & vbCrLf
	mdpickHTML = mdpickHTML + "										<tr>" & vbCrLf
	mdpickHTML = mdpickHTML + "											<td style=""padding:30px 5px;"">" & vbCrLf
	mdpickHTML = mdpickHTML + "												<table width=""690"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:690px; margin:0 auto;"">" & vbCrLf

	for i = LBound(evtList, 2) to UBound(evtList, 2)
		if (i >= 6) and (omail.FOneItem.Fregtype = "4") then
			'// ��ȹ��+������ 6��
			exit for
		end if

		if (i mod 3) = 0 then
			mdpickHTML = mdpickHTML + "													<tr>" & vbCrLf
			mdpickHTML = mdpickHTML + "														<td style=""width:230px; height:313px; padding:15px; vertical-align:top;"">" & vbCrLf
			mdpickHTML = mdpickHTML + "<!-- item 0 -->" & vbCrLf
			mdpickHTML = mdpickHTML + "														</td>" & vbCrLf
			mdpickHTML = mdpickHTML + "														<td style=""width:230px; height:313px; padding:15px; vertical-align:top;"">" & vbCrLf
			mdpickHTML = mdpickHTML + "<!-- item 1 -->" & vbCrLf
			mdpickHTML = mdpickHTML + "														</td>" & vbCrLf
			mdpickHTML = mdpickHTML + "														<td style=""width:230px; height:313px; padding:15px; vertical-align:top;"">" & vbCrLf
			mdpickHTML = mdpickHTML + "<!-- item 2 -->" & vbCrLf
			mdpickHTML = mdpickHTML + "														</td>" & vbCrLf
			mdpickHTML = mdpickHTML + "													</tr>" & vbCrLf
		end if

		' ���̾ ���丮
		if omail.FOneItem.Fregtype = "5" then
			imgURL = evtList(2, i)
			if (evtList(10, i) = "21") then
				imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(2, i)
			else
				imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(0, i)) + "/" + evtList(2, i)
			end if
		else
			imgURL = evtList(1, i)
			if (IsNull(imgURL) = True) then
				if (evtList(10, i) = "21") then
					imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(2, i)
				else
					imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(0, i)) + "/" + evtList(2, i)
				end if
			end if
		end if

		tmpHTML = "															<table width=""200"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:200px; margin:0 auto;"">" & vbCrLf
		tmpHTML = tmpHTML + "																<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																	<td width=""200px;"" style=""width:200px;""><a href=""" & evtList(3, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img width=""200"" height=""200"" src=""" & imgURL & """ style=""width:200px; height:200px; border:0;"" alt="""" /></a></td>" & vbCrLf
		tmpHTML = tmpHTML + "																</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																	<td style=""width:200px; padding:10px 0 0 0; font-size:14px; line-height:1.43; letter-spacing:-0.5px; color:#000; text-align:center; font-family:'MalgunGothic', '�������', verdana, sans-serif;""><a href=""" & evtList(3, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none;"" target=""_blank"">" & evtList(4, i) & "</a></td>" & vbCrLf
		tmpHTML = tmpHTML + "																</tr>" & vbCrLf
		tmpHTML = tmpHTML + "																<tr>" & vbCrLf
		tmpHTML = tmpHTML + "																	<td style=""height:16px; padding:7px 0 0 0; margin:0; text-align:center; vertical-align:top; font-size:16px; line-height:1;"">" & vbCrLf
		if (evtList(7, i) = "Y") or ((evtList(10, i) = "21") and (evtList(13, i) > 0)) then
			tmpHTML = tmpHTML + "																		<span style=""display:inline-block; width:32px; height:16px; letter-spacing:-0.5px; font-size:10px; line-height:16px; font-weight:bold; color:#fff; background:#ff3131 url(http://mailzine.10x10.co.kr/2018/common/tag_sale.png) no-repeat 50% 0; font-family:verdana, sans-serif; text-align:center; vertical-align:top;"">" & CHKIIF((evtList(10, i) = "21") and (evtList(13, i) > 0), evtList(13, i), evtList(8, i)) & "%</span>" & vbCrLf
		end if
		if (evtList(11, i) > 0) then
			tmpHTML = tmpHTML + "																		<span style=""display:inline-block; width:32px; height:16px; letter-spacing:-0.5px; font-size:10px; line-height:16px; font-weight:bold; color:#fff; background:#00b160 url(http://mailzine.10x10.co.kr/2018/common/tag_coupon.png) no-repeat 50% 0; font-family:verdana, sans-serif; text-align:center; vertical-align:top;"">" & evtList(11, i) & "%</span>"
		end if
		if (evtList(9, i) = "Y") then
			tmpHTML = tmpHTML + "																		<span style=""display:inline-block; width:32px; height:16px; vertical-align:top;""><img src=""http://mailzine.10x10.co.kr/2018/common/tag_new.png"" alt=""NEW"" /></span>"
		end if
		tmpHTML = tmpHTML + "																	</td>" & vbCrLf
		tmpHTML = tmpHTML + "																</tr>" & vbCrLf
		tmpHTML = tmpHTML + "															</table>" & vbCrLf

		mdpickHTML = replace(mdpickHTML, "<!-- item " & (i mod 3) & " -->", tmpHTML)
	next

	mdpickHTML = mdpickHTML + "												</table>" & vbCrLf
	mdpickHTML = mdpickHTML + "											</td>" & vbCrLf
	mdpickHTML = mdpickHTML + "										</tr>" & vbCrLf
	mdpickHTML = mdpickHTML + "									</table>" & vbCrLf
	mdpickHTML = mdpickHTML + "								</td>" & vbCrLf
	mdpickHTML = mdpickHTML + "							</tr>" & vbCrLf

end if

if (omail.FOneItem.Fregtype = "2") or (omail.FOneItem.Fregtype = "3") or (omail.FOneItem.Fregtype = "4") or (omail.FOneItem.Fregtype = "5") then
	'// ����Ʈ������ 1��

	if omail.FOneItem.Fimgmap3 = "" then
		'// ����Ʈ������ ���� ��쵵 ����Ʈ������ �����ϰ� ǥ���ϵ��� ����
		''Call PrintErrorAndStop("�߸��� ����Ʈ�������Դϴ�. : '" & omail.FOneItem.Fimgmap3 & "'" & "<br />����Ʈ������ ����.")
	else
		set cEvtCont = new ClsEvent
		cEvtCont.FRectItemid = omail.FOneItem.Fimgmap3
		cEvtCont.FESDay = omail.FOneItem.Fregdate
		if Replace(cEvtCont.FESDay, ".", "-") > "2018-09-18" then
			evtList = cEvtCont.fnGetMailzineJustOneDayData2018
		else
			evtList = cEvtCont.fnGetMailzineJustOneDayData
		end if
		if Not IsArray(evtList) then
			Call PrintErrorAndStop("�߸��� ����Ʈ�������Դϴ�. : '" & cEvtCont.FRectItemid & "'" & "<br />����Ʈ������ ����.")
		end if

		if Replace(cEvtCont.FESDay, ".", "-") > "2018-09-18" and UBound(evtList, 2) <= 1 then
			'// ����Ʈ������ �Ѱ��� ��쵵 ����Ʈ������ �����ϰ� ǥ���ϵ��� ����
			'if (evtList(6, 0) = "21") then
			'	imgURL = evtList(5, 0)
			'else
				imgURL = evtList(5, 0)
				if imgURL="" then imgURL=evtList(11, 0)
			'end if

			just1dayHTML = "							<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "								<td style=""padding-top:10px; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "										<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "											<td background=""http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png"" bgcolor=""#ffffff"" width=""700"" height=""240"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png); background-repeat:no-repeat; background-position:50% 0; background-size:cover; border-top:4px solid #ff3131; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; 240px; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png"" color=""#f5f5f5"" style=""vertical-align:top;"" />" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:textbox style=""mso-fit-shape-to-text:true"" inset=""0,0,0,0"" style=""vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			'just1dayHTML = just1dayHTML + "												<div style=""padding:0; margin:0; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<table height=""240"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; height:240px;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "														<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "															<td style=""height:180px; padding:30px 0; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																<table height=""180"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; height:180px;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																	<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		<td style=""padding:0 0 0 50px; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																			<p style=""padding:0; margin:30px 0 0 0; text-align:left;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																				<a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""vertical-align:top; border:0;"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_just1day.png"" alt="""" style=""vertical-align:top; border:0;"" /></a>" & vbCrLf

			just1dayHTML = just1dayHTML + "																			</p>" & vbCrLf
			just1dayHTML = just1dayHTML + "																			<p style=""margin:15px 0 0 0; padding:0; font-size:16px; line-height:1.5; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;	text-decoration:none;"" target=""_blank"">" & evtList(4, 0) & "</a></p>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		<td style=""padding:0 40px 0 30px; text-align:right; vertical-align:top;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""" & imgURL & """ alt="""" style=""width:180px; height:180px; border:0;"" /></a></td>" & vbCrLf
			just1dayHTML = just1dayHTML + "																	</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "																</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "															</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "														</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</table>" & vbCrLf
			'just1dayHTML = just1dayHTML + "												</div>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</v:textbox>" & vbCrLf
			just1dayHTML = just1dayHTML + "												</v:rect>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			just1dayHTML = just1dayHTML + "											</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "										</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "									</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "								</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "							</tr>" & vbCrLf
		elseif Replace(cEvtCont.FESDay, ".", "-") > "2018-09-18" and UBound(evtList, 2) >= 1 then
			just1dayHTML = "							<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "								<td style=""padding-top:10px; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "										<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "											<td background=""http://mailzine.10x10.co.kr/2018/common/bg_just1day4.png"" bgcolor=""#ffffff"" width=""700"" height=""504"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_just1day4.png); background-repeat:repeat-y; background-position:50% 0; background-size:cover; border-top:4px solid #ff3131; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; 504px; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_just1day4.png"" color=""#f5f5f5"" style=""vertical-align:top;"" />" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:textbox style=""mso-fit-shape-to-text:true"" inset=""0,0,0,0"" style=""vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			'just1dayHTML = just1dayHTML + "												<div style=""padding:0; margin:0; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "														<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "															<td style=""padding:64px 0 0 0; vertical-align:top; text-align:center;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																<a href=""http://www.10x10.co.kr/?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""vertical-align:top; border:0;"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_just1day2.png"" alt="""" style=""vertical-align:top; border:0;"" /></a>" & vbCrLf
			just1dayHTML = just1dayHTML + "															</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "														</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "														<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "															<td style=""padding-top:10px; text-align:center;"">�� �Ϸ�, ���ø� �� ����!</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "														</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "														<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "															<td style=""padding:35px 20px;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																	<tr>" & vbCrLf

			for i = 0 to UBound(evtList, 2)
				imgURL = evtList(5, i)
				if (Trim(imgURL) = "") then
					if (evtList(6, i) = "21") then
						imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(11, i)
					else
						imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(0, i)) + "/" + evtList(11, i)
					end if
				end if
				just1dayHTML = just1dayHTML + "																		<td style=""padding:0; vertical-align:top; text-align:center;"" valign=""top"">" & vbCrLf
				just1dayHTML = just1dayHTML + "																			<table width=""180"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:180px; margin:0 auto;"">" & vbCrLf
				just1dayHTML = just1dayHTML + "																				<tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																					<td width=""180"" style=""width:180px;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""" & imgURL & """ alt="""" style=""width:180px; height:180px; border:0;"" /></a></td>" & vbCrLf
				just1dayHTML = just1dayHTML + "																				</tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																				<tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																					<td style=""padding-top:15px; text-align:center; font-size:14px; color:#000000; font-weight:bold; font-family:MalgunGothic, '�������', verdana, sans-serif; line-height:1.29;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, i) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000000; text-decoration:none;"" target=""_blank"">" & evtList(4, i) & "</a></td>" & vbCrLf
				just1dayHTML = just1dayHTML + "																				</tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																				<tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																					<td style=""padding-top:6px; font-size:20px; color:#ff3131; font-weight:bold; font-family:verdana, sans-serif; text-align:center; vertical-align:top;""><span style=""color:#ff3131;"">" & evtList(9, i) & "</span> <span style=""color:#00b160;"">" & evtList(10, i) & "</span></td>" & vbCrLf
				just1dayHTML = just1dayHTML + "																				</tr>" & vbCrLf
				just1dayHTML = just1dayHTML + "																			</table>" & vbCrLf
				just1dayHTML = just1dayHTML + "																		</td>" & vbCrLf
			next

			just1dayHTML = just1dayHTML + "																	</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "																</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "															</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "														</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</table>" & vbCrLf
			'just1dayHTML = just1dayHTML + "												</div>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</v:textbox>" & vbCrLf
			just1dayHTML = just1dayHTML + "												</v:rect>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			just1dayHTML = just1dayHTML + "											</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "										</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "									</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "								</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "							</tr>" & vbCrLf
		else
			if (evtList(6, 0) = "21") then
				imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(5, 0)
			else
				imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(0, 0)) + "/" + evtList(5, 0)
			end if

			just1dayHTML = "							<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "								<td style=""padding-top:10px; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "										<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "											<td background=""http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png"" bgcolor=""#ffffff"" width=""700"" height=""240"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png); background-repeat:no-repeat; background-position:50% 0; background-size:cover; border-top:4px solid #ff3131; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; 240px; vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_just1day2.png"" color=""#f5f5f5"" style=""vertical-align:top;"" />" & vbCrLf
			just1dayHTML = just1dayHTML + "													<v:textbox style=""mso-fit-shape-to-text:true"" inset=""0,0,0,0"" style=""vertical-align:top;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			'just1dayHTML = just1dayHTML + "												<div style=""padding:0; margin:0; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "													<table height=""240"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; height:240px;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "														<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "															<td style=""height:180px; padding:30px 0; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																<table height=""180"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; height:180px;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																	<tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		<td style=""padding:0 0 0 50px; vertical-align:top;"" valign=""top"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																			<p style=""padding:0; margin:30px 0 0 0; text-align:left;"">" & vbCrLf
			just1dayHTML = just1dayHTML + "																				<a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""vertical-align:top; border:0;"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_just1day.png"" alt="""" style=""vertical-align:top; border:0;"" /></a>" & vbCrLf
			if ((evtList(7, 0) = "Y") and (evtList(3, 0) > 0)) or ((evtList(6, 0) = "21") and (evtList(9, 0) > 0)) then
				just1dayHTML = just1dayHTML + "																				<span style=""display:inline-block; padding-top:6px; font-size:32px; color:#ff3131; font-weight:bold; font-family:verdana, sans-serif; vertical-align:top;"">~" & CHKIIF((evtList(6, 0) = "21") and (evtList(9, 0) > 0), evtList(9, 0), evtList(8, 0)) & "%</span>" & vbCrLf
			elseif ((evtList(7, 0) = "N") and (evtList(2, 0) > 0) and (evtList(3, 0) > 0)) then
				if (evtList(2, 0) > evtList(3, 0)) then
					just1dayHTML = just1dayHTML + "																				<span style=""display:inline-block; padding-top:6px; font-size:32px; color:#ff3131; font-weight:bold; font-family:verdana, sans-serif; vertical-align:top;"">~" & evtList(10, 0) & "%</span>" & vbCrLf
				end if
			end if
			just1dayHTML = just1dayHTML + "																			</p>" & vbCrLf
			just1dayHTML = just1dayHTML + "																			<p style=""margin:15px 0 0 0; padding:0; font-size:16px; line-height:1.5; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;	text-decoration:none;"" target=""_blank"">" & evtList(4, 0) & "</a></p>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "																		<td style=""padding:0 40px 0 30px; text-align:right; vertical-align:top;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(0, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""" & imgURL & """ alt="""" style=""width:180px; height:180px; border:0;"" /></a></td>" & vbCrLf
			just1dayHTML = just1dayHTML + "																	</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "																</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "															</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "														</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</table>" & vbCrLf
			'just1dayHTML = just1dayHTML + "												</div>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<!--[if gte mso 9]>" & vbCrLf
			just1dayHTML = just1dayHTML + "													</v:textbox>" & vbCrLf
			just1dayHTML = just1dayHTML + "												</v:rect>" & vbCrLf
			just1dayHTML = just1dayHTML + "												<![endif]-->" & vbCrLf
			just1dayHTML = just1dayHTML + "											</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "										</tr>" & vbCrLf
			just1dayHTML = just1dayHTML + "									</table>" & vbCrLf
			just1dayHTML = just1dayHTML + "								</td>" & vbCrLf
			just1dayHTML = just1dayHTML + "							</tr>" & vbCrLf
		end if
	end if
end if

if (omail.FOneItem.Fregtype = "2") or (omail.FOneItem.Fregtype = "3") or (omail.FOneItem.Fregtype = "4") or (omail.FOneItem.Fregtype = "5") then
	'// �ٹ����� Ŭ���� 1�� or 3��

	set cEvtCont = new ClsEvent
	cEvtCont.FESDay = omail.FOneItem.Fregdate
	evtList = cEvtCont.fnGetMailzineTenTenClassData
	if Not IsArray(evtList) then
		Call PrintErrorAndStop("�߸��� �ٹ����� Ŭ���� �Դϴ�. : '" & omail.FOneItem.Fregdate & "'" & "<br />�ٹ����� Ŭ���� ����.")
	end if

	if UBound(evtList, 2) = 0 then
		if IsNull(evtList(7, 0)) or IsNull(evtList(13, 0)) then
			'// 1��
			if (evtList(5, 0) = "21") then
				imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(6, 0)
			else
				imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(1, 0)) + "/" + evtList(6, 0)
			end if

			tentenclassHTML = "							<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "								<td style=""padding:20px 0; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "										<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "											<td background=""http://mailzine.10x10.co.kr/2018/common/bg_class3.png"" bgcolor=""#f5f5f5"" width=""700"" height=""240"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_class3.png); background-repeat:no-repeat; background-position:50% 0; background-size:cover; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<!--[if gte mso 9]>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; 240px; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_class3.png"" color=""#f5f5f5"" style=""vertical-align:top;"" />" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<v:textbox style=""mso-fit-shape-to-text:true"" inset=""0,0,0,0"" style=""vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<![endif]-->" & vbCrLf

			tentenclassHTML = tentenclassHTML + "												<table width=""700"" height=""240"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; height:240px;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														<td style=""width:700px; height:210px; padding:30px 0 0 0; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "															<table width=""700"" height=""210"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; height:210px;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																	<td width=""180"" height=""180"" valign=""top"" style=""width:180px; padding:0 40px 0 30px; text-align:left; vertical-align:top;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""" & imgURL & """ alt="""" width=""180"" height=""180"" style=""width:180px; height:180px; border:0; display:block;"" /></a></td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																	<td width=""450"" style=""width:450px; padding:0 30px 0 50px; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<p style=""padding:0; margin:40px 0 0 0; color:#000; font-size:24px; font-family:'MalgunGothic', '�������', verdana, sans-serif; vertical-align:top; text-align:left;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none;"" target=""_blank"">�ٹ����� Ŭ����</a></p>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<p style=""padding:0; margin:0; font-size:24px; color:#000; font-weight:bold; font-family:'MalgunGothic', '�������', verdana, sans-serif; vertical-align:top; text-align:left;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none;"" target=""_blank"">" & evtList(3, 0) & "</a></p>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<p style=""padding:0; margin:20px 0 0 0; font-size:14px; line-height:1.5; color:#000; font-family:'MalgunGothic', '�������', verdana, sans-serif; text-align:left;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1, 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none;"" target=""_blank"">" & evtList(4, 0) & " <span style=""color:#ff3131; font-weight:bold; font-family:verdana, sans-serif;"">" & CHKIIF(evtList(2, 0) <> "" and evtList(2, 0) > 0, evtList(2, 0) & "%", "") & "</span></a></p>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																	</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "															</table>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												</table>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<!--[if gte mso 9]>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													</v:textbox>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												</v:rect>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<![endif]-->" & vbCrLf
			tentenclassHTML = tentenclassHTML + "											</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "										</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "									</table>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "								</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "							</tr>" & vbCrLf
		else
			'// 3��
			tentenclassHTML = "							<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "								<td style=""padding:20px 0; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "										<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "											<td background=""http://mailzine.10x10.co.kr/2018/common/bg_class4.png"" bgcolor=""#f5f5f5"" width=""700"" height=""440"" valign=""top"" style=""background-image:url(http://mailzine.10x10.co.kr/2018/common/bg_class4.png); background-repeat:no-repeat; background-position:50% 0; background-size:cover; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<!--[if gte mso 9]>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<v:rect xmlns:v=""urn:schemas-microsoft-com:vml"" fill=""true"" stroke=""false"" style=""width:700px; 440px; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<v:fill type=""tile"" src=""http://mailzine.10x10.co.kr/2018/common/bg_class4.png"" color=""#f5f5f5"" style=""vertical-align:top;"" />" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<v:textbox style=""mso-fit-shape-to-text:true"" inset=""0,0,0,0"" style=""vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<![endif]-->" & vbCrLf
			'tentenclassHTML = tentenclassHTML + "												<div style=""padding:0; margin:0; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													<table height=""440"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "															<td height=""62"" style=""padding:40px 0 36px 0; text-align:center; vertical-align:top;"" valign=""top""><img src=""http://mailzine.10x10.co.kr/2018/common/tit_class.png"" alt=""�ٹ����� Ŭ���� - �ٹ������� �����ϴ� Ư���� CLASS�� ����������."" style=""vertical-align:top;"" /></td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "															<td style=""padding:0 20px; vertical-align:top;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"" valign=""top"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																	<tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<td style=""width:180px; padding:0 20px; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "<!-- class 0 -->"
			tentenclassHTML = tentenclassHTML + "																		</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<td style=""width:180px; padding:0 20px; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "<!-- class 1 -->"
			tentenclassHTML = tentenclassHTML + "																		</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																		<td style=""width:180px; padding:0 20px; vertical-align:top;"">" & vbCrLf
			tentenclassHTML = tentenclassHTML + "<!-- class 2 -->"
			tentenclassHTML = tentenclassHTML + "																		</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																	</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "																</table>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "															</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "														</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													</table>" & vbCrLf
			'tentenclassHTML = tentenclassHTML + "												</div>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<!--[if gte mso 9]>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "													</v:textbox>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												</v:rect>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "												<![endif]-->" & vbCrLf
			tentenclassHTML = tentenclassHTML + "											</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "										</tr>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "									</table>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "								</td>" & vbCrLf
			tentenclassHTML = tentenclassHTML + "							</tr>" & vbCrLf
		end if

		for i = 0 to 2
			if (evtList(5 + (i*6), 0) = "21") then
				imgURL = "http://webimage.10x10.co.kr/image/icon1/" & evtList(6 + (i*6), 0)
			else
				imgURL = webImgUrlForMAIL & "/image/icon1/" + GetImageSubFolderByItemid(evtList(1 + (i*6), 0)) + "/" + evtList(6 + (i*6), 0)
			end if

			tmpHTML = "																			<table width=""180"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:180px; margin:0 auto;"">" & vbCrLf
			tmpHTML = tmpHTML + "																				<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					<td width=""180"" style=""width:180px;""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1 + (i*6), 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img width=""180"" height=""180"" src=""" & imgURL & """ style=""width:180px; height:180px; border:0;"" alt="""" /></a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																				</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					<td style=""width:180px; padding:17px 0 0 0; font-size:13px; line-height:1.54; letter-spacing:-0.5px; color:#000; text-align:center; font-family:'MalgunGothic', '�������', verdana, sans-serif;""""><a href=""http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" & evtList(1 + (i*6), 0) & "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ style=""color:#000;  text-decoration:none;"" target=""_blank"">" & evtList(3+(i*6), 0) & "</a></td>" & vbCrLf
			tmpHTML = tmpHTML + "																				</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																				<tr>" & vbCrLf
			tmpHTML = tmpHTML + "																					<td style=""padding:9px 0 0 0; margin:0; text-align:center; vertical-align:top; font-size:12px; line-height:1; font-weight:bold; font-family:verdana, sans-serif;""><span style=""color:#ff3131;"">" & CHKIIF(evtList(2+(i*6), 0) <> "" and evtList(2+(i*6), 0) > 0, evtList(2+(i*6), 0) & "%", "") & "</span></td>" & vbCrLf
			tmpHTML = tmpHTML + "																				</tr>" & vbCrLf
			tmpHTML = tmpHTML + "																			</table>" & vbCrLf

			tentenclassHTML = replace(tentenclassHTML, "<!-- class " & i & " -->", tmpHTML)
		next
	end if
end if

if replace(yyyymmdd,".","-") >= "2019-04-01" and replace(yyyymmdd,".","-") < "2019-04-23" then
	tempeventHTML="<tr>" & vbCrLf
	tempeventHTML = tempeventHTML + "	<td style=""padding:0 0 70px;""><a href=""http://www.10x10.co.kr/event/salelife/"" target=""_blank""><img src=""http://webimage.10x10.co.kr/fixevent/event/2019/salabal/bnr_mailzine.jpg"" alt="""" style=""border:0;""></a></td>" & vbCrLf
	tempeventHTML = tempeventHTML + "</tr>" & vbCrLf
end if

headerHTML = "<!DOCTYPE html>" & vbCrLf
headerHTML = headerHTML + "<html>" & vbCrLf
headerHTML = headerHTML + "<head>" & vbCrLf
headerHTML = headerHTML + "<title>(����) " & omail.FOneItem.Ftitle & "</title>" & vbCrLf
headerHTML = headerHTML + "<meta http-equiv=""Content-Type"" content=""text/html; charset=euc-kr"">" & vbCrLf
headerHTML = headerHTML + "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no"" />" & vbCrLf
headerHTML = headerHTML + "</head>" & vbCrLf
headerHTML = headerHTML + "<body>" & vbCrLf
headerHTML = headerHTML + "<div style=""width:100%; margin:0 auto; padding:0; background-color:#fff;"">" & vbCrLf
headerHTML = headerHTML + "	<div style=""width:700px; margin:0 auto; padding:0;"">" & vbCrLf
headerHTML = headerHTML + "		<table width=""700"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; margin-left:auto; margin-right:auto; background-color:#fff"" background=""#fff"">" & vbCrLf
headerHTML = headerHTML + "			<tr>" & vbCrLf
headerHTML = headerHTML + "				<td style=""text-align:center;"" width=""700"">" & vbCrLf
headerHTML = headerHTML + "					<table width=""700"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; margin-left:auto; margin-right:auto;"">" & vbCrLf
headerHTML = headerHTML + "						<!-- ��� ���� -->" & vbCrLf
headerHTML = headerHTML + "						<thead>" & vbCrLf
headerHTML = headerHTML + "							<tr>" & vbCrLf
headerHTML = headerHTML + "								<td style=""padding:20px 0 15px 0;"">" & vbCrLf
headerHTML = headerHTML + "									<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; height:45px;"">" & vbCrLf
headerHTML = headerHTML + "										<tr>" & vbCrLf
headerHTML = headerHTML + "											<td style=""text-align:left;""><a href=""http://www.10x10.co.kr?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/mail_logo.png"" alt=""10x10"" style=""vertical-align:top; border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "											<td style=""padding:0 0 7px 0; text-align:right; vertical-align:bottom; font-size:13px; color:#666; font-family:'MalgunGothic', '�������', verdana, sans-serif;"">[yyyymmdd] �ٹ����� ��õ ����</td>" & vbCrLf
headerHTML = headerHTML + "										</tr>" & vbCrLf
headerHTML = headerHTML + "									</table>" & vbCrLf
headerHTML = headerHTML + "								 </td>" & vbCrLf
headerHTML = headerHTML + "							</tr>" & vbCrLf

headerHTML = headerHTML + "							<tr>" & vbCrLf
headerHTML = headerHTML + "								<td>" & vbCrLf
headerHTML = headerHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px;"">" & vbCrLf
headerHTML = headerHTML + "										<tr>" & vbCrLf
headerHTML = headerHTML + "											<td><a href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_newitem.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2019/txt_gnb_01.gif"" alt=""NEW"" style=""border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "											<td><a href=""http://www.10x10.co.kr/award/awardlist.asp?atype=b&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2019/txt_gnb_02.gif"" alt=""BEST"" style=""border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "											<td><a href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_saleitem.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2019/txt_gnb_03.gif"" alt=""SALE"" style=""border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "											<td><a href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_allevent.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2019/txt_gnb_04.gif"" alt=""�̺�Ʈ"" style=""border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "											<td><a href=""http://www.10x10.co.kr/street/?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2019/txt_gnb_05.gif"" alt=""�귣��"" style=""border:0;"" /></a></td>" & vbCrLf
headerHTML = headerHTML + "										</tr>" & vbCrLf
headerHTML = headerHTML + "									</table>" & vbCrLf
headerHTML = headerHTML + "								</td>" & vbCrLf
headerHTML = headerHTML + "							</tr>" & vbCrLf
headerHTML = headerHTML + "						</thead>" & vbCrLf
headerHTML = headerHTML + "						<!-- //��� ���� -->" & vbCrLf
headerHTML = headerHTML + "						<!-- ������ ���� -->" & vbCrLf
headerHTML = headerHTML + "						<tbody>" & vbCrLf

headerDB = "<div style=""width:100%; margin:0 auto; padding:0; background-color:#fff;"">" & vbCrLf
headerDB = headerDB + "	<div style=""width:700px; margin:0 auto; padding:0;"">" & vbCrLf
headerDB = headerDB + "		<table width=""700"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; margin-left:auto; margin-right:auto; background-color:#fff"" background=""#fff"">" & vbCrLf
headerDB = headerDB + "			<tr>" & vbCrLf
headerDB = headerDB + "				<td style=""text-align:center;"" width=""700"">" & vbCrLf
headerDB = headerDB + "					<table width=""700"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; margin-left:auto; margin-right:auto;"">" & vbCrLf
headerDB = headerDB + "						<!-- ��� ���� -->" & vbCrLf
headerDB = headerDB + "						<thead>" & vbCrLf
headerDB = headerDB + "							<tr>" & vbCrLf
headerDB = headerDB + "								<td style=""padding:20px 0 15px 0;"">" & vbCrLf
headerDB = headerDB + "									<table width=""700"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:700px; height:45px;"">" & vbCrLf
headerDB = headerDB + "										<tr>" & vbCrLf
headerDB = headerDB + "											<td style=""text-align:left;""><a href=""http://www.10x10.co.kr?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ target=""_blank""><img src=""http://mailzine.10x10.co.kr/2018/common/mail_logo.png"" alt=""10x10"" style=""vertical-align:top; border:0;"" /></a></td>" & vbCrLf
headerDB = headerDB + "											<td style=""padding:0 0 7px 0; text-align:right; vertical-align:bottom; font-size:13px; color:#666; font-family:'MalgunGothic', '�������', verdana, sans-serif;"">[yyyymmdd] �ٹ����� ��õ ����</td>" & vbCrLf
headerDB = headerDB + "										</tr>" & vbCrLf
headerDB = headerDB + "									</table>" & vbCrLf
headerDB = headerDB + "								 </td>" & vbCrLf
headerDB = headerDB + "							</tr>" & vbCrLf
headerDB = headerDB + "							<tr>" & vbCrLf
headerDB = headerDB + "								<td><img src=""http://mailzine.10x10.co.kr/2018/common/mail_gnb.png"" alt=""NEW / BEST / SALE / �̺�Ʈ / �귣��� �̵��մϴ�"" style=""vertical-align:top; border:0;"" usemap=""#mailGnbMap"" /></td>" & vbCrLf
headerDB = headerDB + "								<map name=""mailGnbMap"">" & vbCrLf
headerDB = headerDB + "									<area shape=""rect"" coords=""5,1,140,49"" href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_newitem.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ alt=""NEW"" target=""_blank"">" & vbCrLf
headerDB = headerDB + "									<area shape=""rect"" coords=""141,1,280,49"" href=""http://www.10x10.co.kr/award/awardlist.asp?atype=b&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ alt=""BEST"" target=""_blank"">" & vbCrLf
headerDB = headerDB + "									<area shape=""rect"" coords=""281,1,420,49"" href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_saleitem.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ alt=""SALE"" target=""_blank"">" & vbCrLf
headerDB = headerDB + "									<area shape=""rect"" coords=""421,1,560,49"" href=""http://www.10x10.co.kr/shoppingtoday/shoppingchance_allevent.asp?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ alt=""�̺�Ʈ"" target=""_blank"">" & vbCrLf
headerDB = headerDB + "									<area shape=""rect"" coords=""561,1,695,49"" href=""http://www.10x10.co.kr/street/?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & """ alt=""�귣��"" target=""_blank"">" & vbCrLf
headerDB = headerDB + "								</map>" & vbCrLf
headerDB = headerDB + "							</tr>" & vbCrLf
headerDB = headerDB + "						</thead>" & vbCrLf
headerDB = headerDB + "						<!-- //��� ���� -->" & vbCrLf
headerDB = headerDB + "						<!-- ������ ���� -->" & vbCrLf
headerDB = headerDB + "						<tbody>" & vbCrLf

headerHTML = Replace(headerHTML, "[yyyymmdd]", yyyymmdd)
headerDB = Replace(headerDB, "[yyyymmdd]", yyyymmdd)


tailHTML = "							<tr>" & vbCrLf
tailHTML = tailHTML + "								<td style=""height:60px; border-top:1px solid #f4f4f4; background-color:#fff; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px; line-height:1.17; letter-spacing:-1px; text-align:center; color:#808080;"">������ ��� ���� ������ �ǵ��� ����ϰڽ��ϴ�</td>" & vbCrLf
tailHTML = tailHTML + "							</tr>" & vbCrLf
tailHTML = tailHTML + "						</tbody>" & vbCrLf
tailHTML = tailHTML + "						<!-- //������ ���� -->" & vbCrLf
tailHTML = tailHTML + "						<!-- �ϴ� ���� ���� -->" & vbCrLf
tailHTML = tailHTML + "						<tfoot>" & vbCrLf
tailHTML = tailHTML + "							<tr>" & vbCrLf
tailHTML = tailHTML + "								<td style=""padding:35px 0 40px 0; background:#f4f4f4; text-align:center; border-bottom:2px solid #ededed;"">" & vbCrLf
tailHTML = tailHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf
tailHTML = tailHTML + "											<td style=""font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:23px; line-height:30px; color:#000; text-align:center;""><strong>���� ����Ϸ� ���ϰ�!</strong><br />�����ȯ �ٹ�����</td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf
tailHTML = tailHTML + "											<td style=""padding-top:20px; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:23px; line-height:30px; color:#000; text-align:center;"">" & vbCrLf
tailHTML = tailHTML + "												<a href=""https://itunes.apple.com/kr/app/tenbaiten/id864817011?mt=8"" target=""_blank"" title=""��â���� �۽���� ����""><img src=""http://mailzine.10x10.co.kr/2017/btn_appstore.png"" alt=""Download on the App Store"" style=""border:0; vertical-align:top;"" /></a>" & vbCrLf
tailHTML = tailHTML + "												<a href=""https://play.google.com/store/apps/details?id=kr.tenbyten.shopping"" target=""_blank"" title=""��â���� �����÷��� ����""><img src=""http://mailzine.10x10.co.kr/2017/btn_googleplay.png"" alt=""Get it on Google Play"" style=""border:0; vertical-align:top;"" /></a>" & vbCrLf
tailHTML = tailHTML + "											</td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "									</table>" & vbCrLf
tailHTML = tailHTML + "								</td>" & vbCrLf
tailHTML = tailHTML + "							</tr>" & vbCrLf
tailHTML = tailHTML + "							<tr>" & vbCrLf
tailHTML = tailHTML + "								<td style=""background:#f4f4f4; font-size:18px;"">" & vbCrLf
tailHTML = tailHTML + "									<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;"">" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf


tailHTML = tailHTML + "											<td style=""margin:0; padding:67px 37px 8px 37px; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">" & vbCrLf
tailHTML = tailHTML + "												<p style=""margin:0; padding:0; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">* �� ������ ������Ÿ� �̿����� �� ������ȣ � ���� ���������Ģ�� �ǰ� &nbsp;${TODAYSTR} �������� ���ϼ��� �����ϼ̱⿡ �߼۵Ǵ� �߼���������Դϴ�.</p>" & vbCrLf
tailHTML = tailHTML + "												<p style=""margin:0; padding:0; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">* �� ������ �߽� �����̸� ȸ�� �� ������ ���� �� �����ϴ�.</p>" & vbCrLf
tailHTML = tailHTML + "												<p style=""margin:0; padding:0; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">* �� �̻� ������ ������ �����ø� ���Űź� ��ư�� Ŭ�����ּ���.</p>" & vbCrLf
tailHTML = tailHTML + "												<p style=""margin:0; padding:0; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">* �������� ���� �� ���뿡 2~3���� �ҿ�� �� �ִ� �� ���� ��Ź�帳�ϴ�.</p>" & vbCrLf
tailHTML = tailHTML + "												<p style=""margin:0; padding:0; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">&nbsp;&nbsp;[<a href=""${REJECTLINK}"" style=""color:#838383; text-decoration:none; font-style:bold;"" target=""_blank""><b>���Űź�</b></a>] (To unsubscribe this e-mail, click <a href=""${REJECTLINK}"" style=""color:#838383; text-decoration:none; font-style:bold;"" target=""_blank"">HERE</a>)" & vbCrLf
tailHTML = tailHTML + "											</td>" & vbCrLf
''tailHTML = tailHTML + "											<td style=""margin:0; padding:67px 37px 8px 37px; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">* �� ������ ������Ÿ� �̿����� �� ������ȣ � ���� ���������Ģ�� �ǰ� <br />&nbsp;&nbsp;${TODAYSTR} �������� ���ϼ��� �����ϼ̱⿡ �߼۵Ǵ� �߼���������Դϴ�.<br />* �� ������ �߽� �����̸� ȸ�� �� ������ ���� �� �����ϴ�.<br />* �� �̻� ������ ������ �����ø� ���Űź� ��ư�� Ŭ�����ּ���.<br />* �������� ���� �� ���뿡 2~3���� �ҿ�� �� �ִ� �� ���� ��Ź�帳�ϴ�.<br />&nbsp;&nbsp;[<a href=""${REJECTLINK}"" style=""color:#838383; text-decoration:none; font-style:bold;"" target=""_blank""><b>���Űź�</b></a>] (To unsubscribe this e-mail, click <a href=""${REJECTLINK}"" style=""color:#838383; text-decoration:none; font-style:bold;"" target=""_blank"">HERE</a>)</td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf
tailHTML = tailHTML + "											<td style=""margin:0; padding:8px 35px; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px !important; line-height:27px; letter-spacing:-0.3px; color:#838383; text-align:left;"">(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����<br />��ǥ�̻� : ������ / ����ڵ�Ϲ�ȣ : 211-87-00620 / ����Ǹž��Ű� : ��01-1968ȣ / �������� ��ȣ �� û�ҳ� ��ȣå���� : �̹���<br />���ູ���� TEL : <a href=""tel:1644-6030"" style=""color:#838383; text-decoration:none; font-style:bold;""><b>1644-6030</b></a> / E-mail : <a href=""mailto:customer@10x10.co.kr"" style=""color:#838383; text-decoration:none; font-style:bold;""><b>customer@10x10.co.kr</b></a></td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf
tailHTML = tailHTML + "											<td style=""margin:0; padding:8px 35px; font-family:Verdana, sans-serif; font-size:18px; line-height:1.39; letter-spacing:-0.3px; color:#838383; text-align:left;"">COPYRIGHTS 10x10. ALL RIGHTS RESERVED.</td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "										<tr>" & vbCrLf
tailHTML = tailHTML + "											<td style=""padding:35px 35px 72px 35px; line-height:28px; text-align:center;"">" & vbCrLf
tailHTML = tailHTML + "												<a href=""http://www.facebook.com/your10x10/"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2017/ico_facebook.png"" alt=""�ٹ����� ���� Facebook���� �̵�"" style=""margin:0 25px; border:0;"" /></a>" & vbCrLf
tailHTML = tailHTML + "												<a href=""http://www.instagram.com/your10x10/"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2017/ico_instargram.png"" alt=""�ٹ����� ���� Instargram���� �̵�"" style=""margin:0 25px; border:0;"" /></a>" & vbCrLf
tailHTML = tailHTML + "												<a href=""https://www.pinterest.com/your10x10/"" target=""_blank""><img src=""http://mailzine.10x10.co.kr/2017/ico_pinterest.png"" alt=""�ٹ����� ���� Pinterest�� �̵�"" style=""margin:0 25px; border:0;"" /></a>" & vbCrLf
tailHTML = tailHTML + "											</td>" & vbCrLf
tailHTML = tailHTML + "										</tr>" & vbCrLf
tailHTML = tailHTML + "									</table>" & vbCrLf
tailHTML = tailHTML + "								</td>" & vbCrLf
tailHTML = tailHTML + "							</tr>" & vbCrLf
tailHTML = tailHTML + "						</tfoot>" & vbCrLf
tailHTML = tailHTML + "						<!-- //�ϴ� ���� ���� -->" & vbCrLf
tailHTML = tailHTML + "					</table>" & vbCrLf
tailHTML = tailHTML + "				</td>" & vbCrLf
tailHTML = tailHTML + "			</tr>" & vbCrLf
tailHTML = tailHTML + "		</table>" & vbCrLf
tailHTML = tailHTML + "	</div>" & vbCrLf
tailHTML = tailHTML + "</div>" & vbCrLf
tailHTML = tailHTML + "</body>" & vbCrLf
tailHTML = tailHTML + "</html>" & vbCrLf

tailDB = "							<tr>" & vbCrLf
tailDB = tailDB + "								<td style=""height:60px; border-top:1px solid #f4f4f4; background-color:#fff; font-family:'�������','Malgun Gothic','����', dotum, sans-serif; font-size:18px; line-height:1.17; letter-spacing:-1px; text-align:center; color:#808080;"">������ ��� ���� ������ �ǵ��� ����ϰڽ��ϴ�</td>" & vbCrLf
tailDB = tailDB + "							</tr>" & vbCrLf
tailDB = tailDB + "						</tbody>" & vbCrLf
tailDB = tailDB + "						<!-- //������ ���� -->" & vbCrLf
tailDB = tailDB + "					</table>" & vbCrLf
tailDB = tailDB + "				</td>" & vbCrLf
tailDB = tailDB + "			</tr>" & vbCrLf
tailDB = tailDB + "		</table>" & vbCrLf
tailDB = tailDB + "	</div>" & vbCrLf
tailDB = tailDB + "</div>" & vbCrLf

tailHTML = Replace(tailHTML, "${TODAYSTR}", yyyymmddStr)
tailHTML = Replace(tailHTML, "${REJECTLINK}", rejectURL)

''response.write headerHTML
''response.write weekendHTML
''response.write eventList8
''response.write mdpickHTML
''response.write just1dayHTML
''response.write tentenclassHTML
''response.write tempeventHTML
''response.write tailHTML
%>
<% if (typeGubun = "code") then %>
	<font color="red">�� �ڵ� ����</font>
	<form name="frm" onsubmit="return false;" style="margin:0px;">
	<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
	<tr bgcolor="FFFFFF">
		<td>
			<input type="text" name="title" size="100" class="input" readonly value="(����) <% = omail.FOneItem.Ftitle %>"><br>
			<textarea name="mailcontents" rows="35" class="input" readonly style="width:100%;"><%
	response.write headerHTML
	response.write weekendHTML
	response.write maineventHTML
	response.write eventList8
	response.write eventList4
	response.write eventList
	response.write mdpickHTML
	response.write just1dayHTML
	response.write tentenclassHTML
	response.write tempeventHTML
	response.write tailHTML
	%></textarea>
		</td>
	</tr>
	</table>

	<div align="center">
		<br />
		<% if IsNull(omail.FOneItem.FreservationDATE) or (omail.FOneItem.FreservationDATE = "") then %>
			�߼ۿ��� �������Ŀ��� ���Ϸ� ���۰����մϴ�.
		<% else %>
			<input type="button" class="button" style="width=200px; height=70px;" value=" ���Ϸ� �����ϱ� " onClick="jsSendMailer();">
		<% end if %>
		<br />

	</div>
	</form>

	<script type='text/javascript'>

	function jsSendMailer() {
		<% if (Left(Now(),10) >= Replace(omail.FOneItem.Fregdate, ".", "-")) then %>
			alert("���� ������¥�� �̸����� ������ �� �����ϴ�.");
		<% elseif (omail.FOneItem.Farea = "ten_all") and (omail.FOneItem.Fmemgubun = "member_all") and (omail.FOneItem.Fgubun = "5") then %>
			var idx, member;
			var frmAct = document.frmAct;
			if (confirm('�����Ͻðڽ��ϱ�?') == true) {
				frmAct.title.value = document.frm.title.value;
				frmAct.mailcontents.value = document.frm.mailcontents.value;
				frmAct.target = "iframe_mailer";
				frmAct.submit();
			}
		<% else %>
			alert('�ٹ����� ������, ��ü�ɹ�, �����οϼ� ���Ŀ��� ���۰����մϴ�.');
		<% end if %>
	}

	</script>
	<form name="frmAct" method="post" action="mailzine_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="insmailer">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="member" value="<%= member %>">
	<input type="hidden" name="title" value="">
	<input type="hidden" name="mailcontents" value="">
	<input type="hidden" name="yyyymmdd" value="<%= Replace(omail.FOneItem.Fregdate, ".", "-") %>">
	</form>
	<iframe name="iframe_mailer" width="110" height="110" border="0" frameborder="0"></iframe>
<%
else
	response.write headerHTML
	response.write weekendHTML
	response.write maineventHTML
	response.write eventList8
	response.write eventList4
	response.write eventList
	response.write mdpickHTML
	response.write just1dayHTML
	response.write tentenclassHTML
	response.write tempeventHTML
	response.write tailHTML
end if

htmlForDB = headerDB
htmlForDB = htmlForDB & vbCrLf & weekendHTML
htmlForDB = htmlForDB & vbCrLf & maineventHTML
htmlForDB = htmlForDB & vbCrLf & eventList8
htmlForDB = htmlForDB & vbCrLf & eventList4
htmlForDB = htmlForDB & vbCrLf & eventList
htmlForDB = htmlForDB & vbCrLf & mdpickHTML
htmlForDB = htmlForDB & vbCrLf & just1dayHTML
htmlForDB = htmlForDB & vbCrLf & tentenclassHTML
htmlForDB = htmlForDB & vbCrLf & tailDB

htmlForDB = Replace(htmlForDB, "&" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & "", "")
htmlForDB = Replace(htmlForDB, "?" & MailzineTrakingTag(replace(yyyymmdd,".",""),omail.FOneItem.Fregtype) & "", "")

DIM cmd
if (typeGubun = "code") then
	SET cmd = Server.CreateObject("ADODB.Command")
	SET cmd.ActiveConnection = dbget

	'Prepare the stored procedure
	cmd.CommandText = "[db_sitemaster].[dbo].[sp_Ten_Mailzine_SaveHTML]"
	cmd.CommandType = 4  'adCmdStoredProc

	cmd.Parameters.Append cmd.CreateParameter("@IDX", adInteger, adParamInput, , idx)
	cmd.Parameters.Append cmd.CreateParameter("@mailHTML", adLongVarChar, adParamInput, 100000, html2db(htmlForDB))

	cmd.Execute
end if

function PrintErrorAndStop(msg)
	response.write msg
	dbget.close() : response.end
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
