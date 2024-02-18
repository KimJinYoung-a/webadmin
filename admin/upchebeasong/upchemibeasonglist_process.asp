<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 미배송리스트
' History : 이상구 생성
'			2019.01.16 한용민 수정(미출고구분 수기처리 -> 디비화 시킴. 미출고 알림톡 추가.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim refer, tmpcurrstate, sqlStr, i, mode, arrdetailidx, arrbaljudate, chulgomayday, reguserid
dim page, menupos, makerid, sitename, itemid, Dtype, yyyy1, mm1, dd1, yyyy2, mm2, dd2, buyhp, Itemname
dim cdl, dplusOver, MisendReason, MisendState, currState, beasongneedday, beasong_comment, failText
dim maketoorderyn, stockshortyn, reipgostartday, reipgoendday, reipgotype, chulgomaydaystring
dim item_yyyy1,item_yyyy2,item_mm1,item_mm2,item_dd1,item_dd2, osms, oemail, chulgodeleygubun
dim dplusLower, exinmaychulgoday, exinneedchulgoday, sortby, mcancelyn, dcancelyn, failtitle, fullText, btnJson
dim regMisendReason, regMisendipgostartdate, regMisendipgoenddate, regbeasongneedday, prevStateStr
dim IsMisendReasonInserted, detailidx, baljudate, regbeasongdaytype, currorderserial, lastorderserial
dim prevcode, previpgodate, orderserial, preisSendSMS, preisSendEmail, sendsmsmsg, sendmailmsg
dim ckSendSMS, ckSendEmail, ckSendCall, sendState, Sitemid, Sitemoption, itemSoldOut, optSoldOut
dim chulgo_yyyy1,chulgo_yyyy2,chulgo_mm1,chulgo_mm2,chulgo_dd1,chulgo_dd2, chulgoone_yyyy1, chulgoone_mm1
dim tmp_chulgomayday, tmp_chulgomaydaystring, tmp_sendsmsmsg, tmp_sendmailmsg, oneMisend, chulgoone_dd1

reguserid = session("ssBctId")

mode = requestCheckVar(request("mode"),32)
page = requestCheckVar(request("page"),32)
menupos = requestCheckVar(request("menupos"),32)
makerid = requestCheckVar(request("makerid"),32)
sitename = requestCheckVar(request("sitename"),32)
itemid = requestCheckVar(request("itemid"),32)
Dtype = requestCheckVar(request("Dtype"),32)
yyyy1 = requestCheckVar(request("yyyy1"),32)
mm1 = requestCheckVar(request("mm1"),32)
dd1 = requestCheckVar(request("dd1"),32)
yyyy2 = requestCheckVar(request("yyyy2"),32)
mm2 = requestCheckVar(request("mm2"),32)
dd2 = requestCheckVar(request("dd2"),32)
cdl = requestCheckVar(request("cdl"),32)
dplusOver = requestCheckVar(request("dplusOver"),32)
MisendReason = requestCheckVar(request("MisendReason"),32)
MisendState = requestCheckVar(request("MisendState"),32)
currState = requestCheckVar(request("currState"),32)

beasongneedday = requestCheckVar(request("beasongneedday"),32)
beasong_comment = requestCheckVar(html2db(request("beasong_comment")),8000)

dplusLower = requestCheckVar(request("dplusLower"),32)
exinmaychulgoday = requestCheckVar(request("exinmaychulgoday"),32)
exinneedchulgoday = requestCheckVar(request("exinneedchulgoday"),32)
sortby = requestCheckVar(request("sortby"),32)

maketoorderyn = requestCheckVar(request("maketoorderyn"),32)
stockshortyn = requestCheckVar(request("stockshortyn"),32)
reipgotype = requestCheckVar(request("reipgotype"),32)

item_yyyy1 = requestCheckVar(request("item_yyyy1"),32)
item_yyyy2 = requestCheckVar(request("item_yyyy2"),32)
item_mm1 = requestCheckVar(request("item_mm1"),32)
item_mm2 = requestCheckVar(request("item_mm2"),32)
item_dd1 = requestCheckVar(request("item_dd1"),32)
item_dd2 = requestCheckVar(request("item_dd2"),32)

regMisendReason = requestCheckVar(request("regMisendReason"),32)
regMisendipgostartdate = requestCheckVar(request("regMisendipgostartdate"),32)
regMisendipgoenddate = requestCheckVar(request("regMisendipgoenddate"),32)
regbeasongneedday = requestCheckVar(request("regbeasongneedday"),32)

arrdetailidx = requestCheckVar(request("arrdetailidx"),8000)
arrbaljudate = requestCheckVar(request("arrbaljudate"),8000)

if (reipgotype = "1") then
	item_yyyy2= item_yyyy1
	item_mm2 = item_mm1
	item_dd2 = item_dd1
end if

reipgostartday = item_yyyy1 + "-" + item_mm1 + "-" + item_dd1
reipgoendday = item_yyyy2 + "-" + item_mm2 + "-" + item_dd2

refer = "upchemibeasonglist.asp?menupos" + CStr(menupos) + _
									"&page=" + CStr(page) + _
									"&makerid=" + CStr(makerid) + _
									"&sitename=" + CStr(sitename) + _
									"&itemid=" + CStr(itemid) + _
									"&Dtype=" + CStr(Dtype) + _
									"&yyyy1=" + CStr(yyyy1) + _
									"&mm1=" + CStr(mm1) + _
									"&dd1=" + CStr(dd1) + _
									"&yyyy2=" + CStr(yyyy2) + _
									"&mm2=" + CStr(mm2) + _
									"&dd2=" + CStr(dd2) + _
									"&cdl=" + CStr(cdl) + _
									"&dplusOver=" + CStr(dplusOver) + _
									"&dplusLower=" + CStr(dplusLower) + _
									"&exinmaychulgoday=" + CStr(exinmaychulgoday) + _
									"&exinneedchulgoday=" + CStr(exinneedchulgoday) + _
									"&sortby=" + CStr(sortby) + _
									"&MisendReason=" + CStr(MisendReason) + _
									"&MisendState=" + CStr(MisendState) + _
									"&currState=" + CStr(currState)

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

preisSendSMS="N"
preisSendEmail="N"

if mode="getMisendReason" then
	if regMisendReason="" then dbget.close() : response.end

	set osms = New CCSTemplate
		osms.FRectMasterGubun="40"	' 문자
		osms.FRectGubun=regMisendReason
		osms.GetCSTemplateone

	set oemail = New CCSTemplate
		oemail.FRectMasterGubun="41"	' 이메일
		oemail.FRectGubun=regMisendReason
		oemail.GetCSTemplateone

	response.write "<script type='text/javascript'>"
	if osms.FTotalCount>0 then
		response.write "	parent.frmMisendInput.sendsmsmsg.value='" & replace(osms.FOneItem.fcontents,vbcrlf,"\n") & "';"
	end if
	if oemail.FTotalCount>0 then
		response.write "	parent.frmMisendInput.sendmailmsg.value='" & replace(oemail.FOneItem.fcontents,vbcrlf,"\n") & "';"
	end if
	response.write "</script>"

	set osms=nothing
	set oemail=nothing
	dbget.close() : response.end

elseif (mode = "modifybrandmemo") then

	sqlStr = " IF EXISTS (SELECT brandid FROM [db_cs].[dbo].tbl_cs_brand_memo WHERE brandid = '" + CStr(makerid) + "') "
	sqlStr = sqlStr & "	update [db_cs].[dbo].tbl_cs_brand_memo set beasong_modifyday = getdate() "
	sqlStr = sqlStr & " ,beasongneedday = '" & beasongneedday & "' "
	sqlStr = sqlStr & " ,beasong_comment = '" & beasong_comment & "' "
	sqlStr = sqlStr & " ,beasong_reguserid = '" & reguserid & "' "


	sqlStr = sqlStr & " where brandid = '" & makerid & "' "
	sqlStr = sqlStr & " ELSE "
	sqlStr = sqlStr & " insert into [db_cs].[dbo].tbl_cs_brand_memo(brandid, beasongneedday, beasong_comment, beasong_modifyday, beasong_reguserid) "
	sqlStr = sqlStr & "  values('" & makerid & "', " & beasongneedday & ", '" & beasong_comment & "', getdate(), '" & reguserid & "') "
	rsget.Open sqlStr,dbget,1

elseif (mode = "modifyitemmemo") then

	sqlStr = " IF EXISTS (SELECT itemid FROM [db_cs].[dbo].tbl_cs_item_memo WHERE itemid = " + CStr(itemid) + ") "
	sqlStr = sqlStr & "	update [db_cs].[dbo].tbl_cs_item_memo set beasong_modifyday = getdate() "
	sqlStr = sqlStr & " ,beasongneedday = '" & beasongneedday & "' "
	sqlStr = sqlStr & " ,beasong_comment = '" & beasong_comment & "' "
	sqlStr = sqlStr & " ,maketoorderyn = '" & maketoorderyn & "' "
	sqlStr = sqlStr & " ,stockshortyn = '" & stockshortyn & "' "
	sqlStr = sqlStr & " ,reipgostartday = '" & reipgostartday & "' "
	sqlStr = sqlStr & " ,reipgoendday = '" & reipgoendday & "' "
	sqlStr = sqlStr & " ,beasong_reguserid = '" & reguserid & "' "
	sqlStr = sqlStr & " where itemid = '" & itemid & "' "
	sqlStr = sqlStr & " ELSE "
	sqlStr = sqlStr & " insert into [db_cs].[dbo].tbl_cs_item_memo(itemid, beasongneedday, beasong_comment, beasong_modifyday, beasong_reguserid, maketoorderyn, stockshortyn, reipgostartday, reipgoendday) "
	sqlStr = sqlStr & "  values(" & itemid & ", " & beasongneedday & ", '" & beasong_comment & "', getdate(), '" & reguserid & "', '" & maketoorderyn & "', '" & stockshortyn & "', '" & reipgostartday & "', '" & reipgoendday & "') "
	rsget.Open sqlStr,dbget,1

elseif (mode = "regallmisendreason") then
    ''품절출고불가만 ipgodate 널

	regbeasongdaytype = request("regbeasongdaytype")
	chulgo_yyyy1 = request("chulgo_yyyy1")
	chulgo_yyyy2 = request("chulgo_yyyy2")
	chulgo_mm1 = request("chulgo_mm1")
	chulgo_mm2 = request("chulgo_mm2")
	chulgo_dd1 = request("chulgo_dd1")
	chulgo_dd2 = request("chulgo_dd2")

	chulgoone_yyyy1 = request("chulgoone_yyyy1")
	chulgoone_mm1 = request("chulgoone_mm1")
	chulgoone_dd1 = request("chulgoone_dd1")

	sendsmsmsg = request("sendsmsmsg")
	sendmailmsg = request("sendmailmsg")

	chulgomayday = ""
	chulgomaydaystring = ""

	if (regbeasongdaytype = "datearea") then
		if (chulgo_yyyy1 = chulgo_yyyy2) and (chulgo_mm1 = chulgo_mm2) and (chulgo_dd1 = chulgo_dd2) then
			chulgomayday = GetProperDate(chulgo_yyyy1, chulgo_mm1, chulgo_dd1)
			chulgomaydaystring = chulgomayday + " "
		else
			chulgomayday = GetProperDate(chulgo_yyyy2, chulgo_mm2, chulgo_dd2)
			chulgomaydaystring = GetProperDate(chulgo_yyyy1, chulgo_mm1, chulgo_dd1) + " ~ " + GetProperDate(chulgo_yyyy2, chulgo_mm2, chulgo_dd2) + " 중"
		end if
	elseif (regbeasongdaytype = "onedate") then
		chulgomayday = GetProperDate(chulgoone_yyyy1, chulgoone_mm1, chulgoone_dd1)
		chulgomaydaystring = chulgomayday + " "
	end if

    sendState = "2"

    ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
    ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")
    ckSendCall  = CHKIIF(request("ckSendCall")="on","Y","N")

    if (ckSendSMS="Y") or (ckSendEmail="Y") then sendState = "4"

	arrdetailidx = Split(arrdetailidx, ",")
	arrbaljudate = Split(arrbaljudate, ",")

	currorderserial = ""
	lastorderserial = ""
	for i = 0 to UBound(arrdetailidx)

		if (Trim(arrdetailidx(i)) <> "") then

			currorderserial = ""
			tmp_sendsmsmsg = sendsmsmsg
			tmp_sendmailmsg = sendmailmsg

			' 현재주문상태 체크. 출고완료된 주문이 미출고로 처리가 되고 고객에게 문자가 발송됨.		' 2019.09.04 한용민
			tmpcurrstate=""
			mcancelyn=""
			dcancelyn=""
			sqlStr = "select m.orderserial, m.ipkumdiv, d.currstate, d.idx, m.cancelyn as mcancelyn, d.cancelyn as dcancelyn" & vbcrlf
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
			sqlStr = sqlStr & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
			sqlStr = sqlStr & " 	on m.idx = d.masteridx" & vbcrlf
			sqlStr = sqlStr & " where d.idx = "& Trim(arrdetailidx(i)) &"" & vbcrlf

			'response.write sqlStr & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly
			if Not rsget.Eof then
				tmpcurrstate = rsget("currstate")
				mcancelyn = rsget("mcancelyn")
				dcancelyn = rsget("dcancelyn")
			end if
			rsget.close

			' 디테일 출고완료
			if tmpcurrstate="7" then
%>
				<script type='text/javascript'>
					alert('이미 출고된 주문 입니다.');
					<% if refer<>"" then %>
					location.replace('<%= refer %>');
					<% end if %>
				</script>
<%
				dbget.close() : response.end

			' 취소된 주문
			elseif mcancelyn="Y" or dcancelyn="Y" then
%>
				<script type='text/javascript'>
					alert('이미 취소된 주문 입니다.');
					<% if refer<>"" then %>
					location.replace('<%= refer %>');
					<% end if %>
				</script>
<%
				dbget.close() : response.end
			end if

			if chulgomayday = "" then
	            sqlStr = " exec [db_cs].[dbo].[usp_getDayPlusWorkday] '" & Trim(arrbaljudate(i)) & "', " & regbeasongneedday & " " & VbCRLF
	            rsget.CursorLocation = adUseClient
	            rsget.Open sqlStr, dbget, adOpenForwardOnly
	        	if Not rsget.Eof then
	                tmp_chulgomayday = rsget("plusworkday")
	                tmp_chulgomaydaystring = tmp_chulgomayday + " "
	            end if
	        	rsget.close
			else
				tmp_chulgomayday = chulgomayday
				tmp_chulgomaydaystring = chulgomaydaystring
			end if

			if (regMisendReason = "05") then
				response.write "에러 : 시스템팀 문의"
				response.end
			end if

			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[출고예정일]", tmp_chulgomaydaystring)
			tmp_sendsmsmsg = Replace(tmp_sendsmsmsg, "[주문통보일+N일]", tmp_chulgomaydaystring)

			sqlStr = "select top 1 orderserial, itemname, IsNull(itemoptionname, '') as itemoptionname, code, IsNull(isSendSms, 'N') as isSendSms"
			sqlStr = sqlStr & " , IsNull(isSendEmail, 'N') as isSendEmail, IsNull(isSendCall, '') as isSendCall, isnull(convert(varchar(10),ipgodate,121),'') as ipgodate"
			sqlStr = sqlStr & " from [db_temp].dbo.tbl_mibeasong_list with (nolock) where detailidx=" & Trim(arrdetailidx(i)) & " "

			'response.write sqlStr & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly

			IsMisendReasonInserted = False
			if Not rsget.Eof then
				IsMisendReasonInserted = True
				prevcode = rsget("code")
				previpgodate = rsget("ipgodate")
				preisSendSMS = rsget("isSendSMS")
				preisSendEmail = rsget("isSendEmail")
				orderserial = rsget("orderserial")

				prevStateStr = "기존 미출고사유" + vbCrLf
				prevStateStr = prevStateStr + "상품명 : " + CStr(rsget("itemname"))
				prevStateStr = prevStateStr + "[" + CStr(rsget("itemoptionname")) + "]" + vbCrLf
				prevStateStr = prevStateStr + "미출고사유 : " + MiSendCodeToName(rsget("code")) + vbCrLf
				prevStateStr = prevStateStr + "고객알림 : SMS(" + CStr(rsget("isSendSms")) + "), 이메일(" + CStr(rsget("isSendEmail")) + "), 통화(" + CStr(rsget("isSendCall")) + ")" + vbCrLf
				prevStateStr = prevStateStr + "처리예정일 : " + CStr(rsget("ipgodate"))
			end if
			rsget.close

		    sqlStr = " IF Exists(select idx from [db_temp].dbo.tbl_mibeasong_list where detailidx=" & Trim(arrdetailidx(i)) & ")"
		    sqlStr = sqlStr + " BEGIN "
		    sqlStr = sqlStr + "	    update [db_temp].dbo.tbl_mibeasong_list"
		    sqlStr = sqlStr + "	    set code='" & regMisendReason & "'"

		    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
		        sqlStr = sqlStr + "	    ,state='"&sendState&"'"                                         ''상태 변경 (기존 안내메일완료)
		        sqlStr = sqlStr + "	    ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS발송완료
		        sqlStr = sqlStr + "	    ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email발송완료
				sqlStr = sqlStr + "		,sendCount=IsNull(sendCount,0) + 1 "
				sqlStr = sqlStr + "		,lastSendUserid='" + CStr(session("ssBctId")) + "' "
				sqlStr = sqlStr + "		,lastSendDate=getdate() "
		    end if

		    sqlStr = sqlStr + "	,ipgodate='" & tmp_chulgomayday & "'"
			sqlStr = sqlStr + "	,modiuserid = '" + CStr(session("ssBctId")) + "' "
			sqlStr = sqlStr + "	,modidate = getdate() "
		    sqlStr = sqlStr + "	    where detailidx=" & Trim(arrdetailidx(i))
		    sqlStr = sqlStr + " END "
		    sqlStr = sqlStr + " ELSE "
		    sqlStr = sqlStr + " BEGIN "
		    sqlStr = sqlStr + "	    insert into [db_temp].dbo.tbl_mibeasong_list"
		    sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
		    sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "

		    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
		        sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"
				sqlStr = sqlStr + "	sendCount, lastSendUserid, lastSendDate, "
		    end if

		    sqlStr = sqlStr + "	    itemname, itemoptionname, reguserid)"
		    sqlStr = sqlStr + "	    select idx, orderserial, itemid,itemoption,"
		    sqlStr = sqlStr + "	    itemno, itemno, '" & regMisendReason & "',"

			sqlStr = sqlStr + "	'" & tmp_chulgomayday & "','',"
		    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
		        sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
				sqlStr = sqlStr + "	1, '" + CStr(session("ssBctId")) + "', getdate(), "
		    end if
		    sqlStr = sqlStr + "	    itemname, itemoptionname, '" + CStr(session("ssBctId")) + "' "
		    sqlStr = sqlStr + "	    from [db_order].[dbo].tbl_order_detail"
		    sqlStr = sqlStr + "	    where idx=" & Trim(arrdetailidx(i))
		    sqlStr = sqlStr + " END "
			''rw   sqlStr
		    dbget.Execute sqlStr

		    ''SMS 발송 + [CS메모에 저장 -> 같이 되어있음.]
		    if (ckSendSMS="Y") then
		        'if (application("Svr_Info")<>"Dev") then

					'// 중복 SMS 발송 제한
		            sqlStr = " select top 1 orderserial from [db_order].[dbo].tbl_order_detail with (nolock) where idx = " & Trim(arrdetailidx(i)) & " " & VbCRLF
		            rsget.CursorLocation = adUseClient
		            rsget.Open sqlStr, dbget, adOpenForwardOnly
		        	if Not rsget.Eof then
		                currorderserial = rsget("orderserial")
		            end if
		        	rsget.close

					if (currorderserial <> lastorderserial) then
						'// TODO : 상품코드 별로 발송은 하지 않는다.

						lastorderserial = currorderserial

						' 출고지연. 카카오톡 알림톡 발송.   ' 2021.09.17 한용민 생성
						if regMisendReason = "03" then
							set oneMisend = new COldMiSend
								oneMisend.FRectDetailIDx = Trim(arrdetailidx(i))
								oneMisend.getOneOldMisendItem

							buyhp = oneMisend.FOneItem.FBuyHP
							Itemname = replace(oneMisend.FOneItem.FItemname,vbcrlf,"")

							if buyhp<>"" and not(isnull(buyhp)) then
								chulgodeleygubun=""
								sqlStr = "select"
								sqlStr = sqlStr & " l.idx"
								sqlStr = sqlStr & " , (case"
								sqlStr = sqlStr & "     when isnull(l.prevcode,'00')='05' and '"& prevcode &"'<>'03' and convert(varchar(10),ipgodate,121)<>'"& previpgodate &"' then '05_03'"   ' 품절출고불가 상품 출고지연전환 케이스. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
								sqlStr = sqlStr & "     when isnull(l.prevcode,'00')<>'03' and '"& regMisendReason &"'='03' and convert(varchar(10),ipgodate,121)<>'"& previpgodate &"' then '03'"    ' 출고지연 알림톡. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
								sqlStr = sqlStr & "     when '"& prevcode &"'<>'03' and '"& regMisendReason &"'='03' then '03'"   ' 다른사유를 출고지연전환 케이스. 중복 발송을 제거하기 위해 출고예정일이 틀린경우에만 발송.
								sqlStr = sqlStr & "     when '"& prevcode &"'='"& regMisendReason &"' and '"& preisSendSMS &"'='N' and '"& preisSendEmail &"'='N' then '03'"	' 사유는 같으나 알림 발송을 위해 재저장 버튼을 누른 케이스
								sqlStr = sqlStr & "     else '' end) as chulgodeleygubun"
								sqlStr = sqlStr & " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
								sqlStr = sqlStr & " where l.code = '03'"	' 출고지연
								sqlStr = sqlStr & " and l.ipgodate is not null"
								sqlStr = sqlStr & " and l.detailidx="& Trim(arrdetailidx(i)) &""

								'response.write sqlStr & "<br>"
								rsget.CursorLocation = adUseClient
								rsget.Open sqlStr, dbget, adOpenForwardOnly
								if Not rsget.Eof then
									chulgodeleygubun = rsget("chulgodeleygubun")
								end if
								rsget.close

								' 품절출고불가 상품 출고지연전환 케이스
								if chulgodeleygubun="05_03" then
									failtitle = "[텐바이텐]상품출고 안내"
									fullText = "[10x10] 상품출고 안내" & vbCrLf & vbCrLf
									fullText = fullText & "품절취소 안내드렸던 상품의 재고가 확보되어 발송 예정으로, 아래의 예정일까지 출발할 수 있도록 최선의 노력을 다하겠습니다." & vbCrLf & vbCrLf & vbCrLf
									fullText = fullText & "■ 주문번호 : "& oneMisend.FOneItem.Forderserial &"" & vbCrLf
									fullText = fullText & "■ 상품명 : "& Itemname &"" & vbCrLf
									fullText = fullText & "■ 출발예정일 : "& tmp_chulgomayday &"" & vbCrLf & vbCrLf
									fullText = fullText & "감사합니다."
									failText = fullText
									btnJson = "{""button"":[{""name"":""주문내역 바로가기"",""type"":""WL"", ""url_mobile"":""https://tenten.app.link/L1izHiDBdjb""}]}"
									call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0024",fullText,"LMS",failtitle,failText,btnJson,"",oneMisend.FOneItem.Fuserid)

									sqlStr = "update db_temp.dbo.tbl_mibeasong_list set finishstr=N'품절상품 출고지연전환 알림톡 발송완료' where detailidx="& Trim(arrdetailidx(i)) &"" & vbcrlf
									dbget.Execute sqlStr

									Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[품절상품 출고지연전환 알림톡 발송완료 "+ buyhp +"]" + html2db(fullText))

								' 출고지연 알림톡
								elseif chulgodeleygubun="03" then
									failtitle = "[텐바이텐]출고 지연 안내"
									fullText = "[10x10] 출고 지연 안내" & vbCrLf & vbCrLf
									fullText = fullText & "출고지연으로 양해의 말씀 드립니다." & vbCrLf
									fullText = fullText & "주문하신 소중한 상품의 배송지연이 예상되오며," & vbCrLf
									fullText = fullText & "아래의 예정일까지 출발할 수 있도록 최선의 노력을 다하겠습니다." & vbCrLf & vbCrLf
									fullText = fullText & "■ 주문번호 : "& oneMisend.FOneItem.Forderserial &"" & vbCrLf
									fullText = fullText & "■ 상품명 : "& Itemname &"" & vbCrLf
									fullText = fullText & "■ 출발예정일 : "& tmp_chulgomayday &"" & vbCrLf & vbCrLf
									fullText = fullText & "감사합니다."
									failText = fullText
									btnJson = ""
									call SendKakaoCSMsg_LINK("", buyhp,"1644-6030","KC-0009",fullText,"LMS",failtitle,failText,btnJson,oneMisend.FOneItem.Forderserial,oneMisend.FOneItem.Fuserid)

									sqlStr = "update db_temp.dbo.tbl_mibeasong_list set finishstr=N'출고지연 알림톡 발송완료' where detailidx="& Trim(arrdetailidx(i)) &"" & vbcrlf
									dbget.Execute sqlStr

									Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송완료 "+ buyhp +"]" + html2db(fullText))
								else
									' 오류시노티
									Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송실패.미출고사유코드:"& chulgodeleygubun &".이전사유코드:"& prevcode &".이전입고예정일:"&previpgodate&"]")
								end if
							else
								' 오류시노티
								Call AddCsMemo(oneMisend.FOneItem.Forderserial,"1",oneMisend.FOneItem.Fuserid,session("ssBctId"),"[출고지연 알림톡 발송실패.주문자휴대폰번호:"& buyhp &"]")
							end if
							set oneMisend = nothing

						' 문자 발송
						else
							Call SendMiChulgoSMSWithMessage(Trim(arrdetailidx(i)), tmp_sendsmsmsg)
						end if
					end if

		        'end if
		    end if

		    ''EMail발송
		    if (ckSendEmail="Y") then
		        if (application("Svr_Info")<>"Dev") then
		            Call SendMiChulgoMailWithMessage(Trim(arrdetailidx(i)), tmp_sendmailmsg)
		        end if
		    end if

		end if

	next

    if (ckSendSMS="Y") and (ckSendEmail="Y") then
        response.write "<script type='text/javascript'>alert('SMS와 메일이 발송 되었습니다.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script type='text/javascript'>alert('SMS가 발송 되었습니다.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script type='text/javascript'>alert('메일이 발송 되었습니다.');</script>"
    else
        response.write "<script type='text/javascript'>alert('처리 되었습니다.');</script>"
    end if

elseif (mode = "regallmisendstockout") then

    arrdetailidx = Split(arrdetailidx, ",")

	for i = 0 to UBound(arrdetailidx)

		if (Trim(arrdetailidx(i)) <> "") then
			tmpcurrstate=""
			mcancelyn=""
			dcancelyn=""
			sqlStr = "select m.orderserial, m.ipkumdiv, d.currstate, d.idx, m.cancelyn as mcancelyn, d.cancelyn as dcancelyn" & vbcrlf
			sqlStr = sqlStr & " from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
			sqlStr = sqlStr & " join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
			sqlStr = sqlStr & " 	on m.orderserial = d.orderserial" & vbcrlf
			sqlStr = sqlStr & " where d.idx = "& Trim(arrdetailidx(i)) &"" & vbcrlf

			'response.write sqlStr & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly
			if Not rsget.Eof then
				tmpcurrstate = rsget("currstate")
				mcancelyn = rsget("mcancelyn")
				dcancelyn = rsget("dcancelyn")
			end if
			rsget.close

            if tmpcurrstate="7" then
                '
            elseif mcancelyn="Y" or dcancelyn="Y" then
                '
            else
	            sqlStr = "select top 1 orderserial, itemname, IsNull(itemoptionname, '') as itemoptionname, code, IsNull(isSendSms, '') as isSendSms, IsNull(isSendEmail, '') as isSendEmail, IsNull(isSendCall, '') as isSendCall, IsNull(ipgodate, '') as ipgodate  "
	            sqlStr = sqlStr + " from [db_temp].dbo.tbl_mibeasong_list where detailidx=" + Trim(arrdetailidx(i)) + " "
	            rsget.CursorLocation = adUseClient
	            rsget.Open sqlStr, dbget, adOpenForwardOnly

	            IsMisendReasonInserted = False
	            if Not rsget.Eof then
		            IsMisendReasonInserted = True
		            prevcode = rsget("code")
		            orderserial = rsget("orderserial")

		            prevStateStr = "기존 미출고사유" + vbCrLf
		            prevStateStr = prevStateStr + "상품명 : " + CStr(rsget("itemname"))
		            prevStateStr = prevStateStr + "[" + CStr(rsget("itemoptionname")) + "]" + vbCrLf
		            prevStateStr = prevStateStr + "미출고사유 : " + MiSendCodeToName(rsget("code")) + vbCrLf
		            prevStateStr = prevStateStr + "고객알림 : SMS(" + CStr(rsget("isSendSms")) + "), 이메일(" + CStr(rsget("isSendEmail")) + "), 통화(" + CStr(rsget("isSendCall")) + ")" + vbCrLf
		            prevStateStr = prevStateStr + "처리예정일 : " + CStr(rsget("ipgodate"))
	            end if
	            rsget.close

                if Not IsMisendReasonInserted then
		            sqlStr = sqlStr + "	insert into [db_temp].dbo.tbl_mibeasong_list"
		            sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
		            sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
		            sqlStr = sqlStr + "	    itemname, itemoptionname,reqaddstr, reguserid)"
		            sqlStr = sqlStr + "	select idx, orderserial, itemid,itemoption,"
		            sqlStr = sqlStr + "	    itemno, itemno, '05',"
			        sqlStr = sqlStr + "	    NULL,'',"
		            sqlStr = sqlStr + "	    itemname, itemoptionname, NULL, '" + CStr(session("ssBctId")) + "' "
		            sqlStr = sqlStr + "	    from [db_order].[dbo].tbl_order_detail"
		            sqlStr = sqlStr + "	    where idx=" & Trim(arrdetailidx(i))
	                ''rw   sqlStr
	                dbget.Execute sqlStr

                    '// 품절출고불가 담당자 지정
		            sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " & Trim(arrdetailidx(i)) & " "
		            dbget.Execute sqlStr
                elseif prevcode = "05" then
                    '// 사유변경없음 : skip
                else
                    Call AddCsMemo(orderserial, "1", "", session("ssBctId"), prevStateStr)

		            sqlStr = sqlStr + " update [db_temp].dbo.tbl_mibeasong_list"
		            sqlStr = sqlStr + " set code='05' "
			        sqlStr = sqlStr + " , prevcode = '" + CStr(prevcode) + "' "
		            sqlStr = sqlStr + " ,state='0'"
                    sqlStr = sqlStr + " ,sendCount=0"			'// 품절 등록되면 품절알림 문자발송, 2020-02-13, skyer9
		            sqlStr = sqlStr + " ,isSendSMS='N'"
		            sqlStr = sqlStr + " ,isSendEmail='N'"
			        sqlStr = sqlStr + "	,ipgodate=NULL"
		            sqlStr = sqlStr + "	,modiuserid = '" + CStr(session("ssBctId")) + "' "
		            sqlStr = sqlStr + "	,modidate = getdate() "
		            sqlStr = sqlStr + " where detailidx=" & Trim(arrdetailidx(i))
	                ''rw   sqlStr
	                dbget.Execute sqlStr

                    '// 품절출고불가 담당자 지정
		            sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeID] " & Trim(arrdetailidx(i)) & " "
		            dbget.Execute sqlStr
                end if
            end if
        end if

    next
else

end if

function GetProperDate(yyyy, mm, dd)
	dim s, tmpdate

	s = CLng(mm)

	tmpdate = DateSerial(yyyy, mm, dd)

	do while (Month(tmpdate) <> s)
		tmpdate = DateAdd("d", -1, tmpdate)
	loop

	GetProperDate = CStr(tmpdate)

end function

%>
<script type='text/javascript'>
alert('저장 되었습니다.');
<% if referer<>"" then %>
location.replace('<%= referer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
