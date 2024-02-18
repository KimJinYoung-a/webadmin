<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/lib/smsLib.asp"-->
<!-- #include virtual="/apps/academy/lib/maillib.asp"-->
<!-- #include virtual="/apps/academy/lib/mailLib2.asp"-->
<!-- #include virtual="/apps/academy/lib/mailFunc_Designer.asp"-->
<!-- #include virtual="/apps/academy/ordermaster/popup/misendcls.asp"-->
<%
'개발서버인 경우 메일/SMS 발송이 안되도록 되어 있다.
Dim SENDMAIL_YN
if (application("Svr_Info")	= "Dev") then
SENDMAIL_YN = "N"		'Y 인 경우 개발서버에서도 이메일을 발송하게 한다.
Else
SENDMAIL_YN = "Y"
End If

Dim sqlStr,ix, mibeasongSoldOutExists, AssignedRow, GetOrderStateNum
Dim orderserial, MakerID, FailRow, mode, ipgodate, MisendReason, itemSoldOut
orderserial=requestCheckVar(Request.Form("orderserial"),12)
mode=requestCheckVar(Request.Form("mode"),12)
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
If mode="edit" Then SENDMAIL_YN="N" '수정일 경우 메일 발송 안함
MisendReason= requestCheckVar(Request.Form("MisendReason"),2)
ipgodate = requestCheckVar(Request.Form("ipgodate"),10)
itemSoldOut = RequestCheckVar(request("itemSoldOut"),4)

'배열로 처리
ReDim DetailIDX(Request.Form("detailidx").Count)
ReDim Sitemid(Request.Form("Sitemid").Count)
ReDim Sitemoption(Request.Form("Sitemoption").Count)
For ix=1 To Request.Form("detailidx").Count
	DetailIDX(ix) = Request.Form("detailidx")(ix)
	Sitemid(ix) = Request.Form("Sitemid")(ix)
	Sitemoption(ix) = Request.Form("Sitemoption")(ix)
Next

if (mode="misendInput") then
    ''출고 지연 아니면 ipgodate 널
    dim ckSendSMS, ckSendEmail, ckSendCall, sendState, optSoldOut

    sendState = "2"

    ''업체인경우
    if (MisendReason="05") then
        ipgodate    = ""
        ckSendSMS   = "N"
        ckSendEmail = "N"
        ckSendCall  = "N"
    else
        sendState = "4"

        ckSendSMS   = "Y"
        ckSendEmail = "Y"
        ckSendCall  = "N"
    end if

	'DB에 처리
	For ix=1 To Request.Form("detailidx").Count
		If DetailIDX(ix)<>"" And Sitemid(ix)<>"" And Sitemoption(ix)<>"" Then
			if (MisendReason="05") Then
				if (Sitemid(ix)<>"") and (Sitemoption(ix)<>"") then
					if (Sitemoption(ix)="0000") then
						sqlStr = " update db_academy.dbo.tbl_diy_item" & VbCrlf
						sqlStr = sqlStr & " set sellyn='" & itemSoldOut & "'" & VbCrlf
						sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
						sqlStr = sqlStr & " where itemid=" & Sitemid(ix)

						dbACADEMYget.Execute sqlStr
					else
						optSoldOut = "N"

						sqlStr = "update [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
						sqlStr = sqlStr + " set isusing='" + optSoldOut + "'" + VBCrlf
						sqlStr = sqlStr + " , optsellyn='" + optSoldOut + "'" + VBCrlf
						sqlStr = sqlStr + " where itemid=" + CStr(Sitemid(ix))
						sqlStr = sqlStr + " and itemoption='" + Trim(Sitemoption(ix)) + "'"

						dbACADEMYget.Execute sqlStr

						''옵션갯수
						sqlStr = "update [db_academy].[dbo].tbl_diy_item" + VBCrlf
						sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
						sqlStr = sqlStr + " from (" + VBCrlf
						sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
						sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
						sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid(ix)) + VBCrlf
						sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
						sqlStr = sqlStr + " ) T" + VBCrlf
						sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item.itemid=" + CStr(Sitemid(ix)) + VBCrlf
						dbACADEMYget.Execute sqlStr

						''상품한정수량
						sqlStr = "update [db_academy].[dbo].tbl_diy_item" + VBCrlf
						sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
						sqlStr = sqlStr + " from (" + VBCrlf
						sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
						sqlStr = sqlStr + " 	from [db_academy].[dbo].tbl_diy_item_option" + VBCrlf
						sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid(ix)) + VBCrlf
						sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
						sqlStr = sqlStr + " ) T" + VBCrlf
						sqlStr = sqlStr + " where [db_academy].[dbo].tbl_diy_item.itemid=" + CStr(Sitemid(ix)) + VBCrlf
						sqlStr = sqlStr + " and [db_academy].[dbo].tbl_diy_item.optioncnt>0"

						dbACADEMYget.Execute sqlStr

						'' 한정 판매 0 이면 일시 품절 처리
						sqlStr = " update [db_academy].[dbo].tbl_diy_item "
						sqlStr = sqlStr + " set sellyn='S'"
						sqlStr = sqlStr + " where itemid=" + CStr(Sitemid(ix)) + " "
						sqlStr = sqlStr + " and sellyn='Y'"
						sqlStr = sqlStr + " and limityn='Y'"
						sqlStr = sqlStr + " and limitno-limitSold<1"

						dbACADEMYget.Execute sqlStr

						'' 판매중인 옵션이 없으면 품절처리
						sqlStr = " update [db_academy].[dbo].tbl_diy_item "
						sqlStr = sqlStr + " set sellyn='N'"
						sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
						sqlStr = sqlStr + " where itemid=" + CStr(Sitemid(ix)) + " "
						sqlStr = sqlStr + " and optioncnt=0"

						dbACADEMYget.Execute sqlStr

					end if
				end if
			end if

			sqlStr = " IF Exists(select idx from [db_academy].dbo.tbl_academy_mibeasong_list where detailidx=" & DetailIDX(ix) & ")"
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "	    update [db_academy].dbo.tbl_academy_mibeasong_list"
			sqlStr = sqlStr + "	    set code='" & MisendReason & "'"
			if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
				sqlStr = sqlStr + "	    ,state='"&sendState&"'"                                         ''상태 변경 (기존 안내메일완료)
				sqlStr = sqlStr + "	    ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS발송완료
				sqlStr = sqlStr + "	    ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email발송완료
				'''sqlStr = sqlStr + "	    ,isSendCall=(CASE WHEN isSendCall='Y' then 'Y' ELSE '"&ckSendCall&"' END)"  '' CALL완료 : 따로 처리
			end if
			if (ipgodate<>"") then
				sqlStr = sqlStr + "	,ipgodate='" & ipgodate & "'"
			else
				sqlStr = sqlStr + "	,ipgodate=NULL"
			end if
			sqlStr = sqlStr + "	    where detailidx=" & DetailIDX(ix)
			sqlStr = sqlStr + " END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + "	    insert into [db_academy].dbo.tbl_academy_mibeasong_list"
			sqlStr = sqlStr + "	    (detailidx, orderserial, itemid, itemoption,"
			sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
			if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
				sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"             ''상태 변경 (기존 안내메일완료)
				''sqlStr = sqlStr + "	isSendCall,"
			end if
			sqlStr = sqlStr + "	    itemname, itemoptionname)"
			sqlStr = sqlStr + "	    select detailidx, orderserial, itemid,itemoption,"
			sqlStr = sqlStr + "	    itemno, itemno, '" & MisendReason & "',"
			if (ipgodate<>"") then
				sqlStr = sqlStr + "	'" & ipgodate & "','',"
			else
				sqlStr = sqlStr + "	NULL,'',"
			end if
			if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
				sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
				''sqlStr = sqlStr + "	 '"&ckSendCall&"',"
			end if
			sqlStr = sqlStr + "	    itemname, itemoptionname"
			sqlStr = sqlStr + "	    from [db_academy].[dbo].tbl_academy_order_detail"
			sqlStr = sqlStr + "	    where detailidx=" & DetailIDX(ix)
			sqlStr = sqlStr + " END "
			dbACADEMYget.Execute sqlStr

			''SMS 발송 + [CS메모에 저장 -> 같이 되어있음.]
			if (ckSendSMS="Y") then
				if (SENDMAIL_YN = "Y") then
					call SendMiChulgoSMS(DetailIDX(ix))
				end if
			end if
			''EMail발송
			if (ckSendEmail="Y") then
				if (SENDMAIL_YN = "Y") then
					call fcSendMail_SendMiChulgoMail(DetailIDX(ix))
				end if
			end if
		End If
	Next
End If
%>
<script>
<!--
parent.fnMisendReasonInputEnd();
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->