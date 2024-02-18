<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]������>>[CS]��ó��CS����Ʈ
' History : �̻� ����
'			2023.11.15 �ѿ�� ����(�����ϴ� ����ü���� �������� cs������ ���� �̰�)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/mailLib2.asp"-->
<!-- #include virtual="/lib/email/mailFunc_Designer.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<%
dim makerid, mode, csdetailidx, MifinishReason, itemSoldOutFlag, ipgodate, optSoldOut, sqlStr, referer
dim Sitemid, Sitemoption, ckSendSMS, ckSendEmail, ckSendCall, sendState, ischulgostate, finishmemo, asid
    mode = requestCheckVar(request.Form("mode"), 32)
    csdetailidx = Replace(request.Form("csdetailidx"), " ", "")
    MifinishReason = requestCheckVar(request.Form("MifinishReason"), 32)
    itemSoldOutFlag = requestCheckVar(request.Form("itemSoldOutFlag"), 32)
    Sitemid     = RequestCheckVar(request("Sitemid"),10)
    Sitemoption = RequestCheckVar(request("Sitemoption"),4)
    ipgodate = requestCheckVar(request.Form("ipgodate"), 32)
    ckSendSMS = requestCheckVar(request.Form("ckSendSMS"), 32)
    ckSendEmail = requestCheckVar(request.Form("ckSendEmail"), 32)
    ischulgostate = requestCheckVar(request.Form("ischulgostate"), 32)
    finishmemo  = html2db(Replace(request("finishmemo"), " ", ""))
    asid = Replace(request.Form("asid"), " ", "")

referer = request.ServerVariables("HTTP_REFERER")

if (mode = "MiFinishInputOne") then
    sendState = "2"

    ''�����ڰ��
    if (C_ADMIN_USER) then
        ckSendSMS   = CHKIIF(request("ckSendSMS")="on","Y","N")
        ckSendEmail = CHKIIF(request("ckSendEmail")="on","Y","N")
        ckSendCall  = CHKIIF(request("ckSendCall")="on","Y","N")

        if (ckSendCall="Y") then sendState = "4"

        if (MifinishReason="05") then
            ipgodate    = ""
            ckSendSMS   = "N"
            ckSendEmail = "N"
            ckSendCall  = "N"
        else
            sendState = "4"
        end if
    else
        ''��ü�ΰ��
        if (MifinishReason="05") then
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
    end if

    if (MifinishReason="05") then
        if (Sitemid<>"") and (Sitemoption<>"") then
            if (Sitemoption="0000") then
                sqlStr = " update db_item.dbo.tbl_item" & VbCrlf
                sqlStr = sqlStr & " set sellyn='" & itemSoldOut & "'" & VbCrlf
                sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
                sqlStr = sqlStr & " where itemid=" & Sitemid

                dbget.Execute sqlStr
            else
                optSoldOut = "N"

                sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optSoldOut + "'" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(Sitemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(Sitemoption) + "'"

				dbget.Execute sqlStr

				''�ɼǰ���
                sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
                sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
                sqlStr = sqlStr + " from (" + VBCrlf
                sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
                sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
                sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
                sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
                sqlStr = sqlStr + " ) T" + VBCrlf
                sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(Sitemid) + VBCrlf

                dbget.Execute sqlStr

                ''��ǰ��������
            	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
            	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
            	sqlStr = sqlStr + " from (" + VBCrlf
            	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
            	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
            	sqlStr = sqlStr + " 	where itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
            	sqlStr = sqlStr + " ) T" + VBCrlf
            	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(Sitemid) + VBCrlf
            	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"

            	dbget.Execute sqlStr

            	'' ���� �Ǹ� 0 �̸� �Ͻ� ǰ�� ó��
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='S'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " " & VbCrlf
            	sqlStr = sqlStr + " and sellyn='Y'"
            	sqlStr = sqlStr + " and limityn='Y'"
            	sqlStr = sqlStr + " and limitno-limitSold<1"

                dbget.Execute sqlStr

            	'' �Ǹ����� �ɼ��� ������ ǰ��ó��
                sqlStr = " update [db_item].[dbo].tbl_item "
            	sqlStr = sqlStr + " set sellyn='N'"
            	sqlStr = sqlStr & " ,lastupdate=getdate()" & VbCrlf
            	sqlStr = sqlStr + " where itemid=" + CStr(Sitemid) + " "
            	sqlStr = sqlStr + " and optioncnt=0"

                dbget.Execute sqlStr

            end if
        end if
    end if

    sqlStr = " IF Exists(select idx from [db_temp].dbo.tbl_csmifinish_list where csdetailidx=" & csdetailidx & ")"
    sqlStr = sqlStr + " BEGIN "
    sqlStr = sqlStr + "	    update [db_temp].dbo.tbl_csmifinish_list"
    sqlStr = sqlStr + "	    set code='" & MifinishReason & "'"
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	    ,state='"&sendState&"'"                                         ''���� ���� (���� �ȳ����ϿϷ�)
        sqlStr = sqlStr + "	    ,isSendSMS=(CASE WHEN isSendSMS='Y' then 'Y' ELSE '"&ckSendSMS&"' END)" '' SMS�߼ۿϷ�
        sqlStr = sqlStr + "	    ,isSendEmail=(CASE WHEN isSendEmail='Y' then 'Y' ELSE '"&ckSendEmail&"' END)"  '' Email�߼ۿϷ�
    end if
    if (ipgodate<>"") then
        sqlStr = sqlStr + "	,ipgodate='" & ipgodate & "'"
    else
        sqlStr = sqlStr + "	,ipgodate=NULL"
    end if

	sqlStr = sqlStr + "	, reguserid = '" + session("ssBctID") + "' "
	sqlStr = sqlStr + "	, lastupdate = getdate() "

    sqlStr = sqlStr + "	    where csdetailidx=" & csdetailidx
    sqlStr = sqlStr + " END "
    sqlStr = sqlStr + " ELSE "
    sqlStr = sqlStr + " BEGIN "
    sqlStr = sqlStr + "	    insert into [db_temp].dbo.tbl_csmifinish_list"
    sqlStr = sqlStr + "	    (csdetailidx, orderserial, itemid, itemoption,"
    sqlStr = sqlStr + "	    itemno, itemlackno, code, ipgodate, reqstr, "
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	state, isSendSMS, isSendEmail,"             ''���� ���� (���� �ȳ����ϿϷ�)
    end if
    sqlStr = sqlStr + "	    itemname, itemoptionname, reguserid, lastupdate)"
    sqlStr = sqlStr + "	    select d.id, m.orderserial, d.itemid, d.itemoption,"
    sqlStr = sqlStr + "	    d.regitemno, d.regitemno, '" & MifinishReason & "',"

    if (ipgodate<>"") then
        sqlStr = sqlStr + "	'" & ipgodate & "','',"
    else
        sqlStr = sqlStr + "	NULL,'',"
    end if
    if (ckSendSMS<>"N") or (ckSendEmail<>"N") or (ckSendCall<>"N") then
        sqlStr = sqlStr + "	 "&sendState&", '"&ckSendSMS&"', '"&ckSendEmail&"',"
    end if
    sqlStr = sqlStr + "	    d.itemname, d.itemoptionname, '" + session("ssBctID") + "', getdate() "
	sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list m with (nolock)"
	sqlStr = sqlStr & "	join db_cs.dbo.tbl_new_as_detail d with (nolock)"
	sqlStr = sqlStr & " 	on m.id = d.masterid"
    sqlStr = sqlStr + "	    where d.id=" & csdetailidx
    sqlStr = sqlStr + " END "

	'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

	' if (Not C_ADMIN_USER) then
	' 	sqlStr = "update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
	' 	sqlStr = sqlStr + " set finishuser ='" + session("ssBctID") + "'," + VbCrlf
	' 	sqlStr = sqlStr + " contents_finish ='" + finishmemo + "' " + VbCrlf
	' 	sqlStr = sqlStr + " where id =" + asid
	' 	sqlStr = sqlStr + " and makerid='" & session("ssBctID") & "'"
	' 	dbget.Execute sqlStr
	' end if

    ''SMS �߼� + [CS�޸� ���� -> ���� �Ǿ�����.]
    if (ckSendSMS="Y") then
        if (application("Svr_Info")<>"Dev") then
            if (MifinishReason <> "05") and (ischulgostate = "Y") then
            	Call SendMiChulgoSMS_CS(csdetailidx)
            end if
        end if
    end if
    ''EMail�߼�
    if (ckSendEmail="Y") then
        if (application("Svr_Info")<>"Dev") then
            if (MifinishReason <> "05") and (ischulgostate = "Y") then
            	call SendMiChulgoMail_CS(csdetailidx)
            end if
        end if
    end if

	if (MifinishReason="05") and (ischulgostate = "Y") then
        '// ǰ�����Ұ� ����� ����
		sqlStr = " exec db_cs.[dbo].[sp_Ten_MichulgoStockout_SetChargeIDaaa] " & csdetailidx & " "
		''dbget.Execute sqlStr
		response.write "<script type='text/javascript'>alert('TODO : ǰ�����Ұ� ����� ����.');</script>"
    end if

    if (ckSendSMS="Y") and (ckSendEmail="Y") then
        response.write "<script type='text/javascript'>alert('SMS�� ������ �߼� �Ǿ����ϴ�.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script type='text/javascript'>alert('SMS�� �߼� �Ǿ����ϴ�.');</script>"
    elseif (ckSendSMS="Y") then
        response.write "<script type='text/javascript'>alert('������ �߼� �Ǿ����ϴ�.');</script>"
    else
        response.write "<script type='text/javascript'>alert('ó�� �Ǿ����ϴ�.');</script>"
    end if

    response.write "<script type='text/javascript'>opener.location.reload();</script>"
    response.write "<script type='text/javascript'>location.replace('" + CStr(referer) + "')</script>"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
