<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_mail_Function_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim mode, modeflag2, divcd, reguserid, ipkumdiv ,title, orderno, contents_jupsu
dim finishuser, contents_finish ,requireupche, requiremakerid, ForceReturnByTen ,detailitemlist
dim opentitle, opencontents ,newasid ,isCsMailSend ,IsAllCancel ,CancelValidResultMessage
dim sqlStr, i ,ScanErr ,ResultMsg, ReturnUrl, EtcStr ,ProceedFinish ,returnmethod ,masteridxtmp
dim oordermaster ,buf_requiremakerid , masteridx , csmasteridx , cancelorderno ,GC_IsOLDOrder
	masteridx        = requestCheckVar(request.Form("masteridx"),10)
	csmasteridx   = requestCheckVar(request.Form("csmasteridx"),10)
	mode        = requestCheckVar(request.Form("mode"),32)
	modeflag2   = requestCheckVar(request.Form("modeflag2"),32)
	divcd       = requestCheckVar(request.Form("divcd"),4)
	ipkumdiv    = requestCheckVar(request.Form("ipkumdiv"),1)
	reguserid   = session("ssbctid")
	finishuser  = reguserid
	title       = requestCheckVar(html2DB(request.Form("title")),128)
	orderno = requestCheckVar(request.Form("orderno"),16)
	contents_jupsu  = requestCheckVar(html2DB(request.Form("contents_jupsu")),800)
	detailitemlist  = html2db(request.Form("detailitemlist"))
	contents_finish = requestCheckVar(html2DB(request.Form("contents_finish")),32)
	requireupche = requestCheckVar(request.Form("requireupche"),1)
	requiremakerid = requestCheckVar(request.Form("requiremakerid"),32)
	ForceReturnByTen = requestCheckVar(request.Form("ForceReturnByTen"),32)
	buf_requiremakerid  = requestCheckVar(request.Form("buf_requiremakerid"),32)
	isCsMailSend = requestCheckVar((request.Form("csmailsend")="on"),32)
	cancelorderno       = requestCheckVar(request.Form("cancelorderno"),16)

newasid = -1		
if (returnmethod="") then returnmethod="R000"
ScanErr = ""
ProceedFinish = False

''�ֹ� ����Ÿ
set oordermaster = new COrder
	oordermaster.FRectmasteridx = masteridx	
	oordermaster.fQuickSearchOrderMaster

'response.write "mode : " & mode & "<br>"

'/cs����
if (mode="regcsas") then
	'response.write "divcd : " & divcd & "<br>"

	'CS ���� - �ֹ����
	if (divcd="A008") then

		dbget.beginTrans
		
        '' CS Master ���� ''html2db ������� ����.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)  

        'CS Detail ����(���� ��ǰ����)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

		'/�ֹ������� ���̳ʽ��ֹ������� ��ġ �ϴ��� üũ
		'CancelValidResultMessage = GetPartialCancelRegValidResult_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)

		if (CancelValidResultMessage <> "") then
			ScanErr = CancelValidResultMessage
		end if	        		

		'/������ ���ó��
        'Call masterCancelProcess_off(masteridx ,cancelorderno)
            
    	''�ٷ� �Ϸ�ó���� ���� ���� ���� - AsDetail �Է��� �˻�
        ProceedFinish = IsDirectProceedFinish_off(divcd, csmasteridx, masteridx, EtcStr)
        contents_finish = ""

        '' �Ϸ�ó�� ���μ���
        If (ProceedFinish) then
			'/������ ���ó��
            Call CancelProcess_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)
			Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)

			sqlStr = ""
			sqlStr = "select top 1 masteridx , detailidx , orderno" + vbcrlf
			sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail" + vbcrlf
			sqlStr = sqlStr & " where cancelyn='N'"
			sqlStr = sqlStr & " and masteridx = "&masteridx&"" + vbcrlf
		
			'response.write sqlStr &"<br>"
			rsget.open sqlStr ,dbget ,1
		
			if not(rsget.eof) then
				masteridxtmp = false
			else
				masteridxtmp = true
			end if
		
			rsget.close()
		
			'//�������� ���� ��� ��� �����͵� ��� ��Ų��
			if masteridxtmp then
				sqlStr = ""
				sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master set" + vbcrlf
				sqlStr = sqlStr & " cancelyn='Y'" + vbcrlf
				sqlStr = sqlStr & " where masteridx = "&masteridx&""
		
				'response.write sqlStr &"<br>"
				dbget.execute sqlStr
			end if
	
			'//������ ���̺� ���°� �Ϻ���� ���� ���� ��ǰ�� ���� ���� ������ ������ ���̺� ���¸� ���Ϸ�� �ٲ۴�
			'�������
		    sqlStr = " update db_shop.dbo.tbl_shopbeasong_order_master set										" & VbCRLF
		    sqlStr = sqlStr + " ipkumdiv='8', beadaldate=getdate() 														" & VbCRLF
			sqlStr = sqlStr + " where masteridx in ( 																" & VbCRLF
		    sqlStr = sqlStr + " 	select 																	" & VbCRLF
		    sqlStr = sqlStr + " 	m.masteridx 														" & VbCRLF
		    sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
		    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
		    sqlStr = sqlStr + " 		on m.masteridx=d.masteridx 										" & VbCRLF
		    sqlStr = sqlStr + " 	where d.itemid<>0 													" & VbCRLF
		    sqlStr = sqlStr + " 	and m.masteridx in ("&masteridx&") 																	" & VbCRLF
		    sqlStr = sqlStr + " 	group by m.masteridx 														" & VbCRLF
		    sqlStr = sqlStr + " 	having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
		    sqlStr = sqlStr + " ) 																			" & VbCRLF
		
		    'response.write sqlStr &"<br>"
		    dbget.Execute sqlStr
        ELSE
            ResultMsg = ResultMsg + "->. ��ǰ �غ��� ������ ��ǰ�� �����ϹǷ�\n\n �ֹ� ��� ������ ���� �Ǿ����ϴ�.\n\n ��ü ��ȭ Ȯ���� �Ϸ� ó���ϼž� �մϴ�."
        End If
	
        ResultMsg = ResultMsg & "OK"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
        
        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
            'dbget.RollBackTrans
            
			response.write "<script>"
			response.write "	alert('"&ResultMsg&"');"
			response.write "	location.replace('"&ReturnUrl&"');"
			response.write "</script>"
			dbget.close()	:	response.End	            
        Else
            dbget.RollBackTrans
            
            response.write "<script>"
            response.write "	alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")"
            response.write "</script>"
            dbget.close()	:	response.End
        End If
  		
	'CS ���� - ��Ÿ���� / �������ǻ��� / ��ü �߰� �����
    elseif (divcd="A009") or (divcd="A006") or (divcd="A700") then
             
        dbget.beginTrans

        '' CS Master ���� ''html2db ������� ����.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)    

        'CS Detail ����(���� ��ǰ����)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

        '��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        if (requiremakerid<>"") then
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)
        end if

        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
        
        ResultMsg = ResultMsg + "\nOK"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
            
			response.write "<script>location.replace('"&ReturnUrl&"');</script>"
			dbget.close()	:	response.End	            
        Else
            dbget.RollBackTrans
            
            response.write "<script>"
            response.write "	alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")"
            response.write "</script>"
            dbget.close()	:	response.End
        End If

	'CS ���� - ������߼�, ���񽺹߼�
    elseif (divcd="A001") or (divcd="A002") then
  	              
        dbget.beginTrans

        '' CS Master ���� ''html2db ������� ����.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
        
        'CS Detail ����(���� ��ǰ����)
        Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)

    
		'��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        if (requiremakerid<>"") then
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)
        else
        
        	'/�������� ��� '/���� �ٹ����� ����� ������ �����ؼ� �־����
        	'/������ requiremaejang : Y  : �ٹ����ٹ�� requiremaejang : N
        	call RegCSMasterAddmaejang_off(csmasteridx)
        end if

        ResultMsg = "�����Ϸ�"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans

			response.write "<script>alert('OK'); location.replace('"&ReturnUrl&"');</script>"
			dbget.close()	:	response.End
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
        
    'CS ���� - �±�ȯ���
    elseif (divcd="A000") then

        dbget.beginTrans

		'' CS Master ���� ''html2db ������� ����.
        csmasteridx = RegCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu, masteridx)
   
		'CS Detail ����(���� ��ǰ����)
		Call AddCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno ,masteridx)
 
		'��ü����� ��� ���� ��ü �귣�� ���̵� �Է�(requiremakerid)
        if (requiremakerid<>"") then            
            call RegCSMasterAddUpche_off(csmasteridx, requiremakerid)

            ResultMsg = "�±�ȯ �����Ϸ� - ��ü���"        
        
        else
        	'/�������� ��� '/���� �ٹ����� ����� ������ �����ؼ� �־����
      	        	
        	'/�±�ȯ ��� ���� ������ ���� requiremaejang : Y
        	call RegCSMasterAddmaejang_off(csmasteridx)
        	
        	'���� ����� ��� �±�ȯ ȸ�� ����
            '' CS Master ���� ''html2db ������� ����.
            newasid = RegCSMaster_off("A013", orderno, reguserid, "�±�ȯ ȸ������", contents_jupsu, masteridx)
			
			'CS Detail ����(���� ��ǰ����)			
            Call AddCSDetailByArrStr_off(detailitemlist, newasid, orderno ,masteridx)
			
			'/�±�ȯ ȸ�� ���� ������ ���� requiremaejang : Y
        	call RegCSMasterAddmaejang_off(newasid)

             ResultMsg = "�±�ȯ ��� ���� �� ȸ������ �Ϸ� - ���� ���"
        end if

        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"            
            dbget.close()	:	response.End
        End If

        ''�̸��� �߼� �±�ȯ ����
        if (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)

            ''�±�ȯ ȸ���� �������
            if (newasid>0) then
                Call SendCsActionMail_off(newasid)
            end if
        End If    
    else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

''���� ���� ����
elseif (mode="editcsas") then

    dbget.beginTrans

	Call EditCSMaster_off(divcd, orderno, reguserid, title, contents_jupsu ,csmasteridx)

    '' CS Detail ����
    Call EditCSDetailByArrStr_off(detailitemlist, csmasteridx, orderno)

    ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
		
    If (Err.Number = 0) and (ScanErr="") Then
        dbget.CommitTrans
    
    	ResultMsg = ResultMsg + "OK"
    Else
        dbget.RollBackTrans
        response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"        
        dbget.close()	:	response.End
    End If

'CS ���� ���� �Ϸ�ó��
elseif (mode="finishcsas") then	    
	'response.write "divcd : " & divcd & "<br>"
	
	'CS ���� ���� �Ϸ�ó�� - �ֹ����
    if (divcd="A008") then
		
		dbget.beginTrans
		
		'/������ ���ó��
	    Call CancelProcess_off(detailitemlist, csmasteridx, orderno,masteridx ,cancelorderno)
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)	

		sqlStr = ""
		sqlStr = "select top 1 masteridx , detailidx , orderno" + vbcrlf
		sqlStr = sqlStr & " from db_shop.dbo.tbl_shopbeasong_order_detail" + vbcrlf
		sqlStr = sqlStr & " where cancelyn='N'"
		sqlStr = sqlStr & " and masteridx = "&masteridx&"" + vbcrlf
	
		'response.write sqlStr &"<br>"
		rsget.open sqlStr ,dbget ,1
	
		if not(rsget.eof) then
			masteridxtmp = false
		else
			masteridxtmp = true
		end if
	
		rsget.close()
	
		'//�������� ���� ��� ��� �����͵� ��� ��Ų��
		if masteridxtmp then
			sqlStr = ""
			sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_master set" + vbcrlf
			sqlStr = sqlStr & " cancelyn='Y'" + vbcrlf
			sqlStr = sqlStr & " where masteridx = "&masteridx&""
	
			'response.write sqlStr &"<br>"
			dbget.execute sqlStr
		end if

		'//������ ���̺� ���°� �Ϻ���� ���� ���� ��ǰ�� ���� ���� ������ ������ ���̺� ���¸� ���Ϸ�� �ٲ۴�
		'�������
	    sqlStr = " update db_shop.dbo.tbl_shopbeasong_order_master set										" & VbCRLF
	    sqlStr = sqlStr + " ipkumdiv='8', beadaldate=getdate() 														" & VbCRLF
		sqlStr = sqlStr + " where masteridx in ( 																" & VbCRLF
	    sqlStr = sqlStr + " 	select 																	" & VbCRLF
	    sqlStr = sqlStr + " 	m.masteridx 														" & VbCRLF
	    sqlStr = sqlStr + " 	from db_shop.dbo.tbl_shopbeasong_order_master m 							" & VbCRLF
	    sqlStr = sqlStr + " 	join db_shop.dbo.tbl_shopbeasong_order_detail d 					" & VbCRLF
	    sqlStr = sqlStr + " 		on m.masteridx=d.masteridx 										" & VbCRLF
	    sqlStr = sqlStr + " 	where d.itemid<>0 													" & VbCRLF
	    sqlStr = sqlStr + " 	and m.masteridx in ("&masteridx&") 																	" & VbCRLF
	    sqlStr = sqlStr + " 	group by m.masteridx 														" & VbCRLF
	    sqlStr = sqlStr + " 	having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0 " & VbCRLF
	    sqlStr = sqlStr + " ) 																			" & VbCRLF
	
	    'response.write sqlStr &"<br>"
	    dbget.Execute sqlStr
	    
		ResultMsg = "ó�� �Ϸ�"
		ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If
				
	'CS ���� ���� �Ϸ�ó�� - �±�ȯ ��� / ���� / ���� �߼� / ��Ÿ /  ���� ���ǻ���
    elseif  (divcd="A000") or (divcd="A001") or (divcd="A002") or (divcd="A009") or (divcd="A006") or (divcd="A005") or (divcd="A700") then
	
		dbget.beginTrans
		
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)
		
		ResultMsg = "ó�� �Ϸ�"
		ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
		
		If (Err.Number = 0) and (ScanErr="") Then
		    dbget.CommitTrans
		Else
		    dbget.RollBackTrans
		    response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"	        
		    dbget.close()	:	response.End
		End If
		
		If (isCsMailSend) then
		    if ((divcd="A000") or (divcd="A001") or (divcd="A002")) then
		        
		        ''�±�ȯ/����/���� �Ϸ� ����
		        Call SendCsActionMail_off(csmasteridx)
		    end if
		End If
	
	'CS ���� ���� �Ϸ�ó�� - �±�ȯȸ��(������)
    elseif (divcd="A013") then    	
    	        
        dbget.beginTrans
        
		Call FinishCSMaster_off(csmasteridx, reguserid, contents_finish)

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd

        If (Err.Number = 0) and (ScanErr="") Then
            dbget.CommitTrans
        Else
            dbget.RollBackTrans
            response.write "<script>alert(" + Chr(34) + "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " + CStr(errcode) + ":" + Err.Description + "|" + ScanErr + ")" + Chr(34) + ")</script>"
            dbget.close()	:	response.End
        End If

        ''�±�ȯ �Ϸ� ����
        If (isCsMailSend) then
            Call SendCsActionMail_off(csmasteridx)
        End If
	        
	else
        ResultMsg = "���ǵ��� �ʾҽ��ϴ�. : mode=" + mode + " , divcd=" + divcd
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if

'' ��ü ó���Ϸ� => �������·κ���
elseif (mode="upcheconfirm2jupsu") then
	    
    sqlStr = " select top 1 currstate from db_shop.dbo.tbl_shopbeasong_cs_master"
    sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
	
	'response.write sqlStr &"<br>"
    rsget.Open sqlStr,dbget,1
	    if not rsget.Eof then
	        ResultMsg = ""
	        if (rsget("currstate")<>"B006") then
	            ResultMsg = "��ü ó�� �Ϸ� ���°� �ƴմϴ�. ���� �Ұ�"
	        end if
		else
		    ResultMsg = "�ڵ����. ���� �Ұ�"
		end if
	rsget.Close

    if (ResultMsg="") then
        sqlStr = " update db_shop.dbo.tbl_shopbeasong_cs_master" + VbCrlf
        sqlStr = sqlStr + "set currstate='B001'" + VbCrlf
        sqlStr = sqlStr + ",contents_jupsu='" + (contents_jupsu) + "'" + VbCrlf
        sqlStr = sqlStr + " where masteridx=" + CStr(csmasteridx)
        dbget.Execute sqlStr

        ResultMsg = "ó�� �Ϸ�"
        ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd
    else
        response.write "<script>alert('" + ResultMsg + "');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
    
'CS ����
elseif (mode="deletecsas") then
	
    ''Check Valid Delete - ����� B006 ��üó���Ϸ� , B007 �Ϸ� ������ ���(����) �Ұ�
    if (NOT ValidDeleteCS_off(csmasteridx)) then
        response.write "<script>alert(" + Chr(34) + "���� ��� ���� ���°� �ƴմϴ�. ������ ���� ���." + Chr(34) + ")</script>"
        response.write "<script>history.back()</script>"
        dbget.close()	:	response.End
    end if

    If Not DeleteCSProcess_off(csmasteridx, reguserid) then
        ResultMsg = ResultMsg + "������ ������ ����"
    else
        ResultMsg = ResultMsg + "OK"
    End if
    
    ReturnUrl = "/admin/offshop/cscenter/action/pop_cs_action_new.asp?csmasteridx="  + CStr(csmasteridx) + "&divcd=" + divcd + "&mode=editreginfo"
   
end if

%>

<%
response.write "<script>alert('" + ResultMsg + "');</script>"
response.write "<script>location.replace('" + ReturnUrl + "');</script>"
response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->