<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü���
' Hieditor : 2007.04.07 ������ ����
'			 2011.04.15 �ѿ�� ����
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
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp" -->
<%
dim Makerid ,mode ,orderserialArr ,songjangnoArr ,songjangdivArr ,detailidxArr ,MisendReason, ipgodate, detailidx
dim sqlStr,i ,Overlap ,RectdetailidxArr, RectOrderSerialArr, RectSongjangnoArr, RectSongjangdivArr, OrderCount
dim TotAssignedRow, AssignedRow, FailRow
	Makerid = session("ssBctID")
	orderserialArr = Replace(request.Form("orderserialArr"), " ", "")
	songjangnoArr  = Replace(request.Form("songjangnoArr"), " ", "")
	songjangdivArr = Replace(request.Form("songjangdivArr"), " ", "")
	detailidxArr   = Replace(request.Form("detailidxArr"), " ", "")
	mode            = requestCheckVar(request.Form("mode"), 32)
	MisendReason    = requestCheckVar(request.Form("MisendReason"), 32)
	ipgodate        = requestCheckVar(request.Form("ipgodate"), 32)
	detailidx       = Replace(request.Form("detailidx"), " ", "")

	TotAssignedRow = 0
	AssignedRow    = 0
	FailRow        = 0

if (mode="SongjangInputCSV") then
    ''CSV �Է��� ���� , �� �ϳ� ����. �޸� ���̿� ���� ����
    orderserialArr = Replace(orderserialArr," ","") & ","
    songjangnoArr  = Replace(songjangnoArr," ","") & ","
    songjangdivArr = Replace(songjangdivArr," ","") & ","
    detailidxArr   = Replace(detailidxArr," ","") & ","
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

'' SongjangInputCSV CSV�� �Է� �߰�

if (mode="SongjangInput") or (mode="SongjangInputCSV") then

	if detailidxArr = "" then
        response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
        dbget.close()	:	response.end
	end if

    RectdetailidxArr   = split(detailidxArr,",")
    RectOrderSerialArr = split(orderserialArr,",")
    RectSongjangnoArr  = split(songjangnoArr,",")
    RectSongjangdivArr = split(songjangdivArr,",")

    if IsArray(RectdetailidxArr) then
        OrderCount = Ubound(RectdetailidxArr)

        ''2010-05-26 �߰�
        if (OrderCount<>Ubound(RectOrderSerialArr)) or (OrderCount<>Ubound(RectSongjangnoArr)) or (OrderCount<>Ubound(RectSongjangdivArr)) then
            response.write "<script>alert('���۵� �����Ͱ� ��ġ���� �ʽ��ϴ�.');</script>"
            dbget.close()	:	response.end
        end if
    end if

    if Right(detailidxArr,1)="," then detailidxArr = Left(detailidxArr,Len(detailidxArr)-1)
    if (Right(orderserialArr,1)=",") then orderserialArr=Left(orderserialArr,Len(orderserialArr)-1)
    orderserialArr = replace(orderserialArr,",","','")

    ''#################################################
    ''�����ȣ�Է� ����
    ''#################################################
    ''2009 ��� �ҿ��� passday �߰�.
    for i=0 to OrderCount - 1
        if (Trim(RectdetailidxArr(i))<>"") then

			'// ===============================================================
			'// CS Detail
			sqlStr = " update d " & VbCrLf
			sqlStr = sqlStr + " set currstate='B006' " & VbCrLf
        	sqlStr = sqlStr + " ,songjangno='" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF
        	sqlStr = sqlStr + " ,songjangdiv='" & Trim(RectSongjangdivArr(i)) & "'" & VbCRLF
			sqlStr = sqlStr + " from " & VbCrLf
			sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
			sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		m.id=d.masterid " & VbCrLf
			sqlStr = sqlStr + " where " & VbCrLf
			sqlStr = sqlStr + " 	1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 	and d.id =" & Trim(RectdetailidxArr(i))  & VbCRLF
			sqlStr = sqlStr + " 	and m.requireupche = 'Y' " & VbCrLf
			sqlStr = sqlStr + " 	and m.makerid = '" & Makerid & "' " & VbCrLf
			sqlStr = sqlStr + " 	and m.deleteyn='N' " & VbCrLf
			sqlStr = sqlStr + " 	and m.currstate < 'B006' " & VbCrLf
        	if (mode="SongjangInputCSV") then
        	    sqlStr = sqlStr + " and IsNULL(d.currstate,'B001')<'B006'"   ''�Ϸ��� �����ȣ ���� �� �� ����.. :: �����Է¸� �����ϵ���.
            end if

			'rw sqlStr
            dbget.Execute sqlStr, AssignedRow

            TotAssignedRow = TotAssignedRow + AssignedRow

            if (AssignedRow=0) then FailRow = FailRow + 1


			'// ===============================================================
			'// CS Master
			sqlStr = " update m " & VbCrLf
			sqlStr = sqlStr + " set songjangno='" & Trim(RectSongjangnoArr(i)) & "'" & VbCRLF
        	sqlStr = sqlStr + " ,songjangdiv='" & Trim(RectSongjangdivArr(i)) & "'" & VbCRLF
			sqlStr = sqlStr + " from " & VbCrLf
			sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
			sqlStr = sqlStr + " 	JOIN [db_cs].[dbo].tbl_new_as_detail d " & VbCrLf
			sqlStr = sqlStr + " 	on " & VbCrLf
			sqlStr = sqlStr + " 		m.id=d.masterid " & VbCrLf
			sqlStr = sqlStr + " where " & VbCrLf
			sqlStr = sqlStr + " 	1 = 1 " & VbCrLf
			sqlStr = sqlStr + " 	and d.id =" & Trim(RectdetailidxArr(i))  & VbCRLF
			sqlStr = sqlStr + " 	and m.requireupche = 'Y' " & VbCrLf
			sqlStr = sqlStr + " 	and m.makerid = '" & Makerid & "' " & VbCrLf
			sqlStr = sqlStr + " 	and m.deleteyn='N' " & VbCrLf
			sqlStr = sqlStr + " 	and m.currstate < 'B006' " & VbCrLf
        	if (mode="SongjangInputCSV") then
        	    sqlStr = sqlStr + " and IsNULL(m.currstate,'B001')<'B006'"   ''�Ϸ��� �����ȣ ���� �� �� ����.. :: �����Է¸� �����ϵ���.
            end if

			'rw sqlStr
            dbget.Execute sqlStr, AssignedRow

        end if
    next

	'' currstate B004 �߰�
	sqlStr = " update m " & VbCrLf
	sqlStr = sqlStr + " set currstate='B004' " & VbCrLf
	sqlStr = sqlStr + " from " & VbCrLf
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
	sqlStr = sqlStr + " where " & VbCrLf
	sqlStr = sqlStr + " 	1 = 1 " & VbCrLf
	sqlStr = sqlStr + " 	and m.id in ( " & VbCrLf
	sqlStr = sqlStr + " 		select " & VbCrLf
	sqlStr = sqlStr + " 		m.id " & VbCrLf
	sqlStr = sqlStr + " 		from " & VbCrLf
	sqlStr = sqlStr + " 			[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
	sqlStr = sqlStr + " 			JOIN [db_cs].[dbo].tbl_new_as_detail d " & VbCrLf
	sqlStr = sqlStr + " 			on " & VbCrLf
	sqlStr = sqlStr + " 				m.id=d.masterid " & VbCrLf
	sqlStr = sqlStr + " 		where " & VbCrLf
	sqlStr = sqlStr + " 			1 = 1 " & VbCrLf
	sqlStr = sqlStr + " 			and d.id in (" & Trim(detailidxArr) & ") "  & VbCRLF
	sqlStr = sqlStr + " 			and m.requireupche = 'Y' " & VbCrLf
	sqlStr = sqlStr + " 			and m.makerid = '" & Makerid & "' " & VbCrLf
	sqlStr = sqlStr + " 			and m.deleteyn='N' " & VbCrLf
	sqlStr = sqlStr + " 			and m.currstate < 'B006' " & VbCrLf
	sqlStr = sqlStr + " 		group by m.id " & VbCrLf
	sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'B001')<'B006' then 1 else 0 end )>0" & VbCRLF
	sqlStr = sqlStr + " ) " & VbCrLf

    'rw sqlStr
	dbget.Execute sqlStr


	'' currstate B006 �߰�
	sqlStr = " update m " & VbCrLf
	sqlStr = sqlStr + " set currstate='B006' " & VbCrLf
	sqlStr = sqlStr + " ,finishuser='" & Makerid & "' " & VbCrLf
	sqlStr = sqlStr + " ,finishdate=getdate() " & VbCrLf
	sqlStr = sqlStr + " from " & VbCrLf
	sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
	sqlStr = sqlStr + " where " & VbCrLf
	sqlStr = sqlStr + " 	1 = 1 " & VbCrLf
	sqlStr = sqlStr + " 	and m.id in ( " & VbCrLf
	sqlStr = sqlStr + " 		select " & VbCrLf
	sqlStr = sqlStr + " 		m.id " & VbCrLf
	sqlStr = sqlStr + " 		from " & VbCrLf
	sqlStr = sqlStr + " 			[db_cs].[dbo].tbl_new_as_list m " & VbCrLf
	sqlStr = sqlStr + " 			JOIN [db_cs].[dbo].tbl_new_as_detail d " & VbCrLf
	sqlStr = sqlStr + " 			on " & VbCrLf
	sqlStr = sqlStr + " 				m.id=d.masterid " & VbCrLf
	sqlStr = sqlStr + " 		where " & VbCrLf
	sqlStr = sqlStr + " 			1 = 1 " & VbCrLf
	sqlStr = sqlStr + " 			and d.id in (" & Trim(detailidxArr) & ") "  & VbCRLF
	sqlStr = sqlStr + " 			and m.requireupche = 'Y' " & VbCrLf
	sqlStr = sqlStr + " 			and m.makerid = '" & Makerid & "' " & VbCrLf
	sqlStr = sqlStr + " 			and m.deleteyn='N' " & VbCrLf
	sqlStr = sqlStr + " 			and m.currstate < 'B006' " & VbCrLf
	sqlStr = sqlStr + " 		group by m.id " & VbCrLf
	sqlStr = sqlStr + " 		having sum(case when IsNull(d.currstate,'B001')<'B006' then 1 else 0 end )=0" & VbCRLF
	sqlStr = sqlStr + " ) " & VbCrLf

    ''rw sqlStr
	dbget.Execute sqlStr

    dim AlertMsg
    AlertMsg = TotAssignedRow & "�� ó�� �Ǿ����ϴ�."
    if (FailRow>0) then
        AlertMsg = AlertMsg & "\n\n(" & FailRow & "�� �Է� ����)"
    end if

    response.write "<script language='javascript'>alert('" & AlertMsg & "')</script>"

    if (mode="SongjangInputCSV") then
        response.write "<script language='javascript'>opener.location.reload();</script>"
        response.write "<script language='javascript'>window.close();</script>"
    else
        response.write "<script language='javascript'>location.replace('" + CStr(referer) + "')</script>"
    end if
    dbget.close()	:	response.End

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->