<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<%

dim id,finishmemo, finishuser,songjangdiv, songjangno, customerreceiveyn, customerrealbeasongpay
dim refasid, receiveonly
dim targetid, needChkYN, needRefChkYN

id          	= requestCheckVar(request("id"),32)
refasid         = requestCheckVar(request("refasid"),32)
receiveonly     = requestCheckVar(request("receiveonly"),32)

finishmemo  = html2db(requestCheckVar(request("finishmemo"),3200))
finishuser  = requestCheckVar(request("finishuser"),32)
songjangdiv = requestCheckVar(request("songjangdiv"),32)
songjangno  = requestCheckVar(request("songjangno"),32)

customerreceiveyn  		= requestCheckVar(request("customerreceiveyn"),32)			'// ���߰���ۺ� ����
customerrealbeasongpay 	= requestCheckVar(request("customerrealbeasongpay"),32)		'// ���߰���ۺ� Ȯ�ξ�

needChkYN 	= requestCheckVar(request("needChkYN"),32)
needRefChkYN 	= requestCheckVar(request("needRefChkYN"),32)

dim sqlStr
dim currstate, divcd, currsongjangdiv, currsongjangno
dim IsSongjangChanged

'// ===========================================================================
'// 1. �±�ȯȸ�� �Ϸ�ó��						- receiveonly = "Y"
'// 2. �±�ȯ���(�±�ȯȸ�� ��ϵ� ���)		- refasid <> 0 and receiveonly <> "Y"
'// 3. ������									- refasid = 0
'// ===========================================================================


'// ===========================================================================
if (receiveonly <> "Y") then
	targetid = id
else
	targetid = refasid
end if

if (targetid = 0) then
	response.write "<script>alert('�߸��� �����Դϴ�. - �ٹ����� �ý����� ���ǿ��');history.back();</script>"
	response.end
end if

sqlStr = "select currstate, divcd, IsNull(songjangdiv, '') as currsongjangdiv, IsNull(songjangno, '') as currsongjangno "
sqlStr = sqlStr + " from [db_cs].[dbo].tbl_new_as_list" + VbCrlf
sqlStr = sqlStr + " where id =" + targetid

rsget.Open sqlStr,dbget,1
    currstate 		= rsget("currstate")
    divcd 			= rsget("divcd")

	IsSongjangChanged = (songjangdiv <> currsongjangdiv) or (songjangno <> currsongjangno)
rsget.Close


'// ===========================================================================
if (currstate="B007") then

	response.write "<script>alert('�̹� ó�� �Ϸ�� �����Դϴ�. - �Ϸ�ó���� ���� �� �� �����ϴ�.');history.back();</script>"
	response.end

else

	sqlStr = "update [db_cs].[dbo].tbl_new_as_list" + VbCrlf
	sqlStr = sqlStr + " set finishuser ='" + finishuser + "'," + VbCrlf
	sqlStr = sqlStr + " contents_finish ='" + finishmemo + "'," + VbCrlf
	sqlStr = sqlStr + " songjangdiv ='" + songjangdiv + "'," + VbCrlf
	sqlStr = sqlStr + " songjangno ='" + songjangno + "'," + VbCrlf
	if (IsSongjangChanged = True) then
		sqlStr = sqlStr + " songjangRegGubun ='U'," + VbCrlf
		sqlStr = sqlStr + " songjangRegUserID ='" + session("ssBctID") + "'," + VbCrlf
	end if
	sqlStr = sqlStr + " finishdate=getdate()," + VbCrlf
	if (needChkYN <> "") then
		sqlStr = sqlStr + " needChkYN='" & needChkYN & "'," + VbCrlf
	end if
	sqlStr = sqlStr + " currstate='B006' " + VbCrlf
	sqlStr = sqlStr + " where id =" + targetid
	sqlStr = sqlStr + " and makerid='" & session("ssBctID") & "'"
	rsget.Open sqlStr,dbget,1

	'// ���� ó���� ���̵� ����
	Call SaveCSListHistory(targetid)

	if receiveonly = "Y" and customerreceiveyn = "Y" then

		sqlStr = " update "
		sqlStr = sqlStr + " 	c "
		sqlStr = sqlStr + " set "
		sqlStr = sqlStr + " 	c.receiveyn = '" + CStr(customerreceiveyn) + "' "
		sqlStr = sqlStr + " 	, c.realbeasongpay = " + CStr(customerrealbeasongpay) + " "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	[db_cs].[dbo].tbl_new_as_list a "
		sqlStr = sqlStr + " 	join [db_cs].[dbo].tbl_as_customer_addbeasongpay_info c "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		a.id = c.asid "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and a.id = " + CStr(id) + " "
		sqlStr = sqlStr + " 	and a.makerid = '" & session("ssBctID") & "' "
		rsget.Open sqlStr,dbget,1

	end if

	'// ��ÿϷ� ó��
	if (needChkYN = "N") then
		'CS���������� ��������
		dim ocsaslist, orderserial
		set ocsaslist = New CCSASList
		ocsaslist.FRectCsAsID = targetid
		ocsaslist.GetOneCSASMaster

		'// ��ÿϷ�ó��
		'// ���񽺹߼�(A002) �� ���
		if InStr(",A000,A100,A001,A002,A009,A006,A012,", ocsaslist.FOneItem.Fdivcd) > 0 And InStr(",A000,A002,A006,", ocsaslist.FOneItem.Fdivcd) > 0 then
			'==============================================================================
			dim oordermaster
			set oordermaster = new COrderMaster
			oordermaster.FRectOrderSerial = ocsaslist.FOneItem.Forderserial
			oordermaster.QuickSearchOrderMaster

			orderserial = ocsaslist.FOneItem.Forderserial

			'' ���� 6���� ���� ���� �˻�
			if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
				oordermaster.FRectOldOrder = "on"
				oordermaster.QuickSearchOrderMaster
			end if

			'==============================================================================
			'// �±�ȯȸ��
			dim ioneRefas, IsRefASExist, IsRefASFinished

			IsRefASExist = False
			IsRefASFinished = False

			if (divcd = "A000") or (divcd = "A100") then
				set ioneRefas = new CCSASList
				ioneRefas.FRectCsRefAsID = targetid
				ioneRefas.GetOneCSASMaster

				if (ioneRefas.FResultCount>0) then
					IsRefASExist = True
					if (ioneRefas.FOneItem.Fcurrstate = "B007") then
			    		IsRefASFinished = True
					end If

					If (divcd = "A000") And (needRefChkYN = "N") And (IsRefASFinished = False) Then
						'// ��ȯȸ���� �Ϸ�ó��
						Call FinishCSMaster(ioneRefas.FOneItem.Fid, session("ssBctID"), "��ȯ��� �� ��ȯȸ�� ���ÿϷ�ó��")
						IsRefASFinished = True
					End If
				end if
			end if

			'==============================================================================
			''�Ϸ�ó�� �Ұ��� �޼���
			dim FinishInValidMsg

			''�Ϸ�ó�� ���� ����
			dim IsFinishProcessAvail

			FinishInValidMsg = ""
			IsFinishProcessAvail = True

			if (IsRefASExist) and (IsRefASFinished = False) and (ocsaslist.FOneItem.Frequireupche = "Y") then
    			FinishInValidMsg = "��ü����� ��� �±�ȯȸ���� ���� �Ϸ�ó���ؾ� �±�ȯ��� �Ϸ�ó���� �� �ֽ��ϴ�."
    			IsFinishProcessAvail = False
			end if

			if (IsFinishProcessAvail = True) then
				Call FinishCSMaster(targetid, session("ssBctID"), finishmemo)

				if (datediff("d", ocsaslist.FOneItem.Fregdate, now()) <= 21) and oordermaster.FOneItem.FSiteName="10x10" then
					'// 21�� �̳��̰�, �ٹ����� �ֹ��̸� ���Ϲ߼�
					if ((divcd="A000") or (divcd="A100") or (divcd="A001") or (divcd="A002")) then
						''�±�ȯ/����/���� �Ϸ� ����
						Call SendCsActionMail(targetid)
					end if
				end if
			end if

			if (IsFinishProcessAvail = False) And (needRefChkYN = "") then
				response.write "<script>alert('����!!\n\n" & FinishInValidMsg & "')</script>"
			end if

		end if

	end if

end if

%>

<script>

alert('����Ǿ����ϴ�.');

if (window.opener) {
	opener.location.reload();
	opener.focus();
	window.close();
} else {
	//location.replace('upchecslist.asp');
	history.back();
}

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
