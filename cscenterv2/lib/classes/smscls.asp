<%
Class CSMSClass
	public function CheckHpOk(byval irechp)
		CheckHpOk = false
		if Len(irechp)<3 then exit function
		if (Left(irechp,3)="013") or (Left(irechp,3)="011") or (Left(irechp,3)="016") or (Left(irechp,3)="017") or (Left(irechp,3)="018") or (Left(irechp,3)="019") or (Left(irechp,3)="010") then
			CheckHpOk = true
		end if
	end function

	public Sub SendJumunOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]���������� �����Ϸ� �Ǿ����ϴ�. �ֹ���ȣ : " + iorderserial + "')"

		rsget.Open sqlStr,dbget,1
	end Sub

	public sub SendAcctJumunOkMsg2(byval irechp, byval iorderserial, byval iacct, byval iprice)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]�ֹ����� �Ǿ����ϴ�. ����:" + iacct + " �ݾ�:" + iprice + "��')"

		rsget.Open sqlStr,dbget,1
	end sub

	public Sub SendAcctJumunOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]�ֹ������� �Աݴ�����Դϴ�.���¾ȳ�:" & CS_RECEIVE_BANK_INFO & ".��" & CS_MAIL_SITENAME & "')"

		rsget.Open sqlStr,dbget,1
	end Sub

    public Sub SendAcctIpkumCancelMsg(byval irechp, byval iorderserial)
        dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]�Ա� �� ����� ������ ��� �Ǿ����ϴ�. ����Ȯ���� �� �Ա� �� �ּ���')"

		dbget.Execute sqlStr
	end Sub

	public Sub SendAcctIpkumOkMsg(byval irechp, byval iorderserial)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]�Ա�Ȯ�� �Ǿ����ϴ�. �ֹ���ȣ�� " + iorderserial + "�Դϴ�.�����մϴ�.')"

		rsget.Open sqlStr,dbget,1
	end Sub

	public Sub SendBeaSongOkMsg(byval irechp, byval isongjangno)
		dim sqlStr
		dim delivercoper

		if Not CheckHpOk(irechp) then Exit sub

        delivercoper = "�ù�� �����ù�"
        if Left(isongjangno,1)="6" then
        	delivercoper = "�ù�� CJ�ù�"
        end if

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]��ǰ�� ���Ǿ����ϴ�.  " + delivercoper + " �����ȣ " + isongjangno + "')"

		rsget.Open sqlStr,dbget,1
	end Sub

	public Sub SendJikjupWaitMsg(byval irechp)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]�ֹ��� ��ǰ�� �غ�Ǿ����ϴ�.�������� �൵�� Ȩ������ �� �������ּ���.')"

		rsget.Open sqlStr,dbget,1
	end Sub

	public Sub SendNormalMsg(byval imsg,byval irechp)
		dim sqlStr

		if Not CheckHpOk(irechp) then Exit sub

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '" + imsg + "')"

		rsget.Open sqlStr,dbget,1
	end Sub

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
end Class
%>