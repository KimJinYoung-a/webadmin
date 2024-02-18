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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]정상적으로 결제완료 되었습니다. 주문번호 : " + iorderserial + "')"

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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]주문접수 되었습니다. 계좌:" + iacct + " 금액:" + iprice + "원')"

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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]주문접수후 입금대기중입니다.계좌안내:" & CS_RECEIVE_BANK_INFO & ".㈜" & CS_MAIL_SITENAME & "')"

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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]입금 후 전산상 오류로 취소 되었습니다. 계좌확인후 재 입금 해 주세요')"

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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]입금확인 되었습니다. 주문번호는 " + iorderserial + "입니다.감사합니다.')"

		rsget.Open sqlStr,dbget,1
	end Sub

	public Sub SendBeaSongOkMsg(byval irechp, byval isongjangno)
		dim sqlStr
		dim delivercoper

		if Not CheckHpOk(irechp) then Exit sub

        delivercoper = "택배사 현대택배"
        if Left(isongjangno,1)="6" then
        	delivercoper = "택배사 CJ택배"
        end if

		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		sqlStr = sqlStr + " values('" + irechp + "',"
		sqlStr = sqlStr + " '" & CS_MAIN_PHONENO & "',"
		sqlStr = sqlStr + " '1',"
		sqlStr = sqlStr + " getdate(),"
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]상품이 출고되었습니다.  " + delivercoper + " 송장번호 " + isongjangno + "')"

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
		sqlStr = sqlStr + " '[" & CS_MAIL_SITENAME & "]주문한 상품이 준비되었습니다.직접수령 약도는 홈페이지 를 참고해주세요.')"

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