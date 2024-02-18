<%
Class CPayForm
	public Fempno
	public Fdefaultpay
	public Ffoodpay
	public Fjobpay
	public Fstartdate
	public Fenddate
	public FinBreakTime
	public FOverTime

	public FStartTime(8)
	public FEndTime(8)
	public FBreakSTime(8)
	public FBreakETime(8)
	public FStartHour(8)
	public FStartMinute(8)
	public FEndHour(8)
	public FEndMinute(8)
	public FBreakSHour(8)
	public FBreakSMinute(8)
	public FBreakEHour(8)
	public FBreakEMinute(8)
	public FDutyTime(8)
 	public FNightTime(8)
	public Fworktype(8)

	public Fholidaywdtime
	public Fregdate
	public Flastupdate
	public Fadminid

 	public FTotDutyTime
 	public FTotNightTime
 	public FtotPaySum

 	public FpatternSeq
 	public Fpart_sn
 	public Fpatternname

 	public FPageSize
	public FCurrPage
	public FTotCnt
	public FSPageNo
	public FEPageNo
	public FdefaultPaySeq
	public Fino

	public Fpredefaultpay
	public FpreFoodpay

	public Fposit_sn
	public Fposit_name
	public Fdepartment_id
	public FdepartmentNameFull
	public Fjobkind
	public Fplacekind

	'//������� ����Ϻ� ����Ʈ
		public Function fnGetDefaultPayList
			Dim strSql
			IF Fempno = "" THEN Exit Function

			strSql ="[db_partner].[dbo].sp_Ten_defaultpay_getListCnt('"&Fempno&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_partner].[dbo].sp_Ten_defaultpay_getList('"&Fempno&"',"&FSPageNo&","&FEPageNo&")"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetDefaultPayList = rsget.getRows()
			END IF
			rsget.close
			END IF
		End Function

	'����� �⺻ �޿��� �ٹ��� ����(�������)
	public Function fnGetDefaultPayData
	IF Fempno = "" THEN Exit Function
	IF Fino = "" THEN Fino = 0
		Dim strSql
		Dim intLoop
		Dim NST(8) ,NET(8), NBST(8), NBET(8)

		strSql ="db_partner.dbo.sp_Ten_user_defaultpay_GetData('"&Fempno&"',"&Fino&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FdefaultPaySeq=rsget("defaultPayseq")
				Fempno		= rsget("empno")
				Fstartdate		= rsget("startdate")
				Fenddate		= rsget("enddate")
				Fdefaultpay = rsget("defaultpay")
				Ffoodpay	= rsget("foodpay")
				Fjobpay		= rsget("jobpay")
				FinBreakTime = rsget("inBreakTime")
				FOverTime	= rsget("overtime")/60

				FStartTime(1)   = rsget("sunStart")
				FEndTime(1)     = rsget("sunEnd")
				FBreakSTime(1)   = rsget("sunBreakS")
				FBreakETime(1)   = rsget("sunBreakE")
				Fworktype(1)	= rsget("sunworktype")

				FStartTime(2)   = rsget("monStart")
				FEndTime(2)     = rsget("monEnd")
				FBreakSTime(2)   = rsget("monBreakS")
				FBreakETime(2)   = rsget("monBreakE")
				Fworktype(2)	= rsget("monworktype")

				FStartTime(3)   = rsget("tueStart")
				FEndTime(3)     = rsget("tueEnd")
				FBreakSTime(3)   = rsget("tueBreakS")
				FBreakETime(3)   = rsget("tueBreakE")
				Fworktype(3)	= rsget("tueworktype")

				FStartTime(4)   = rsget("wedStart")
				FEndTime(4)     = rsget("wedEnd")
				FBreakSTime(4)   = rsget("wedBreakS")
				FBreakETime(4)   = rsget("wedBreakE")
				Fworktype(4)	= rsget("wedworktype")

				FStartTime(5)   = rsget("thuStart")
				FEndTime(5)     = rsget("thuEnd")
				FBreakSTime(5)   = rsget("thuBreakS")
				FBreakETime(5)   = rsget("thuBreakE")
				Fworktype(5)	= rsget("thuworktype")

				FStartTime(6)   = rsget("friStart")
				FEndTime(6)     = rsget("friEnd")
				FBreakSTime(6)   = rsget("friBreakS")
				FBreakETime(6)   = rsget("friBreakE")
				Fworktype(6)	= rsget("friworktype")

				FStartTime(7)   = rsget("satStart")
				FEndTime(7)     = rsget("satEnd")
				FBreakSTime(7)   = rsget("satBreakS")
				FBreakETime(7)   = rsget("satBreakE")
				Fworktype(7)	= rsget("satworktype")

				FTotDutyTime = 0

				For intLoop = 1 To 7

				FStartHour(intLoop) = format00(2,Fix(FStartTime(intLoop)/60))
				FStartMinute(intLoop) = format00(2,(FStartTime(intLoop) mod 60))

				FEndHour(intLoop) = format00(2,Fix(FEndTime(intLoop)/60))
				FEndMinute(intLoop) = format00(2,(FEndTime(intLoop) mod 60))

				FBreakSHour(intLoop) = format00(2,Fix(FBreakSTime(intLoop)/60))
				FBreakSMinute(intLoop) = format00(2,(FBreakSTime(intLoop) mod 60))
				FBreakEHour(intLoop) = format00(2,Fix(FBreakETime(intLoop)/60))
				FBreakEMinute(intLoop) = format00(2,(FBreakETime(intLoop) mod 60))

				IF  FinBreakTime  THEN  '�ް�ð��� �ٹ��ð���  �����ϸ�..	�ް�ð��� ���� �ʴ´�.
					FDutyTime(intLoop)	= format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop))/60)) + ":"+Format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop)) mod 60))
					FTotDutyTime = FTotDutyTime + FEndTime(intLoop)  - FStartTime(intLoop)
				ELSE
					FDutyTime(intLoop)	= format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop) - (FBreakETime(intLoop)-FBreakSTime(intLoop)))/60)) + ":"+Format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop) - (FBreakETime(intLoop)-FBreakSTime(intLoop))) mod 60))
					FTotDutyTime = FTotDutyTime + FEndTime(intLoop)  - FStartTime(intLoop) -(FBreakETime(intLoop)-FBreakSTime(intLoop))
				END IF

				'�߰��ٹ��ð�
					IF FEndTime(intLoop) >= (22*60) AND FStartTime(intLoop)< (30*60) THEN	'����ð��� 10�� �����̰� ���۽ð��� ��ħ 6�� �����϶� - �߰��ٹ�
						IF FStartTime(intLoop) < (22*60) THEN
							NST(intLoop) = 22*60
						ELSE
							NST(intLoop) = FStartTime(intLoop)
						END IF

						IF FEndTime(intLoop) > (30*60) THEN
							NET(intLoop) = 30*60
						ELSE
							NET(intLoop) = FEndTime(intLoop)
						END IF

						IF FBreakSTime(intLoop) < (22*60) THEN
							NBST(intLoop) = 22*60
						ELSEIF 	FBreakSTime(intLoop) >= (30*60) THEN
							NBST(intLoop) = 0
						ELSE
							NBST(intLoop) = FBreakSTime(intLoop)
						END IF

						IF FBreakETime(intLoop) < (22*60) THEN
							NBET(intLoop) = 22*60
						ELSEIF FBreakETime(intLoop) > (30*60) THEN
							NBET(intLoop) = 0
						ELSE
							NBET(intLoop) = FBreakETime(intLoop)
						END IF

						IF FinBreakTime THEN
							FNightTime(intLoop) = NET(intLoop)   - NST(intLoop)
						ELSE
							FNightTime(intLoop) =NET(intLoop)   - NST(intLoop) -(NBET(intLoop)-NBST(intLoop))
						END IF
					END IF

					FTotNightTime = FTotNightTime + FNightTime(intLoop)
				Next


				Fholidaywdtime	= rsget("holidaywdtime")
				FtotPaySum	=rsget("totPaySum")
				Fregdate    = rsget("regdate")
				Flastupdate = rsget("lastupdate")
				Fadminid    = rsget("adminid")
				Fino		=rsget("ino")
				Fposit_sn 	= rsget("posit_sn")
				Fposit_name = rsget("posit_name")
				Fdepartment_id = rsget("department_id")
				FdepartmentNameFull = rsget("departmentfullname")
				Fjobkind	= rsget("jobkind")
				Fplacekind		= rsget("placekind")

			END IF
		rsget.close
		END Function

		'//������� �� ���ϸ���Ʈ
		public Function fnGetPayPatternList
			Dim strSql

			IF Fpart_sn = "" THEN Fpart_sn = 0

			strSql ="[db_partner].[dbo].sp_Ten_user_defaultpay_pattern_GetListCnt("&Fpart_sn&",'"&Fpatternname&"')"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close

			IF FTotCnt > 0 THEN
			FSPageNo = (FPageSize*(FCurrPage-1)) + 1
			FEPageNo = FPageSize*FCurrPage

			strSql ="[db_partner].[dbo].sp_Ten_user_defaultpay_pattern_GetList("&Fpart_sn&",'"&Fpatternname&"',"&FSPageNo&","&FEPageNo&")"
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				fnGetPayPatternList = rsget.getRows()
			END IF
			rsget.close
			END IF
		End Function


		'//������� �� ���ϵ�����
		public Function fnGetPayPatternData
		IF FpatternSeq = "" THEN Exit Function
		Dim strSql
		Dim intLoop
		Dim NST(8) ,NET(8), NBST(8), NBET(8)
		strSql ="db_partner.dbo.sp_Ten_user_defaultpay_pattern_GetData("&FpatternSeq&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				Fpart_sn		= rsget("part_sn")
				FpatternName		= rsget("patternName")
				Fdefaultpay = rsget("defaultpay")
				Ffoodpay	= rsget("foodpay")
				Fjobpay		= rsget("jobpay")
				FinBreakTime = rsget("inBreakTime")
				FOverTime	= rsget("overtime")/60

				FStartTime(1)   = rsget("sunStart")
				FEndTime(1)     = rsget("sunEnd")
				FBreakSTime(1)  = rsget("sunBreakS")
				FBreakETime(1)  = rsget("sunBreakE")
				Fworktype(1)	= rsget("sunworktype")

				FStartTime(2)   = rsget("monStart")
				FEndTime(2)     = rsget("monEnd")
				FBreakSTime(2)   = rsget("monBreakS")
				FBreakETime(2)   = rsget("monBreakE")
				Fworktype(2)	= rsget("monworktype")

				FStartTime(3)   = rsget("tueStart")
				FEndTime(3)     = rsget("tueEnd")
				FBreakSTime(3)   = rsget("tueBreakS")
				FBreakETime(3)   = rsget("tueBreakE")
				Fworktype(3)	= rsget("tueworktype")

				FStartTime(4)   = rsget("wedStart")
				FEndTime(4)     = rsget("wedEnd")
				FBreakSTime(4)   = rsget("wedBreakS")
				FBreakETime(4)   = rsget("wedBreakE")
				Fworktype(4)	= rsget("wedworktype")

				FStartTime(5)   = rsget("thuStart")
				FEndTime(5)     = rsget("thuEnd")
				FBreakSTime(5)   = rsget("thuBreakS")
				FBreakETime(5)   = rsget("thuBreakE")
				Fworktype(5)	= rsget("thuworktype")

				FStartTime(6)   = rsget("friStart")
				FEndTime(6)     = rsget("friEnd")
				FBreakSTime(6)   = rsget("friBreakS")
				FBreakETime(6)   = rsget("friBreakE")
				Fworktype(6)	= rsget("friworktype")

				FStartTime(7)   = rsget("satStart")
				FEndTime(7)     = rsget("satEnd")
				FBreakSTime(7)   = rsget("satBreakS")
				FBreakETime(7)   = rsget("satBreakE")
				Fworktype(7)	= rsget("satworktype")

				FTotDutyTime = 0
				FTotNightTime = 0

				For intLoop = 1 To 7
				FStartHour(intLoop) = format00(2,Fix(FStartTime(intLoop)/60))
				FStartMinute(intLoop) = format00(2,(FStartTime(intLoop) mod 60))

				FEndHour(intLoop) = format00(2,Fix(FEndTime(intLoop)/60))
				FEndMinute(intLoop) = format00(2,(FEndTime(intLoop) mod 60))

				FBreakSHour(intLoop) = format00(2,Fix(FBreakSTime(intLoop)/60))
				FBreakSMinute(intLoop) = format00(2,(FBreakSTime(intLoop) mod 60))
				FBreakEHour(intLoop) = format00(2,Fix(FBreakETime(intLoop)/60))
				FBreakEMinute(intLoop) = format00(2,(FBreakETime(intLoop) mod 60))

				IF  FinBreakTime  THEN  '�ް�ð��� �ٹ��ð���  �����ϸ�..	�ް�ð��� ���� �ʴ´�.
					FDutyTime(intLoop)	= format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop))/60)) + ":"+Format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop)) mod 60))
					FTotDutyTime = FTotDutyTime + FEndTime(intLoop)  - FStartTime(intLoop)
				ELSE

					FDutyTime(intLoop)	= format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop) - (FBreakETime(intLoop)-FBreakSTime(intLoop)))/60)) + ":"+Format00(2,Fix((FEndTime(intLoop)  - FStartTime(intLoop) - (FBreakETime(intLoop)-FBreakSTime(intLoop))) mod 60))
					FTotDutyTime = FTotDutyTime + FEndTime(intLoop)  - FStartTime(intLoop) -(FBreakETime(intLoop)-FBreakSTime(intLoop))
				END IF

					'�߰��ٹ��ð�
					IF FEndTime(intLoop) >= (22*60) AND FStartTime(intLoop)< (30*60) THEN	'����ð��� 10�� �����̰� ���۽ð��� ��ħ 6�� �����϶� - �߰��ٹ�
						IF FStartTime(intLoop) < (22*60) THEN
							NST(intLoop) = 22*60
						ELSE
							NST(intLoop) = FStartTime(intLoop)
						END IF

						IF FEndTime(intLoop) > (30*60) THEN
							NET(intLoop) = 30*60
						ELSE
							NET(intLoop) = FEndTime(intLoop)
						END IF

						IF FBreakSTime(intLoop) < (22*60) THEN
							NBST(intLoop) = 22*60
						ELSEIF 	FBreakSTime(intLoop) >= (30*60) THEN
							NBST(intLoop) = 0
						ELSE
							NBST(intLoop) = FBreakSTime(intLoop)
						END IF

						IF FBreakETime(intLoop) < (22*60) THEN
							NBET(intLoop) = 22*60
						ELSEIF FBreakETime(intLoop) > (30*60) THEN
							NBET(intLoop) = 0
						ELSE
							NBET(intLoop) = FBreakETime(intLoop)
						END IF

						IF FinBreakTime THEN
							FNightTime(intLoop) = NET(intLoop)   - NST(intLoop)
						ELSE
							FNightTime(intLoop) =NET(intLoop)   - NST(intLoop) -(NBET(intLoop)-NBST(intLoop))
						END IF
					END IF

					FTotNightTime = FTotNightTime + FNightTime(intLoop)
				Next



				Fholidaywdtime	= rsget("holidaywdtime")
				FtotPaySum		=rsget("totPaySum")
				Fregdate    		= rsget("regdate")
				Flastupdate 		= rsget("lastupdate")
				Fadminid    		= rsget("adminid")
			END IF
		rsget.close
		End Function

End Class

 '�޿�����
Class CPay
	public FSearchText
	public Fworktime
	public Fextendtime
	public Finighttime
	public Fholidaytime
	public Ftimepay
	public Fextendpay
	public Fnightpay
	public Fholidaypay
	public Fwholidaypay
	public Ffoodpay
	public Fjobpay
	public Foutstandingpay
	public Ftotpay
	public Fnpensionpay
	public Fhealthinspay
	public Frecuinspay
	public Funempinspay
	public Ftaxtotpay
	public Frealtotpay
	public Fregdate
	public Fadminid
	public Fstate
	public Fempno
	public Fyyyymmdd
	public Fweekday
	public Fstartwork
	public Fendwork
	public Fbreaktime
	public Fyyyymm
	public FSyyyymm
	public FEyyyymm
	public Fyearpay
	public Fbonuspay
	public FPreyyyymmdd

	public FPageSize
	public FCurrPage
	public FSPageNo
	public FEPageNo
	public FTotCnt

	public FSearchType
	public Fpart_sn
	public Fposit_sn
	public Fjob_sn
	public Fstatediv
	public Forderby
	public Fusername
	public Fjoinday
	public Fretireday
	public Fposit_name
	public fshopid
	public Fdefaultpay
	public Fstartdate
	public Fenddate
	public FinBreakTime
	public FOverTime
	public Fnight
	public FStartTime(8)
	public FEndTime(8)
	public FBreakSTime(8)
	public FBreakETime(8)
	public FStartHour(8)
	public FStartMinute(8)
	public FEndHour(8)
	public FEndMinute(8)
	public FBreakSHour(8)
	public FBreakSMinute(8)
	public FBreakEHour(8)
	public FBreakEMinute(8)
	public FDutyTime(8)
	public FNightTime(8)
	public FdefaultTime(8)
	public Fworktype(8)

	public Fholidaywdtime
	public FTotDutyTime
	public FTotNightTime
	public FTotpaySum
	public Flongtimepay
	public Faddpay
	public FIsMonth
	public FiNo
	public Fdefaultpayseq
	public FWeekWorkDay		'// ��������(�ñ����� �� �ٹ��ϼ��� �Ѵ�.)
	public Fworkday

	public Fdepartment_id
	public Finc_subdepartment

	public FReworktime
	public FReextendtime
	public FRenighttime
	public FReholidaytime
	public FRetimepay
	public FpreDefaultpay
	public FpreFoodpay

	public  FReFoodtime
	public  FReExtimepay
	public FReNTtimepay
	public FReHDtimepay
	public FReFtimepay
	public FReTotpay
  public FReWorkday
	public FSearchDate
	public Function fnGetMonthlypayList
		Dim strSql

		IF Fpart_sn = "" THEN Fpart_sn = 0
		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fstate = "" THEN Fstate = 9				'�ۼ����� 0 �̹Ƿ� 9 �� ������ ����.
		IF FIsMonth = "" THEN FIsMonth = 0
		IF FSearchDate = "" THEN FSearchDate ="Y"
		strSql ="[db_partner].[dbo].sp_Ten_user_monthlypay_GetListCnt("&Fpart_sn&","&Fposit_sn&","&Fstate&",'"&FSearchType&"','"&FSearchText&"','"&FSyyyymm&"','"&FEyyyymm&"','"&FIsMonth&"','"&fshopid&"', '" & CStr(Fdepartment_id) & "', '" &CStr(Finc_subdepartment)& "','"&FSearchDate&"')"
		  rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close

		IF FTotCnt > 0 THEN
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage

		strSql ="[db_partner].[dbo].sp_Ten_user_monthlypay_GetList("&Fpart_sn&","&Fposit_sn&","&Fstate&",'"&FSearchType&"','"&FSearchText&"','"&FSyyyymm&"','"&FEyyyymm&"','"&FIsMonth&"','"&fshopid&"','"&Forderby&"',"&FSPageNo&","&FEPageNo&", '" & CStr(Fdepartment_id) & "', '" & CStr(Finc_subdepartment) &"','"&FSearchDate&"')"
	    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMonthlypayList = rsget.getRows()
		END IF
		rsget.close
		END IF
	End Function

	public Function fnGetMonthlypayListCSV
		Dim strSql

		IF Fpart_sn = "" THEN Fpart_sn = 0
		IF Fposit_sn = "" THEN Fposit_sn = 0
		IF Fstate = "" THEN Fstate = 9				'�ۼ����� 0 �̹Ƿ� 9 �� ������ ����.
		IF FIsMonth = "" THEN FIsMonth = 0
		IF FSearchDate = "" THEN FSearchDate ="Y"

		strSql ="[db_partner].[dbo].sp_Ten_user_monthlypay_GetList_CSV("&Fpart_sn&","&Fposit_sn&","&Fstate&",'"&FSearchType&"','"&FSearchText&"','"&FSyyyymm&"','"&FEyyyymm&"','"&FIsMonth&"','"&fshopid&"','"&Forderby&"', '" & CStr(Fdepartment_id) & "', '" & CStr(Finc_subdepartment) &"','"&FSearchDate&"')"
	    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetMonthlypayListCSV = rsget.getRows()
		END IF
		rsget.close

	End Function


	'�˻� ���ǿ� �ش��ϴ� �� �޿����� ��������
	public Function fnGetDailypayData
	Dim strSql
		strSql ="db_partner.dbo.sp_Ten_user_dailypay_GetData('"&Fempno&"','"&FSyyyymm&"','"&FEyyyymm&"')"

		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				 fnGetDailypayData = rsget.getRows()
			END IF
		rsget.close
	End Function

	'������ 26~���� ����� ����Ʈ ��������
	public function fnGetPreReDailypayData
	Dim strSql
		strSql ="db_partner.dbo.[sp_Ten_user_dailypay_GetPreReData]('"&Fempno&"','"&FPreyyyymmdd&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				 fnGetPreReDailypayData = rsget.getRows()
			END IF
		rsget.close
	End Function

	'�˻� ���ǿ� �ش��ϴ� �� ���� �޿����� ��������
	public Function fnGetPreDailypayData
	Dim strSql
		strSql ="db_partner.dbo.sp_Ten_user_dailypay_GetPreData('"&Fempno&"','"&FPreyyyymmdd&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				 fnGetPreDailypayData = rsget.getRows()
			END IF
		rsget.close
	End Function

	'����� �ش��ϴ� ������� �� ������� ��������
	public Function fnGetUserPayData
		IF Fempno = "" THEN Exit Function
		IF Fyyyymm = "" THEN Exit Function
		IF Fino = "" THEN Fino = 1
		Dim strSql, intLoop
		Dim NST(8) ,NET(8), NBST(8), NBET(8)

		strSql ="db_partner.dbo.Sp_Ten_User_Tenbyten_Defaultpay_Getdata('"&Fempno&"','"&Fyyyymm&"',"&Fino&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				Fempno		=rsget("empno")
				Fusername 	=rsget("username")
				Fjoinday 		=rsget("joinday")
				Fstatediv 		=rsget("statediv")
				Fretireday		=rsget("retireday")
				Fpart_sn		=rsget("part_sn")
				Fposit_sn		=rsget("posit_sn")
				Fjob_sn		=rsget("job_sn")
				Fposit_name	=rsget("posit_name")

				FiNo			=rsget("ino")
				Fstartdate		=rsget("startdate")
				Fenddate		=rsget("enddate")
				Fdefaultpay = rsget("defaultpay")
				Ffoodpay	= rsget("foodpay")
				Fjobpay		= rsget("jobpay")
				FinBreakTime = rsget("inBreakTime")
				FOverTime	= rsget("overtime")/60

				FStartTime(1)   = rsget("sunStart")
				FEndTime(1)     = rsget("sunEnd")
				FBreakSTime(1)   = rsget("sunBreakS")
				FBreakETime(1)   = rsget("sunBreakE")
				Fworktype(1)	= rsget("sunworktype")

				FStartTime(2)   = rsget("monStart")
				FEndTime(2)     = rsget("monEnd")
				FBreakSTime(2)   = rsget("monBreakS")
				FBreakETime(2)   = rsget("monBreakE")
				Fworktype(2)	= rsget("monworktype")

				FStartTime(3)   = rsget("tueStart")
				FEndTime(3)     = rsget("tueEnd")
				FBreakSTime(3)   = rsget("tueBreakS")
				FBreakETime(3)   = rsget("tueBreakE")
				Fworktype(3)	= rsget("tueworktype")

				FStartTime(4)   = rsget("wedStart")
				FEndTime(4)     = rsget("wedEnd")
				FBreakSTime(4)   = rsget("wedBreakS")
				FBreakETime(4)   = rsget("wedBreakE")
				Fworktype(4)	= rsget("wedworktype")

				FStartTime(5)   = rsget("thuStart")
				FEndTime(5)     = rsget("thuEnd")
				FBreakSTime(5)   = rsget("thuBreakS")
				FBreakETime(5)   = rsget("thuBreakE")
				Fworktype(5)	= rsget("thuworktype")

				FStartTime(6)   = rsget("friStart")
				FEndTime(6)     = rsget("friEnd")
				FBreakSTime(6)   = rsget("friBreakS")
				FBreakETime(6)   = rsget("friBreakE")
				Fworktype(6)	= rsget("friworktype")

				FStartTime(7)   = rsget("satStart")
				FEndTime(7)     = rsget("satEnd")
				FBreakSTime(7)   = rsget("satBreakS")
				FBreakETime(7)   = rsget("satBreakE")
				Fworktype(7)	= rsget("satworktype")

				Fholidaywdtime	=rsget("holidaywdtime")
				Fdefaultpayseq	= rsget("defaultpayseq")

				FpreDefaultpay = rsget("predefaultpay")
				FpreFoodpay = rsget("prefoodpay")
				FTotDutyTime = 0
				FWeekWorkDay = 0

				For intLoop = 1 To 7
					IF FinBreakTime  THEN
						FdefaultTime(intLoop) = FEndTime(intLoop) - FStartTime(intLoop)
					ELSE
						FdefaultTime(intLoop) = FEndTime(intLoop) - FStartTime(intLoop) - (FBreakETime(intLoop) - FBreakSTime(intLoop))
					END IF

					if (FdefaultTime(intLoop) >= 60) then
						'60�� �̻� �ٹ��ϸ� �ٹ��ϼ��� �����Ѵ�.(�Ĵ�����)
						FWeekWorkDay = FWeekWorkDay + 1
					end if

					FTotDutyTime = FTotDutyTime + FdefaultTime(intLoop)
					FdefaultTime(intLoop) = format00(2,Fix(FdefaultTime(intLoop)/60)) &":"&format00(2,(FdefaultTime(intLoop) mod 60))

					'�߰��ٹ��ð�
					IF FEndTime(intLoop) >= (22*60) AND FStartTime(intLoop)< (30*60) THEN	'����ð��� 10�� �����̰� ���۽ð��� ��ħ 6�� �����϶� - �߰��ٹ�
						IF FStartTime(intLoop) < (22*60) THEN
							NST(intLoop) = 22*60
						ELSE
							NST(intLoop) = FStartTime(intLoop)
						END IF

						IF FEndTime(intLoop) > (30*60) THEN
							NET(intLoop) = 30*60
						ELSE
							NET(intLoop) = FEndTime(intLoop)
						END IF

						IF FBreakSTime(intLoop) < (22*60) THEN
							NBST(intLoop) = 22*60
						ELSEIF 	FBreakSTime(intLoop) >= (30*60) THEN
							NBST(intLoop) = 0
						ELSE
							NBST(intLoop) = FBreakSTime(intLoop)
						END IF

						IF FBreakETime(intLoop) < (22*60) THEN
							NBET(intLoop) = 22*60
						ELSEIF FBreakETime(intLoop) > (30*60) THEN
							NBET(intLoop) = 0
						ELSE
							NBET(intLoop) = FBreakETime(intLoop)
						END IF

						IF FinBreakTime THEN
							FNightTime(intLoop) = NET(intLoop)   - NST(intLoop)
						ELSE
							FNightTime(intLoop) =NET(intLoop)   - NST(intLoop) -(NBET(intLoop)-NBST(intLoop))
						END IF
					END IF

					FTotNightTime = FTotNightTime + FNightTime(intLoop)
				Next


				FtotPaySum	=rsget("totPaySum")
			END IF
		rsget.close
	End Function


	'�˻� ���ǿ� �ش��ϴ� ���� �޿� ���� ��������
	public Function fnGetmonthlypayData
	Dim strSql
		strSql ="db_partner.dbo.sp_Ten_user_monthlypay_GetData('"&Fempno&"','"&Fyyyymm&"',"&Fino&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN

				Fworktime               = rsget("worktime")
				Fextendtime             = rsget("extendtime")
				Fnight             		= rsget("nighttime")
				Fholidaytime            = rsget("holidaytime")
				Ftimepay                = rsget("timepay")
				Fextendpay              = rsget("extendpay")
				Fnightpay               = rsget("nightpay")
				Fholidaypay            	= rsget("holidaypay")
				Ffoodpay                = rsget("foodpay")			'// �ѽĴ�(�հ�ݾ�)
				Fjobpay                 = rsget("jobpay")
				Foutstandingpay         = rsget("outstandingpay")
				Flongtimepay			= rsget("longtimepay")
				Faddpay					= rsget("addpay")
				Fyearpay				= rsget("yearpay")
				Fbonuspay				= rsget("bonuspay")
				Ftotpay                 = rsget("totpay")
				Fnpensionpay            = rsget("npensionpay")
				Fhealthinspay           = rsget("healthinspay")
				Frecuinspay             = rsget("recuinspay")
				Funempinspay            = rsget("unempinspay")
				Ftaxtotpay              = rsget("taxtotpay")
				Frealtotpay             = rsget("realtotpay")
				Fregdate                = rsget("regdate")
				Fadminid                = rsget("adminid")
				Fstate                  = rsget("paystate")
				Fworkday                = rsget("workday")

				FReworktime             = rsget("recaltime")
				FReextendtime           = rsget("recalexttime")
				FRenighttime            = rsget("recalnttime")
				FReholidaytime          = rsget("recalhdtime")
				FReFoodtime							= rsget("recalftime")
				FRetimepay              = rsget("recalpay")
				FReExtimepay            = rsget("recalexpay")
				FReNTtimepay            = rsget("recalntpay")
				FReHDtimepay            = rsget("recalhdpay")
				FReFtimepay             = rsget("recalfpay")
				FReTotpay								= rsget("recaltotpay")
				FReWorkday							= rsget("recalworkday")
			END IF
		rsget.close
	End Function

	 public Function fnGetDailyPayState
	 Dim objCmd
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_user_dailypay_GetState]('"&sEmpNo&"','"&Fyyyymm&"')}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With
		    fnGetDailyPayState = objCmd(0).Value
		Set objCmd = nothing
	End Function
End Class

Function fnSetStateDesc(ByVal iState)
	SELECT CASE iState
	Case "1"
		fnSetStateDesc = "�ۼ��Ϸ�"
	Case "5"
		fnSetStateDesc = "�濵����Ȯ�οϷ�"
	Case "8"
		fnSetStateDesc = "�ԱݿϷ�"
	Case ELSE
		fnSetStateDesc = "�ۼ���"
	END SELECT
end Function

'���� ���� -> �ؽ�Ʈ
 Function fnGetStringWD(ByVal iWD)
	 SELECT CASE iWD
	 CASE 1
	 	fnGetStringWD = "��"
	 CASE 2
	 	fnGetStringWD = "��"
	 CASE 3
	 	fnGetStringWD = "ȭ"
	 CASE 4
	 	fnGetStringWD = "��"
	 CASE 5
	 	fnGetStringWD = "��"
	 CASE 6
	 	fnGetStringWD = "��"
	 CASE 7
	 	fnGetStringWD = "��"
	END SELECT
 End Function

 '�޿����� ��ϻ��� ->�ؽ�Ʈ
 Function fnGetStateDesc(ByVal iState)
 	SELECT CASE iState
 	 CASE 0
	 	fnGetStateDesc = "<font color=""red"">�ۼ���</font>"
	 CASE 1
	 	fnGetStateDesc = "�ۼ��Ϸ�"
	 CASE 5
	 	fnGetStateDesc = "Ȯ�οϷ�"
	 CASE 7
	 	fnGetStateDesc = "�ԱݿϷ�"
	 CASE ELSE
	 	 fnGetStateDesc = "<font color=""blue"">�Է�</font>"
	END SELECT
End Function

Function fnSetTimeFormat(ByVal iMinute)
	IF iMinute = "" or iMinute =0  THEN
		fnSetTimeFormat = "00:00"
	ELSEIF iMinute < 0 then
		fnSetTimeFormat = "-"&format00(2,Fix((iMinute/60)*-1)) &":"& format00(2,((iMinute mod 60)*-1))
	ELSE
 		fnSetTimeFormat = format00(2,Fix(iMinute/60)) &":"& format00(2,(iMinute mod 60))
	END IF
End Function

Function fnSetMinuteFromTimeForm(ByVal sTime)
	Dim returnValue,arrValue
	IF sTime = "" or isNull(sTime) or sTime="0"  THEN
		returnValue = 0
	ELSE
		arrValue = split(sTime,":")
		if left(arrValue(0),1)="-"   THEN
				returnValue = arrValue(0)*60+arrValue(1)*-1
		else
				returnValue = arrValue(0)*60+arrValue(1)
		end if
	END IF
	fnSetMinuteFromTimeForm = returnValue
End Function
%>
