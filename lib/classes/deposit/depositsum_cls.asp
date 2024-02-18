<%
'###########################################################
' Description :  ��ġ�� ���� ��� Ŭ����
' History : 2012.12.05 �ѿ�� ����
'###########################################################

class cdepositsum_oneitem
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub

	public fidx
	public fuserid
	public fdeposit
	public fjukyocd
	public fjukyo
	public forderserial
	public fdeleteyn
	public freguserid
	public fdeluserid
	public fregdate
	public ffixyyyymmdd
	public fyyyymm
	public freforeremaincash
	public fsellCash
	public fuseCash
	public frefundCash
	public fuseroutCash
	public fdelcash
	public fremaincash
	public fyyyymmdd
	public fsitename
	public fshopid
end class

class cdepositsum_list
	Private Sub Class_Initialize()
		FCurrPage =1
		FPageSize = 50
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()

	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	public FItemList()
	public FCurrPage
	public FPageSize
	public FResultCount
	public FTotalCount
	public FScrollCount
	public FTotalPage

	public FRectonoffgubun
	public FRectStartdate
	public FRectEndDate
	public frectjukyocd

	'/admin/maechul/managementsupport/depositsum_month.asp
	public function fdepositsum_sell_month
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymm"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " , db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--�̿��ܾ�
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(use10,0)+isnull(use200,0)+isnull(use210,0)+isnull(use100,0)"
			sql = sql & " 	+isnull(use300,0)+isnull(use900,0)"
			sql = sql & " ) as remaincash"		'/--�ܾ�
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--��ġ������
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- ���
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--������ȯ��
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--ȸ��Ż��

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " ,0 as reforeremaincash"		'/--�̿��ܾ�
			sql = sql + " ,0 as remaincash"
			sql = sql + " ,0 as sellCash"
			sql = sql + " ,0 as useCash"
			sql = sql + " ,0 as refundCash"
			sql = sql & " ,0 as useroutCash"		'/--ȸ��Ż��
		else
			sql = sql & " , db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"') as reforeremaincash"		'/--�̿��ܾ�
			sql = sql & " ,("
			sql = sql & " 	db_user.dbo.uf_deposit_beforeremaincash(t.yyyymm,'"&FRectonoffgubun&"')"
			sql = sql & " 	+isnull(use10,0)+isnull(use200,0)+isnull(use210,0)+isnull(use100,0)"
			sql = sql & " 	+isnull(use300,0)+isnull(use900,0)"
			sql = sql & " ) as remaincash"		'/--�ܾ�
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--��ġ������
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- ���
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--������ȯ��
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--ȸ��Ż��
		end if

		sql = sql & " ,0 as delcash"		'/--�Ҹ�
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(7),L.fixyyyymmdd,121) as yyyymm"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='10' then l.deposit"
		sql = sql & " 		else 0 end) as use10"		'/--��ġ��ȯ��(���ȯ��), ��ǰ�������ȯ��	(��ġ������ �������� �ٽ� ��ġ������ ��ȯ)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='100' then l.deposit"
		sql = sql & " 		else 0 end) as use100"		'/--��ǰ����
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='200' then l.deposit"
		sql = sql & " 		else 0 end) as use200"		'/--��ġ����ȯ(CS), ��ǰ ó�� �� ��ġ�� ȯ��, �ֹ� ��� �� ��ġ�� ȯ�� (���� �����ε� ��ġ������ �����Ѱ�)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='210' then l.deposit"
		sql = sql & " 		else 0 end) as use210"		'/--��ġ����ȯ(������ ��ǰǰ���� �ֹ��Ұ�..��ġ������ ��ȯ���� �ذ�)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='300' then l.deposit"
		sql = sql & " 		else 0 end) as use300"		'/--������ȯ��, ��ġ�� ���������� ȯ��
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd in ('900','9999') then l.deposit"
		sql = sql & " 		else 0 end) as use900"		'/--ȸ��Ż�� ---9999 ���?/2013/11/12
		sql = sql & " 	from [db_user].dbo.tbl_depositlog l"
		sql = sql & " 	where L.deleteyn='N' " & sqlsearch
		sql = sql & " 	and l.fixyyyymmdd is not null"
		sql = sql & " 	group by convert(varchar(7),L.fixyyyymmdd,121)"
		sql = sql & " ) as t"
		sql = sql & " order by t.yyyymm desc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fYYYYMM			= rsget("YYYYMM")
				FItemList(i).freforeremaincash			= rsget("reforeremaincash")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")
				FItemList(i).fremaincash			= rsget("remaincash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'/admin/maechul/managementsupport/depositsum_day.asp
	public function fdepositsum_sell_day
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) >='" + CStr(FRectStartdate) + "'"
		end if
		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and convert(varchar(10),l.fixyyyymmdd,21) <'" + CStr(FRectEndDate) + "'"
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " t.yyyymmdd"

		if FRectonoffgubun="ONLINE" then
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--��ġ������
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- ���
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--������ȯ��
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--ȸ��Ż��

		elseif FRectonoffgubun="OFFLINE" then
			sql = sql & " ,0 as sellCash"		'/--��ġ������
			sql = sql & " ,0 as useCash"		'/-- ���
			sql = sql & " ,0 as refundCash"		'/--������ȯ��
			sql = sql & " ,0 as useroutCash"		'/--ȸ��Ż��
		else
			sql = sql & " ,("
			sql = sql & " 	isnull(use10,0) + isnull(use200,0) + isnull(use210,0)"
			sql = sql & " 	) as sellCash"		'/--��ġ������
			sql = sql & " ,("
			sql = sql & " 	isnull(use100,0)"
			sql = sql & " 	) as useCash"		'/-- ���
			sql = sql & " ,("
			sql = sql & " 	isnull(use300,0)"
			sql = sql & " 	) as refundCash"		'/--������ȯ��
			sql = sql & " ,isnull(use900,0) as useroutCash"		'/--ȸ��Ż��
		end if

		sql = sql & " ,0 as delcash"		'/--�Ҹ�
		sql = sql & " from ("
		sql = sql & " 	select"
		sql = sql & " 	convert(varchar(10),L.fixyyyymmdd,121) as yyyymmdd"
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='10' then l.deposit"
		sql = sql & " 		else 0 end) as use10"		'/--��ġ��ȯ��(���ȯ��), ��ǰ�������ȯ��	(��ġ������ �������� �ٽ� ��ġ������ ��ȯ)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='100' then l.deposit"
		sql = sql & " 		else 0 end) as use100"		'/--��ǰ����
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='200' then l.deposit"
		sql = sql & " 		else 0 end) as use200"		'/--��ġ����ȯ(CS), ��ǰ ó�� �� ��ġ�� ȯ��, �ֹ� ��� �� ��ġ�� ȯ�� (���� �����ε� ��ġ������ �����Ѱ�)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='210' then l.deposit"
		sql = sql & " 		else 0 end) as use210"		'/--��ġ����ȯ(������ ��ǰǰ���� �ֹ��Ұ�..��ġ������ ��ȯ���� �ذ�)
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd='300' then l.deposit"
		sql = sql & " 		else 0 end) as use300"		'/--������ȯ��, ��ġ�� ���������� ȯ��
		sql = sql & " 	,sum(CASE"
		sql = sql & " 		when jukyocd in ('900','9999') then l.deposit"
		sql = sql & " 		else 0 end) as use900"		'/--ȸ��Ż��
		sql = sql & " 	from [db_user].dbo.tbl_depositlog l"
		sql = sql & " 	where L.deleteyn='N' " & sqlsearch
		sql = sql & " 	and l.fixyyyymmdd is not null"
		sql = sql & " 	group by convert(varchar(10),L.fixyyyymmdd,121)"
		sql = sql & " ) as t"
		sql = sql & " order by t.yyyymmdd asc"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")
				FItemList(i).fsellCash			= rsget("sellCash")
				FItemList(i).fuseCash			= rsget("useCash")
				FItemList(i).frefundCash			= rsget("refundCash")
				FItemList(i).fuseroutCash			= rsget("useroutCash")
				FItemList(i).fdelcash			= rsget("delcash")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function

	'//admin/maechul/managementsupport/depositsum_use_list.asp
	public function fdepositsum_use_list
		dim i , sql , sqlsearch

		if FRectStartdate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd >='" + CStr(FRectStartdate) + "'"
		end if

		if FRectEndDate<>"" then
			sqlsearch = sqlsearch + " and l.fixyyyymmdd <'" + CStr(FRectEndDate) + "'"
		end if

		if frectjukyocd = "sellCash" then	'/��ġ������
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(jukyocd='10' or jukyocd='200' or jukyocd='210')"
				sqlsearch = sqlsearch + " )"
			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//�Ⱥ��̱� ����, 0���� �ӽ�ó��
			else
				sqlsearch = sqlsearch + " and ("
				sqlsearch = sqlsearch + " 	(jukyocd='10' or jukyocd='200')"
				sqlsearch = sqlsearch + " )"
			end if

		elseif frectjukyocd = "useCash" then	'/������
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd='100'"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//�Ⱥ��̱� ����, 0���� �ӽ�ó��

			else
				sqlsearch = sqlsearch + " and jukyocd='100'"
			end if

		elseif frectjukyocd = "refundCash" then		'/ȯ��
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd='300'"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//�Ⱥ��̱� ����, 0���� �ӽ�ó��

			else
				sqlsearch = sqlsearch + " and jukyocd='300'"
			end if

		elseif frectjukyocd = "useroutCash" then		'/ȸ��Ż��
			if FRectonoffgubun="ONLINE" then
				sqlsearch = sqlsearch + " and jukyocd in ('900','9999')"

			elseif FRectonoffgubun="OFFLINE" then
				sqlsearch = sqlsearch + " and jukyocd='0'"		'//�Ⱥ��̱� ����, 0���� �ӽ�ó��

			else
				sqlsearch = sqlsearch + " and jukyocd in ('900','9999')"
			end if

		elseif frectjukyocd = "delcash" then		'/�Ҹ�
			sqlsearch = sqlsearch + " and jukyocd='0'"		'//�Ⱥ��̱� ����, 0���� �ӽ�ó��
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage)
		sql = sql & " L.userid, L.deposit, L.jukyocd, L.jukyo, L.orderserial, L.deleteYn"
		sql = sql & " , L.fixyyyymmdd as yyyymmdd"
		sql = sql & " , isnull(m.sitename,lm.sitename) as sitename"
		sql = sql & " from [db_user].dbo.tbl_depositlog l"
		sql = sql & " left join db_order.dbo.tbl_order_master m"
		sql = sql & " 	on l.orderserial=m.orderserial"
		sql = sql & " left join db_log.dbo.tbl_old_order_master_2003 lm"
		sql = sql & " 	on l.orderserial=lm.orderserial"
		sql = sql & " where L.deleteyn='N' " & sqlsearch
		sql = sql & " and l.fixyyyymmdd is not null"
		sql = sql & " order by yyyymmdd"

		'response.write sql & "<Br>"
		rsget.open sql,dbget,1

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		redim FItemList(FresultCount)
		i = 0
		If Not rsget.Eof Then
			Do Until rsget.Eof
				set FItemList(i) = new cdepositsum_oneitem

				FItemList(i).fsitename			= rsget("sitename")
				FItemList(i).fuserid			= rsget("userid")
				FItemList(i).fdeposit		= rsget("deposit")
				FItemList(i).fjukyocd			= rsget("jukyocd")
				FItemList(i).fjukyo			= rsget("jukyo")
				FItemList(i).forderserial			= rsget("orderserial")
				FItemList(i).fdeleteYn			= rsget("deleteYn")
				FItemList(i).fyyyymmdd			= rsget("yyyymmdd")

			rsget.movenext
			i = i + 1
			Loop
		End If

		rsget.close
	end function
end class

'//����
function drawdepositjukyocd(selBoxName, selVal, chplg)
%>
    <select name="<%= selBoxName %>" <%= chplg %>>
		<option value="" <% if selVal="" then response.write " selected" %>>��ü</option>
		<option value="sellCash" <% if selVal="sellCash" then response.write " selected" %>>��ġ������</option>
		<option value="useCash" <% if selVal="useCash" then response.write " selected" %>>������</option>
		<option value="refundCash" <% if selVal="refundCash" then response.write " selected" %>>ȯ��</option>
		<option value="useroutCash" <% if selVal="useroutCash" then response.write " selected" %>>ȸ��Ż��</option>
		<option value="delcash" <% if selVal="delcash" then response.write " selected" %>>�Ҹ�</option>
	</select>
<%
end function

'//����
function getdepositjukyocd(selVal)

	if selVal = "" then exit function

	if selVal = "sellCash" then
		getgiftcardaccountdiv = "��ġ������"
	elseif selVal = "useCash" then
		getgiftcardaccountdiv = "������"
	elseif selVal = "refundCash" then
		getgiftcardaccountdiv = "ȯ��"
	elseif selVal = "useroutCash" then
		getgiftcardaccountdiv = "ȸ��Ż��"
	elseif selVal = "delcash" then
		getgiftcardaccountdiv = "�Ҹ�"
	end if
end function
%>