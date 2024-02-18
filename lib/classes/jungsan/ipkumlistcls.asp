<%
Sub drawSelectBoxBankList(selectBoxName,selectedId)
    dim tmp_str,strSql, arrrows, buf
    strSql = "select top 100 accountno,divcode,bkname,description,altName,inoutGbn,sortord from db_order.dbo.tbl_bank_div"
    strSql = strSql & " where sortord is Not NULL"
    strSql = strSql & " order by sortord"

    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) THEN
		arrrows = rsget.getRows()
	END IF
	rsget.close

	buf = "<select class='select' name='"&selectBoxName&"'>"
	buf= buf&"<option value=''>전체</option>"
	IF isArray(arrrows) then
	    For i =0 To UBOund(arrrows,2)
	         buf= buf&"<option value='"&arrrows(0,i)&"' "&CHKIIF(selectedId = arrrows(0,i),"selected","")&">"&arrrows(0,i)&" / "&arrrows(3,i)&" / "&arrrows(2,i)&"</option>"
	    next
	End IF
	buf= buf&"</select>"

	response.write buf
end Sub

Sub drawSelectBoxBankList_OLD(selectBoxName,selectedId)
	dim tmp_str,query1
   %>
	<select class="select" name="<%= selectBoxName %>">
		<option value="">전체</option>
		<option value="06505649001011" <% if selectedId = "06505649001011" then response.write " selected" %>>기업 06505649001011</option>
		<option value="">--------------</option>
		<option value="06505649001028" <% if selectedId = "06505649001028" then response.write " selected" %>>기업 06505649001028</option>
		<option value="06505649001042" <% if selectedId = "06505649001042" then response.write " selected" %>>기업 06505649001042</option>
		<!--<option value="06505649001050" <% if selectedId = "06505649001050" then response.write " selected" %>>기업 06505649001050</option>-->
   		<option value="27702818201021" <% if selectedId = "27702818201021" then response.write " selected" %>>기업 27702818201021</option>
		<!--<option value="27702818201039" <% if selectedId = "27702818201039" then response.write " selected" %>>기업 27702818201039</option>-->
		<!--<option value="27702818204011" <% if selectedId = "27702818204011" then response.write " selected" %>>기업 27702818204011</option>-->
		<!--<option value="53401017868" <% if selectedId = "53401017868" then response.write " selected" %>>신한 53401017868</option>-->
		<!--<option value="06022038611" <% if selectedId = "06022038611" then response.write " selected" %>>외환 06022038611</option>-->
		<option value="">--------------</option>
		<option value="09227549513001" <% if selectedId = "09227549513001" then response.write " selected" %>>우리 09227549513001</option>
		<option value="47030101014754" <% if selectedId = "47030101014754" then response.write " selected" %>>국민 47030101014754</option>
		<option value="14691000928804" <% if selectedId = "14691000928804" then response.write " selected" %>>하나 14691000928804</option>
		<!--<option value="53401016039" <% if selectedId = "53401016039" then response.write " selected" %>>조흥 53401016039</option>-->
		<option value="100016523130" <% if selectedId = "100016523130" then response.write " selected" %>>신한 100016523130</option>
		<option value="02901246118" <% if selectedId = "02901246118" then response.write " selected" %>>농협 02901246118</option>
		<option value="27702818201046" <% if selectedId = "27702818201046" then response.write " selected" %>>기업 27702818201046</option>
		<option value="">--------------</option>
		<option value="27702818201078" <% if selectedId = "27702818201078" then response.write " selected" %>>기업 27702818201078</option>
		<option value="27703783604018" <% if selectedId = "27703783604018" then response.write " selected" %>>기업 27703783604018</option>
		<option value="1140049124201" <% if selectedId = "1140049124201" then response.write " selected" %>>씨티은행 1140049124201</option>
		<option value="">--------------</option>
		<option value="27703918804031" <% if selectedId = "27703918804031" then response.write " selected" %>>기업 27703918804031</option>
		<option value="27703783604032" <% if selectedId = "27703783604032" then response.write " selected" %>>기업 27703783604032</option>
	</select>
<%
End Sub

Class ipkumListitem
	public Fidx
	public FBankdate
	public Fgubun
	public Fipkumuser
	public Fipkumsum
	public Fchulkumsum
	public Fremainsum
	public Fbankname
	public Forderserial
	public Ffinishstr
	public Fipkumstate
	public Fregdate
	public Ffinishuser
	public Ftenbank

	public Fpaperexist

	Private sub Class_Intialize()

	end sub

	Private sub Class_Terminate()

	end Sub

end Class

Class IpkumListNetteller

	public Finoutidx
	public Fmatchstate
	public Ftotmatchedprice
	public Fmatchprice
	public Fmatchdetailidx

	public Fbkno
	public Fbkacctno
	public Fbkname
	public Fbkdate
	public Fbkjukyo
	public Fbkcontent
	public Fbketc
	public Fbkinput
	public Fbkoutput
	public Fbkbalance
	public Fbkjango
	public Fbkxferdatetime
	public Ftlno
	public Fbkmemo1
	public Fbkmemo2
	public Fbkmemo3
	public Fbkmemo4
	public Fbktag1
	public Fbktag2
	public Fbktag3
	public finout_gubun

	public Fjungsanidx
	public Fjungsancnt
	public Ftotmatchprice

	public Forderserial
	public Fmatchmemo

	public function GetMatchStateName()
		if Fmatchstate="N" or IsNull(Fmatchstate) or Fmatchstate = "" then
			GetMatchStateName = "매칭이전"
		elseif Fmatchstate="H" then
			GetMatchStateName = "일부매칭"
		elseif Fmatchstate="Y" then
			GetMatchStateName = "매칭완료"
		elseif Fmatchstate="X" then
			GetMatchStateName = "매칭제외"
		else
			GetMatchStateName = Fmatchstate
		end if
	end function

	public function GetMatchStateColor()
		if Fmatchstate="N" or IsNull(Fmatchstate) or Fmatchstate = "" then
			GetMatchStateColor = "black"
		elseif Fmatchstate="H" then
			GetMatchStateColor = "green"
		elseif Fmatchstate="Y" then
			GetMatchStateColor = "blue"
		elseif Fmatchstate="X" then
			GetMatchStateColor = "gray"
		else
			GetMatchStateColor = "red"
		end if
	end function

	Private sub Class_Intialize()

	end sub

	Private sub Class_Terminate()

	end Sub

end Class

Class IpkumChecklist
	public Fipkumitem()
	public FipkumListNetteller()
	public Fipkumoneitem

	public Fckdate
	public FRectRegStart
	public FRectRegEnd
	public yyyy1,yyyy2 ''검색 날짜
	public mm1,mm2
	public dd1,dd2

	public FRectTXDayStart
	public FRectTXDayEnd
	public FRectJeokyo
	public FRectTXAmmount
	public FRectAcctNo
	public FRectInOutGubun
	public FRectExcluudeMatchFinish
	public FRectExcluudeCustomer
	public FRectExcluude10X10
	public FRectShowDismatch

	public FRectInOutIDX
	public FRectJungsanIDX

	public FSearchtype01
	public FOrderby	'검색순서(최근일)
	public FTotalCount
	public CTenbank ''검색하는 은행
	public ipkumname '입금자명
	public ipkumstate  '입금 미확인. 미처리

	public FScrollCount
	public FCurrpage
	public FPagesize
	public FTotalPage
	public FResultCount
	public cksheep
	public Frectidx
	Private Sub Class_Initialize()
		dim Fipkumitem()
		FScrollCount=5
		FPagesize=200


	end Sub

	Private Sub Class_Terminate()

	end Sub

	Public Sub Getipkumlist()
	dim strSql,othersql,i
	if Ctenbank<> "" then
		otherSql = otherSql + " and i.tenbank='" & Ctenbank & "'" + vbcrlf
	end if

	if ipkumname<>"" then
		otherSql= otherSql + " and i.jukyo like '%" & ipkumname & "%'" + vbcrlf
	end if

	if ipkumstate="1" then '매칭실패
		otherSql = otherSql + " and i.ipkumstate='1'" + vbcrlf
	elseif ipkumstate="0" then '미처리
		otherSql = otherSql + " and i.ipkumstate='0'" + vbcrlf
	else
	end if

	strSql = "select Count(i.idx) as cnt" + vbcrlf
	strSql = strSql + " from [db_order].[dbo].tbl_ipkum_list i " + vbcrlf
	strSql = strSql + " where i.bankdate between '" & dateserial(CStr(yyyy1),CStr(mm1),CStr(dd1)) & "' and '" & dateserial(CStr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
	strSql = strSql + othersql
	rsget.open strSql,dbget,1

	FtotalCount=rsget("cnt")

	rsget.close

	if cksheep=1 then
		strSql = "select " + vbcrlf
	else
		strSql = "select top " & CStr(FCurrpage*FPagesize) + vbcrlf
	end if
	strSql = strSql + " i.idx,i.bankdate,i.gubun,i.jukyo,i.ipkumsum,i.chulkumsum,i.remainsum" + vbcrlf
	strSql = strSql + ",i.bankname,i.orderserial,i.finishstr,i.ipkumstate,i.regdate,i.finishuser,i.tenbank" + vbcrlf
	strSql = strSql + ", (case when IsNull(o.orderserial, '') = '' then 'N' else 'Y' end) as paperexist " + vbcrlf
	strSql = strSql + " from " + vbcrlf
	strSql = strSql + " 	[db_order].[dbo].tbl_ipkum_list i " + vbcrlf
	strSql = strSql + " 	left join [db_order].[dbo].tbl_order_master o " + vbcrlf
	strSql = strSql + " 	on " + vbcrlf
	strSql = strSql + " 		1 = 1 " + vbcrlf
	strSql = strSql + " 		and i.orderserial = o.orderserial " + vbcrlf
	strSql = strSql + " 		and o.accountdiv = '7' " + vbcrlf
	strSql = strSql + " 		and ( " + vbcrlf
	strSql = strSql + " 			IsNull(o.authcode, '') <> '' " + vbcrlf
	strSql = strSql + " 			or " + vbcrlf
	strSql = strSql + " 			IsNull(o.cashreceiptreq, '') in ('R', 'T') " + vbcrlf
	strSql = strSql + " 	) " + vbcrlf

	strSql = strSql + " where i.bankdate between '" & dateserial(CStr(yyyy1),CStr(mm1),CStr(dd1)) & "' and '" & dateserial(CStr(yyyy2),CStr(mm2),CStr(dd2)) & "'" + vbcrlf
	strSql = strSql + othersql
	strSql = StrSql + " order by i.regdate desc "

	rsget.pagesize=FPagesize
	rsget.open strSql,dbget,1

	FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

	if FResultCount<1 then FResultCount=0

	FTotalPage = FTotalCount\FPagesize
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FTotalPage = FtotalPage +1
		end if


	i=0
	redim preserve Fipkumitem(FResultCount)

	if not rsget.eof then
	rsget.absolutepage=FCurrPage

	Do until rsget.eof
		set Fipkumitem(i) = new ipkumlistitem

			Fipkumitem(i).Fidx  =rsget("idx")
			Fipkumitem(i).FBankdate=rsget("bankdate")
			Fipkumitem(i).Fgubun=rsget("gubun")
			Fipkumitem(i).Fipkumuser=rsget("jukyo")
			Fipkumitem(i).Fipkumsum=rsget("ipkumsum")
			Fipkumitem(i).Fchulkumsum=rsget("chulkumsum")
			Fipkumitem(i).Fremainsum=rsget("remainsum")
			Fipkumitem(i).Fbankname=rsget("bankname")
			Fipkumitem(i).Forderserial=rsget("orderserial")
			Fipkumitem(i).Ffinishstr=rsget("finishstr")
			Fipkumitem(i).Fipkumstate=rsget("ipkumstate")
			Fipkumitem(i).Fregdate=rsget("regdate")
			Fipkumitem(i).Ffinishuser=rsget("finishuser")
			Fipkumitem(i).Ftenbank=rsget("tenbank")

			Fipkumitem(i).Fpaperexist=rsget("paperexist")
		rsget.movenext
		i=i+1
		loop
		rsget.close
	end if
	end Sub


	Public Sub GetipkumlistByIdx()
		dim strSql
		dim i

		strSql = " select idx,bankdate,gubun,jukyo,ipkumsum,chulkumsum,remainsum" + vbcrlf
		strSql = strSql + ",bankname,orderserial,finishstr,ipkumstate,regdate,finishuser,tenbank" + vbcrlf
		strSql = strSql + " from [db_order].[dbo].tbl_ipkum_list" + vbcrlf
		strSql = strSql + " where idx = '" + CStr(idx) + "'" + vbcrlf

		rsget.Open strSql,dbget,1

		set Fipkumoneitem = new ipkumlistitem

		if Not rsget.Eof then
			Fipkumoneitem.Fidx  =rsget("idx")
			Fipkumoneitem.FBankdate=rsget("bankdate")
			Fipkumoneitem.Fgubun=rsget("gubun")
			Fipkumoneitem.Fipkumuser=rsget("jukyo")
			Fipkumoneitem.Fipkumsum=rsget("ipkumsum")
			Fipkumoneitem.Fchulkumsum=rsget("chulkumsum")
			Fipkumoneitem.Fremainsum=rsget("remainsum")
			Fipkumoneitem.Fbankname=rsget("bankname")
			Fipkumoneitem.Forderserial=rsget("orderserial")
			Fipkumoneitem.Ffinishstr=rsget("finishstr")
			Fipkumoneitem.Fipkumstate=rsget("ipkumstate")
			Fipkumoneitem.Fregdate=rsget("regdate")
			Fipkumoneitem.Ffinishuser=rsget("finishuser")
			Fipkumoneitem.Ftenbank=rsget("tenbank")
		end if

		rsget.close

	end Sub

	'//admin/accounts/paymentist_accounts.asp
	Public Sub GetipkumlistAccounts()
		dim strSql,i

		strSql = "select Count(site_no) as cnt "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT L "
		strSql = strSql + " 	left join [db_order].[dbo].tbl_bank_div A "
		strSql = strSql + " 	on "
		strSql = strSql + " 		L.acct_no=A.accountno "
		strSql = strSql + " 	left join db_order.dbo.tbl_ipkum_list i "
		strSql = strSql + " 	on "
		strSql = strSql + " 	i.inoutidx = L.inoutidx "
		strSql = strSql + " where 1 = 1 "

		if (FRectInOutIDX <> "") then
			strSql = strSql + " and L.inoutidx='" & FRectInOutIDX & "'"
		end if

		if Ctenbank<> "" then
		strSql = strSql + " and L.acct_no='" & Ctenbank & "'"
		end if

		if (FSearchtype01<>"") then
			strSql= strSql + " and L.jeokyo like '%" & ipkumname & "%'"
		end if

		if (Fckdate<>"") then
			strSql = strSql + " and acct_txday >='" + CStr(Replace(FRectRegStart,"-","")) + "'"
			strSql = strSql + " and acct_txday <'" + CStr(Replace(FRectRegEnd,"-","")) + "'"
		end if

		if (FRectInOutGubun<>"") then
			strSql = strSql + " and INOUT_GUBUN = " + CStr(FRectInOutGubun) + " "
		end if

		if (FRectExcluudeMatchFinish<>"") then
			strSql = strSql + " and (IsNull(matchstate, 'N') <> 'Y' and IsNull(i.ipkumstate, 0) <> 7) "
		end if

		if (FRectExcluudeCustomer<>"") then
			strSql = strSql + " and IsNull(A.divcode, 'XX') not in ('00', '30') "
		end if

		if (FRectExcluude10X10<>"") then
			strSql = strSql + " and branch <> '03 0277' "
		end if

		if (FRectTXAmmount<>"") then
			strSql = strSql + " and tx_amt = " + CStr(FRectTXAmmount) + " "
		end if

		if (FRectJeokyo<>"") then
			strSql = strSql + " and L.jeokyo like '%" & html2db(FRectJeokyo) & "%'"
		end if

		if (FRectShowDismatch = "") then
			strSql = strSql + " and IsNull(L.matchstate, 'N') <> 'X' "
		end if

		''response.write strSql &"<br>"

		rsget.open strSql,dbget,1
		    FtotalCount=rsget("cnt")
		rsget.close

		strSql = "select top " & CStr(FCurrpage*FPagesize)
		strSql = strSql + " L.* "

		strSql = strSql + " , ( "
		strSql = strSql + " 	select "
		strSql = strSql + " 		IsNull(max(m.jungsanidx), 0) "
		strSql = strSql + " 	from "
		strSql = strSql + " 		db_jungsan.dbo.tbl_ipkum_match_master m "
		strSql = strSql + " 		join db_jungsan.dbo.tbl_ipkum_match_detail d "
		strSql = strSql + " 		on "
		strSql = strSql + " 			m.idx = d.masteridx "
		strSql = strSql + " 	where "
		strSql = strSql + " 		d.ipkummethod = 'BNK' and d.ipkumidx = L.inoutidx and d.useyn = 'Y' "
		strSql = strSql + " ) as jungsanidx "
		strSql = strSql + " , ( "
		strSql = strSql + " 	select "
		strSql = strSql + " 		count(m.jungsanidx) "
		strSql = strSql + " 	from "
		strSql = strSql + " 		db_jungsan.dbo.tbl_ipkum_match_master m "
		strSql = strSql + " 		join db_jungsan.dbo.tbl_ipkum_match_detail d "
		strSql = strSql + " 		on "
		strSql = strSql + " 			m.idx = d.masteridx "
		strSql = strSql + " 	where "
		strSql = strSql + " 		d.ipkummethod = 'BNK' and d.ipkumidx = L.inoutidx and d.useyn = 'Y' "
		strSql = strSql + " ) as jungsancnt "
		strSql = strSql + " , ( "
		strSql = strSql + " 	SELECT sum(matchprice) "
		strSql = strSql + " 	FROM db_jungsan.dbo.tbl_ipkum_match_master m "
		strSql = strSql + " 	INNER JOIN db_jungsan.dbo.tbl_ipkum_match_detail d ON m.idx = d.masteridx "
		strSql = strSql + " 	WHERE d.ipkummethod = 'BNK' "
		strSql = strSql + " 		AND d.ipkumidx = L.inoutidx "
		strSql = strSql + " 		AND d.useyn = 'Y' "
		strSql = strSql + " ) as totmatchprice "
		strSql = strSql + " , IsNull(i.orderserial, i.finishstr) as orderserial "
		strSql = strSql + " from "
		strSql = strSql + " 	[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT L "
		strSql = strSql + " 	left join [db_order].[dbo].tbl_bank_div A "
		strSql = strSql + " 	on "
		strSql = strSql + " 		L.acct_no=A.accountno "
		strSql = strSql + " 	left join db_order.dbo.tbl_ipkum_list i "
		strSql = strSql + " 	on "
		strSql = strSql + " 	i.inoutidx = L.inoutidx "

		strSql = strSql + " where 1 = 1 "

		if (FRectInOutIDX <> "") then
			strSql = strSql + " and L.inoutidx='" & FRectInOutIDX & "'"
		end if

		if Ctenbank<> "" then
		strSql = strSql + " and L.acct_no='" & Ctenbank & "'"
		end if

		if (FSearchtype01<>"") then
			strSql= strSql + " and L.jeokyo like '%" & ipkumname & "%'"
		end if

		if (Fckdate<>"") then
			strSql = strSql + " and acct_txday >='" + CStr(Replace(FRectRegStart,"-","")) + "'"
			strSql = strSql + " and acct_txday <'" + CStr(Replace(FRectRegEnd,"-","")) + "'"
		end if

		if (FRectInOutGubun<>"") then
			strSql = strSql + " and INOUT_GUBUN = " + CStr(FRectInOutGubun) + " "
		end if

		if (FRectExcluudeMatchFinish<>"") then
			strSql = strSql + " and (IsNull(matchstate, 'N') <> 'Y' and IsNull(i.ipkumstate, 0) <> 7) "
		end if

		if (FRectExcluudeCustomer<>"") then
			strSql = strSql + " and IsNull(A.divcode, 'XX') not in ('00', '30') "
		end if

		if (FRectExcluude10X10<>"") then
			strSql = strSql + " and branch <> '03 0277' "
		end if

		if (FRectTXAmmount<>"") then
			strSql = strSql + " and tx_amt = " + CStr(FRectTXAmmount) + " "
		end if

		if (FRectJeokyo<>"") then
			strSql = strSql + " and L.jeokyo like '%" & html2db(FRectJeokyo) & "%'"
		end if

		if (FRectShowDismatch = "") then
			strSql = strSql + " and IsNull(L.matchstate, 'N') <> 'X' "
		end if

		if (FOrderby="") then
			strSql = strSql + " order by L.acct_txday asc"
		else
			strSql = strSql + " order by L.acct_txday desc"  '최근일순
		end if

		''response.write strSql &"<br>"
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve Fipkumitem(FResultCount)
		i=0

		if not rsget.Eof then

		rsget.absolutepage = FCurrPage

		Do until rsget.eof
			set Fipkumitem(i) = new IpkumListNetteller

				Fipkumitem(i).Fbkno				= rsget("acct_txday_seq")
				Fipkumitem(i).Fbkacctno         = rsget("acct_no")
				Fipkumitem(i).Fbkname           = rsget("acct_nm")
				Fipkumitem(i).Fbkdate           = rsget("acct_txday")
				Fipkumitem(i).Fbkjukyo          = rsget("jeokyo")
				Fipkumitem(i).Fbkcontent        = rsget("branch")
				Fipkumitem(i).Fbketc            = rsget("bigo")
				Fipkumitem(i).Fbkinput          = rsget("tx_amt")
				Fipkumitem(i).finout_gubun		= rsget("INOUT_GUBUN")
				'Fipkumitem(i).Fbkoutput         = rsget("bkoutput")
				'Fipkumitem(i).Fbkbalance        = rsget("bkbalance")
				Fipkumitem(i).Fbkjango          = rsget("tx_cur_bal")
				Fipkumitem(i).Fbkxferdatetime   = rsget("erp_datetime")
				Fipkumitem(i).Ftlno             = rsget("curr_status")
				'Fipkumitem(i).Fbkmemo1          = rsget("bkmemo1")
				'Fipkumitem(i).Fbkmemo2          = rsget("bkmemo2")
				'Fipkumitem(i).Fbkmemo3          = rsget("bkmemo3")
				'Fipkumitem(i).Fbkmemo4          = rsget("bkmemo4")
				'Fipkumitem(i).Fbktag1           = rsget("bktag1")
				'Fipkumitem(i).Fbktag2           = rsget("bktag2")
				'Fipkumitem(i).Fbktag3           = rsget("bktag3")

				Fipkumitem(i).Fjungsanidx		= rsget("jungsanidx")
				Fipkumitem(i).Fjungsancnt		= rsget("jungsancnt")

				Fipkumitem(i).Finoutidx			= rsget("inoutidx")
				Fipkumitem(i).Ftotmatchprice	= rsget("totmatchprice")

				Fipkumitem(i).Forderserial		= rsget("orderserial")
				Fipkumitem(i).Fmatchstate		= rsget("matchstate")
				Fipkumitem(i).Fmatchmemo		= rsget("matchmemo")

			rsget.movenext
			i=i+1
		loop
		end if
		rsget.close
	end Sub

	'/admin2009scm/admin/offshop/pop_ipkum_search.asp
	Public Sub GetipkumlistAccountsNew()
		dim strSql, addSql, i

		'// ===================================================================

		addSql = " from "
		addSql = addSql + " 	[db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT L "
		addSql = addSql + " 	LEFT JOIN [db_order].[dbo].tbl_bank_div A "
		addSql = addSql + " 	on "
		addSql = addSql + " 		L.acct_no=A.accountno "
		addSql = addSql + " 	left join db_order.dbo.tbl_ipkum_list i "
		addSql = addSql + " 	on "
		addSql = addSql + " 	i.inoutidx = L.inoutidx and i.ipkumstate = 7 "
		addSql = addSql + " where 1 = 1 "

		if (FRectJeokyo<>"") then
			addSql = addSql + " and L.jeokyo like '%" & html2db(FRectJeokyo) & "%'"
		end if

		if (FRectTXDayStart<>"") then
			addSql = addSql + " and acct_txday >= '" + CStr(Replace(FRectTXDayStart,"-","")) + "'"
		end if

		if (FRectTXDayEnd<>"") then
			addSql = addSql + " and acct_txday < '" + CStr(Replace(FRectTXDayEnd,"-","")) + "'"
		end if

		if (FRectTXAmmount<>"") then
			addSql = addSql + " and tx_amt = " + CStr(FRectTXAmmount) + " "
			addSql = addSql + " and tx_amt <> 0 "
		end if

		if (FRectInOutGubun<>"") then
			addSql = addSql + " and INOUT_GUBUN = " + CStr(FRectInOutGubun) + " "
		end if

		if (FRectExcluudeMatchFinish<>"") then
			addSql = addSql + " and (IsNull(matchstate, 'N') <> 'Y' and IsNull(i.ipkumstate, 0) <> 7) "
		end if

		if (FRectAcctNo<>"") then
			addSql = addSql + " and acct_no = '" + CStr(FRectAcctNo) + "' "
		end if

		'// ===================================================================
		strSql = "select Count(site_no) as cnt "
		strSql = strSql + addSql
		rsget.open strSql,dbget,1

		FtotalCount = rsget("cnt")
		rsget.close


		'// ===================================================================
		strSql = "select top " & CStr(FCurrpage*FPagesize)
		strSql = strSql + " L.* "
		strSql = strSql + " , ( "
		strSql = strSql + " 	select "
		strSql = strSql + " 		IsNull(sum(matchprice), 0) as totmatchprice "
		strSql = strSql + " 	from "
		strSql = strSql + " 		db_jungsan.dbo.tbl_ipkum_match_detail d "
		strSql = strSql + " 	where "
		strSql = strSql + " 		d.ipkummethod = 'BNK' and d.ipkumidx = L.inoutidx and d.useyn = 'Y' "
		strSql = strSql + " ) as totmatchedprice "

		strSql = strSql + addSql

		if (FOrderby="") then
			strSql = strSql + " order by L.acct_txday asc"
		else
			strSql = strSql + " order by L.acct_txday desc"  '최근일순
		end if

		'response.write strSql &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve Fipkumitem(FResultCount)
		i=0

		if not rsget.Eof then

		rsget.absolutepage = FCurrPage

		Do until rsget.eof
			set Fipkumitem(i) = new IpkumListNetteller

				Fipkumitem(i).Finoutidx			= rsget("inoutidx")
				Fipkumitem(i).Fmatchstate		= rsget("matchstate")

				if IsNull(rsget("totmatchedprice")) then
					Fipkumitem(i).Ftotmatchedprice	= 0
				else
					Fipkumitem(i).Ftotmatchedprice	= rsget("totmatchedprice")
				end if

				Fipkumitem(i).Fbkno				= rsget("acct_txday_seq")
				Fipkumitem(i).Fbkacctno         = rsget("acct_no")
				Fipkumitem(i).Fbkname           = rsget("acct_nm")
				Fipkumitem(i).Fbkdate           = rsget("acct_txday")
				Fipkumitem(i).Fbkjukyo          = rsget("jeokyo")
				Fipkumitem(i).Fbkcontent        = rsget("branch")
				Fipkumitem(i).Fbketc            = rsget("bigo")

				Fipkumitem(i).Fbkinput          = rsget("tx_amt")

				Fipkumitem(i).finout_gubun		= rsget("INOUT_GUBUN")
				'Fipkumitem(i).Fbkoutput         = rsget("bkoutput")
				'Fipkumitem(i).Fbkbalance        = rsget("bkbalance")
				Fipkumitem(i).Fbkjango          = rsget("tx_cur_bal")
				Fipkumitem(i).Fbkxferdatetime   = rsget("erp_datetime")
				Fipkumitem(i).Ftlno             = rsget("curr_status")
				'Fipkumitem(i).Fbkmemo1          = rsget("bkmemo1")
				'Fipkumitem(i).Fbkmemo2          = rsget("bkmemo2")
				'Fipkumitem(i).Fbkmemo3          = rsget("bkmemo3")
				'Fipkumitem(i).Fbkmemo4          = rsget("bkmemo4")
				'Fipkumitem(i).Fbktag1           = rsget("bktag1")
				'Fipkumitem(i).Fbktag2           = rsget("bktag2")
				'Fipkumitem(i).Fbktag3           = rsget("bktag3")

			rsget.movenext
			i=i+1
		loop

		end if
        rsget.close
	end Sub

	'/admin2009scm/admin/offshop/pop_ipkum_list.asp
	Public Sub GetMatchedIpkumlistAccounts()
		dim strSql, addSql, i

		'// ===================================================================
		addSql = " from "
		addSql = addSql + " 	db_jungsan.dbo.tbl_ipkum_match_master m "
		addSql = addSql + " 	JOIN db_jungsan.dbo.tbl_ipkum_match_detail d "
		addSql = addSql + " 	on "
		addSql = addSql + " 		m.idx = d.masteridx "
		addSql = addSql + " 	JOIN [db_log].[dbo].tbl_IBK_ISS_ACCT_INOUT L "
		addSql = addSql + " 	on "
		addSql = addSql + " 		d.ipkummethod = 'BNK' and d.ipkumidx = L.inoutidx and d.useyn = 'Y' "
		addSql = addSql + " 	LEFT JOIN [db_order].[dbo].tbl_bank_div A "
		addSql = addSql + " 	on "
		addSql = addSql + " 		L.acct_no=A.accountno "
		addSql = addSql + " where 1 = 1 "

		addSql = addSql + " and m.jungsanidx = " + CStr(FRectJungsanIDX) + " "

		'// ===================================================================
		strSql = "select Count(site_no) as cnt "
		strSql = strSql + addSql
		rsget.open strSql,dbget,1

		FtotalCount = rsget("cnt")
		rsget.close


		'// ===================================================================
		strSql = "select top " & CStr(FCurrpage*FPagesize)
		strSql = strSql + " L.*, d.matchprice, d.idx as matchdetailidx "
		strSql = strSql + addSql

		if (FOrderby="") then
			strSql = strSql + " order by L.acct_txday asc"
		else
			strSql = strSql + " order by L.acct_txday desc"  '최근일순
		end if

		'response.write strSql &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.Open strSql,dbget,1

		FResultCount = rsget.RecordCount

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

		redim preserve Fipkumitem(FResultCount)
		i=0

		if not rsget.Eof then

		rsget.absolutepage = FCurrPage

		Do until rsget.eof
			set Fipkumitem(i) = new IpkumListNetteller

				Fipkumitem(i).Finoutidx			= rsget("inoutidx")
				Fipkumitem(i).Fmatchstate		= rsget("matchstate")
				Fipkumitem(i).Fmatchprice		= rsget("matchprice")
				Fipkumitem(i).Fmatchdetailidx	= rsget("matchdetailidx")

				Fipkumitem(i).Fbkno				= rsget("acct_txday_seq")
				Fipkumitem(i).Fbkacctno         = rsget("acct_no")
				Fipkumitem(i).Fbkname           = rsget("acct_nm")
				Fipkumitem(i).Fbkdate           = rsget("acct_txday")
				Fipkumitem(i).Fbkjukyo          = rsget("jeokyo")
				Fipkumitem(i).Fbkcontent        = rsget("branch")
				Fipkumitem(i).Fbketc            = rsget("bigo")
				Fipkumitem(i).Fbkinput          = rsget("tx_amt")
				Fipkumitem(i).finout_gubun		= rsget("INOUT_GUBUN")
				'Fipkumitem(i).Fbkoutput         = rsget("bkoutput")
				'Fipkumitem(i).Fbkbalance        = rsget("bkbalance")
				Fipkumitem(i).Fbkjango          = rsget("tx_cur_bal")
				Fipkumitem(i).Fbkxferdatetime   = rsget("erp_datetime")
				Fipkumitem(i).Ftlno             = rsget("curr_status")
				'Fipkumitem(i).Fbkmemo1          = rsget("bkmemo1")
				'Fipkumitem(i).Fbkmemo2          = rsget("bkmemo2")
				'Fipkumitem(i).Fbkmemo3          = rsget("bkmemo3")
				'Fipkumitem(i).Fbkmemo4          = rsget("bkmemo4")
				'Fipkumitem(i).Fbktag1           = rsget("bktag1")
				'Fipkumitem(i).Fbktag2           = rsget("bktag2")
				'Fipkumitem(i).Fbktag3           = rsget("bktag3")

			rsget.movenext
			i=i+1
		loop

		end if
        rsget.close
	end Sub


	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	public Function StarScrollPage()
		StarScrollPage = ((FCurrPage-1)\FScrollCount)*FScrollCount +1
	end Function
end class
%>
