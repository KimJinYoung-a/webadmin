<%
'####################################################
' Description : 제휴몰 수수료정산 클래스
' History : 2017.04.06 한용민 생성
'####################################################

Class Ccommission_item
	public frDate
	public ffixedDate
	public forderserial
	public fitemname
	public fitemno
	public fsuppPrc
	public fcommpro
	public fcommissoin
	public fordStatName
	public fcancelDT
	public fdevice
	public frdsite
	public fisCharge
	public fexplain

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class

Class Ccommission
	public FItemList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FPageCount

	public FRectyyyymm
	public frectorderserial
	public frectitemname
	public farrlist
	public frectrdsite
	public frectismobile

	'//admin/jungsan/commission/commissionjungsan_between.asp
	public function Getcommissionjungsan_between_notpaging
		dim i , sql , sqlsearch

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql & " jd.rDate, jd.fixedDate, jd.orderserial, replace(replace(replace(replace(jd.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName"
		sql = sql & " , jd.itemno, jd.suppPrc, jd.commpro, jd.commissoin, jd.ordStatName , jd.cancelDT"
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)"
		sql = sql & " left join db_order.dbo.tbl_order_master m with (nolock)"
		sql = sql & " 	on jd.orderserial = m.orderserial"
		sql = sql & " where jd.rdsite='betweenshop' " & sqlsearch
		sql = sql & " order by jd.rDate asc, jd.orderserial asc"

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		i = 0
		If Not rsget.Eof Then
			farrlist			= rsget.getrows()
		End If

		rsget.close
	end function

	'//admin/jungsan/commission/commissionjungsan_between.asp
	public sub Getcommissionjungsan_between_paging()
		dim sql,i, sqlsearch

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if

		sql = "select count(jd.rDate) as cnt" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)"
		sql = sql & " where jd.rdsite='betweenshop' " & sqlsearch

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sql = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sql = sql & " jd.rDate, jd.fixedDate, jd.orderserial, replace(replace(replace(replace(jd.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName"
		sql = sql & " , jd.itemno, jd.suppPrc, jd.commpro, jd.commissoin, jd.ordStatName , jd.cancelDT"
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)"
		sql = sql & " where jd.rdsite='betweenshop' " & sqlsearch
		sql = sql & " order by jd.rDate asc, jd.orderserial asc"

		'response.write sql &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Ccommission_item
				
				FItemList(i).frDate = rsget("rDate")
				FItemList(i).ffixedDate = rsget("fixedDate")						
				FItemList(i).forderserial = rsget("orderserial")	
				FItemList(i).fitemname = db2html(rsget("itemNOptionName"))
				FItemList(i).fitemno = rsget("itemno")
				FItemList(i).fsuppPrc = rsget("suppPrc")									
				FItemList(i).fcommpro = rsget("commpro")							
				FItemList(i).fcommissoin = rsget("commissoin")
				FItemList(i).fordStatName = rsget("ordStatName")							
				FItemList(i).fcancelDT = rsget("cancelDT")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/jungsan/commission/commissionjungsan_daum.asp	'//admin/jungsan/commission/commissionjungsan_nate.asp
	public function Getcommissionjungsan_daum_notpaging
		dim i , sql , sqlsearch

		if frectrdsite="" then exit function

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and D.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and D.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and D.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if frectrdsite<>"" then
			sqlsearch = sqlsearch & " and R.gubun='" & frectrdsite & "'"
		end if

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql & " D.rDate,D.fixedDate,D.rdsite"
		sql = sql & " ,(CASE WHEN D.ismobile=1 THEN'모바일' ELSE'웹' END) as device"
		sql = sql & " ,D.orderserial"
		sql = sql & " , replace(replace(replace(replace(D.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName"
		sql = sql & " ,D.itemno,D.suppPrc,D.commpro,D.commissoin,D.ordStatName,D.cancelDT"
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail D with (nolock)"
		sql = sql & " Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)"
		sql = sql & " 	on D.rdsite=R.rdsite"
		sql = sql & " where 1=1  " & sqlsearch	'D.ismobile=0
		sql = sql & " order by D.rDate,D.orderserial"

		'response.write sql & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		i = 0
		If Not rsget.Eof Then
			farrlist			= rsget.getrows()
		End If

		rsget.close
	end function

	'//admin/jungsan/commission/commissionjungsan_daum.asp	'//admin/jungsan/commission/commissionjungsan_nate.asp
	public sub Getcommissionjungsan_daum_paging()
		dim sql,i, sqlsearch

		if frectrdsite="" then exit sub

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and D.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and D.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and D.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if frectrdsite<>"" then
			sqlsearch = sqlsearch & " and R.gubun='" & frectrdsite & "'"
		end if

		sql = "select count(D.rDate) as cnt" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail D with (nolock)"
		sql = sql & " Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)"
		sql = sql & " 	on D.rdsite=R.rdsite"
		sql = sql & " where 1=1  " & sqlsearch	'D.ismobile=0

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql & " D.rDate,D.fixedDate,D.rdsite"
		sql = sql & " ,(CASE WHEN D.ismobile=1 THEN'모바일' ELSE'웹' END) as device"
		sql = sql & " ,D.orderserial"
		sql = sql & " , replace(replace(replace(replace(D.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName"
		sql = sql & " ,D.itemno,D.suppPrc,D.commpro,D.commissoin,D.ordStatName,D.cancelDT"
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail D with (nolock)"
		sql = sql & " Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)"
		sql = sql & " 	on D.rdsite=R.rdsite"
		sql = sql & " where 1=1  " & sqlsearch	'D.ismobile=0
		sql = sql & " order by D.rDate,D.orderserial"

		'response.write sql &"<br>"
		'response.END
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Ccommission_item

				FItemList(i).frDate = rsget("rDate")
				FItemList(i).ffixedDate = rsget("fixedDate")						
				FItemList(i).frdsite = rsget("rdsite")
				FItemList(i).fdevice = rsget("device")
				FItemList(i).forderserial = rsget("orderserial")
				FItemList(i).fitemname = db2html(rsget("itemNOptionName"))
				FItemList(i).fitemno = rsget("itemno")
				FItemList(i).fsuppPrc = rsget("suppPrc")									
				FItemList(i).fcommpro = rsget("commpro")							
				FItemList(i).fcommissoin = rsget("commissoin")
				FItemList(i).fordStatName = rsget("ordStatName")							
				FItemList(i).fcancelDT = rsget("cancelDT")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/jungsan/commission/commissionjungsan_naver.asp
	public function Getcommissionjungsan_naver_notpaging
		dim i , sql , sqlsearch

		if FRectismobile="" then exit function

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if FRectismobile<>"" then
			sqlsearch = sqlsearch & " and jd.ismobile=" & CStr(FRectismobile) & ""
		end if

		sql = "SELECT TOP " & Cstr(FPageSize * FCurrPage)
		sql = sql & " jd.rDate,jd.fixedDate,jd.orderserial" + vbcrlf
		sql = sql & " , replace(replace(replace(replace(jd.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName" + vbcrlf
		sql = sql & " ,jd.itemno,jd.suppPrc,jd.commpro,jd.commissoin,jd.ordStatName ,jd.cancelDT" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
		sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
		sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
		sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
		sql = sql & " where 1=1 " & sqlsearch
		sql = sql & " order by jd.rDate,jd.orderserial" + vbcrlf

		'response.write sql & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		i = 0
		If Not rsget.Eof Then
			farrlist			= rsget.getrows()
		End If

		rsget.close
	end function

	'//admin/jungsan/commission/commissionjungsan_naver.asp
	public sub Getcommissionjungsan_naver_paging()
		dim sql,i, sqlsearch

		if FRectismobile="" then exit sub

		if frectorderserial<>"" then
			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
		end if
		if frectitemname<>"" then
			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if FRectismobile<>"" then
			sqlsearch = sqlsearch & " and jd.ismobile=" & CStr(FRectismobile) & ""
		end if

		sql = "select count(jd.rDate) as cnt" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
		sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
		sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
		sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
		sql = sql & " where 1=1 " & sqlsearch

		'response.write sql &"<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
		rsget.Close

		if FTotalCount < 1 then exit sub

		sql = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sql = sql & " jd.rDate,jd.fixedDate,jd.orderserial" + vbcrlf
		sql = sql & " , replace(replace(replace(replace(jd.itemNOptionName,char(9),''),char(10),''),char(13),''),'""','''') as itemNOptionName" + vbcrlf
		sql = sql & " ,jd.itemno,jd.suppPrc,jd.commpro,jd.commissoin,jd.ordStatName ,jd.cancelDT" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
		sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
		sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
		sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
		sql = sql & " where 1=1 " & sqlsearch
		sql = sql & " order by jd.rDate,jd.orderserial" + vbcrlf

		'response.write sql &"<br>"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		if (FCurrPage * FPageSize < FTotalCount) then
			FResultCount = FPageSize
		else
			FResultCount = FTotalCount - FPageSize*(FCurrPage-1)
		end if

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Ccommission_item
				
				FItemList(i).frDate = rsget("rDate")
				FItemList(i).ffixedDate = rsget("fixedDate")						
				FItemList(i).forderserial = rsget("orderserial")	
				FItemList(i).fitemname = db2html(rsget("itemNOptionName"))
				FItemList(i).fitemno = rsget("itemno")
				FItemList(i).fsuppPrc = rsget("suppPrc")									
				FItemList(i).fcommpro = rsget("commpro")							
				FItemList(i).fcommissoin = rsget("commissoin")
				FItemList(i).fordStatName = rsget("ordStatName")							
				FItemList(i).fcancelDT = rsget("cancelDT")

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	'//admin/jungsan/commission/commissionjungsan_naver.asp
	public function Getcommissionjungsan_naver_sum_notpaging
		dim i , sql , sqlsearch

		if FRectismobile="" then exit function

'		if frectorderserial<>"" then
'			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
'		end if
'		if frectitemname<>"" then
'			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
'		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if FRectismobile<>"" then
			sqlsearch = sqlsearch & " and jd.ismobile=" & CStr(FRectismobile) & ""
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sql = sql & " jd.rdsite, sum(jd.itemno) as itemno, sum(jd.suppPrc) as suppPrc ,sum(jd.commissoin) as commissoin" + vbcrlf
		sql = sql & " ,R.isCharge,R.explain" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
		sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
		sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
		sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
		sql = sql & " where jd.rdsite<>'betweenshop' " & sqlsearch
		sql = sql & " group by jd.rdsite,R.isCharge,R.explain" + vbcrlf

		'response.write sql & "<Br>"
		'response.end
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FresultCount = rsget.recordcount

		i = 0
		If Not rsget.Eof Then
			farrlist			= rsget.getrows()
		End If

		rsget.close
	end function

	'//admin/jungsan/commission/commissionjungsan_naver.asp
	public sub Getcommissionjungsan_naver_sum()
		dim sql,i, sqlsearch

		if FRectismobile="" then exit sub

'		if frectorderserial<>"" then
'			sqlsearch = sqlsearch & " and jd.orderserial='" & CStr(frectorderserial) & "'"
'		end if
'		if frectitemname<>"" then
'			sqlsearch = sqlsearch & " and jd.itemNOptionName like '%" & CStr(frectitemname) & "%'"
'		end if
		if FRectyyyymm<>"" then
			sqlsearch = sqlsearch & " and jd.jmonth='" & CStr(FRectyyyymm) & "'"
		end if
		if FRectismobile<>"" then
			sqlsearch = sqlsearch & " and jd.ismobile=" & CStr(FRectismobile) & ""
		end if

		sql = "select top " & Cstr(FPageSize * FCurrPage) + vbcrlf
		sql = sql & " jd.rdsite, sum(jd.itemno) as itemno, sum(jd.suppPrc) as suppPrc ,sum(jd.commissoin) as commissoin" + vbcrlf
		sql = sql & " ,R.isCharge,R.explain" + vbcrlf
		sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
		sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
		sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
		sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
		sql = sql & " where jd.rdsite<>'betweenshop' " & sqlsearch
		sql = sql & " group by jd.rdsite,R.isCharge,R.explain" + vbcrlf

		'response.write sql &"<br>"
		'response.end
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

		FTotalCount = rsget.recordcount
		FResultCount = rsget.recordcount

		FTotalPage = (FTotalCount\FPageSize)
		if (FTotalPage<>FTotalCount/FPageSize) then FTotalPage = FTotalPage +1

		redim preserve FItemList(FResultCount)

		FPageCount = FCurrPage - 1

		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.EOF
				set FItemList(i) = new Ccommission_item
				
				FItemList(i).frdsite = rsget("rdsite")
				FItemList(i).fitemno = rsget("itemno")						
				FItemList(i).fsuppPrc = rsget("suppPrc")	
				FItemList(i).fcommissoin = rsget("commissoin")
				FItemList(i).fisCharge = rsget("isCharge")									
				FItemList(i).fexplain = db2html(rsget("explain"))

				rsget.movenext
				i=i+1
			loop
		end if
		rsget.Close
	end sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	Private Sub Class_Initialize()
		redim  FItemList(0)
		redim  FCountList(0)
		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub
	Private Sub Class_Terminate()
	End Sub
end Class
%>