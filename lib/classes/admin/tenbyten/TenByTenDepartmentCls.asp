<%

Class CTenByTenDepartmentItem
	public Fcid1
	public FdepartmentName1
	public FdispOrderNo1
	public FuseYN1
	public Fregdate1
	public Flastupdate1

	public Fcid2
	public FdepartmentName2
	public FdispOrderNo2
	public FuseYN2
	public Fregdate2
	public Flastupdate2

	public Fcid3
	public FdepartmentName3
	public FdispOrderNo3
	public FuseYN3
	public Fregdate3
	public Flastupdate3

	public Fcid4
	public FdepartmentName4
	public FdispOrderNo4
	public FuseYN4
	public Fregdate4
	public Flastupdate4

	public Fcid5
	public FdepartmentName5
	public FdispOrderNo5
	public FuseYN5
	public Fregdate5
	public Flastupdate5

	public Fcid6
	public FdepartmentName6
	public FdispOrderNo6
	public FuseYN6
	public Fregdate6
	public Flastupdate6

	public Fcid
	public FdepartmentName
	public FdepartmentNameFull
	public FdispOrderNo
	public FuseYN
	public Fregdate
	public Flastupdate

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class


Class CTenByTenDepartment
	public FOneItem
	public FItemList()

	public FPageSize
	public FTotalPage
    public FPageCount
	public FTotalCount
	public FResultCount
    public FScrollCount
	public FCurrPage
	public FMaxPage

	public FRectCID

	public FRectUseYN
	public FRectSearchKey
	public FRectSearchString

	Private Sub Class_Initialize()
		redim  FitemList(0)

		FCurrPage 		= 1
		FPageSize 		= 20
		FResultCount 	= 0
		FScrollCount 	= 10
		FTotalCount 	= 0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	'##### 직책 목록 #####
	public Sub GetList()
		dim SQL, AddSQL, i, strTemp

		AddSQL = AddSQL & " from "
		AddSQL = AddSQL & " 	db_partner.dbo.vw_user_department T "
		AddSQL = AddSQL & " where "
		AddSQL = AddSQL & " 	1 = 1 "

		if FRectUseYN = "Y" then
			AddSQL = AddSQL & " and IsNull(useYN1, 'Y') = 'Y' "
			AddSQL = AddSQL & " and IsNull(useYN2, 'Y') = 'Y' "
			AddSQL = AddSQL & " and IsNull(useYN3, 'Y') = 'Y' "
			AddSQL = AddSQL & " and IsNull(useYN4, 'Y') = 'Y' "
			AddSQL = AddSQL & " and IsNull(useYN5, 'Y') = 'Y' "
			AddSQL = AddSQL & " and IsNull(useYN6, 'Y') = 'Y' "
		end if

		if (FRectCID <> "") then
			AddSQL = AddSQL & " and (IsNull(cid1, -1) = " + CStr(FRectCID) + " or IsNull(cid2, -1) = " + CStr(FRectCID) + " or IsNull(cid3, -1) = " + CStr(FRectCID) + " or IsNull(cid4, -1) = " + CStr(FRectCID) + ") "
		end if


		'// 개수 파악 //
		SQL =	"Select count(cid1), CEILING(CAST(Count(cid1) AS FLOAT)/" & FPageSize & ") "
		SQL = SQL & AddSQL

		rsget.Open SQL,dbget,1
			FTotalCount = rsget(0)
			FtotalPage = rsget(1)
		rsget.Close

		'// 목록 접수 //
		SQL =	"Select top " & CStr(FPageSize*FCurrPage) & " T.* "

		SQL = SQL & AddSQL

		SQL = SQL & " order by "
		SQL = SQL & " 	dispOrderNo1, dispOrderNo2, dispOrderNo3, dispOrderNo4, dispOrderNo5, dispOrderNo6 " 
		rsget.pagesize = FPageSize
		rsget.Open SQL,dbget,1
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CTenByTenDepartmentItem

				FItemList(i).Fcid1				= rsget("cid1")
				FItemList(i).FdepartmentName1	= rsget("departmentName1")
				FItemList(i).FdispOrderNo1		= rsget("dispOrderNo1")
				FItemList(i).FuseYN1			= rsget("useYN1")
				FItemList(i).Fregdate1			= rsget("regdate1")
				FItemList(i).Flastupdate1		= rsget("lastupdate1")
				if IsNull(FItemList(i).FuseYN1) then
					FItemList(i).FuseYN1 = "Y"
				end if

				FItemList(i).Fcid2				= rsget("cid2")
				FItemList(i).FdepartmentName2	= rsget("departmentName2")
				FItemList(i).FdispOrderNo2		= rsget("dispOrderNo2")
				FItemList(i).FuseYN2			= rsget("useYN2")
				FItemList(i).Fregdate2			= rsget("regdate2")
				FItemList(i).Flastupdate2		= rsget("lastupdate2")
				if IsNull(FItemList(i).FuseYN2) then
					FItemList(i).FuseYN2 = "Y"
				end if

				FItemList(i).Fcid3				= rsget("cid3")
				FItemList(i).FdepartmentName3	= rsget("departmentName3")
				FItemList(i).FdispOrderNo3		= rsget("dispOrderNo3")
				FItemList(i).FuseYN3			= rsget("useYN3")
				FItemList(i).Fregdate3			= rsget("regdate3")
				FItemList(i).Flastupdate3		= rsget("lastupdate3")
				if IsNull(FItemList(i).FuseYN3) then
					FItemList(i).FuseYN3 = "Y"
				end if

				FItemList(i).Fcid4				= rsget("cid4")
				FItemList(i).FdepartmentName4	= rsget("departmentName4")
				FItemList(i).FdispOrderNo4		= rsget("dispOrderNo4")
				FItemList(i).FuseYN4			= rsget("useYN4")
				FItemList(i).Fregdate4			= rsget("regdate4")
				FItemList(i).Flastupdate4		= rsget("lastupdate4")
				if IsNull(FItemList(i).FuseYN4) then
					FItemList(i).FuseYN4 = "Y"
				end if

				FItemList(i).Fcid5				= rsget("cid5")
				FItemList(i).FdepartmentName5	= rsget("departmentName5")
				FItemList(i).FdispOrderNo5		= rsget("dispOrderNo5")
				FItemList(i).FuseYN5			= rsget("useYN5")
				FItemList(i).Fregdate5			= rsget("regdate5")
				FItemList(i).Flastupdate5		= rsget("lastupdate5")
				if IsNull(FItemList(i).FuseYN5) then
					FItemList(i).FuseYN5 = "Y"
				end if

				FItemList(i).Fcid6				= rsget("cid6")
				FItemList(i).FdepartmentName6	= rsget("departmentName6")
				FItemList(i).FdispOrderNo6		= rsget("dispOrderNo6")
				FItemList(i).FuseYN6			= rsget("useYN6")
				FItemList(i).Fregdate6			= rsget("regdate6")
				FItemList(i).Flastupdate6		= rsget("lastupdate6")
				if IsNull(FItemList(i).FuseYN6) then
					FItemList(i).FuseYN6 = "Y"
				end if

				FItemList(i).Fcid				= rsget("cid")
				FItemList(i).FdepartmentName	= rsget("departmentName")
				FItemList(i).FdepartmentNameFull	= rsget("departmentNameFull")
				FItemList(i).FdispOrderNo		= rsget("dispOrderNo")
				FItemList(i).FuseYN				= rsget("useYN")
				FItemList(i).Fregdate			= rsget("regdate")
				FItemList(i).Flastupdate		= rsget("lastupdate")

				rsget.moveNext
				i=i+1
			loop
		end if

		rsget.Close

	end Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function


	public Sub GetInfo()
		dim SQL

		set FOneItem = new CTenByTenDepartmentItem

		SQL =	"Select * " &_
				"From db_partner.dbo.tbl_user_department " &_
				"Where cid = " & FRectCID
		rsget.Open SQL,dbget,1

		if Not(rsget.EOF or rsget.BOF) then
			FResultCount = 1

			FOneItem.Fcid				= rsget("cid")
			FOneItem.FdepartmentName	= db2html(rsget("departmentName"))
			FOneItem.FdispOrderNo		= rsget("dispOrderNo")
			FOneItem.FuseYN				= rsget("useYN")
			FOneItem.Fregdate			= rsget("regdate")
			FOneItem.Flastupdate		= rsget("lastupdate")
		else
			FResultCount = 0
		end if

		rsget.Close

	end Sub
end Class

public function drawSelectBoxDepartment(cidfrmname, cidval)
	dim i
	dim strOutput, strSelected
	dim oCTenByTenDepartment
	set oCTenByTenDepartment = new CTenByTenDepartment

	oCTenByTenDepartment.FPageSize = 500
	oCTenByTenDepartment.FCurrPage = 1
	oCTenByTenDepartment.FRectUseYN = "Y"

	oCTenByTenDepartment.GetList

	strOutput = "<select class='select' name='" + CStr(cidfrmname) + "'>"
	strOutput = strOutput + "<option></option>"
	if oCTenByTenDepartment.FResultCount > 0 then
		for i = 0 to oCTenByTenDepartment.FResultcount - 1
			strSelected = ""

			if (cidval <> "") then
				if (CLng(cidval) = oCTenByTenDepartment.FItemList(i).Fcid) then
					strSelected = "selected"
				end if
			end if

			strOutput = strOutput + "<option value='" + CStr(oCTenByTenDepartment.FItemList(i).Fcid) + "' " + CStr(strSelected)  + ">" + CStr(oCTenByTenDepartment.FItemList(i).FdepartmentNameFull) + "</option>"
		next
	end if
	strOutput = strOutput + "</select>"

	drawSelectBoxDepartment = strOutput
end function

public function drawChSelectBoxDepartment(cidfrmname, cidval,strText)
	dim i
	dim strOutput, strSelected
	dim oCTenByTenDepartment
	set oCTenByTenDepartment = new CTenByTenDepartment

	oCTenByTenDepartment.FPageSize = 500
	oCTenByTenDepartment.FCurrPage = 1
	oCTenByTenDepartment.FRectUseYN = "Y"

	oCTenByTenDepartment.GetList
 
	strOutput = "<select class='select' name='" + CStr(cidfrmname) + "' " +strText+" >"
	strOutput = strOutput + "<option></option>"
	if oCTenByTenDepartment.FResultCount > 0 then
		for i = 0 to oCTenByTenDepartment.FResultcount - 1
			strSelected = ""

			if (cidval <> "") then
				if (CLng(cidval) = oCTenByTenDepartment.FItemList(i).Fcid) then
					strSelected = "selected"
				end if
			end if

			strOutput = strOutput + "<option value='" + CStr(oCTenByTenDepartment.FItemList(i).Fcid) + "' " + CStr(strSelected)  + ">" + CStr(oCTenByTenDepartment.FItemList(i).FdepartmentNameFull) + "</option>"
		next
	end if
	strOutput = strOutput + "</select>"

	drawChSelectBoxDepartment = strOutput
end function

'/admin/dataanalysis/salesissue/salesissue.asp	'/admin/dataanalysis/salesissue/salesissue_edit.asp
public function getDepartmentALL(cidval)
	dim i
	dim strOutput, strSelected, strNotUse
	dim oCTenByTenDepartment
	
	if cidval="" then exit function
	
	set oCTenByTenDepartment = new CTenByTenDepartment

	oCTenByTenDepartment.FPageSize = 1000
	oCTenByTenDepartment.FCurrPage = 1
	oCTenByTenDepartment.FRectUseYN = "Y"

	oCTenByTenDepartment.GetList

	if oCTenByTenDepartment.FResultCount > 0 then
		for i = 0 to oCTenByTenDepartment.FResultcount - 1
			if (oCTenByTenDepartment.FItemList(i).FuseYN = "N") then
				strNotUse = "XXX "
			end if

			if (cidval <> "") then
				if (CStr(cidval) = CStr(oCTenByTenDepartment.FItemList(i).Fcid)) then
					strOutput = CStr(strNotUse) + CStr(oCTenByTenDepartment.FItemList(i).FdepartmentNameFull)
					exit for
				end if
			end if
		next
	end if

	getDepartmentALL = strOutput
end function

'/admin/dataanalysis/salesissue/salesissue.asp	'/admin/dataanalysis/salesissue/salesissue_edit.asp
public function drawSelectBoxDepartmentALL(cidfrmname, cidval)
	dim i
	dim strOutput, strSelected, strNotUse
	dim oCTenByTenDepartment
	set oCTenByTenDepartment = new CTenByTenDepartment

	oCTenByTenDepartment.FPageSize = 1000
	oCTenByTenDepartment.FCurrPage = 1
	oCTenByTenDepartment.FRectUseYN = "Y"

	oCTenByTenDepartment.GetList

	strOutput = "<select class='select' name='" + CStr(cidfrmname) + "'>"
	strOutput = strOutput + "<option></option>"
	if oCTenByTenDepartment.FResultCount > 0 then
		for i = 0 to oCTenByTenDepartment.FResultcount - 1
			strSelected = ""
			strNotUse = ""

			if (cidval <> "") then
				if (CStr(cidval) = CStr(oCTenByTenDepartment.FItemList(i).Fcid)) then
					strSelected = "selected"
				end if
			end if

			if (oCTenByTenDepartment.FItemList(i).FuseYN = "N") then
				strNotUse = "XXX "
			end if

			strOutput = strOutput + "<option value='" + CStr(oCTenByTenDepartment.FItemList(i).Fcid) + "' " + CStr(strSelected)  + ">" + CStr(strNotUse) + CStr(oCTenByTenDepartment.FItemList(i).FdepartmentNameFull) + "</option>"
		next
	end if
	strOutput = strOutput + "</select>"

	drawSelectBoxDepartmentALL = strOutput
end function

public function drawSelectBoxMyDepartment(userid, cidfrmname, cidval)
	dim i
	dim strOutput, strSelected
	dim strSql

	strSql = " select top 500 isNull(dv.cid,0) as cid, dv.departmentNameFull "
	strSql = strSql + " from "
	strSql = strSql + " 	(select department_id "
 	strSql = strSql + " from db_partner.dbo.tbl_user_tenbyten  "
	strSql = strSql + "  where userid = '" + CStr(userid) + "'  and department_id is not NULL "
 	strSql = strSql + " union all "
	strSql = strSql + "  select departmentid as department_id "
	strSql = strSql + " from db_partner.dbo.tbl_partner_addDepartment  "
	strSql = strSql + "  where userid ='" + CStr(userid) + "'  and departmentid is not NULL and isusing=1 "
	strSql = strSql + "  ) as t "
	strSql = strSql + " 	join db_partner.dbo.vw_user_department dv "
	strSql = strSql + " 	on "
	strSql = strSql + " 		1 = 1 "
	strSql = strSql + " 		and ( "
	strSql = strSql + " 			dv.cid1 = t.department_id or dv.cid2 = t.department_id or dv.cid3 = t.department_id or dv.cid4 = t.department_id or dv.cid5 = t.department_id or dv.cid6 = t.department_id "
	strSql = strSql + " 		) "
	strSql = strSql + " 		and dv.useYN = 'Y' "
	strSql = strSql + " order by "
	strSql = strSql + " 	dispOrderNo1, dispOrderNo2, dispOrderNo3, dispOrderNo4, dispOrderNo5, dispOrderNo6 " 
	 
	rsget.Open strSql, dbget, 1

	strOutput = "<select class='select' name='" + CStr(cidfrmname) + "'>"
	''strOutput = strOutput + "<option></option>"

 
	i=0
	if  not rsget.EOF  then
		do until rsget.eof
			strSelected = ""

			if (cidval <> "") then
				if (CLng(cidval) = rsget("cid")) then
					strSelected = "selected"
				end if
			end if

			strOutput = strOutput + "<option value='" + CStr(rsget("cid")) + "' " + CStr(strSelected)  + ">" + CStr(db2html(rsget("departmentNameFull"))) + "</option>"
			rsget.moveNext
			i = i + 1
		loop
	end if
	rsget.Close

	strOutput = strOutput + "</select>"

	drawSelectBoxMyDepartment = strOutput
end function

function GetUserDepartmentID(empno , userid)
	dim sqlStr ,sqlsearch

	GetUserDepartmentID = -1

	if empno = "" and userid = "" then exit function

	if empno <> "" then
		sqlsearch = sqlsearch & " and t.empno = '"&empno&"'"
	end if
	if userid <> "" then
		sqlsearch = sqlsearch & " and p.id = '"&userid&"'"
	end if

	sqlStr = "select top 1 "
	sqlStr = sqlStr & " t.department_id "
	sqlStr = sqlStr & " from db_partner.dbo.tbl_user_tenbyten t"
	sqlStr = sqlStr & " left join db_partner.dbo.tbl_partner p"
	sqlStr = sqlStr & " 	on t.userid = p.id"
	sqlStr = sqlStr & " 	and p.isusing = 'Y'"
	sqlStr = sqlStr & " where" & vbcrlf

	' 퇴사예정자 처리	' 2018.10.16 한용민
	sqlStr = sqlStr & "	(t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0)) " & sqlsearch

	'response.write sqlStr &"<br>"
	rsget.Open sqlStr,dbget,1
	if not rsget.EOF  then
		GetUserDepartmentID = rsget("department_id")
	end if
	rsget.close
end function

%>
