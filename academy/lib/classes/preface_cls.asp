<%
'##### �Ӹ��� ���ڵ�¿� Ŭ���� #####
class CprfItem

	public Fprfid
	public FcommCd
	public FgroupCd
	public FprfDiv
	public FcommNm
	public FgroupNm
	public FprfCont
	public Fisusing
	public Fregdate

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### �Ӹ��� Ŭ���� #####
Class Cprf

	public FprfList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectprfid
	public FRectsearchDiv
	public FRectsearchString

	'// �⺻ ������ ����
	Private Sub Class_Initialize()
		redim preserve FprfList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// �Ӹ��� �з��� ��� ���
	public Sub GetprfList()
		dim SQL, AddSQL, lp

		if FRectsearchString<>"" then
			AddSQL = AddSQL & " and t1.prfCont like '%" & FRectsearchString & "%' "
		end if

		if FRectsearchDiv<>"" then
			AddSQL = AddSQL & " and t3.groupCd='" & FRectsearchDiv & "' "
		end if

		'@ �ѵ����ͼ�
		SQL =	" Select count(prfid) as cnt " &_
				" From db_academy.dbo.tbl_preface as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Join db_academy.dbo.tbl_groupCd as t3 on t1.groupCd=t3.groupCd " &_
				" Where 1=1 " & AddSQL

		rsACADEMYget.Open sql, dbACADEMYget, 1
			FTotalCount = rsACADEMYget("cnt")
		rsACADEMYget.close


		'@ ������
		SQL =	" Select top " & CStr(FPageSize*FCurrPage) &_
				"		prfid, t1.commCd, t2.commNm, t3.groupNm " &_
				"		,Case t1.isusing When 'Y' Then '<font color=darkblue>���</font>' When 'N' Then '<font color=darkred>����</font>' End isusing " &_
				"		,t1.regdate " &_
				" From db_academy.dbo.tbl_preface as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Join db_academy.dbo.tbl_groupCd as t3 on t1.groupCd=t3.groupCd " &_
				" Where 1=1 " & AddSQL &_
				" Order by t1.groupCd, t1.commCd, prfid "

		rsACADEMYget.pagesize = FPageSize
		rsACADEMYget.Open sql, dbACADEMYget, 1

		FtotalPage =  CInt(FTotalCount\FPageSize)
		if  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) then
			FtotalPage = FtotalPage +1
		end if
		FResultCount = rsACADEMYget.RecordCount-(FPageSize*(FCurrPage-1))

		redim FprfList(FResultCount)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

		    lp = 0
			rsACADEMYget.absolutepage = FCurrPage
			do until rsACADEMYget.eof
				set FprfList(lp) = new CprfItem

				FprfList(lp).Fprfid			= rsACADEMYget("prfid")
				FprfList(lp).FcommCd		= rsACADEMYget("commCd")
				FprfList(lp).FcommNm		= rsACADEMYget("commNm")
				FprfList(lp).FgroupNm		= rsACADEMYget("groupNm")
				FprfList(lp).Fisusing		= rsACADEMYget("isusing")
				FprfList(lp).Fregdate		= rsACADEMYget("regdate")

				lp=lp+1
				rsACADEMYget.moveNext
			loop
		end if
		rsACADEMYget.close

	end Sub


	'// prf ���� ����
	public Sub GetprfRead()
		dim SQL

		SQL =	" Select prfid, prfCont, t1.isusing " &_
				"		, t1.commCd, commNm, t1.groupCd, groupNm, t1.regdate " &_
				" From db_academy.dbo.tbl_preface as t1 " &_
				"		Join db_academy.dbo.tbl_commCd as t2 on t1.commCd=t2.commCd " &_
				"		Join db_academy.dbo.tbl_groupCd as t3 on t1.groupCd=t3.groupCd " &_
				" Where prfid = " & FRectprfid

		rsACADEMYget.Open sql, dbACADEMYget, 1

		redim FprfList(0)

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then

			set FprfList(0) = new CprfItem

			FprfList(0).Fprfid		= rsACADEMYget("prfid")
			FprfList(0).FprfCont	= rsACADEMYget("prfCont")
			FprfList(0).FcommCd		= rsACADEMYget("commCd")
			FprfList(0).FcommNm		= rsACADEMYget("commNm")
			FprfList(0).FgroupCd	= rsACADEMYget("groupCd")
			FprfList(0).FgroupNm	= rsACADEMYget("groupNm")
			FprfList(0).Fisusing	= rsACADEMYget("isusing")
			FprfList(0).Fregdate	= rsACADEMYget("regdate")

		end if
		rsACADEMYget.close

	end sub


	'// �����ڵ� �ɼ� ���� //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select commCd, commNm From db_academy.dbo.tbl_commCd Where groupCd in (" & grpCd & ")"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("commCd") & "' "

				if nowCd=rsACADEMYget("commCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("commNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optCommCd = strOpt

	end function


	'// �׷��ڵ� �ɼ� ���� //
	function optGroupCd(nowCd)
		dim SQL, strOpt

		SQL =	"Select groupCd, groupNm From db_academy.dbo.tbl_groupCd Where groupCd in ('A000', 'C000', 'D000')"
		rsACADEMYget.Open sql, dbACADEMYget, 1

		if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
			Do Until rsACADEMYget.EOF
				strOpt = strOpt & "<option value='" & rsACADEMYget("groupCd") & "' "

				if nowCd=rsACADEMYget("groupCd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsACADEMYget("groupNm") & "</option>"
				rsACADEMYget.MoveNext
			Loop
		end if

		rsACADEMYget.Close

		optGroupCd = strOpt

	end function


	public FPrevID
	public FNextID

	'// ���� ������ �˻�
	public Function HasPreScroll()
		HasPreScroll = StarScrollPage > 1
	end Function

	'// ���� ������ �˻�
	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StarScrollPage + FScrollCount -1
	end Function

	'// ù������ ����
	public Function StarScrollPage()
		StarScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

end Class

%>