<%
Class CAgitPoint
	public FPageSize
	public FCurrPage
	public FSPageNo
	public FEPageNo
	public FTotCnt
	public FTotPage

	public FRectposit_sn
	public FRectSearchKey
	public FRectSearchString
	public FRectStateDiv
	public Fdepartment_id
	public Finc_subdepartment 
	
	public FRectDateType
	public FRectStartDate
	public FRectEndDate
	public FRectStatus
	public FRectYYYY
	public FRectPenaltyKind
	
	'// 아지트 포인트 등록 리스트
	public Function fnAgitGetList
		Dim strSql , strSqlAdd
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage
		
		'// 검색어 쿼리 //
		strSqlAdd = ""
		if (FRectposit_sn <> "") then
			if (FRectposit_sn = "99") then
				strSqlAdd = strSqlAdd & "and u.posit_sn <= '11' " & vbCrlf
			else
				strSqlAdd = strSqlAdd & "and u.posit_sn = '" & FRectposit_sn & "' " & vbCrlf
			end if
		end if

	
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			if FRectSearchKey = "1" then 
				strSqlAdd = strSqlAdd & " and u.userid like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey ="2"	then
				strSqlAdd = strSqlAdd & " and u.username like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="3" then
				strSqlAdd = strSqlAdd & " and u.empno like '%" & FRectSearchString & "%' "
			end if
		end if
 

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				strSqlAdd = strSqlAdd & " AND u.department_id = '" & Fdepartment_id & "' "
			else
				strSqlAdd = strSqlAdd & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if
		
		if (FRectStateDiv <> "") then
			strSqlAdd = strSqlAdd & "and u.statediv = '" & FRectStateDiv & "' " & vbCrlf
		end if

		if FRectYYYY <> "" then
			strSqlAdd = strSqlAdd & " and m.yyyy = '"&FRectYYYY&"'"
		end if
		
		strSql = " SELECT count(pidx)  "
		strSql = strSql & " FROM db_partner.dbo.tbl_TenAgit_Point as m " & vbCrlf
		strSql =	strSql &"	 			inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno " & vbCrlf
		strSql =	strSql &"				left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql =  strSql & "			left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql =	strSql &"  where m.isusing = 1 and u.isusing = 1 " & vbCrlf
		strSql =  strSql & strSqlAdd
			
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
			 
		IF FTotCnt > 0 THEN
			strSql =  " SELECT pidx, empno, userid, username, joinday,departmentNameFull, posit_name, startday, endday, totPoint, usePoint , statediv,adminid, regdate " & vbCrlf
			strSql =	strSql &" FROM ( " & vbCrlf
			strSql =	strSql &" 		SELECT ROW_NUMBER() OVER (ORDER BY u.joinday DESC) as RowNum " & vbCrlf
			strSql =	strSql &" 			,pidx, m.empno, u.userid, u.username, u.joinday, isNull(dv.departmentNameFull,'') AS departmentNameFull, po.posit_name " & vbCrlf
			strSql =  strSql &"				, m.startday, m.endday, m.totPoint, m.usePoint , u.statediv, m.adminid, m.regdate " & vbCrlf
			strSql =	strSql &" 		FROM db_partner.dbo.tbl_TenAgit_Point as m " & vbCrlf
			strSql =	strSql &"	 			inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno " & vbCrlf
			strSql =	strSql &"				left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
			strSql =  strSql & "			left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
			strSql =	strSql &"  where m.isusing = 1 and u.isusing = 1 " & vbCrlf
			strSql  = strSql & strSqlAdd
			strSql =	strSql &" ) AS TB " & vbCrlf
			strSql =	strSql &" WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				fnagitGetList = rsget.getRows()
			End IF
			rsget.Close
		END IF
	End Function
	
	'// 미등록 리스트
	public Function fnGetNonRegAgitList
		dim strSql, strSqlAdd
		dim dRectyyyy, dRectday
		dRectyyyy = year(date())	
		
		'// 검색어 쿼리 //
		strSqlAdd = ""
		if (FRectposit_sn <> "") then
			if (FRectposit_sn = "99") then
				strSqlAdd = strSqlAdd & "and u.posit_sn <= '11' " & vbCrlf
			else
				strSqlAdd = strSqlAdd & "and u.posit_sn = '" & FRectposit_sn & "' " & vbCrlf
			end if
		end if

	
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			if FRectSearchKey = "1" then 
				strSqlAdd = strSqlAdd & " and u.userid like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey ="2"	then
				strSqlAdd = strSqlAdd & " and u.username like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="3" then
				strSqlAdd = strSqlAdd & " and u.empno like '%" & FRectSearchString & "%' "
			end if
		end if
 

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				strSqlAdd = strSqlAdd & " AND u.department_id = '" & Fdepartment_id & "' "
			else
				strSqlAdd = strSqlAdd & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if
		
		if (FRectStateDiv <> "") then
			strSqlAdd = strSqlAdd & "and u.statediv = '" & FRectStateDiv & "' "  
		end if
			
		strSql = " SELECT count(u.empno)  " & vbCrlf
		strSql= strSql & " FROM db_partner.dbo.tbl_user_tenbyten as u " & vbCrlf
		strSql= strSql & "	LEFT OUTER JOIN db_partner.dbo.tbl_TenAgit_Point as m on u.empno = m.empno and m.yyyy = '"&dRectyyyy&"' and m.isusing=1" & vbCrlf
		strSql=	strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql= strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql= strSql & " WHERE  u.isusing = 1 " & vbCrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		strSql = strSql & "	and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
		strSql= strSql & "	and u.part_sn <> 17   " & vbCrlf
		strSql = strSql &"    and u.posit_sn < 13" & vbCrlf '시급계약직 이상
		strSql= strSql & "  and m.empno is null "
		strSql = strSql & strSqlAdd  
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			End IF
			rsget.Close
			 
		
		IF 	FTotCnt > 0 THEN
		strSql = " SELECT  u.empno, u.userid, username, joinday,departmentNameFull, posit_name " & vbCrlf
		strSql= strSql & " FROM db_partner.dbo.tbl_user_tenbyten as u " & vbCrlf
		strSql= strSql & "	LEFT OUTER JOIN db_partner.dbo.tbl_TenAgit_Point as m on u.empno = m.empno and m.yyyy = '"&dRectyyyy&"' and m.isusing=1" & vbCrlf
		strSql=	strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql= strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql= strSql & " WHERE  u.isusing = 1 " & vbCrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		strSql = strSql & "	and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
		strSql= strSql & "		and u.part_sn <> 17  " & vbCrlf
		strSql = strSql &"    and u.posit_sn < 13" & vbCrlf '시급계약직 이상
		strSql= strSql & "  and m.empno is null "
		strSql = strSql & strSqlAdd
		strSql = strSql & " order by u.joinday desc" 

			'response.write strSql & "<br>"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				fnGetNonRegAgitList = rsget.getRows()
			End IF
			rsget.Close
		 
		END IF	
	end Function

	'텐바이텐 아지트 패널티 목록
	public Function fnGetAgitPenaltyList()
		Dim strSql, addSql

		if FRectSearchKey<>"" and FRectSearchString<>"" then
			if FRectSearchKey = "1" then 
				addSql = addSql & " and u.userid like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey ="2"	then
				addSql = addSql & " and u.username like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="3" then
				addSql = addSql & " and u.empno like '%" & FRectSearchString & "%' "
			end if
		end if

		if FRectPenaltyKind<>"" then
			addSql = addSql & " and p.penaltykind='" & FRectPenaltyKind & "' "
		end if

		if FRectStateDiv<>"1" then
			addSql = addSql & " and getdate() between p.startdate and p.enddate "
		end if

		strSql = "select count(p.pidx) as totCnt, CEILING(CAST(Count(p.pidx) AS FLOAT)/" & FPageSize & ") as totPg "
		strSql = strSql & " from db_partner.dbo.tbl_TenAgit_penalty as p with (noLock) "
		strSql = strSql & " 	join db_partner.dbo.tbl_user_tenbyten as u with (noLock) "
		strSql = strSql & " 		on p.empno = u.empno "
		strSql = strSql & " 	left join db_partner.dbo.tbl_TenAgit_Booking as b with (noLock) "
		strSql = strSql & " 		on p.idx=b.idx "
		strSql = strSql & " Where u.isusing=1 " & addSql
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FTotCnt = rsget("totCnt")
			FTotPage = rsget("totPg")
		End IF
		rsget.Close

		IF FTotCnt > 0 THEN
			strSql = "select p.pidx, p.idx, p.penaltykind, p.startdate, p.enddate, p.regdate, p.empno "
			strSql = strSql & " 	,u.userid, u.username "
			strSql = strSql & " 	,b.AreaDiv, p.penaltyCause, p.penaltyPoint "
			strSql = strSql & " from db_partner.dbo.tbl_TenAgit_penalty as p with (noLock) "
			strSql = strSql & " 	join db_partner.dbo.tbl_user_tenbyten as u with (noLock) "
			strSql = strSql & " 		on p.empno = u.empno "
			strSql = strSql & " 	left join db_partner.dbo.tbl_TenAgit_Booking as b with (noLock) "
			strSql = strSql & " 		on p.idx=b.idx "
			strSql = strSql & " Where u.isusing=1 " & addSql
			strSql = strSql & " order by p.idx desc "
			strSql = strSql & "	OFFSET " & ((FCurrPage-1)*FPageSize) & " ROWS FETCH NEXT " & FPageSize & " ROWS ONLY"
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly

			IF not rsget.EOF THEN
				fnGetAgitPenaltyList =rsget.getRows()
			END IF
			rsget.close
		End if
	end Function

End Class

Class CMyAgit
 
public FRectEmpno
public FRectChkStart
public Fpidx
public FtotPoint
public FusePoint
public Fstartday
public Fendday
public Fpenaltykind
public Fpenaltysdate
public Fpenaltyedate
public FpenaltyCause
public FpenaltyPoint

	public Function fnGetMyAgit
		Dim strSql
		strSql = " SELECT top 1 A.pidx, A.totPoint, A.usePoint, A.startday, A.endday, isNull(p.penaltykind,0) penaltykind, p.startdate as psdate, p.enddate as pedate, p.penaltyCause, p.penaltyPoint "
		strSql = strSql & " FROM db_partner.dbo.tbl_TenAgit_Point as A  "
		strSql = strSql & "  left outer join db_partner.dbo.tbl_TenAgit_penalty as P "
		strSql = strSql & "			on A.empno = P.empno and P.startdate <= getdate() and p.enddate >=getdate() "
		strSql = strSql & " WHERE A.isusing = 1 "
		strSql = strSql & "  and  A.yyyy = year('"&FRectChkStart&"') "
		strSql = strSql & "  and A.empno ='"&FRectEmpno&"' "  
		 
			rsget.CursorLocation = adUseClient
			rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
			IF not rsget.EOF THEN
				Fpidx = rsget("pidx")
				FtotPoint = rsget("totPoint")
				FusePoint = rsget("usePoint")
				Fstartday = rsget("startday")
				Fendday = rsget("endday")
				Fpenaltykind = rsget("penaltykind")
				Fpenaltysdate = rsget("psdate")
				Fpenaltyedate = rsget("pedate")
				FpenaltyCause = rsget("penaltyCause")
				FpenaltyPoint = rsget("penaltyPoint")
		END IF
		rsget.close
	End Function
	
	public Function fnGetMyInfoAgit
		Dim strSql
		strSql = " SELECT  A.pidx, A.yyyy, A.totPoint, A.usePoint, A.startday, A.endday, isNull(p.penaltykind,0) penaltykind, p.startdate as psdate, p.enddate as pedate "
		strSql = strSql & " FROM db_partner.dbo.tbl_TenAgit_Point as A  "
		strSql = strSql & "  left outer join db_partner.dbo.tbl_TenAgit_penalty as P "
		strSql = strSql & "			on A.empno = P.empno and P.startdate <= getdate() and p.enddate >=getdate() "
		strSql = strSql & " WHERE A.isusing = 1 "
		strSql = strSql & "  and  A.yyyy >= year(getdate()) "
		strSql = strSql & "  and A.empno ='"&FRectEmpno&"' "  
		strSql = strSql & " order by A.yyyy "
		 
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			fnGetMyInfoAgit =rsget.getRows()
		END IF
		rsget.close
	End Function

End Class

Class CAgitUse
public FPageSize
public FCurrPage
public FSPageNo
public FEPageNo
public FTotCnt

public FRectposit_sn
public FRectSearchKey
public FRectSearchString
public FRectStateDiv
public Fdepartment_id
public Finc_subdepartment 

public FRectDateType
public FRectStartDate
public FRectEndDate
public FRectStatus
public FRectYYYY
	
public FRectChkTerm	
public FRectSYYYYMM
public FRectEYYYYMM
public FRectIpkum
public FRectreturnkey
public FRectUsing 
public FRectRefund
public FRectAreadiv
	
	public Function FnAgitUseList
		Dim strSql , strSqlAdd
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage
		
		'// 검색어 쿼리 //
		strSqlAdd = ""
		if (FRectposit_sn <> "") then
			if (FRectposit_sn = "99") then
				strSqlAdd = strSqlAdd & "and u.posit_sn <= '11' " & vbCrlf
			else
				strSqlAdd = strSqlAdd & "and u.posit_sn = '" & FRectposit_sn & "' " & vbCrlf
			end if
		end if

	
		if FRectSearchKey<>"" and FRectSearchString<>"" then
			if FRectSearchKey = "1" then 
				strSqlAdd = strSqlAdd & " and b.userid like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey ="2"	then
				strSqlAdd = strSqlAdd & " and u.username like '%" & FRectSearchString & "%' "
			elseif FRectSearchKey="3" then
				strSqlAdd = strSqlAdd & " and b.empno like '%" & FRectSearchString & "%' "
			end if
		end if
 

		if (Fdepartment_id <> "") then
			if (Finc_subdepartment = "N") then
				strSqlAdd = strSqlAdd & " AND u.department_id = '" & Fdepartment_id & "' "
			else
				strSqlAdd = strSqlAdd & " AND (IsNull(dv.cid1, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid2, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid3, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid4, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid5, -1) = '" & Fdepartment_id & "' or IsNull(dv.cid6, -1) = '" & Fdepartment_id & "') "
			end if
		end if
		
		if (FRectStateDiv <> "") then
			strSqlAdd = strSqlAdd & "and u.statediv = '" & FRectStateDiv & "' " & vbCrlf
		end if
		
		if (FRectAreadiv<>"") then
			strSqlAdd = strSqlAdd & "and b.areadiv = '" & FRectAreadiv & "' " & vbCrlf
	end if

		 if FRectChkTerm <> "" then
		 	strSqlAdd = strSqlAdd & " and ( ( convert(varchar(7),chkstart,121) >= '"&FRectSYYYYMM&"' and  convert(varchar(7),chkstart,121) <='"&FRectEYYYYMM&"' )  "
		 	strSqlAdd = strSqlAdd & "  			or ( convert(varchar(7),chkend,121) >= '"&FRectSYYYYMM&"' and  convert(varchar(7),chkend,121) <='"&FRectEYYYYMM&"' )  )"
		 end if
		 
		 if FRectIpkum <> "" then 
		 	strSqlAdd = strSqlAdd & " and isipkum = "&FRectIpkum
		 end if
		 
		 if FRectreturnkey <> "" then 
		 	strSqlAdd = strSqlAdd & " and isreturnkey = "&FRectreturnkey
		 end if
		 
		  if FRectUsing <> "" then 
		 	strSqlAdd = strSqlAdd & " and b.isUsing = '"&FRectUsing&"'"
		 end if
		 
		 if FRectRefund <>"" then 
		 		strSqlAdd = strSqlAdd & " and  b.isUsing = 'N' and isrefund = "&FRectRefund			 
		 end if
		
		strSql = " select count(b.idx) " 
		strSql = strSql & "  	from db_partner.dbo.tbl_TenAgit_Booking as b "
		strSql = strSql & "			left outer join db_partner.dbo.tbl_TenAgit_Point as a on (b.empno = a.empno or b.userid = a.userid) and a.yyyy = convert(varchar(4),chkstart,121) "
		strSql = strSql & "			left outer join db_partner.dbo.tbl_user_tenbyten as u on (b.empno = u.empno or b.userid = u.userid) and u.isusing =1 "
		strSql = strSql & " 		left outer join db_partner.dbo.tbl_TenAgit_penalty as p on b.idx = p.idx  "
		strSql = strSql &  " where 1=1 " &strSqlAdd
	 
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FTotCnt = rsget(0)
		END IF
		rsget.close
		
		if FTotCnt > 0 then
		strSql = " SELECT idx, empno, userid, username, joinday, departmentNameFull, posit_name, statediv"
		strSql = strSql & "	 , areadiv, chkstart, chkend, usepersonno, usepoint, lpoint,usemoney"
		strSql = strSql & "  , isipkum, ipkumdate, isreturnkey, isUsing, refunddate, canceldate   "	
		strSql = strSql & "  , penaltykind,  psdate,  pedate, regdate "
		strSql = strSql & " FROM (" 
		strSql = strSql & " 		select  ROW_NUMBER() OVER (ORDER BY b.idx DESC) as RowNum "
		strSql = strSql & " 		,b.idx, b.empno, isNull(u.userid,b.userid) as userid, u.username,  joinday,departmentNameFull, posit_name , u.statediv "
		strSql = strSql & "		 , b.areadiv, b.chkstart, b.chkend, usepersonno, b.usepoint,(a.totpoint-a.usepoint) as lpoint, isNull(b.usemoney,0) as usemoney"
		strSql = strSql & "    , b.isipkum, b.ipkumdate, b.isreturnkey, b.isUsing, b.refunddate, b.canceldate   "	
		strSql = strSql & "		 , isNull(p.penaltykind,0) penaltykind, p.startdate as psdate, p.enddate as pedate, b.regdate  "
		strSql = strSql & "  	from db_partner.dbo.tbl_TenAgit_Booking as b "		
		strSql = strSql & "			left outer join db_partner.dbo.tbl_TenAgit_Point as a on (b.empno = a.empno or b.userid = a.userid ) and a.yyyy = convert(varchar(4),chkstart,121) "
		strSql = strSql & "			left outer join db_partner.dbo.tbl_user_tenbyten as u on (b.empno = u.empno or b.userid = u.userid) and u.isusing =1 "
		strSql = strSql & "			left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql = strSql & "			left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql = strSql & " 		left outer join db_partner.dbo.tbl_TenAgit_penalty as p on b.idx = p.idx  "
		strSql = strSql & "	where 1= 1 " &strSqlAdd
		strSql =	strSql &" ) AS TB " & vbCrlf
		strSql =	strSql &" WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo		  

        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			FnAgitUseList = rsget.getRows()
		END IF
		rsget.close
	end if
	End Function

	'// 아지트 안내 문자 접수
	Function fnGetAgitSmsCont()
		Dim strSql
		strSql = "Select top 1 * from db_partner.dbo.tbl_TenAgit_smsInfo Where AreaDiv=" & FRectAreadiv
        rsget.CursorLocation = adUseClient
        rsget.Open strSql,dbget,adOpenForwardOnly, adLockReadOnly
		IF not rsget.EOF THEN
			fnGetAgitSmsCont = rsget.getRows()
		end if
		rsget.Close
	End Function
End Class


public Function FnWeekName(iWeekNo)
	if iWeekNo = "" then exit Function

	dim strWeek
	if iWeekNo = 1 THEN
		strWeek = "일"
	elseif 	iWeekNo = 2 THEN
		strWeek = "월"
	elseif 	iWeekNo = 3 THEN
		strWeek = "화"
	elseif 	iWeekNo = 4 THEN
		strWeek = "수"
	elseif 	iWeekNo = 5 THEN
		strWeek = "목"
	elseif 	iWeekNo = 6 THEN
		strWeek = "금"
	elseif 	iWeekNo = 7 THEN
		strWeek = "토"
	end if
	FnWeekName = 	strWeek
End Function

public Function AgitName(areaDiv)
	Select Case areaDiv
		Case "1"
			AgitName = "제주"
		Case "2"
			AgitName = "양평"
		Case "3"
			AgitName = "속초"
	end Select
end Function

public Function PenaltyKindName(penaltyKind)
	Select Case penaltyKind
		Case "1"
			PenaltyKindName = "투숙일 5일전 취소"
		Case "2"
			PenaltyKindName = "투숙 당일 취소"
		Case "3"
			PenaltyKindName = "No-Show"
		Case "4"
			PenaltyKindName = "관리자 패널티"
	end Select
end Function
%>
 	 