<%
Class Cvitamin
  public FPageSize
	public FCurrPage
	public FSPageNo
	public FEPageNo
	public FTotCnt
	
	 
	public FRectposit_sn
	public FRectSearchKey
	public FRectSearchString
	public FRectStateDiv
	public FRectIdx
	public Fdepartment_id
	public Finc_subdepartment 
	
	public FRectDateType
	public FRectStartDate
	public FRectEndDate
	public FRectStatus
	public FRectyyyy
	public FRectOrderby
	
	'// 비타민 등록 리스트
	public Function fnvitaminGetList
		Dim strSql , strSqlAdd, strOrderby
		FSPageNo = (FPageSize*(FCurrPage-1)) + 1
		FEPageNo = FPageSize*FCurrPage
		
		'// 검색어 쿼리 //
		strSqlAdd = ""
		if FRectIdx<>"" then
			strSqlAdd = strSqlAdd & "and m.midx = " & FRectIdx & vbCrlf
		end if

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
		 
		 if FRectOrderby ="CA" then
		 	 strOrderby = " m.empno asc "
		 elseif 	  FRectOrderby ="ND" then
		 	 strOrderby = " u.username desc, m.empno desc "
		 	elseif 	  FRectOrderby ="NA" then
		 	 strOrderby = "u.username asc, m.empno desc "
		 	else 
		 	 strOrderby = " m.empno desc "  
		  end if
		
		 
		  strSql = " SELECT count(midx)  "
		  strSql = strSql & " FROM db_partner.dbo.tbl_vitamin_master as m " & vbCrlf
			strSql =	strSql &"	 			inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno " & vbCrlf
			strSql =	strSql &"				left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
			strSql =  strSql & "			left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
			strSql =	strSql &"  where m.isusing = 1 and u.isusing = 1 " & vbCrlf
			strSql =  strSql & strSqlAdd
			
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			END IF
			rsget.close
			 
			IF FTotCnt > 0 THEN
			strSql =  " SELECT midx, empno, userid, username, joinday,departmentNameFull, posit_name, startday, endday, totvm, usevm , statediv,adminid, regdate " & vbCrlf
			strSql =	strSql &" FROM ( " & vbCrlf
			strSql =	strSql &" 		SELECT ROW_NUMBER() OVER (ORDER BY "&strOrderby&") as RowNum " & vbCrlf
			strSql =	strSql &" 			,midx, m.empno, u.userid, u.username, u.joinday, isNull(dv.departmentNameFull,'') AS departmentNameFull, po.posit_name " & vbCrlf
			strSql =  strSql &"				, m.startday, m.endday, m.totvm, m.usevm , u.statediv, m.adminid, m.regdate " & vbCrlf
			strSql =	strSql &" 		FROM db_partner.dbo.tbl_vitamin_master as m " & vbCrlf
			strSql =	strSql &"	 			inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno " & vbCrlf
			strSql =	strSql &"				left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
			strSql =  strSql & "			left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
			strSql =	strSql &"  where m.isusing = 1 and u.isusing = 1 " & vbCrlf
			strSql  = strSql & strSqlAdd
			strSql =	strSql &" ) AS TB " & vbCrlf
			strSql =	strSql &" WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnvitaminGetList = rsget.getRows()
			End IF
			rsget.Close
			END IF
	End Function
	
	'// 미등록 리스트
	public Function fnGetNonRegVMList
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
		strSql= strSql & "	LEFT OUTER JOIN db_partner.dbo.tbl_vitamin_master as m on u.empno = m.empno and m.yyyy = '"&dRectyyyy&"' and m.isusing=1" & vbCrlf
		strSql=	strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql= strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql= strSql & " WHERE  u.isusing = 1 " & vbCrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		strSql = strSql & "	and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
		strSql= strSql & "	and u.posit_sn <= 11 and u.part_sn <> 17 " & vbCrlf
		strSql= strSql & "  and m.empno is null "
		strSql = strSql & strSqlAdd  
		rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				FTotCnt = rsget(0)
			End IF
			rsget.Close
			 
		
		IF 	FTotCnt > 0 THEN
		strSql = " SELECT  u.empno, u.userid, username, joinday,departmentNameFull, posit_name " & vbCrlf
		strSql= strSql & " FROM db_partner.dbo.tbl_user_tenbyten as u " & vbCrlf
		strSql= strSql & "	LEFT OUTER JOIN db_partner.dbo.tbl_vitamin_master as m on u.empno = m.empno and m.yyyy = '"&dRectyyyy&"' and m.isusing=1" & vbCrlf
		strSql=	strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql= strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
		strSql= strSql & " WHERE  u.isusing = 1 " & vbCrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		strSql = strSql & "	and (u.statediv ='Y' or (u.statediv ='N' and datediff(dd,u.retireday,getdate())<=0))" & vbcrlf
		strSql= strSql & "	and u.posit_sn <= 11  and u.part_sn <> 17 " & vbCrlf
		strSql= strSql & "  and m.empno is null "
		strSql = strSql & strSqlAdd
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnGetNonRegVMList = rsget.getRows()
			End IF
			rsget.Close
		 
		END IF	
	end Function
	
		 
		
		
	'//비타민 상세리스트
	public Function fnGetDetailList 
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

		if FRectStatus <> "" then
			if FRectStatus =8 then
				strSqlAdd = strSqlAdd & "and d.vmstatus = 0 and e.reportidx is null "& vbCrlf				
			else
				strSqlAdd = strSqlAdd & "and d.vmstatus = "&FRectStatus&" and e.reportidx is not null "& vbCrlf
			end if
		end if
		 
		if FRectDateType ="1" then
				strSqlAdd = strSqlAdd & " and d.regdate >='"&FRectStartdate&"' and d.regdate < '"&dateadd("d",1,FRectEndDate)&"'"
		else
				strSqlAdd = strSqlAdd & " and d.paydate >='"&FRectStartdate&"' and d.paydate < '"&dateadd("d",1,FRectEndDate)&"'"
		end if
		 
		 strSql = " SELECT count(d.didx) "
	  strSql = strSql & " FROM db_partner.dbo.tbl_vitamin_detail as d "
	  strSql = strSql & "  inner join db_partner.dbo.tbl_vitamin_master as m on d.midx = m.midx "	  
	  strSql = strSql & "	 inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno "
	  strSql=	strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql= strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
	  strSql = strSql & " left outer join db_partner.dbo.tbl_eappreport as e on d.didx = e.scmlinkno  and e.edmsidx = 33 "
	  strSql = strSql & " 	where m.isusing = 1 and d.isusing =1 "
	  strSql = strSql & strSqlAdd
	  rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				FtotCnt = rsget(0)
			End IF
			rsget.Close
	  
	 ' if FtotCnt > 0 then
	  strSql = " SELECT didx, vmmoney, regdate, paydate, vmstatus, empno, userid, username, joinday, departmentNameFull, posit_name, isNull(reportidx,0) as reportidx , reportstate  " & vbCrlf
	  strSql = strSql & " FROM ( " & vbCrlf
	  strSql = strSql & " SELECT ROW_NUMBER() OVER (ORDER BY d.didx DESC) as RowNum " & vbCrlf 
	  strSql = strSql & " , d.didx, d.vmmoney, d.regdate, d.paydate, d.vmstatus,u.empno, u.userid, u.username, joinday, departmentNameFull, posit_name, isNull(e.reportidx,0) as reportidx , e.reportstate  " & vbCrlf
	  strSql = strSql & " FROM db_partner.dbo.tbl_vitamin_detail as d " & vbCrlf
	  strSql = strSql & "  inner join db_partner.dbo.tbl_vitamin_master as m on d.midx = m.midx "	   & vbCrlf
	  strSql = strSql & "	 inner join db_partner.dbo.tbl_user_tenbyten as u on m.empno = u.empno " & vbCrlf
	  strSql = strSql & "	left join db_partner.dbo.vw_user_department as dv on u.department_id = dv.cid	and dv.useYN = 'Y'" & vbCrlf
		strSql = strSql & "	left join [db_partner].dbo.tbl_positInfo as po on u.posit_sn = po.posit_sn " & vbCrlf
	  strSql = strSql & " left outer join db_partner.dbo.tbl_eappreport as e on d.didx = e.scmlinkno  and e.edmsidx = 33 " & vbCrlf
	  strSql = strSql & " 	where m.isusing = 1 and d.isusing =1 " & vbCrlf
	  strSql = strSql & strSqlAdd
	  strSql =	strSql &" ) AS TB " & vbCrlf
		strSql =	strSql &" WHERE TB.RowNum Between  "&FSPageNo&" AND  "&FEPageNo	  
	 
	  	rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnGetDetailList = rsget.getRows()
			End IF
			rsget.Close
	'	end IF
	End Function
End Class

Class CMyVitamin

public FRectEmpno
public FRectyyyy

public Ftotvm
public Fusevm
 


	public Function fnGetMyVitamin
		dim strSql
		FRectyyyy = year(date())
		strSql = " SELECT totvm, usevm  "
		strSql = strSql & " FROM db_partner.dbo.tbl_vitamin_master "
		strSql = strSql & " WHERE empno = '"&FRectEmpno& "' and yyyy ='"& FRectyyyy&"' and isusing =1  "
			rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				Ftotvm = rsget("totvm")
				Fusevm = rsget("usevm") 
		END IF
		rsget.close
	End Function
	 
	public Function fnGetMyVitaminList
	dim strSql
	strSql = "SELECT d.didx, d.vmmoney, d.regdate, d.paydate, d.vmstatus, isNull(e.reportidx,0) as reportidx , e.reportstate "
	strSql = strSql & " FROM db_partner.dbo.tbl_vitamin_master as m"
	strSql = strSql & " inner join  db_partner.dbo.tbl_vitamin_detail as d on m.midx = d.midx and d.isusing =1 "
	strSql = strSql & "	left outer join db_partner.dbo.tbl_eappreport as e on d.didx = e.scmlinkno and e.edmsidx = 33 "
	strSql = strSql & " WHERE m.empno =  '"&FRectEmpno& "' and m.yyyy ='"& FRectyyyy&"' and m.isusing =1 "
	strSql = strSql & " order by m.midx desc , d.didx desc"
	rsget.Open strSql,dbget,1
			IF not rsget.EOF THEN
				fnGetMyVitaminList = rsget.getRows()
		END IF
		rsget.close
	End Function
End Class

function fnMyStatusDesc(istatus, ireportidx, didx, reqvm)
	if istatus = "0" and ireportidx ="0" THEN
		%>신청
		<input type="button" class="button" value="품의서작성" onClick="jsRegEapp('<%=didx%>','<%=reqvm%>');">
		<input type="button" class="button" value="삭제" onClick="jsDelVM('<%=didx%>');">
		<%
	elseif istatus ="0" and ireportidx <> "0"	 THEN
				%>승인대기
		<input type="button" class="button" value="품의서보기" onClick="jsViewEapp(<%=ireportidx%>);">
		<%
		elseif istatus ="3" and ireportidx <> "0"	 THEN
				%>승인보류
				<input type="button" class="button" value="품의서보기" onClick="jsViewEapp(<%=ireportidx%>);">
		<%
		elseif istatus ="5" and ireportidx <> "0"	 THEN
				%>승인반려
				<input type="button" class="button" value="품의서보기" onClick="jsViewEapp(<%=ireportidx%>);">
		<%
		elseif istatus ="1" and ireportidx <> "0"	 THEN
				%>승인완료				
		<input type="button" class="button" value="품의서보기" onClick="jsViewEapp(<%=ireportidx%>);">
		<%	elseif istatus ="7" and ireportidx <> "0"	 THEN
				%>지급완료 
		<%
	end if
End Function

function fnStatusDesc(istatus, ireportidx)
	if istatus = "0" and ireportidx ="0" THEN
		%>신청 
		<%
	elseif istatus ="0" and ireportidx <> "0"	 THEN
		%>승인대기 
		<%
		elseif istatus ="3" and ireportidx <> "0"	 THEN
		%>승인보류 
		<%
			elseif istatus ="5" and ireportidx <> "0"	 THEN
		%>
		승인반려
		<%
	elseif istatus ="1" THEN
		%>승인완료 
		<%	
	elseif istatus ="7"  THEN
		%>지급완료 
		<%
	end if
End Function
%>