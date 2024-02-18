<%
'####################################################
' Description :  파트관리자 클래스
' History : 2011.01.25 김진영 생성
'####################################################
%>
<%
Class Partlist
	public idx
	public gubun
	public sabun
	public FGubun
	public FTeam

	'카테고리 중 상카테고리만 가져오기
	public Function fnGetlist
		Dim strSql, arrList, subStr
		If FGubun <> "" Then
			subStr = "and isusing = '"& FGubun &"'"
		end if
		strSql ="select idx, category1, gubun, isusing from db_partner.dbo.tbl_partperson_category where gubun = 0 " & subStr & " order by isusing desc , sortNo asc"
		rsget.Open strSql, dbget, 1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList = rsget.GetRows()
			fnGetlist = arrList
		END IF
		rsget.Close
	End Function

	'카테고리 테이블과 데이터테이블을 구분에 의해 조인
    '정렬 수정 2018.12.24 김광일
	public Function fnGetlist2
		Dim strSql2, arrList2
		strSql2 = ""
		strSql2 = strSql2 & "select D.idx, D.category1, D.gubun, D.isusing, M.idx, M.cidx, M.category1, M.sabun, M.isusing from db_partner.dbo.tbl_partperson_category as D "
		strSql2 = strSql2 &	"inner join db_partner.dbo.tbl_partperson_category2 as M on D.idx = M.cidx where D.gubun ='"& idx &"' Order by d.sortno asc  "
		''response.write "fnGetlist2 : " & strSql2
		rsget.Open strSql2, dbget, 1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList2 = rsget.GetRows()
			fnGetlist2 = arrList2
		END IF
		rsget.Close
	End Function

	'상/하 카테고리 중 번호에 맞는 내용 뽑아오기
	public Function fnGetmolist
		Dim strSql, arrList
		strSql ="select idx, category1, gubun, isusing from db_partner.dbo.tbl_partperson_category where idx='"& idx &"'"
		rsget.Open strSql, dbget, 1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList = rsget.GetRows()
			fnGetmolist = arrList
		END IF
		rsget.Close
	End Function

	'사번데이터가 들어있는 테이블의 데이터 뽑아오기
	public Function fnGetmolist2
		Dim strSql, arrList, subStr
		If FGubun = "Y" Then
			subStr = "and c2.isusing = 'Y'"
		end if
		strSql = ""
		strSql = strSql & "select c2.category1, c2.cidx, c1.category1, M.username, M.direct070, M.extension, M.usermail, c2.sabun, "
		strSql = strSql & "Case When P.posit_sn = '12' Then '사원' "
		strSql = strSql & "	When P.posit_sn = '13' Then '사원' "
		strSql = strSql & "Else P.posit_name End as posit_name "
		strSql = strSql & ", M.userid, c1.sortno, c2.isusing "
		strSql = strSql & "From db_partner.dbo.tbl_partperson_category AS c1 "
		strSql = strSql & "inner join db_partner.dbo.tbl_partperson_category2 AS c2 ON c2.cidx = c1.idx AND c1.gubun <> '0' "
		strSql = strSql & "inner join db_partner.dbo.tbl_user_tenbyten AS M ON c2.sabun = M.empno "
		strSql = strSql & "inner join db_partner.dbo.tbl_positInfo AS P ON m.posit_sn = P.posit_sn and P.posit_isDel = 'N' "
		strSql = strSql & "where c1.gubun = '" & idx & "' "& subStr &" order by c1.sortno asc"
		''response.write "fnGetmolist2 : " & strSql
		rsget.Open strSql, dbget, 1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList = rsget.GetRows()
			fnGetmolist2 = arrList
		END IF
		rsget.Close
	End Function

	'사번데이터가 들어있는 테이블의 데이터 뽑아오기
	public Function fnGetmolist3
		Dim strSql, arrList
		strSql = ""
		strSql = strSql & "select M.username, M.direct070, M.extension, M.usermail, D.sabun, P.posit_name, D.isusing, C.category1 from db_partner.dbo.tbl_partperson_category2 as D "
		strSql = strSql & "inner join db_partner.dbo.tbl_partperson_category C on C.idx = D.cidx "
		strSql = strSql & "inner join db_partner.dbo.tbl_user_tenbyten as M on D.sabun = M.empno "
		strSql = strSql & "inner join db_partner.dbo.tbl_positInfo P on m.posit_sn = P.posit_sn and P.posit_isDel = 'N'"
		strSql = strSql & "where D.cidx='"& idx &"'"
		rsget.Open strSql, dbget, 1
		IF Not (rsget.EOF OR rsget.BOF) THEN
			arrList = rsget.GetRows()
			fnGetmolist3 = arrList
		END IF
		rsget.Close
	End Function

	'####### 직원리스트 #######
	public Function fnMemberList
		Dim strSql

		strSql = "select T.userid, I.part_name, P.posit_name, T.username, T.part_sn, T.empno from db_partner.dbo.tbl_user_tenbyten as T "
		strSql = strSql & " inner join db_partner.dbo.tbl_positInfo P on T.posit_sn = P.posit_sn and P.posit_isDel = 'N' "
		strSql = strSql & " inner join db_partner.dbo.tbl_partInfo as I on T.part_sn = I.part_sn and I.part_isDel = 'N' "
		strSql = strSql & " where T.userid != '' and T.isusing = '1'" & vbcrlf

		' 퇴사예정자 처리	' 2018.10.16 한용민
		strSql = strSql & "	and (t.statediv ='Y' or (t.statediv ='N' and datediff(dd,t.retireday,getdate())<=0))" & vbcrlf
		if FTeam<>"" then
			strSql = strSql & " and T.part_sn IN(" & FTeam &") "
		end if
		strSql = strSql & " ORDER BY T.part_sn ASC, T.posit_sn ASC, T.regdate ASC "

		'response.write strSql & "<br>"
		rsget.Open strSql,dbget,1

		IF not rsget.EOF THEN
			fnMemberList = rsget.GetRows()
		END IF
		rsget.Close
	End Function

	public Function fnPartList
		Dim strSql, i
		strSql = "SELECT part_sn, part_name From [db_partner].[dbo].[tbl_partInfo] " &_
				 " WHERE part_isDel = 'N' " & _
				 "ORDER BY part_sort "
		rsget.Open strSql,dbget,1
		'response.write strSql
		IF not rsget.EOF THEN
			fnPartList = rsget.getRows()
		END IF
		rsget.Close
	End Function
End Class
%>
