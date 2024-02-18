<%
Class CAddDep

public FEmpno
public Function fnGetAddDepList

 dim strSql
 strSql = "SELECT a.empno, a.userid, a.departmentid, b.departmentnameFull "
 strSql = strSql &  " FROM db_partner.dbo.tbl_partner_addDepartment as a "
 strsql = strSql & " inner join db_partner.[dbo].[vw_user_department] as b on a.departmentid = b.cid "
 strSql = strSql & " where a.isusing = 1 and a.empno ='"&FEmpno&"' "
 strSql = strSql & " order by a.depidx "
  rsget.Open strSql,dbget,1
 	if not rsget.eof then
 		fnGetAddDepList = rsget.getRows()
	end if
	rsget.close
End Function

End Class
%>