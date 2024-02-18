<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"--> 
<%
dim ChkStart, ChkEnd, usePersonNo, usepoint, useMoney
dim strSql,arrList, intLoop
dim empno
 	ChkStart	= requestCheckvar(request("ChkStart"),10)  
	ChkEnd		= requestCheckvar(request("ChkEnd"),10)  
	usePersonNo	= requestCheckvar(request("usePersonNo"),8)
	empno		= requestCheckvar(request("empno"),16)  
	strSql = " SELECT solar_date, holiday from db_sitemaster.dbo.LunarToSolar where solar_date > '"&ChkStart&"' and solar_date<='"&ChkEnd&"'"
	rsget.Open strSql, dbget, 1
	if  not rsget.EOF  then
		arrList = rsget.getRows()
	end if
	rsget.close
		
	if isArray(arrList)	 then
		usepoint = 0
		 useMoney = 0
		 for intLoop = 0 To uBound(arrList,2)
		 	  
			  'if fnPeakSeason(arrList(0,intLoop)) or arrList(1,intLoop) >= 1  then '성수기 정책 폐지(2020.01.17)
			  if arrList(1,intLoop) >= 1  then
		 	  	usepoint = usepoint + 2 
		 	  else
		 	  	 
		 	  	if   uBound(arrList,2) >=3 and (intLoop+3) <= uBound(arrList,2) then 
		 	  		if arrList(1,intLoop+1) >= 1 and arrList(1,intLoop+2) >= 1  and arrList(1,intLoop+3) >= 1 then
		 	  			usepoint = usepoint + 2
		 	  		 
		 	  		else
		 	  			usepoint = usepoint + 1	
		 	  		end if
		 	  	else
		 	  		usepoint = usepoint + 1	 
		 	  	end if			 	  	
		 	  end if	 
		 	  useMoney =  useMoney + 15000
		 next
	end if
	
	if usePersonNo>=5 then
		useMoney = useMoney + (usePersonNo-4)*10000
	end if	
	
	dim totap,useap
	strSql = "select totPoint, usePoint  from db_partner.dbo.tbl_TenAgit_Point where isusing = 1  and  yyyy = year('"&ChkStart&"')  and empno ='"&empno&"' " 
	 
	rsget.Open strSql, dbget, 1
	if  not rsget.EOF  then
		totap = rsget("totPoint")
		useap = rsget("usePoint")
	end if
	rsget.close
	
	 
%>
<script type="text/javascript">
	parent.document.getElementById("iPoint").value = "<%=usepoint%>";
	parent.document.getElementById("sMoney").value = "<%=useMoney%>";	
	parent.document.getElementById("avPoint").value ="<%=totap-useap%>";
	parent.document.getElementById("chkp").value = 1;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->