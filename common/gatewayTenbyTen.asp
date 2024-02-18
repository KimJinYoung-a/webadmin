<script language="javascript" >document.domain="10x10.co.kr";</script>
<%
	Dim url , uid , uname
	
	url = request("url")
	uid = request("uid")
	uname = request("uname")

	Response.AddHeader "P3P", "CP=ALL CURa ADMa DEVa TAIa OUR BUS IND PHY ONL UNI PUR FIN COM NAV INT DEM CNT STA POL HEA PRE LOC OTC"
   
	if session("ssBctId") = "" then
		session("ssBctId") = uid
		session("ssBctDiv") = 9
		session("ssBctCname") = uname
		Session.Timeout=20 
	end If

	If session("ssBctId") <> ""   Then 
	 	response.redirect(url)  
	End if
	 	response.End 
%>

