<%
'Response.Buffer = TRUE
%>
<html>
<title>GeoIP Test</title>
<body>
<h1><center>GeoIP Test</center></h1>
<br><center>
<%
	set geoip = Server.CreateObject("GeoIPCOM.GeoIP")

	hostname = Request.Form("hostname")

	if Request.Form("submit") = "Submit" then
		
		set geoip = Server.CreateObject("GeoIPCOM.GeoIP")
		geoip.loadDataFile("C:\GeoIP\GeoIP.dat")
	
		country_code = geoip.country_code_by_name(hostname)
		
		country_name = geoip.country_name_by_name(hostname)
		set geoip = nothing
		
		Response.Write("<table cellpadding=2 border=1><tr><th colspan=2>Results</th></tr>")
		Response.Write("<tr><td>Hostname</td><td>" + hostname + "</td></tr>")
		Response.Write("<tr><td>ISO 3166 Country Code</td><td>" + country_code + "</td></tr>")
		Response.Write("<tr><td>Full Country Name</td><td>" + country_name + "</td></tr>")
		Response.Write("</table>")
	
	end if
%>
<br>
<form action="GeoIPTest.asp" method="POST">
<table border=0>
<tr><td>hostname:</td><td><input type=text value="<%= hostname%>" name=hostname></td></tr>
</table>
<input type=submit value="Submit" name=submit>
</form>
</center>

</body>
</html>
