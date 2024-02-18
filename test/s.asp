<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%

dim str
str = "Mozilla/5.0 (Linux; Android 4.4; Nexus 5 Build/_BuildID_)"

response.write str


Dim re, matches, Img, item
Set re = new RegExp     ' creates the RegExp object
re.IgnoreCase = true
re.Global = true
re.Pattern = "Android [0-9]+\.[0-9]+"
Set Matches  = re.Execute(str)
Img = ""

''response.write UBound(Matches)


For each Item in Matches
   response.write "aaaa" & Item.Value
Next

%>
