<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/EmoDic/EmoDicCls.asp" -->
<%

dim eNumber
eNumber = request("eno")
dim eType 
eType = request("etp")
dim eTitle
eTitle = request("etlt")
IF eNumber="" or eType="" Then
	response.write "肋给等 立辟涝聪促"
	dbget.close()	:	response.End
End if

dim oWord,iLp
set oWord =	new EmodicCls
oWord.FRectEmoNumber = eNumber
oWord.FRectEmoType = eType
oWord.FRectEmoTitle = eTitle
oWord.getEmoCommentList

IF oWord.FResultCount<0 Then
	response.write "坷幅 包府磊俊霸 巩狼秦林技夸"
	dbget.close()	:	response.End
End if


function getETypeStr(eTp)
	dim eStr 
	Select Case etp
	Case "1"
		eStr = "掺傣掺傣"	
	Case"2"
		eStr = "倔贩蹲顶"	
	Case "3"
		eStr = "教件积件"		
	Case "4"
		eStr = "尝府尝府"
	End Select
	getETypeStr=eStr
End Function

%>
<table width="350" border="0" class="a" cellpadding="2" cellspacing="1" align="left" bgcolor="<%=adminColor("tablebg") %>">
<form name="" method="post" action="<%= uploadUrl %>/linkweb/Emodic_Upload.asp" Enctype="multiPart/form-data"> 
<input type="hidden" name="eno" value="<%= eNumber %>">
<input type="hidden" name="etp" value="<%= eType %>">
<input type="hidden" name="etlt" value="<%= eTitle %>">
<input type="hidden" name="rgflag" value="edit">
<tr bgcolor="#FFFFFF">
	<td colspan="2" bgcolor="<%=adminColor("tablebar")%>">
		<b><%=eNumber %>瞒 - <%=getETypeStr(eType) %></b> 览葛磊 府胶飘 
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="40">窜绢</td>
	<td width="295"><b><%= eTitle%></b></td>
</tr>
<% for iLp= 0 To oWord.FResultCount -1 %>
<tr  bgcolor="#FFFFFF">
	<td><%=oWord.FList(iLp).Userid %></td>
	<td><%=oWord.FList(iLp).Comment %></td>
</tr>
<% Next %>
</form>
</table>

<% SET oWord = nothing %>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->