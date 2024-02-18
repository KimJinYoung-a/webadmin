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

IF eNumber="" then eNumber="1"
IF eType="" Then eType="1"
	
dim oWord,iLp
set oWord =	new EmodicCls
oWord.FRectEmoNumber = eNumber
oWord.FRectEmoType = eType
oWord.getEmoWordsList
%>
<script language="javascript" type="text/javascript">

function fncgsel(){
	document.rFrm.submit();
}
function fnbatIn(){
	var en,et 
	en= document.rFrm.eno.value;
	et= document.rFrm.etp.value;
	var inwin = window.open('EmoBatchInput.asp?eno='+ en +'&etp='+ et ,'inwinn','resizable=yes,scrollbars=yes');
	//document.rFrm.target='inwin';
	//document.rFrm.action='';
	//document.rFrm.submit();
}
function fnEdit(stt){
	document.rWordFrm.etlt.value=stt;
	document.rWordFrm.submit()
}
function fnUsingupdate(uv){
	document.rWordFrm.action='EmoDic_Proc.asp';
	document.rWordFrm.ius.value=uv;
	document.rWordFrm.mode.value='allUsing';
	document.rWordFrm.submit()
	
}
function fncomlist(stt){
	document.rWordFrm.action='EmoDicCommList.asp';
	document.rWordFrm.etlt.value=stt;
	document.rWordFrm.submit()
}
function fnarrEdit(){
	var frm = document.rWordFrm;
	
	for(i=0 ;i<<%=oWord.FResultCount%>;i++){
		var tus = document.getElementsByName('ius'+i);
		//alert(tus);
		for(j=0; j<tus.length; j++){
			if(tus[j].checked){
				frm.ius.value= frm.ius.value  + tus[j].value + ",";
			}
		}
	}
	frm.mode.value='arrEdit';
	frm.action='EmoDic_Proc.asp';
	frm.submit();
}
</script>

<table width="800" height="450" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">

	<td bgcolor="#FFFFFF" valign="top" width="150">
		<!-- 차수 분류 선택 // -->
		<form name="rFrm" method="get" action="">
		<table border="0" cellpadding="5" class="a" cellspacing="0">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<tr>
			<td valign="top"><b>차수</b> : </td>
			<td valign="top">
				<select name="eno"  style="width:70" size="4" onchange="fncgsel()">
					<option value="1" <% IF eNumber="1" THEN response.write "selected" %>>1차</option>
					<option value="2" <% IF eNumber="2" THEN response.write "selected" %>>2차</option>
					<option value="3" <% IF eNumber="3" THEN response.write "selected" %>>3차</option>
					<option value="4" <% IF eNumber="4" THEN response.write "selected" %>>4차</option>
				</select><br>
				
			</td>
		</tr>
		<tr>
			<td valign="top"><b>분류</b> : </td>
			<td valign="top">
				<select name="etp"  style="width:70" size="4" onchange="fncgsel()">
					<option value="1" <% IF eType="1" THEN response.write "selected" %>>끄덕끄덕</option>
					<option value="2" <% IF eType="2" THEN response.write "selected" %>>얼렁뚱땅</option>
					<option value="3" <% IF eType="3" THEN response.write "selected" %>>싱숭생숭</option>
					<option value="4" <% IF eType="4" THEN response.write "selected" %>>끼리끼리</option>
				</select>
			</td>
		</tr>
		</form>
		</table>
		<!-- // 차수 분류 선택 -->
		
	</td>
	<td align="left" valign="top">
		<!-- 단어리스트  // -->
		<table width="250" border="0" cellspacing="1" cellpadding="1" class="a" bgcolor="<%=adminColor("tablebg") %>">
		<form name="rWordFrm" method="post" action="ifr_EmoEdit.asp" target="regFrame">
		<input type="hidden" name="eno" value="<%=eNumber%>">
		<input type="hidden" name="etp" value="<%=eType%>">
		<input type="hidden" name="etlt" value="">
		<input type="hidden" name="ius" value="">
		<input type="hidden" name="mode" value="">
		<tr bgcolor="<%=adminColor("sky") %>">	
			<td align="center" width="100"><b>단어</b></td>
			<td align="center" width="65"><b>사용여부</b></td>
			<td align="center" width="45"><b>응모자</b></td>
			<td align="center" width="40"><b>순서</b></td>
		</tr>
		<% IF oWord.FResultCount>0 Then %>
			
			<% For iLp = 0 To oWord.FResultCount -1 %>
			<input type="hidden" name="awrd" value="<%= oWord.FList(iLp).EmoTitle%>">
			<tr bgcolor="#FFFFFF">	
				<td>
					<span onclick="fnEdit('<%= oWord.FList(iLp).EmoTitle%>');" style="cursor:pointer"><b><%= oWord.FList(iLp).EmoTitle%></b></span>
				</td>
				<td align="center">
					<input type="radio" value="Y" name="ius<%=iLp%>" <% IF oWord.FList(iLp).EmoUsing="Y" Then Response.write "checked"%>>Y
					<input type="radio" value="N" name="ius<%=iLp%>" <% IF oWord.FList(iLp).EmoUsing="N" Then Response.write "checked"%>>N
				</td>
				<td align="center"><input type="button" class="button" value="보기" onclick="fncomlist('<%= oWord.FList(iLp).EmoTitle%>');"></td>
				<td align="center"><input type="text" name="srtno" size="2" value="<%= oWord.FList(iLp).EmoSortNo%>"></td>
			</tr>
			<% Next %>
		<% End IF %>
			<tr height="10" bgcolor="<%= adminColor("dgray") %>">
				<td colspan="4" align="right">
					<input type="button" class="button" value="전체-Y" onclick="fnUsingupdate('Y');">
					<input type="button" class="button" value="전체-N" onclick="fnUsingupdate('N');">
					
					<input type="button" class="button" value="수정" onclick="fnarrEdit();">&nbsp;&nbsp;
					<input type="button" class="button" value="추가" onclick="fnbatIn();">
				</td>
			</tr>
		</form>
		</table>
		<!-- // 단어리스트  -->
	</td>
	<td align="left" valign="top">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="left" valign="top"> <iframe name="regFrame" id="regFrame" frameborder="0" width="400" height="400"></td>
		</tr>
		</table>
	</td>
</tr>
</table>



<% SET oWord = nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->