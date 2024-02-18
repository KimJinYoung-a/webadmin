<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

dim userid
dim i, page
dim strIsue

userid = request("userid")



'==============================================================================
dim oTax
set oTax = new CTax
oTax.FCurrPage = 1
oTax.FPageSize = 100
'oTax.FRectsearchDiv = "Y"					'발행된 내역만
oTax.FRectsearchBilldiv = "01"				'소비자매출
oTax.FRectsearchKey = "t1.userid"

if (userid <> "") then
	oTax.FRectsearchString = userid
else
	oTax.FRectsearchString = "----"
end if

oTax.GetTaxList

%>
<script language="javascript">
function doInputData(frm) {
	opener.frm.socno.value = frm.socno.value;
	opener.frm.socname.value = frm.socname.value;
	opener.frm.ceoname.value = frm.ceoname.value;
	opener.frm.socaddr.value = frm.socaddr.value;
	opener.frm.socstatus.value = frm.socstatus.value;
	opener.frm.socevent.value = frm.socevent.value;

	opener.frm.managername.value = frm.managername.value;
	opener.frm.managerphone.value = frm.managerphone.value;
	opener.frm.managermail.value = frm.managermail.value;
}
</script>

<style type="text/css">
.Readonlybox { border:0px; }
.writebox { border:10px; background:#E6E6E6; }
</style>



<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td>

		<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
			<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td align="left">
					<b>세금계산서 발행내역</b>
				</td>
			</tr>
		</table>

	</td>
</tr>
<tr height="20">
	<td>
	</td>
</tr>
<tr>
	<td>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oTax.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="80">사업자번호</td>
		<td>상호</td>
		<td>대표자</td>
		<td>담당자</td>
		<td width="80">등록자</td>
		<td width="65">등록일</td>
		<td width="65">비고</td>
	</tr>
	<%
		for i=0 to oTax.FResultCount - 1
			'발급여부
			if oTax.FTaxList(i).FisueYn="Y" then
				strIsue = "<font color=darkblue>발급</font>"
			else
				strIsue = "<font color=darkred>미발급</font>"
			end if
	%>
	<form name="frm<%= i %>" method="Post" action="">
	<input type="hidden" name="socno" 			value="<%= oTax.FTaxList(i).FbusiNo %>">
	<input type="hidden" name="socname" 		value="<%= oTax.FTaxList(i).FbusiName %>">
	<input type="hidden" name="ceoname" 		value="<%= oTax.FTaxList(i).FbusiCEOName %>">
	<input type="hidden" name="socaddr" 		value="<%= oTax.FTaxList(i).FbusiAddr %>">
	<input type="hidden" name="socstatus" 		value="<%= oTax.FTaxList(i).FbusiType %>">
	<input type="hidden" name="socevent" 		value="<%= oTax.FTaxList(i).FbusiItem %>">
	<input type="hidden" name="managername" 	value="<%= oTax.FTaxList(i).FrepName %>">
	<input type="hidden" name="managerphone" 	value="<%= oTax.FTaxList(i).FrepTel %>">
	<input type="hidden" name="managermail" 	value="<%= oTax.FTaxList(i).FrepEmail %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oTax.FTaxList(i).FbusiNo %></td>
		<td><%= db2html(oTax.FTaxList(i).FbusiName)%></td>
		<td><%= oTax.FTaxList(i).FbusiCEOName %></td>
		<td><%= oTax.FTaxList(i).FrepName %></td>
		<td><%= oTax.FTaxList(i).Fuserid %></td>
		<td><%= FormatDate(oTax.FTaxList(i).Fregdate,"0000-00-00") %></td>
		<td><input type="button" class="button" value="입력" onClick="doInputData(frm<%= i %>)"></td>
	</tr>
	</form>
	<%
		next
	%>

</table>
	</td>
</tr>
</table>

<p>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->