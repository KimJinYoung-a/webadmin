<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/inspectstockcls.asp"-->
<%
if not (C_ADMIN_AUTH) then
	response.write "<script>alert('권한이 없습니다.');</script>"
	dbget.close()	:	response.End
end if


dim code, makerid
code = request("code")
makerid = request("makerid")

dim oipchuldetail
set oipchuldetail = new CInspectStock
oipchuldetail.FRectIpchulCode = code
oipchuldetail.FRectMakerid = makerid
oipchuldetail.GetIpChulDetail

dim i, totitemno
dim totbaljuitemno, totrealitemno
%>
<script language='javascript'>
function EditItemCodeName(frm){
	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">코드</td>
      <td width="60">브랜드</td>
      <td width="20">구분</td>
      <td width="40">상품코드</td>
      <td width="40">옵션코드</td>
      <td>아이템명</td>
      <td>옵션</td>
      <td width="50">소비자가</td>
      <td width="50">매입가</td>
      <td width="50">공급가</td>
      <td width="30">수량</td>
      <td width="30">수정</td>
<!--      <td width="30">삭제</td>   -->
    </tr>
    <% for i=0 to oipchuldetail.FResultCount-1 %>
    <%
    totitemno = totitemno + oipchuldetail.FItemList(i).FItemNo
    %>
    <form name="frmold_<%= i %>" method="post" action="editadminipchuledit.asp">
    <input type=hidden name="mode" value="editipchuldetailwithjungsan">
    <input type=hidden name="code" value="<%= code %>">
    <input type=hidden name="ipchulflag" value="<%= oipchuldetail.FItemList(i).Fipchulflag %>">
    <input type=hidden name="ipchuldetailid" value="<%= oipchuldetail.FItemList(i).Fid %>">
    <input type=hidden name="itemgubun" value="<%= oipchuldetail.FItemList(i).FIItemgubun %>">
    <input type=hidden name="itemid" value="<%= oipchuldetail.FItemList(i).FItemID %>">
    <input type=hidden name="itemoption" value="<%= oipchuldetail.FItemList(i).FItemOption %>">
    <tr bgcolor="#FFFFFF">
      <td ><%= oipchuldetail.FItemList(i).FCode %></td>
      <td ><%= oipchuldetail.FItemList(i).FIMakerid %></td>
      <td ><input type=text name="newitemgubun" value="<%= oipchuldetail.FItemList(i).FIItemgubun %>" size=2 maxlength=2 style="border:1px #999999 solid; text-align=left"></td>
      <td ><input type=text name="newitemid" value="<%= oipchuldetail.FItemList(i).FItemID %>" size=9 maxlength=9 style="border:1px #999999 solid; text-align=left"></td>
      <td ><input type=text name="newitemoption" value="<%= oipchuldetail.FItemList(i).FItemOption %>" size=4 maxlength=4 style="border:1px #999999 solid; text-align=left"></td>
      <td >
      	<input type=text name="newitemname" value="<%= oipchuldetail.FItemList(i).FIItemName %>" size=32 maxlength=64 style="border:1px #999999 solid; text-align=left">
      </td>
      <td >
      	<input type=text name="newitemoptionname" value="<%= oipchuldetail.FItemList(i).FIItemOptionName %>" size=16 maxlength=64 style="border:1px #999999 solid; text-align=left">
      </td>
      <td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).FSellCash,0) %></td>
      <td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).FbuyCash,0) %></td>
      <td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).FsuplyCash,0) %></td>
      <td align="center">
      	<input type=text name="newitemno" value="<%= oipchuldetail.FItemList(i).FItemNo %>" size=4 maxlength=5 style="border:1px #999999 solid; text-align=center" readonly >
      </td>
      <td align="center"><input type=button value="수정" onclick="EditItemCodeName(frmold_<%= i %>)"></td>
<!--      <td align="center"><input type=button value="삭제" onclick="DELItemCodeName(frmold_<%= i %>)"></td> -->
    </tr>
    </form>
	<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align=center>Total</td>
	  <td colspan=9></td>
	  <td align=center><%= FormatNumber(totitemno,0) %></td>
	  <td ></td>
<!--	  <td ></td>	-->
	</tr>
</table>
<%

dim oordersheetdetail
set oordersheetdetail = new CInspectStock
oordersheetdetail.FRectIpchulCode = code
oordersheetdetail.FRectMakerid = makerid
oordersheetdetail.GetOrderSheetDetail

%>
<br><br>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">코드</td>
      <td width="60">브랜드</td>
      <td width="20">구분</td>
      <td width="40">상품코드</td>
      <td width="40">옵션코드</td>
      <td>아이템명</td>
      <td>옵션</td>
      <td width="50">소비자가</td>
      <td width="50">매입가</td>
      <td width="50">공급가</td>
      <td width="30">주문량</td>
      <td width="30">수량</td>
      <td width="30">수정</td>
<!--      <td width="30">삭제</td>  -->
    </tr>
    <% for i=0 to oordersheetdetail.FResultCount-1 %>
    <%
    totbaljuitemno = totbaljuitemno + oordersheetdetail.FItemList(i).FBaljuItemNo
    totrealitemno = totrealitemno + oordersheetdetail.FItemList(i).FRealItemNo
    %>
    <form name="frmipchul_<%= i %>" method="post" action="editadminipchuledit.asp">
    <input type=hidden name="mode" value="editordersheetdetail">
    <input type=hidden name="code" value="<%= code %>">
    <input type=hidden name="sheetdetailid" value="<%= oordersheetdetail.FItemList(i).Fidx %>">
    <input type=hidden name="itemgubun" value="<%= oordersheetdetail.FItemList(i).FItemgubun %>">
    <input type=hidden name="itemid" value="<%= oordersheetdetail.FItemList(i).FItemID %>">
    <input type=hidden name="itemoption" value="<%= oordersheetdetail.FItemList(i).FItemOption %>">
    <tr bgcolor="#FFFFFF">
      <td ><%= oordersheetdetail.FItemList(i).FBaljucode %></td>
      <td ><%= oordersheetdetail.FItemList(i).FMakerid %></td>
      <td ><input type=text name="newitemgubun" value="<%= oordersheetdetail.FItemList(i).FItemgubun %>" size=2 maxlength=2 style="border:1px #999999 solid; text-align=left"></td>
      <td ><input type=text name="newitemid" value="<%= oordersheetdetail.FItemList(i).FItemID %>" size=9 maxlength=9 style="border:1px #999999 solid; text-align=left"></td>
      <td ><input type=text name="newitemoption" value="<%= oordersheetdetail.FItemList(i).FItemOption %>" size=4 maxlength=4 style="border:1px #999999 solid; text-align=left"></td>
      <td >
      	<input type=text name="newitemname" value="<%= oordersheetdetail.FItemList(i).FItemName %>" size=32 maxlength=64 style="border:1px #999999 solid; text-align=left">
      </td>
      <td >
      	<input type=text name="newitemoptionname" value="<%= oordersheetdetail.FItemList(i).FItemOptionName %>" size=16 maxlength=64 style="border:1px #999999 solid; text-align=left">
      </td>
      <td align="right"><%= FormatNumber(oordersheetdetail.FItemList(i).FSellCash,0) %></td>
      <td align="right"><%= FormatNumber(oordersheetdetail.FItemList(i).FbuyCash,0) %></td>
      <td align="right"><%= FormatNumber(oordersheetdetail.FItemList(i).FsuplyCash,0) %></td>
      <td align="center">
      	<input type=text name="newbaljuitemno" value="<%= oordersheetdetail.FItemList(i).Fbaljuitemno %>" size=4 maxlength=5 style="border:1px #999999 solid; text-align=center" readonly>
      </td>
      <td align="center">
      	<input type=text name="newrealitemno" value="<%= oordersheetdetail.FItemList(i).Frealitemno %>" size=4 maxlength=5 style="border:1px #999999 solid; text-align=center" readonly>
      </td>
      <td align="center"><input type=button value="수정" onclick="EditItemCodeName(frmipchul_<%= i %>)"></td>
<!--      <td align="center"><input type=button value="삭제" onclick="DELItemCodeName(frmipchul_<%= i %>)"></td>  -->
    </tr>
    </form>
	<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align=center>Total</td>
	  <td colspan=9></td>
	  <td align=center><%= FormatNumber(totbaljuitemno,0) %></td>
	  <td align=center><%= FormatNumber(totrealitemno,0) %></td>
	  <td ></td>
<!--	  <td ></td>	-->
	</tr>
</table>

<%
set oipchuldetail = Nothing
set oordersheetdetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->