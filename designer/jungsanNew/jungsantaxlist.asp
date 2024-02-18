<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, page
makerid = session("ssBctID")
page    = requestCheckvar(request("page"),10)

if page="" then page=1

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 30
ojungsanTax.FCurrPage = page
ojungsanTax.FRectMakerid = makerid
ojungsanTax.getJungsanTaxListByMakerid2GroupID

dim i
dim commCnt : commCnt =0
%>
<script language='javascript'>
function goMonthJungsan(yyyy,mm,jid){
    location.href='/designer/jungsanNew/monthjungsan.asp?menupos=1647&yyyy1='+yyyy+'&mm1='+mm;
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("/designer/jungsan/red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=800,height=700,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}



function NextPage(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}
</script>
<form name="frm" method="get">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">

<tr align="center" bgcolor="#FFFFFF">
    <td colspan="14" align="right">�� <%= FormatNumber(ojungsanTax.FTotalCount,0) %> �� <%=page%>/<%=ojungsanTax.FTotalPage%> page</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60" >�����</td>
    <td width="60" >����ó</td>
    <td width="70" >������</td>
    <td width="100" >��꼭����</td>
    <td width="80" >�귣��ID</td>
    <td width="120" >����</td>
    <td width="80" >������</td>
    <td width="80" >���ް���</td>
    <td width="80" >�ΰ���</td>
    <td width="80" >�հ�</td>
    <td width="90" >�������</td>
    <td width="70">���곻��</td>
    <td width="60">������<br>(������)</td>
    <td >���</td>

</tr>
<% for i=0 to ojungsanTax.FResultCount-1 %>
<%
if (ojungsanTax.FItemList(i).IsCommissionTax) then
    commCnt=commCnt+1
end if
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%=ojungsanTax.FItemList(i).FYYYYMM%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxJungsanGubun%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxTypeStrUpcheView%></td>
    <td><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td><%= ojungsanTax.FItemList(i).Ftaxregdate %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSuply,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxVat,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSum,0) %></td>
    <td><%= ojungsanTax.FItemList(i).GetTaxEvalStateName %></td>
    <td><img src="/images/icon_search.jpg" onClick="goMonthJungsan('<%=Left(ojungsanTax.FItemList(i).FYYYYMM,4)%>','<%=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>','<%=ojungsanTax.FItemList(i).Fid%>');" style="cursor:pointer"></td>
    <td><%= ojungsanTax.FItemList(i).getTaxEvalStyleStr %></td>
    <td>
        <% if ojungsanTax.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTax.FItemList(i).Fneotaxno %>');">���
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="14" align="center">
        <% if ojungsanTax.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojungsanTax.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojungsanTax.StartScrollPage to ojungsanTax.FScrollCount + ojungsanTax.StartScrollPage - 1 %>
			<% if i>ojungsanTax.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojungsanTax.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
</table>
<%
set ojungsanTax = Nothing
%>
<script language='javascript'>
<% if (commCnt>0 and page=1) then %>
alert('2014�� 1�� ������� ������ ����п� ���ؼ���\n\n�ٹ����ٿ��� ��꼭�� ���� �Ͽ���\n\n�̼��� ���� ���� ���� �������� ���� �ֽñ� �ٶ��ϴ�.');
<% end if %>
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->