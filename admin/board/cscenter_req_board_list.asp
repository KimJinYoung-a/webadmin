<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%

dim i, j,ix
dim page,gubun, onlymifinish
dim research, searchkey,catevalue
dim ipjumYN
page = request("pg")
gubun = request("gubun")
onlymifinish = request("onlymifinish")
research = request("research")
searchkey = request("searchkey")
catevalue=request("catevalue")
ipjumYN=request("ipjumYN")
if research="" and onlymifinish="" then onlymifinish="on"

if gubun="" then gubun="01"

if (page = "") then page = "1"

'==============================================================================
'��ü���Խ���
dim companyrequest
set companyrequest = New CCompanyRequest

companyrequest.PageSize = 20
companyrequest.CurrPage = CInt(page)
companyrequest.ScrollCount = 10
companyrequest.FReqcd=gubun
companyrequest.FOnlyNotFinish = onlymifinish
companyrequest.FRectSearchKey = searchkey
companyrequest.FRectCatevalue = catevalue
companyrequest.FipjumYN = ipjumYN
companyrequest.list

%>
<script>
function delitem(id){
	
	if (confirm("�����Ͻðڽ��ϱ�?.") ==true)
		frmdel.mode.value="del";
		frmdel.id.value=id;
		frmdel.submit();
}
function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="cscenter_req_board_list.asp";
	frm.submit();
}

function ViewPage(id){
	frm.id.value=id;
	frm.pg.value=<%=page%>;
	frm.research.value="<%=research %>";
	frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="cscenter_req_board_view.asp";
	frm.submit();
}

function changecontent() {}
</script>

<table width="790" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="id" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="pg" value="">
	<tr>
		<td class="a" >
		<input type="radio" name="gubun" value="01" <% if gubun="01" then response.write "checked" %> >�����Ƿڼ�
		<input type="radio" name="gubun" value="02" <% if gubun="02" then response.write "checked" %> >������޼�
		<input type="radio" name="gubun" value="03" <% if gubun="03" then response.write "checked" %> >Ư����ǰ�Ƿ�
		<input type="radio" name="gubun" value="04" <% if gubun="04" then response.write "checked" %> >��õ��ǰ�Ƿ�
		&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> >ó���ȵȸ��
		<br>
		ī�װ���
		<% call DrawSelectBoxCategoryLarge("catevalue",catevalue) %>
		�Ϸᱸ��
		<select name="ipjumYN" class="a">
			<option value="">��ü</option>
			<option value="Y" <% if ipjumYN="Y" then response.write "selected" %>>�����Ϸ�</option>
			<option value="N" <% if ipjumYN="N" then response.write "selected" %>>�̿Ϸ�</option>
		</select>
		��ü�� <input type="text" name="searchkey" value="<%= searchkey %>">
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="790" cellspacing="1" class="a" bgcolor=#3d3d3d>
  <tr bgcolor="#DDDDFF">
    <td width="80" align="center">��û��</td>
    <td width="370" align="center">����</td>
    <td width="80" align="center">ó����</td>
    <td width="60" align="center">��������</td>
    <td width="120" align="center">ī�װ�����</td>
    <td width="60" align="center">�亯����</td>
  </tr>
<% for i = 0 to (companyrequest.ResultCount - 1) %>

<tr bgcolor="#FFFFFF">
    <td align="center"><%= FormatDate(companyrequest.results(i).regdate, "0000-00-00") %></td>
    <td><a href="javascript:ViewPage(<%= companyrequest.results(i).id %>);">[<%= companyrequest.code2name(companyrequest.results(i).reqcd) %>] <%= companyrequest.results(i).companyname %></a></td>
    <td align="center">
        <% if (IsNull(companyrequest.results(i).finishdate) = true) then %>
      <font color="red">�̿Ϸ�</font>
        <% else %>
      <%= FormatDate(companyrequest.results(i).finishdate, "0000-00-00") %>
        <% end if %>
    </td>
    <td align="center">
    	<%if companyrequest.results(i).ipjumYN="Y" then response.write "�����Ϸ�" %>
    	<%if companyrequest.results(i).ipjumYN="N" then response.write "N" %>
    	</td>
  	<td align="center"><%= GetCategoryName(companyrequest.results(i).categubun) %></td>
  	<td align="center">
  		<% if companyrequest.commentcheck(companyrequest.results(i).replycomment)="Y" then %>
  		Y
  		<% else %>
  		<font color="red">N</font>
  		<% end if %>
  	</td>
<!--  	<td width="30"><input type="button" onclick="javascript:delitem('<%= companyrequest.results(i).id %>');" value="����"></td> -->
  </tr>

<% next %>
</table>
<table width="790" cellspacing="1" class="a" bgcolor=#3d3d3d>
  <tr bgcolor="#FFFFFF">
    <td align="center">
		<% if companyrequest.HasPreScroll then %>
			<a href="javascript:MovePage(<%= companyrequest.StartScrollPage-1 %>);">[prev]</a>
		<% else %>
			[prev]
		<% end if %>

		<% for ix=0 + companyrequest.StartScrollPage to companyrequest.ScrollCount + companyrequest.StartScrollPage - 1 %>
			<% if ix>companyrequest.Totalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
					<font color="red">[<%= ix%>]</font>
				<% else %>
					<a href="javascript:MovePage(<%=ix%>);">[<%= ix %>]</a>
				<% end if %>
		<% next %>

		<% if companyrequest.HasNextScroll then %>
			<a href="javascript:MovePage(<%=ix%>);">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
  </tr>
</table>

<form name="frmdel" method="get" action="cscenter_req_board_act.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="id" value="">
<input type="hidden" name="page" value="<%=page%>">
</form>
<br><br>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->