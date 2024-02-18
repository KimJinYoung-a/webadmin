<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/upchemargincls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%

dim mwdiv, page, makerid, showType, brandUsingYn, catecode
makerid   = requestCheckVar(request("makerid"),32)
catecode  = requestCheckVar(request("catecode"),3)
mwdiv    = requestCheckVar(request("mwdiv"),1)
page     = requestCheckVar(request("page"),9)
showType = requestCheckVar(request("showType"),9)
brandUsingYn = requestCheckVar(request("brandUsingYn"),1)

if (page="") then page=1
if (showType="") then showType="ononly"

dim oUpcheMargin
set oUpcheMargin = new CUpcheMargin
oUpcheMargin.FPageSize = 30
oUpcheMargin.FCurrPage = page
oUpcheMargin.FRectMwDiv = mwdiv
oUpcheMargin.FRectMakerid = makerid
oUpcheMargin.FRectCateCode = catecode
oUpcheMargin.FRectbrandUsingYn = brandUsingYn

if (showType="onoff") then
    oUpcheMargin.GetUpcheTotalMarginList
else
    oUpcheMargin.GetUpcheOnlineMarginList
end if


'==============================================================================
dim i, j, k, tmp

dim S00Exists, P00Exists, Q00Exists

%>
<script language='javascript'>
function PopUpcheInfo(v){
	window.open("/admin/lib/popupcheinfo.asp?designer=" + v,"popupcheinfo","width=640 height=540");
}

function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function NextPage(page){
    document.frm.page.value=page;
    document.frm.submit();

}

function popItemList(makerid,mwdiv){
    var popUrl = "/admin/itemmaster/itemlist.asp?menupos=594&page=1&makerid=" + makerid + "&sellyn=&usingyn=Y&danjongyn=&limityn=&mwdiv=" + mwdiv + "&vatyn=&sailyn="

    var popwin = window.open(popUrl,'popItemList','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus()
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a">
		�귣��ID
		<% drawSelectBoxDesignerwithName "makerid", makerid %>
		&nbsp;
		ī�װ� : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		<select name="mwdiv">
		<option value="">��ü�⺻��������
		<option value="M" <%= ChkIIF(mwdiv="M","selected","") %> >����
		<option value="W" <%= ChkIIF(mwdiv="W","selected","") %> >��Ź
		<option value="U" <%= ChkIIF(mwdiv="U","selected","") %> >��ü
		</select>

		<select name="brandUsingYn">
		<option value="">�귣���뿩��
		<option value="Y" <%= ChkIIF(brandUsingYn="Y","selected","") %> >���
		<option value="N" <%= ChkIIF(brandUsingYn="N","selected","") %> >������
		</select>

		&nbsp;
		�˻����� :
		<input type="radio" name="showType" value="ononly" <%= chkIIF(showType="ononly","checked","") %> >�¶���
		<input type="radio" name="showType" value="onoff" <%= chkIIF(showType="onoff","checked","") %> >�¶��� + ��������
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<!-- �����ޱ� -->
<%
	Dim exlPsz, exlPg
	exlPsz = 5000
	exlPg = ceil(oUpcheMargin.FTotalCount/exlPsz)
%>
<script>
	function fnGetExcel(pg) {
		window.open("designermargin_excel.asp?page="+pg+"&makerid=<%=makerid%>&catecode=<%=catecode%>&mwdiv=<%=mwdiv%>&brandUsingYn=<%=brandUsingYn%>&showType=<%=showType%>");
	}
</script>
<div style="text-align:right; margin:10px 5px;">
	<select id="exlPage" class="select" style="vertical-align: middle;">
	<% for i=1 to exlPg %>
	<option value="<%=i%>"><%=((i-1)*exlPsz)+1%>~<%=chkIIF(i*exlPsz<oUpcheMargin.FTotalCount,i*exlPsz,oUpcheMargin.FTotalCount)%></option>
	<% next %>
	</select>
	<img src="/images/btn_excel.gif" onClick="fnGetExcel(document.getElementById('exlPage').value)" style="cursor:pointer;vertical-align: middle;" />
</div>
<!-- �˻���� -->
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
  <td colspan="25">
  Total : <%= FormatNumber(oUpcheMargin.FTotalCount,0) %>�� Page:<%= page %>/<%= oUpcheMargin.FTotalPage %>
  </td>
</tr>
<tr bgcolor="#DDDDFF">
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td colspan=6 align=center>�¶���</td>
  <% if (showType="onoff") then %>
  <td colspan=4 align=center>��������</td>
  <% end if %>
  <td></td>
  <td></td>
  <td></td>
</tr>
<tr bgcolor="#DDDDFF">
  <td align=center width="100">�귣��ID</td>
  <td align=center width="120">�귣���</td>
  <td align=center width="60">�׷��ڵ�</td>
  <td align=center width="100">��ü��</td>
  <td align=center width="100">�����å</td>
  <td align=center width="50">����<br>����</td>
  <td align=center width="50">�⺻<br>����</td>
  <td align=center width="50">����<br>��ǰ</td>
  <td align=center width="50">��Ź<br>��ǰ</td>
  <td align=center width="50">��ü<br>��ǰ</td>
  <% if (showType="onoff") then %>
  <td align=center>��ID</td>
  <td align=center>����</td>
  <td align=center>����</td>
  <td align=center>����</td>
  <% end if %>
  <td align=center width="50">��뿩��</td>
  <td align=center width="50">����</td>
  <td align=center>���</td>
</tr>
<% for i=0 to oUpcheMargin.FResultCount - 1 %>
<%
S00Exists = False
P00Exists = False
Q00Exists = False
%>
<tr bgcolor="#FFFFFF" align=center>
  <td align=left>
    <a href="javascript:PopUpcheInfo('<%= oUpcheMargin.FItemList(i).FMakerid %>')"><%= oUpcheMargin.FItemList(i).FMakerid %></a>
  </td>
  <td align=left><%= oUpcheMargin.FItemList(i).FBrandName %></td>
  <td align=center><%= oUpcheMargin.FItemList(i).FGroupID %></td>
  <td align=center><%= oUpcheMargin.FItemList(i).FCompany_name %></td>
  <td align=center><%= oUpcheMargin.FItemList(i).getOnlinedefaultDlvTypeName %></td>
 <td align=center><font color="<%= mwdivColor(oUpcheMargin.FItemList(i).FDefaultOnlineMwDiv) %>"><%= mwdivName(oUpcheMargin.FItemList(i).FDefaultOnlineMwDiv) %></font></td>
  <td align=center><%= oUpcheMargin.FItemList(i).FDefaultOnlineMargin %> %</td>
  <td align=center>
       <% if oUpcheMargin.FItemList(i).FOnlineMCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineMAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineMAvgMargin %> %</font> <br>
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','M');"><%= oUpcheMargin.FItemList(i).FOnlineMCount %> ��</a>)
       <% end if %>
  </td>
  <td align=center>
       <% if oUpcheMargin.FItemList(i).FOnlineWCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineWAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineWAvgMargin %> %</font> <br>
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','W');"><%= oUpcheMargin.FItemList(i).FOnlineWCount %> ��</a>)
       <% end if %>
  </td>
  <td align=center>
       <% if oUpcheMargin.FItemList(i).FOnlineUCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineUAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineUAvgMargin %> %</font> <br>
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','U');"><%= oUpcheMargin.FItemList(i).FOnlineUCount %> ��</a>)
       <% end if %>
  </td>
  <% if (showType="onoff") then %>
  <td >
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS000comm_cd),"","����") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS800comm_cd),"","<p>����") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS870comm_cd),"","<p>����") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS700comm_cd),"","<p>�ؿ�") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FT000comm_cd),"","<p>���̶��") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FY000comm_cd),"","<p>����") %>
  </td>
  <td >
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FS000comm_cd) %><p>
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FS800comm_cd) %><p>
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FS870comm_cd) %><p>
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FS700comm_cd) %><p>
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FT000comm_cd) %><p>
       <%= GetJungsanGubunName(oUpcheMargin.FItemList(i).FY000comm_cd) %>
  </td>
  <td >
       <%= oUpcheMargin.FItemList(i).FS000defaultmargin %><p>
       <%= oUpcheMargin.FItemList(i).FS800defaultmargin %><p>
       <%= oUpcheMargin.FItemList(i).FS870defaultmargin %><p>
       <%= oUpcheMargin.FItemList(i).FS700defaultmargin %><p>
       <%= oUpcheMargin.FItemList(i).FT000defaultmargin %><p>
       <%= oUpcheMargin.FItemList(i).FY000defaultmargin %>
  </td>
  <td >
       <%= oUpcheMargin.FItemList(i).FS000defaultsuplymargin %><p>
       <%= oUpcheMargin.FItemList(i).FS800defaultsuplymargin %><p>
       <%= oUpcheMargin.FItemList(i).FS870defaultsuplymargin %><p>
       <%= oUpcheMargin.FItemList(i).FS700defaultsuplymargin %><p>
       <%= oUpcheMargin.FItemList(i).FT000defaultsuplymargin %><p>
       <%= oUpcheMargin.FItemList(i).FY000defaultsuplymargin %>
  </td>
  <% end if %>
  <td <%= ChkIIF(oUpcheMargin.FItemList(i).FBrandUsingYn="N","bgcolor='#CCCCCC'","") %> >
    <a href="javascript:PopBrandAdminUsingChange('<%= oUpcheMargin.FItemList(i).Fmakerid %>');"><%= ChkIIF(oUpcheMargin.FItemList(i).FBrandUsingYn="Y","O","X") %></a>
  </td>

  <td <%= ChkIIF(oUpcheMargin.FItemList(i).FPartnerUsingYn="N","bgcolor='#CCCCCC'","") %> >
    <a href="javascript:PopBrandAdminUsingChange('<%= oUpcheMargin.FItemList(i).Fmakerid %>');"><%= ChkIIF(oUpcheMargin.FItemList(i).FPartnerUsingYn="Y","O","X") %></a>
  </td>
  <td align=center></td>
</tr>

<% next %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="25" align="center">
    <% if oUpcheMargin.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oUpcheMargin.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oUpcheMargin.StartScrollPage to oUpcheMargin.FScrollCount + oUpcheMargin.StartScrollPage - 1 %>
		<% if i>oUpcheMargin.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oUpcheMargin.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->