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
		브랜드ID
		<% drawSelectBoxDesignerwithName "makerid", makerid %>
		&nbsp;
		카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		<select name="mwdiv">
		<option value="">업체기본마진선택
		<option value="M" <%= ChkIIF(mwdiv="M","selected","") %> >매입
		<option value="W" <%= ChkIIF(mwdiv="W","selected","") %> >위탁
		<option value="U" <%= ChkIIF(mwdiv="U","selected","") %> >업체
		</select>

		<select name="brandUsingYn">
		<option value="">브랜드사용여부
		<option value="Y" <%= ChkIIF(brandUsingYn="Y","selected","") %> >사용
		<option value="N" <%= ChkIIF(brandUsingYn="N","selected","") %> >사용안함
		</select>

		&nbsp;
		검색범위 :
		<input type="radio" name="showType" value="ononly" <%= chkIIF(showType="ononly","checked","") %> >온라인
		<input type="radio" name="showType" value="onoff" <%= chkIIF(showType="onoff","checked","") %> >온라인 + 오프라인
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<!-- 엑셀받기 -->
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
<!-- 검색결과 -->
<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
  <td colspan="25">
  Total : <%= FormatNumber(oUpcheMargin.FTotalCount,0) %>건 Page:<%= page %>/<%= oUpcheMargin.FTotalPage %>
  </td>
</tr>
<tr bgcolor="#DDDDFF">
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td colspan=6 align=center>온라인</td>
  <% if (showType="onoff") then %>
  <td colspan=4 align=center>오프라인</td>
  <% end if %>
  <td></td>
  <td></td>
  <td></td>
</tr>
<tr bgcolor="#DDDDFF">
  <td align=center width="100">브랜드ID</td>
  <td align=center width="120">브랜드명</td>
  <td align=center width="60">그룹코드</td>
  <td align=center width="100">업체명</td>
  <td align=center width="100">배송정책</td>
  <td align=center width="50">매입<br>구분</td>
  <td align=center width="50">기본<br>마진</td>
  <td align=center width="50">매입<br>상품</td>
  <td align=center width="50">위탁<br>상품</td>
  <td align=center width="50">업체<br>상품</td>
  <% if (showType="onoff") then %>
  <td align=center>샵ID</td>
  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <% end if %>
  <td align=center width="50">사용여부</td>
  <td align=center width="50">어드민</td>
  <td align=center>비고</td>
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
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','M');"><%= oUpcheMargin.FItemList(i).FOnlineMCount %> 건</a>)
       <% end if %>
  </td>
  <td align=center>
       <% if oUpcheMargin.FItemList(i).FOnlineWCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineWAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineWAvgMargin %> %</font> <br>
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','W');"><%= oUpcheMargin.FItemList(i).FOnlineWCount %> 건</a>)
       <% end if %>
  </td>
  <td align=center>
       <% if oUpcheMargin.FItemList(i).FOnlineUCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineUAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineUAvgMargin %> %</font> <br>
       (<a href="javascript:popItemList('<%= oUpcheMargin.FItemList(i).FMakerid %>','U');"><%= oUpcheMargin.FItemList(i).FOnlineUCount %> 건</a>)
       <% end if %>
  </td>
  <% if (showType="onoff") then %>
  <td >
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS000comm_cd),"","직영") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS800comm_cd),"","<p>가맹") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS870comm_cd),"","<p>도매") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FS700comm_cd),"","<p>해외") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FT000comm_cd),"","<p>아이띵소") %>
        <%= CHKIIF(isNULL(oUpcheMargin.FItemList(i).FY000comm_cd),"","<p>대행") %>
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