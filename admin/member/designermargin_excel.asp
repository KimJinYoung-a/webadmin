<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/upchemargincls.asp"-->
<%

dim mwdiv, page, pagesize, makerid, showType, brandUsingYn, catecode
makerid   = requestCheckVar(request("makerid"),32)
catecode  = requestCheckVar(request("catecode"),3)
mwdiv    = requestCheckVar(request("mwdiv"),1)
page     = requestCheckVar(request("page"),9)
showType = requestCheckVar(request("showType"),9)
brandUsingYn = requestCheckVar(request("brandUsingYn"),1)

pagesize = 5000
if (page="") then page=1
if (showType="") then showType="ononly"

dim oUpcheMargin
set oUpcheMargin = new CUpcheMargin
oUpcheMargin.FPageSize = pagesize
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

'Excel Header
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=brandMargin_page" & CStr(page) & ".xls"
Response.CacheControl = "public"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="application/vnd.ms-excel;charset=euc-kr">
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table border="0" cellspacing="1" cellpadding="2">
<tr>
  <td colspan="<%=chkIIF(showType="onoff","16","12")%>">
  Total : <%= FormatNumber(oUpcheMargin.FTotalCount,0) %>�� Page:<%= page %>/<%= oUpcheMargin.FTotalPage %>
  </td>
</tr>
<tr bgcolor="#DDDDFF">
  <td align="center" rowspan="2">�귣��ID</td>
  <td align="center" rowspan="2">�귣���</td>
  <td align="center" rowspan="2">�׷��ڵ�</td>
  <td align="center" rowspan="2">��ü��</td>
  <td colspan="6" align="center">�¶���</td>
  <% if (showType="onoff") then %>
  <td colspan="4" align="center">��������</td>
  <% end if %>
  <td align="center" rowspan="2">��뿩��</td>
  <td align="center" rowspan="2">����</td>
</tr>
<tr bgcolor="#DDDDFF">
  <td align="center">�����å</td>
  <td align="center">���Ա���</td>
  <td align="center">�⺻����</td>
  <td align="center">���Ի�ǰ</td>
  <td align="center">��Ź��ǰ</td>
  <td align="center">��ü��ǰ</td>
  <% if (showType="onoff") then %>
  <td align="center">��ID</td>
  <td align="center">����</td>
  <td align="center">����</td>
  <td align="center">����</td>
  <% end if %>
</tr>
<% for i=0 to oUpcheMargin.FResultCount - 1 %>
<tr>
  <td align="left"><%= oUpcheMargin.FItemList(i).FMakerid %></td>
  <td align="left"><%= oUpcheMargin.FItemList(i).FBrandName %></td>
  <td align="center"><%= oUpcheMargin.FItemList(i).FGroupID %></td>
  <td align="center"><%= oUpcheMargin.FItemList(i).FCompany_name %></td>
  <td align="center"><%= oUpcheMargin.FItemList(i).getOnlinedefaultDlvTypeName %></td>
 <td align="center"><font color="<%= mwdivColor(oUpcheMargin.FItemList(i).FDefaultOnlineMwDiv) %>"><%= mwdivName(oUpcheMargin.FItemList(i).FDefaultOnlineMwDiv) %></font></td>
  <td align="center"><%= oUpcheMargin.FItemList(i).FDefaultOnlineMargin %> %</td>
  <td align="center">
       <% if oUpcheMargin.FItemList(i).FOnlineMCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineMAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineMAvgMargin %> %</font>
       (<%= oUpcheMargin.FItemList(i).FOnlineMCount %> ��)
       <% end if %>
  </td>
  <td align="center">
       <% if oUpcheMargin.FItemList(i).FOnlineWCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineWAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineWAvgMargin %> %</font>
       (<%= oUpcheMargin.FItemList(i).FOnlineWCount %> ��)
       <% end if %>
  </td>
  <td align="center">
       <% if oUpcheMargin.FItemList(i).FOnlineUCount>0 then %>
       <font color="<%= ChkIIF(oUpcheMargin.FItemList(i).FDefaultOnlineMargin<>oUpcheMargin.FItemList(i).FOnlineUAvgMargin,"#CC0000","#000000") %>"><%= oUpcheMargin.FItemList(i).FOnlineUAvgMargin %> %</font>
       (<%= oUpcheMargin.FItemList(i).FOnlineUCount %> ��)
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
  <td <%= ChkIIF(oUpcheMargin.FItemList(i).FBrandUsingYn="N","bgcolor='#CCCCCC'","") %> ><%= ChkIIF(oUpcheMargin.FItemList(i).FBrandUsingYn="Y","O","X") %></td>
  <td <%= ChkIIF(oUpcheMargin.FItemList(i).FPartnerUsingYn="N","bgcolor='#CCCCCC'","") %> ><%= ChkIIF(oUpcheMargin.FItemList(i).FPartnerUsingYn="Y","O","X") %></td>
</tr>

<% next %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->