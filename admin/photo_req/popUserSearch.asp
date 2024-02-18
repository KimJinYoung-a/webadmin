<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/usercls.asp"-->

<%
dim i,page
dim rectDesigner
dim usingonly, research, userdiv, rect, crect, mrect, mduserid
dim catecode

dim frmName, compName, userName, socname_kr

frmName= request("frmName")
compName= request("compName")
userName	 = request("userName")

rectDesigner= request("rectDesigner")
usingonly   = request("usingonly")

research    = request("research")
userdiv     = request("userdiv")
rect        = requestCheckVar(request("rect"),60)
mduserid    = request("mduserid")
catecode    = request("catecode")
crect       = requestCheckVar(request("crect"),60)
mrect       = requestCheckVar(request("mrect"),60)
socname_kr  = requestCheckVar(request("socname_kr"),60)

page        = request("page")


if ((research="") and (usingonly="")) then usingonly="all"
if ((research="") and (userdiv="")) then userdiv="02"

if page="" then page=1

dim opartner
set opartner = new CPartnerUser
opartner.FCurrpage = page
opartner.FPageSize = 100
opartner.FRectDesignerID = rectDesigner
opartner.FrectIsUsing = usingonly
opartner.FRectDesignerDiv = userdiv
opartner.FRectMdUserID = mduserid
opartner.FRectInitial = rect
opartner.FRectCompanyname = crect
opartner.FRectManagerName = mrect
opartner.FRectCatecode = catecode
opartner.FRectSOCName  = socname_kr
opartner.GetPartnerNUserCList
%>

<script language='javascript'>
function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function selectThis(selval,selval2){
    opener.<%= frmName %>.<%= compName %>.value = selval;
	opener.<%= frmName %>.<%= userName %>.value = selval2;
    window.close();
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="rectorder" value="">
	
	<input type="hidden" name="frmName" value="<%= frmName %>">
	<input type="hidden" name="compName" value="<%= compName %>">
	<input type="hidden" name="userName" value="<%= userName%>">
	
	
	
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
				ID : <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				�̸� : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
	        </td>
	        <td align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>

</table>

<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="#BABABA"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        		<tr>
        		</tr>
        	</table>
        </td>
       	<td align="right">
       		�� <%= opartner.FtotalCount %>��
       	</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ �߰��� ��-->



<table width="500" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50%" >ID</td>
		<td width="50%" >�̸�</td>
	</tr>
	<% for i=0 to opartner.FresultCount-1 %>
	<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>','<%= opartner.FPartnerList(i).Fcompany_name %>')"><%= opartner.FPartnerList(i).FID %></a></td>
		<td>
			<%= opartner.FPartnerList(i).Fcompany_name %>
		</td>
		
	</tr>
	<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if opartner.HasPreScroll then %>
			<a href="javascript:NextPage('<%= opartner.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + opartner.StartScrollPage to opartner.FScrollCount + opartner.StartScrollPage - 1 %>
    			<% if i>opartner.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if opartner.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>

        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<%
set opartner = Nothing
%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->