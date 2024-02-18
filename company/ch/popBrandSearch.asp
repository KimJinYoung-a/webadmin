<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim i,page
dim rectDesigner
dim usingonly, research, userdiv, rect, crect, mrect, mduserid
dim catecode

dim frmName, compName, socname_kr

frmName= request("frmName")
compName= request("compName")

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

function selectThis(selval){
    opener.<%= frmName %>.<%= compName %>.value = selval;
    window.close();
}
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="rectorder" value="">
	
	<input type="hidden" name="frmName" value="<%= frmName %>">
	<input type="hidden" name="compName" value="<%= compName %>">
	
	
	
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td><input type="hidden" name="usingonly"  value="all">
				카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>
				&nbsp;
				<input type="hidden" name="usingonly" value="<%= usingonly %>">
				<input type="hidden" name="userdiv" value="<%= userdiv %>">
				브랜드ID : <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				스트리트명(한글) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
	        </td>
	        <td align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>

</table>

<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="#BABABA"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
        		<tr>
        			<TD width="8"><A href="?rect=a&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">A</A></TD>
					<TD width="8"><A href="?rect=b&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">B</A></TD>
					<TD width="8"><A href="?rect=c&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">C</A></TD>
					<TD width="8"><A href="?rect=d&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">D</A></TD>
					<TD width="8"><A href="?rect=e&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">E</A></TD>
					<TD width="8"><A href="?rect=f&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">F</A></TD>
					<TD width="8"><A href="?rect=g&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">G</A></TD>
					<TD width="8"><A href="?rect=h&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">H</A></TD>
					<TD width="8"><A href="?rect=i&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">I</A></TD>
					<TD width="8"><A href="?rect=j&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">J</A></TD>
					<TD width="8"><A href="?rect=k&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">K</A></TD>
					<TD width="8"><A href="?rect=l&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">L</A></TD>
					<TD width="8"><A href="?rect=m&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">M</A></TD>
					<TD width="8"><A href="?rect=n&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">N</A></TD>
					<TD width="8"><A href="?rect=o&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">O</A></TD>
					<TD width="8"><A href="?rect=p&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">P</A></TD>
					<TD width="8"><A href="?rect=q&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">Q</A></TD>
					<TD width="8"><A href="?rect=r&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">R</A></TD>
					<TD width="8"><A href="?rect=s&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">S</A></TD>
					<TD width="8"><A href="?rect=t&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">T</A></TD>
					<TD width="8"><A href="?rect=u&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">U</A></TD>
					<TD width="8"><A href="?rect=v&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">V</A></TD>
					<TD width="8"><A href="?rect=w&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">W</A></TD>
					<TD width="8"><A href="?rect=x&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">X</A></TD>
					<TD width="8"><A href="?rect=y&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">Y</A></TD>
					<TD width="8"><A href="?rect=z&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">Z</A></TD>
					<TD width="8"><A href="?rect=etc&frmName=<%= frmName %>&compName=<%= compName %>&usingonly=<%= usingonly %>&userdiv=<%= userdiv %>">etc</A></TD>
					<TD></TD>
        		</tr>
        	</table>
        </td>
       	<td align="right">
       		총 <%= opartner.FtotalCount %>건
       	</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 중간바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >브랜드ID</td>
		<td >스트리트명(한글)<br>스트리트명(영문)</td>
	</tr>
	<% for i=0 to opartner.FresultCount-1 %>
	<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>')"><%= opartner.FPartnerList(i).FID %></a></td>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>')">
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
			</a>
		</td>
		
	</tr>
	<% next %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
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
<!-- 표 하단바 끝-->

<%
set opartner = Nothing
%>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->