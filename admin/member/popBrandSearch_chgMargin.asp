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
dim mgnfName, evtjs

frmName= request("frmName")
compName= request("compName")
mgnfName= request("mgnfName")
evtjs= request("evtjs")

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

function qsearch(alp){
	frm.page.value="1";
	frm.rect.value=alp;
	frm.submit();
}

function selectThis(selval,selmgn,jungsangubun, companyno){
    opener.<%= frmName %>.<%= compName %>.value = selval;
    opener.<%= frmName %>.<%= mgnfName %>.value = selmgn;
    
    //상품등록시 브랜드에 따른 조건 확인을 위해 추가 2014.02.19 정윤정----------
    if(typeof(opener.<%= frmName %>.jungsangubun)=="object"){
    	opener.<%= frmName %>.jungsangubun.value = jungsangubun;
    } 
    if(typeof(opener.<%= frmName %>.companyno)=="object"){
    	opener.<%= frmName %>.companyno.value = companyno;
    }
    //--------------------------------------------------------------------------
    opener.<%=evtjs & "()"%>;
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
	<input type="hidden" name="mgnfName" value="<%= mgnfName %>">
	<input type="hidden" name="evtjs" value="<%= evtjs %>">
	
	
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	            <% if (Not C_IS_FRN_SHOP) THen %>
	        	<input type=radio name=usingonly  value="all" <% if usingonly="all" then response.write "checked" %> >전체
				<input type=radio name=usingonly  value="on" <% if usingonly="on" then response.write "checked" %> >사용함
				
				<input type=radio name=usingonly  value="off_new" <% if usingonly="off_new" then response.write "checked" %> >사용안함(신규)
				<input type=radio name=usingonly  value="off_old" <% if usingonly="off_old" then response.write "checked" %> >사용안함(한달이전)
				<br>
				담당자 : <% drawSelectBoxCoWorker "mduserid", mduserid %>
				&nbsp;
				카테고리 : <% SelectBoxBrandCategory "catecode", catecode %>
				&nbsp;
				업체구분 : <% DrawBrandGubunCombo "userdiv", userdiv %>
				&nbsp;
				<br>
				<% else %>
				<input type="hidden" name="usingonly" value="<%= usingonly %>">
				<input type="hidden" name="userdiv" value="<%= userdiv %>">
				<% end if %>
				브랜드ID : <input type="text" class="text" name="rect" value="<%= rect %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				스트리트명(한글) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
				&nbsp;&nbsp;
				회사명 : <input type="text" class="text" name="crect" value="<%= crect %>" Maxlength="32" size="16">
				<% if (Not C_IS_FRN_SHOP) THen %>
				&nbsp;&nbsp;
				담당자명 : <input type="text" class="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="16">
				<% end if %>
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
        			<TD width="8"><A href="javascript:qsearch('a')">A</A></TD>
					<TD width="8"><A href="javascript:qsearch('b')">B</A></TD>
					<TD width="8"><A href="javascript:qsearch('c')">C</A></TD>
					<TD width="8"><A href="javascript:qsearch('d')">D</A></TD>
					<TD width="8"><A href="javascript:qsearch('e')">E</A></TD>
					<TD width="8"><A href="javascript:qsearch('f')">F</A></TD>
					<TD width="8"><A href="javascript:qsearch('g')">G</A></TD>
					<TD width="8"><A href="javascript:qsearch('h')">H</A></TD>
					<TD width="8"><A href="javascript:qsearch('i')">I</A></TD>
					<TD width="8"><A href="javascript:qsearch('j')">J</A></TD>
					<TD width="8"><A href="javascript:qsearch('k')">K</A></TD>
					<TD width="8"><A href="javascript:qsearch('l')">L</A></TD>
					<TD width="8"><A href="javascript:qsearch('m')">M</A></TD>
					<TD width="8"><A href="javascript:qsearch('n')">N</A></TD>
					<TD width="8"><A href="javascript:qsearch('o')">O</A></TD>
					<TD width="8"><A href="javascript:qsearch('p')">P</A></TD>
					<TD width="8"><A href="javascript:qsearch('q')">Q</A></TD>
					<TD width="8"><A href="javascript:qsearch('r')">R</A></TD>
					<TD width="8"><A href="javascript:qsearch('s')">S</A></TD>
					<TD width="8"><A href="javascript:qsearch('t')">T</A></TD>
					<TD width="8"><A href="javascript:qsearch('u')">U</A></TD>
					<TD width="8"><A href="javascript:qsearch('v')">V</A></TD>
					<TD width="8"><A href="javascript:qsearch('w')">W</A></TD>
					<TD width="8"><A href="javascript:qsearch('x')">X</A></TD>
					<TD width="8"><A href="javascript:qsearch('y')">Y</A></TD>
					<TD width="8"><A href="javascript:qsearch('z')">Z</A></TD>
					<TD width="8"><A href="javascript:qsearch('etc')">etc</A></TD>
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
		<td >구분</td>
		<td >브랜드ID</td>
		<td >스트리트명(한글)<br>스트리트명(영문)</td>
		<td >회사명</td>
		<td >담당자</td>
	</tr>
	<% for i=0 to opartner.FresultCount-1 %>
	<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
		<td align="center"><%= opartner.FPartnerList(i).GetUserDivName %></a></td>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>','<%=opartner.FPartnerList(i).Fdefaultmargine & "," & opartner.FPartnerList(i).Fmaeipdiv & "," & opartner.FPartnerList(i).FdefaultFreeBeasongLimit & "," & opartner.FPartnerList(i).FdefaultDeliverPay & "," & opartner.FPartnerList(i).FdefaultDeliveryType %>','<%=opartner.FPartnerList(i).Fjungsan_gubun%>','<%=opartner.FPartnerList(i).Fcompany_no%>')"><%= opartner.FPartnerList(i).FID %></a></td>
		<td>
			<%= opartner.FPartnerList(i).FSocName_Kor %><br>
			<%= opartner.FPartnerList(i).FSocName %>
		</td>
		<td><a href="javascript:selectThis('<%= opartner.FPartnerList(i).FID %>','<%=opartner.FPartnerList(i).Fdefaultmargine & "," & opartner.FPartnerList(i).Fmaeipdiv & "," & opartner.FPartnerList(i).FdefaultFreeBeasongLimit & "," & opartner.FPartnerList(i).FdefaultDeliverPay & "," & opartner.FPartnerList(i).FdefaultDeliveryType %>','<%=opartner.FPartnerList(i).Fjungsan_gubun%>','<%=opartner.FPartnerList(i).Fcompany_no%>')"><%= opartner.FPartnerList(i).Fcompany_name %></a></td>
		<td align="center"><%= opartner.FPartnerList(i).Fmanager_name %></td>
		
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