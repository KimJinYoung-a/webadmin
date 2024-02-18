<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->

<%
dim makerid, yyyy1, mm1, page, jungsan_date, taxtype
'''dim taxdate , ck_taxdate
dim txdayorjday

makerid      = RequestCheckVar(request("makerid"),32)
yyyy1        = RequestCheckVar(request("yyyy1"),4)
mm1          = RequestCheckVar(request("mm1"),2)
taxtype      = RequestCheckVar(request("taxtype"),10)
jungsan_date = RequestCheckVar(request("jungsan_date"),10)
''ck_taxdate   = RequestCheckVar(request("ck_taxdate"),10)
''taxdate      = RequestCheckVar(request("taxdate"),10)
txdayorjday  = RequestCheckVar(request("txdayorjday"),10)
dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

if (txdayorjday="") then txdayorjday="T"

dim jungsanlist
set jungsanlist = new CUpcheJungsan
jungsanlist.FRectYYYYMM      = yyyy1 + "-" + mm1
jungsanlist.FRectDesigner    = makerid
jungsanlist.FRectJungsanDate = jungsan_date
jungsanlist.FRectfinishflag  = "NFixInclude"

''jungsanlist.FRectOnlyIncludeWonChon = "on"
jungsanlist.FRectTaxType     = taxtype

if (txdayorjday="T") then
    jungsanlist.FRectYYYYMM = ""
    jungsanlist.FRectTaxYYYYMM = yyyy1 + "-" + mm1
else
    jungsanlist.FRectYYYYMM = yyyy1 + "-" + mm1
    jungsanlist.FRectTaxYYYYMM = ""
end if

if (makerid<>"") then
    jungsanlist.FRectYYYYMM = ""
    jungsanlist.FRectTaxYYYYMM = ""
end if
jungsanlist.FRectOnlyIncludeWonChon="on"
jungsanlist.JungsanFixedList ''LecJungsanMasterList


dim i
dim realsellmargin
%>
<script language='javascript'>
function checkComp(comp){
    comp.form.yyyy1.disabled = comp.checked;
    comp.form.mm1.disabled = comp.checked;
}

function research(frm,makerid){
    
    frm.makerid.value = makerid;
    frm.submit();
}

function PopDetail(v){
    var popwin = window.open('popjungsandetail.asp?id=' + v , 'popjungsandetail','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function popSimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=500,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    <select name="txdayorjday">
		    <option value="T" <%= CHKIIF(txdayorjday="T","selected","") %> >������
		    <option value="J" <%= CHKIIF(txdayorjday="J","selected","") %> >��������
		    </select>
			 : <% DrawYMBox yyyy1,mm1 %>
			&nbsp;
			�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<!--
	     	�������� : 
            <select class="select" name="taxtype">
	            <option value="" <% if taxtype="" then response.write "selected" %> >����
	            <option value="01" <% if taxtype="01" then response.write "selected" %> >����
	            <option value="02" <% if taxtype="02" then response.write "selected" %> >�鼼
	            <option value="03" <% if taxtype="03" then response.write "selected" %> >����
            </select>
         
	     	&nbsp;
	     	������ : 
            <select class="select" name="jungsan_date">
	            <option value="" <% if jungsan_date="" then response.write "selected" %> >����
	            <option value="15��" <% if jungsan_date="15��" then response.write "selected" %> >15��
	            <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
	            <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
	            <option value="NULL" <% if jungsan_date="NULL" then response.write "selected" %> >������
            </select>
            &nbsp;
        
            <input type="checkbox" name="ck_taxdate" onclick="checkComp(this);" >��꼭������
            <input type="text" class="text" name="taxdate" value="" size=10 readonly ><a href="javascript:calendarOpen(frm.taxdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> 
             -->
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= FormatNumber(jungsanlist.FTotalCount,0) %></b>
			&nbsp;
<!--			
			������ : <b><%= page %> / <%= jungsanlist.FTotalPage %></b>
			&nbsp;
-->

		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="45">�����</td>
		<td width="80">������</td>
		<td width="25">����</td>
		<td width="25">����</td>
		<td width="90">�귣��ID</td>
		<td width="60">��ǥ�ڸ�</td>
		<td width="90">�ֹ�/����ڹ�ȣ</td>
		<td>�ּ�</td>

		<td width="60">�Ѹ����</td>
		<td width="40">����<br>����</td>
		<td width="60">�������</td>
		<td width="50">��õ¡��<br>����</td>
		<td width="60">�Աݱݾ�</td>
		<td width="65">����</a></td>
	    <% if (FALSE) then %><td width="65">�Ա���</td><% end if %>
    </tr>
    <% if jungsanlist.FResultCount>0 then %>
    <% for i=0 to jungsanlist.FResultCount-1 %>
    <%
    dim osum,ipsum
	osum = osum + fix(jungsanlist.FItemList(i).GetTotalSuplycash)
	ipsum = ipsum + jungsanlist.FItemList(i).GetTotalWithHoldingJungSanSum
	%>
    <% 
        'if (jungsanlist.FItemList(i).Ftot_orgsellprice<>0) then
        '    orgsellmargin = CLng((jungsanlist.FItemList(i).Ftot_orgsellprice-jungsanlist.FItemList(i).Ftot_jungsanprice)/jungsanlist.FItemList(i).Ftot_orgsellprice*100*100)/100 
        'else
        '    orgsellmargin = 0
        'end if
        
        if (jungsanlist.FItemList(i).GetTotalSellcash<>0) then
            realsellmargin = CLng((jungsanlist.FItemList(i).GetTotalSellcash-jungsanlist.FItemList(i).GetTotalSuplycash)/jungsanlist.FItemList(i).GetTotalSellcash*100*100)/100 
        else
            realsellmargin = 0
            ''realsellmargin = CLng((jungsanlist.FItemList(i).GetTotalSellcash-jungsanlist.FItemList(i).GetTotalSuplycash)/jungsanlist.FItemList(i).GetTotalSellcash*100*100)/100  
        end if
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td ><%= jungsanlist.FItemList(i).FYYYYMM %></td>
      	<td ><%= jungsanlist.FItemList(i).Ftaxregdate %></td>
      	<td ><%= jungsanlist.FItemList(i).Fdifferencekey %></td>
      	<td ><font color="<%= jungsanlist.FItemList(i).GetTaxtypeNameColor %>"><%= jungsanlist.FItemList(i).GetSimpleTaxtypeName %></font></td>
      	<td align="left"><a href="javascript:popSimpleBrandInfo('<%= jungsanlist.FItemList(i).Fdesignerid %>')"><%= jungsanlist.FItemList(i).Fdesignerid %></a></td>
    	<td><a href="javascript:PopUpcheBrandInfoEdit('<%= jungsanlist.FItemList(i).Fdesignerid %>')"><%= jungsanlist.FItemList(i).Fceoname %></a></td>
    	<td>
    	    <% if Len(jungsanlist.FItemList(i).Fcompany_no)=12 then %>
    	    <%= jungsanlist.FItemList(i).Fcompany_no %>
    	    <% else %>
    	    
        	    <% if C_ADMIN_AUTH or (C_MngPowerUser) or (session("ssBctID")="jennygo" or session("ssBctID")="aimcta") then %>
        	        <% if (session("ssBctID")="jennygo" or session("ssBctID")="aimcta" or session("ssBctID")="icommang") then %>
        	            <%=getDecSOCNobyGroupID(jungsanlist.FItemList(i).Fgroupid,jungsanlist.FItemList(i).Fcompany_no)%>
        	        <% else %>
        	            <%=  jungsanlist.FItemList(i).Fcompany_no %>
        	        <% end if %>
        	    <% else %>
        	        <%=  Left(jungsanlist.FItemList(i).Fcompany_no,7) %>*******
        	    <% end if %>
    	    <% end if %>
    	
    	</td>
    	<td align="left"><%= jungsanlist.FItemList(i).Fcompany_address %><br><%= jungsanlist.FItemList(i).Fcompany_address2 %></td>

        <td align="right"><%= FormatNumber(jungsanlist.FItemList(i).GetTotalSellcash,0) %></td>
      	<td align="center"><%= realsellmargin %> %</td>
        <td align="right"><b><%= FormatNumber(jungsanlist.FItemList(i).GetTotalSuplycash,0) %></b></td>
        <% if jungsanlist.FItemList(i).Ftaxtype="03" then %>
		<td align="right"><%= FormatNumber(jungsanlist.FItemList(i).GetTotalSuplycash-jungsanlist.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
		<td align="right"><%= FormatNumber(jungsanlist.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
		<% else %>
		<td align="right">0]</td>
		<td align="right"><%= FormatNumber(jungsanlist.FItemList(i).GetTotalSuplycash,0) %></td>
		<% end if %>
      	<td><font color="<%= jungsanlist.FItemList(i).GetStateColor %>"><%= jungsanlist.FItemList(i).GetStateName %></font></td>
    <% if (FALSE) then %><td><%= jungsanlist.FItemList(i).Fipkumdate %></td><% end if %>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="10"></td>
		<td align="right"><b><%= FormatNumber(osum,0) %></b></td>
		<td align="right"><%= FormatNumber(osum-ipsum,0) %></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td  ></td>
		
	</tr>
    <% else %>
    <tr bgcolor="#FFFFFF">
		<td colspan=20 align="center">[ �˻������ �����ϴ�. ]</td>
	</tr>
	<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<!--
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<% if jungsanlist.HasPreScroll then %>
				<a href="javascript:NextPage('<%= jungsanlist.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + jungsanlist.StarScrollPage to jungsanlist.FScrollCount + jungsanlist.StarScrollPage - 1 %>
				<% if i>jungsanlist.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if jungsanlist.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table> 
-->
<!-- ǥ �ϴܹ� ��-->

<%
set jungsanlist = Nothing
%>
<script language='javascript'>
//function getonLoad(){
//checkComp(frm.ck_taxdate)
//}
//window.onload = getonLoad;
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
