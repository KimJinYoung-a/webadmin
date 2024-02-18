<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���
' History : �̻� ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/shopbatchstockcls.asp"-->
<%
'jobgubun
'10     ����ľ�
'90     ��ǰ

dim shopid, page, jobgubun, jobstate, assignshopid
dim research
	shopid      = RequestCheckVar(request("shopid"),32)
	page        = RequestCheckVar(request("page"),10)
	jobgubun    = RequestCheckVar(request("jobgubun"),10)
	jobstate    = RequestCheckVar(request("jobstate"),10)
	assignshopid= RequestCheckVar(request("assignshopid"),32)
	research    = RequestCheckVar(request("research"),2)

if (page = "") then page = "1"
if (research="") and (jobstate="") then jobstate="M"

dim oshoporder
set oshoporder = new CShopOrder
oshoporder.FPageSize = 300
oshoporder.FCurrPage = page
oshoporder.FRectShopID = shopid
oshoporder.FRectJobGubun = jobgubun
oshoporder.FRectjobState = jobstate
oshoporder.GetShopOrderList

dim i
dim totalsum
%>

<script language='javascript'>
function GotoPage(page){
	frm.page.value = page;
	frm.submit();
}

function OpenWinDetail(idx) {
    var popwin = window.open("pop_batchjaegolist.asp?idx=" + idx, "popwin", "width=1000,height=600,scrollbars=yes");
    popwin.focus();
}

function OpenWinInsert(idx) {
    var popwin = window.open("pop_batchjaegoinsert.asp?idx=" + idx, "popwin", "width=1000,height=600,scrollbars=yes");
    popwin.focus();
}

function ExcelSheet(idx){
	window.open('pop_batchjaegosheet.asp?idx=' + idx);
}

function MakeJobArr(){
	var frmlist = document.frmlist;
	var idxarr = "";
	var upfrm = document.frmArrupdate;

	for (var i=0;i<frmlist.elements.length;i++){
		if ((frmlist.elements[i].name=="ck_all") && (frmlist.elements[i].checked)){
        	idxarr = idxarr + frmlist.elements[i+1].value + "|";
      	}
	}

	if (idxarr==""){
		alert('����ľ� ������ �����ϼ���.');
		return;
	}
	if (document.frm.assignshopid.selectedIndex == 0){
		alert('���� �����ϼ���.');
		document.frm.assignshopid.focus();
		return;
	}
	
	//if (document.frm.jobgubun.selectedIndex == 0){
	//	alert('�۾������� �����ϼ���.');
	//	return;
	//}
    
	upfrm.jobshopid.value = document.frm.assignshopid[document.frm.assignshopid.selectedIndex].value;
	upfrm.jobgubun.value = "10"; //����ľ�
	//upfrm.jobgubun.value = document.frm.jobgubun[document.frm.jobgubun.selectedIndex].value;
	upfrm.idx.value = idxarr;
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        upfrm.submit();
    }
}

function CancelJobArr(){
	var frm = document.frmlist;
	var idxarr = "";
	var upfrm = document.frmArrupdate;

	for (var i=0;i<frm.elements.length;i++){
		if ((frm.elements[i].name=="ck_all") && (frm.elements[i].checked)){
        	idxarr = idxarr + frm.elements[i+1].value + "|";
      	}
	}

    if (frm.ck_all.length){
        for (i=0;i<frm.ck_all.length;i++){
    	    if ((frm.ck_all[i].checked)&&(frm.ismaster[i].value=="0")){
    	        alert('�۾� ������ ������ �۾�������� �����մϴ�.['+ i+ ']' + frm.idx[i].value);
    	        frm.ck_all[i].focus();
    	        return;
    	    }
    	}
	}else{
	    if ((frm.ck_all.checked)&&(frm.ismaster.value=="0")){
	        alert('�۾� ������ ������ �۾�������� �����մϴ�.');
	        frm.ck_all.focus();
	        return;
	    }
	}
        	
	if (idxarr==""){
		alert('����ľ� ������ �����ϼ���.');
		return;
	}

	upfrm.jobshopid.value = document.frm.shopid[document.frm.shopid.selectedIndex].value;
	upfrm.jobgubun.value = document.frm.jobgubun[document.frm.jobgubun.selectedIndex].value;
	upfrm.idx.value = idxarr;
	upfrm.mode.value = "cancelarr";
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        upfrm.submit();
    }
}

function popShopStockByJob(idx){
    //var popwin = window.open('/common/offshop/popShopStockByBatch.asp?idx=' + idx,'popShopStockByJob','width=900,height=700,scrollbars=yes,resizable=yes');
    var popwin = window.open('/common/offshop/popShopStockByBatchGroupBrand.asp?idx=' + idx,'popShopStockByBatchGroupBrand','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popJobStateChange(idx){
    var popwin = window.open('popJobStateChange.asp?idx=' + idx,'popJobStateChange','width=400,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    �� : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
        	&nbsp;&nbsp;�۾����� :
        	<select name="jobgubun">
        	  <option value="">����</option>
        	  <option value="10" <% if (jobgubun = "10") then %>selected<% end if %>>����ľ�</option>
        	  <!-- option value="90" <% if (jobgubun = "90") then %>selected<% end if %>>��ǰ</option -->
        	</select>
	        &nbsp;&nbsp;������� :
	        <select name="jobstate">
	            <option value="">����</option>
	            <option value="M" <%= ChkIIf(jobstate = "M","selected","") %> >��ó����ü</option>
	            <option value="0" <%= ChkIIf(jobstate = "0","selected","") %> >��ó��</option>
	            <option value="3" <%= ChkIIf(jobstate = "3","selected","") %> >ó����</option>
	            <!-- option value="5" <%= ChkIIf(jobstate = "5","selected","") %> >Ȯ��</option -->
	            <option value="7" <%= ChkIIf(jobstate = "7","selected","") %> >�Ϸ�</option>
	        </select>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	
</table>
<!-- �˻� �� -->



<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <% 'drawSelectBoxOffShop "assignshopid",assignshopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("assignshopid",assignshopid, "21") %>
            <input type=button value="�۾��ϰ�����" onclick="MakeJobArr()">&nbsp;&nbsp;<input type=button value="�۾��������" onclick="CancelJobArr()"><p></td>
        <td align="right">
        �˻���� : <%= oshoporder.FTotalCount %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="20"></td>
    	<td width="40">�۾�<br>��ȣ</td>
    	<td width="110">����ľǹ�ȣ</td>
    	<td width="100">ShopID</td>
    	<td width="50">PosID</td>
    	<td width="80">�ѱݾ�</td>
    	<td width="80">���԰�</td>
    	<td width="60">�۾�<br>����</td>
    	<td width="80">��¥</td>
    	<td width="80">�������</td>
    	<td width="80">����Է�</td>
    	<td width="40">���<br>����</td>
    	<td></td>
    	<!--
    	<td>�����ڵ�</td>
    	-->
    	<td width="42"></td>
    </tr>
    <form name=frmlist method=post>
    <input type="hidden" name="idxarr" >
<% for i=0 to oshoporder.FresultCount-1 %>
	<%
	totalsum = totalsum + oshoporder.FItemList(i).Frealsum
	%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td >
    	    <%  IF (oshoporder.FItemList(i).IsJobCheckAvali) then %>
    	    <input type=checkbox name=ck_all onclick="AnCheckClick(this);"  >
    	    <% else %>
    	    <input type=hidden name=ck_all >
    	    <% end if %>
    	</td>
    	<input type="hidden" name="idx" value="<%= oshoporder.FItemList(i).Fidx %>">
    	<input type="hidden" name="ismaster" value="<%= ChkIIF(oshoporder.FItemList(i).IsMasterJob,"1","0") %>">
    	
    	<td><%= oshoporder.FItemList(i).Fjobkey %></td>
    	<td>
    	    <%  IF Not (oshoporder.FItemList(i).IsSubJob) then %>
    	    <%= oshoporder.FItemList(i).Forderno %>
    	    <% end if %>
    	</td>
    	<td>
    		<%= oshoporder.FItemList(i).Fjobshopid %>
    	<br><font color="#CCCCCC"><%= oshoporder.FItemList(i).Fshopid %></font>
    	</td>
    	<td align="center">
    	    <%  IF  (oshoporder.FItemList(i).IsSubJob) then %>
    	    <%= Right(Left(oshoporder.FItemList(i).Forderno,11),2) %>
    	    - <%= Right(oshoporder.FItemList(i).Forderno,5) %> 
    	    <% else %>
    	    <% if false then %>
    	    <%= Right(Left(oshoporder.FItemList(i).Forderno,11),2) %>
    	    - <%= Right(oshoporder.FItemList(i).Forderno,5) %> 
    	    <% end if %>
    	    <% end if %>
    	    <br>
    	    <%= oshoporder.FItemList(i).FCasherid %>
    	    <br>
    	    <%= oshoporder.FItemList(i).Fpointuserno %>
    	</td>
    	<td align="right"><%= FormatNumber(oshoporder.FItemList(i).Frealsum, 0) %></td>
    	<td align="right"><%= FormatNumber(oshoporder.FItemList(i).Fsuplysum, 0) %></td>
    	<td><%= oshoporder.FItemList(i).GetJobGubunName %></td>
    	<td><%= oshoporder.FItemList(i).Fshopregdate %></td>
    	<td>
    	    <% if oshoporder.FItemList(i).IsJobStateChangeAvali then %>
    	    <a href="javascript:popJobStateChange('<%= oshoporder.FItemList(i).Fidx %>');"><%= oshoporder.FItemList(i).GetJobStateName %></a>
    	    <% else %>
    	    <%= oshoporder.FItemList(i).GetJobStateName %>
    	    <% end if %>
    	</td>
    	<td>
    	    <%  IF (oshoporder.FItemList(i).IsMasterJob) and (oshoporder.FItemList(i).FjobState="3") then %>
    	    <a href="javascript:popShopStockByJob('<%= oshoporder.FItemList(i).Fidx %>');">�Է� <img src="/images/icon_arrow_link.gif" border="0" width="14" align="absmiddle"></a>
    	    <% end if %>    
    	</td>
    	<td>
            <% if (oshoporder.FItemList(i).Fcancelyn = "Y") then %>
              <font color=red>���</font>
            <% else %>
              ����
            <% end if %>
        </td>
        <td>&nbsp;</td>
        <!--
    	<td>
          <% if ((isnull(oshoporder.FItemList(i).Fjoblinkcode) = true) and (oshoporder.FItemList(i).Fjobgubun = "90")) then %>
            <a href="javascript:OpenWinInsert('<%= oshoporder.FItemList(i).Fidx %>')">�ֹ����ۼ�</a>
          <% end if %>
          <% if ((isnull(oshoporder.FItemList(i).Fjoblinkcode) <> true) and (oshoporder.FItemList(i).Fjobgubun = "90")) then %>
            <%= oshoporder.FItemList(i).Fjoblinkcode %>(<a href="/admin/fran/jumunlist.asp?menupos=520&baljucode=<%= oshoporder.FItemList(i).Fjoblinkcode %>" target="_frnOrder">[����]</a> <a href="javascript:OpenWinInsert('<%= oshoporder.FItemList(i).Fidx %>')">[���ۼ�]</a>)
          <% end if %>
        </td>
        -->
    	<td>
    		<a href="javascript:OpenWinDetail('<%= oshoporder.FItemList(i).Fidx %>')"><img src="/images/iexplorer.gif" width=21 border=0></a>
    		<a href="javascript:ExcelSheet('<%= oshoporder.FItemList(i).Fidx %>');"><img src="/images/iexcel.gif" width=21 border=0>
    	</td>
    </tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="5"></td>
		<td align=right><%= FormatNumber(totalsum, 0) %></td>
		<td align=right></td>
		<td colspan="7"></td>
	</tr>
	</form>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td valign="bottom" align="center">
	<% if oshoporder.HasPreScroll then %>
		<a href="javascript:GotoPage('<%= oshoporder.StarScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oshoporder.StarScrollPage to oshoporder.FScrollCount + oshoporder.StarScrollPage - 1 %>
		<% if i>oshoporder.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:GotoPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oshoporder.HasNextScroll then %>
		<a href="javascript:GotoPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
        </td>
    </tr>
</table>
<form name="frmArrupdate" method="post" action="batchjaegolist_process.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="jobshopid" value="">
<input type="hidden" name="jobgubun" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->