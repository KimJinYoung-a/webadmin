<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemMaeipSaleMarginShareCls.asp"-->
<%

dim page, i
dim makerid
dim yyyy1, mm1, grpon

makerid     = requestCheckvar(request("makerid"),32)
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
grpon = requestCheckvar(request("grpon"),10)

page = requestCheckvar(request("page"),10)
if (page="") then page=1
dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if


'==============================================================================
dim oCItemMaeipSaleMarginShare

set oCItemMaeipSaleMarginShare = new CItemMaeipSaleMarginShare

oCItemMaeipSaleMarginShare.FPageSize         = 30
oCItemMaeipSaleMarginShare.FCurrPage         = page
oCItemMaeipSaleMarginShare.FRectMakerid      = makerid
oCItemMaeipSaleMarginShare.FRectYYYYMM 		 = yyyy1+"-"+mm1
if (grpon<>"") then
	oCItemMaeipSaleMarginShare.SearchMaeipSaleMarginShareJungsanListGrp
else
	oCItemMaeipSaleMarginShare.GetMasterList
end if

%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

function popDtl(iidx){
	var popwin = window.open("","maeipSaleMarginModi","width=1200,height=800,scrollbars=yes,resizable=yes,status=yes");
	popwin.location.href="maeipSaleMarginModi.asp?menupos=<%= menupos %>&idx="+iidx;

	popwin.focus();

}

function jsEtcSaleMarginJungsan(makerid){
	var upfrm1 = document.frmEtcJOne;
    upfrm1.makerid.value=makerid;

    if (confirm("�ۼ� �Ͻðڽ��ϱ�?")){
        upfrm1.submit();    
    }
}
</script>
<style>
.hnd {
    cursor: pointer;
}
</style>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %>

			&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;

			<input type="checkbox" name="grpon" <% if grpon="on" then response.write "checked" %>  >�����󺸱�
			(��������:<% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;)

		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="NextPage(1);">
		</td>
	</tr>
	</form>
</table>

<p />

<% if (grpon<>"") then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="12">
			�˻���� : <b><%= oCItemMaeipSaleMarginShare.FTotalCount%></b>
			
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">�귣��ID</td>
		<td width="80">����ݾ�</td>
		
		<td width="20"></td>
		<td>����TITLE</td>
		<td width="80">�������</td>
		<td width="80">���걸��</td>
		<td width="50">����</td>

		<td width="80">�������</td>
		<td width="80">�����ǸŰ���</td>
		<td width="80">������԰���</td>
		<td width="50">����</td>
		<td>���</td>
    </tr>
<% if oCItemMaeipSaleMarginShare.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="12" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oCItemMaeipSaleMarginShare.FresultCount > 0 then %>
   	<% for i=0 to oCItemMaeipSaleMarginShare.FresultCount-1 %>
	<tr height="25"  bgcolor="#FFFFFF" >
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Fmakerid %>
		</td>
		<td align="right">
			<%= FormatNumber(oCItemMaeipSaleMarginShare.FItemList(i).FmaySum,0) %>
		</td>
		<td align="right"></td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Ftitle %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Ffinishflag %>
		</td>
		<td align="center"></td>

		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Fjgubun %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Fet_cnt %>
		</td>
		<td align="right">
			<% if NOT isNULL(oCItemMaeipSaleMarginShare.FItemList(i).Fdlv_totalsuplycash) then %>
			<%= FormatNumber(oCItemMaeipSaleMarginShare.FItemList(i).Fdlv_totalsuplycash,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if NOT isNULL(oCItemMaeipSaleMarginShare.FItemList(i).Fdlv_totalsuplycash) then %>
			<%= FormatNumber(oCItemMaeipSaleMarginShare.FItemList(i).Fdlv_totalsuplycash,0) %>
			<% end if %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Fmaydiff %>
		</td>
		<td>
		<% if (oCItemMaeipSaleMarginShare.FItemList(i).Fmaydiff=1) then %>
        <input type="button" value="�ۼ�" onClick="jsEtcSaleMarginJungsan('<%= oCItemMaeipSaleMarginShare.FItemList(i).Fmakerid %>')">
        <% end if %>
		</td>
	</tr>
	<% next %>
	
<% end if %>
</table>
<% else %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="25" valign="bottom">
        <td align="left">
        	<input type="button" value="���ε��" class="button" onclick="javascript:location.href='maeipSaleMarginModi.asp?menupos=<%= menupos %>';" >
	    </td>
	    <td align="right"></td>
	</tr>
</table>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11">
			�˻���� : <b><%= oCItemMaeipSaleMarginShare.FTotalCount%></b>
			&nbsp;
			������ : <b><%= page %> /<%=  oCItemMaeipSaleMarginShare.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">IDX</td>
		<td width="80">�����ڵ�</td>
		<td width="120">�귣��</td>
		<td width="150">�Ⱓ</td>
		<td width="80">����</td>
		<td width="60">�⺻����</td>
		<td width="60">���θ���</td>
		<td width="150">�����</td>
		<td width="100">�����</td>
		<td width="100">��������</td>
		<td>���</td>
    </tr>
<% if oCItemMaeipSaleMarginShare.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="11" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oCItemMaeipSaleMarginShare.FresultCount > 0 then %>
   	<% for i=0 to oCItemMaeipSaleMarginShare.FresultCount-1 %>
	<tr height="25"  bgcolor="<%= CHKIIF(oCItemMaeipSaleMarginShare.FItemList(i).Fuseyn="Y", "#FFFFFF", "#EEEEEE")%>" >
		<td align="center">
			<a href="#" onClick="popDtl('<%= oCItemMaeipSaleMarginShare.FItemList(i).Fidx %>'); return false;"><%= oCItemMaeipSaleMarginShare.FItemList(i).Fidx %></a>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).FsaleCode %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Fmakerid %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).FstartDate %> ~ <%= oCItemMaeipSaleMarginShare.FItemList(i).FendDate %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).GetMeachulGubun %>
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).FdefaultMargin %>%
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).FsaleMargin %>%
		</td>
		<td align="center">
			<%= oCItemMaeipSaleMarginShare.FItemList(i).Freguserid %>
		</td>
		<td align="center">
			<%= Left(oCItemMaeipSaleMarginShare.FItemList(i).Fregdate,10) %>
		</td>
		<td align="center">
			<%= Left(oCItemMaeipSaleMarginShare.FItemList(i).Flastupdate,10) %>
		</td>
		<td></td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11" align="center">
			<% if oCItemMaeipSaleMarginShare.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oCItemMaeipSaleMarginShare.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oCItemMaeipSaleMarginShare.StartScrollPage to oCItemMaeipSaleMarginShare.FScrollCount + oCItemMaeipSaleMarginShare.StartScrollPage - 1 %>
    			<% if i>oCItemMaeipSaleMarginShare.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oCItemMaeipSaleMarginShare.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
<% end if %>
</table>
<% end if %>
<form name="frmEtcJOne" method="post" action="/admin/upchejungsan/dobatch.asp">
<input type="hidden" name="mode" value="etcSaleMarginJOne">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<input type="hidden" name="makerid" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
