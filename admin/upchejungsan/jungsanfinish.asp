<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim research, segumtype
dim thismonth

research = request("research")
segumtype = request("segumtype")


thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)
%>


<script language='javascript'>


function PopJungsanUpload(){
	var popwin = window.open("/admin/upchejungsan/pop_jungsan_upload.asp","PopJungsanUpload","width=800 height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	<input type="button" value="������ε�����" onclick="PopJungsanUpload();">
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<%
dim ipkumregdate
ipkumregdate = request("ipkumregdate")


dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"
ojungsan.FRectbankingupFile = "Y"

ojungsan.JungsanFixedList

dim ipsum,i
ipsum =0
%>
<script language='javascript'>
function ipkumfinish(frm,iidx){
	if (frm.ipkumregdate.value.length<1){
		alert('�Ա����� �Է��ϼ���.');
		frm.ipkumregdate.focus();
		return;
	}

	frm.idx.value= iidx;

	var ret = confirm('�����Ͻðڽ��ϱ�?');

	if (ret){
		var popwin = window.open("","regipkumfinish","width=300 height=300");
		popwin.focus();
		frm.target = "regipkumfinish";
		frm.submit();
	}
}

function delbankingup(iidx){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		var popwin = window.open("dobankingupflag.asp?mode=delflag&id=" + iidx,"regipkumfinish","width=100 height=100");
		popwin.focus();
	}
}

function batchipkumfinish(frm){
	if (frmip.ipkumregdate.value.length<1){
		alert('�Ա����� �Է��ϼ���.');
		calendarOpen(frmip.ipkumregdate);
		return;
	}


	if (confirm(frmip.ipkumregdate.value + '�� �Ա�Ȯ�� ���� �Ͻðڽ��ϱ�?')){
		frm.ipkumregdate.value=frmip.ipkumregdate.value;
		frm.submit();
	}
}

</script>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name=frmip method=post action="dodesignerjungsan.asp">
    <input type=hidden name=rd_state value=7>
    <input type="hidden" name="mode" value="ipkumfinish">
    <input type="hidden" name="idx" value="">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        �Ա��� : <input type=text name=ipkumregdate value="<%= ipkumregdate %>" size=10 maxlength=10 readonly >
	    	<a href="javascript:calendarOpen(frmip.ipkumregdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	    	(2004-06-30)
    		<input type="button" value="��ü�ԱݿϷ�����" onclick="batchipkumfinish(frmbatch);">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ �߰��� ��-->





<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >�ݿ�(<%= thismonth %>) ���ݰ�꼭 (<%= ojungsan.FresultCount %>��)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">�����</td>
		<td width="70">������</td>
		<td width="40">������</td> 
		<td width="120">�귣��ID</td>
      	<td width="150">������</td>
		<td width="60">����</td>
		<td width="60">����</td>
		<td width="80">����</td>
		<td width="80">����ݾ�</td>
		<td>��ü��</td>
		<td width="30">����</td>
		<td width="30">FileNo</td>
	</tr>
<form name="frmbatch" method="post" action="dobankingupflag.asp">
<input type="hidden" name="mode" value="ipkumfinish">
<input type="hidden" name="ipkumregdate" value="">
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>
	<input type=hidden name="checkone" value="<%= ojungsan.FItemList(i).FId %>">
	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ojungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ojungsan.FItemList(i).Fjungsan_date %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
		<td><%= ojungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
		<td><%= ojungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ojungsan.FItemList(i).Fid %>')">
		x
		</a>
		</td>
		<td><%= ojungsan.FItemList(i).FipFileNo %></td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectNotYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"
ojungsan.FRectbankingupFile = "Y"

ojungsan.JungsanFixedList



ipsum =0
%>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >���� ���ݰ�꼭 (<%= ojungsan.FresultCount %>��)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">�����</td>
		<td width="70">������</td>
		<td width="40">������</td> 
		<td width="120">�귣��ID</td>
      	<td width="150">������</td>
		<td width="60">����</td>
		<td width="60">����</td>
		<td width="80">����</td>
		<td width="80">����ݾ�</td>
		<td>��ü��</td>
		<td width="30">����</td>
		<td width="30">FileNo</td>
     </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>
	<input type=hidden name="checkone" value="<%= ojungsan.FItemList(i).FId %>">
	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ojungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ojungsan.FItemList(i).Fjungsan_date %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
		<td><%= ojungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
		<td><%= ojungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ojungsan.FItemList(i).Fid %>')">
		x
		</a>
		</td>
		<td><%= ojungsan.FItemList(i).FipFileNo %></td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = ""
ojungsan.FRectOnlyIncludeWonChon = "on"
ojungsan.FRectbankingupflag = "Y"
ojungsan.FRectbankingupFile = "Y"

ojungsan.JungsanFixedList

ipsum =0
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" >��õ¡�� ����� (<%= ojungsan.FresultCount %>��)</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">�����</td>
		<td width="70">������</td>      																																																																																																
		<td width="40">������</td>        																																																																																																
		<td width="120">�귣��ID</td>
      	<td width="100">������</td>																																																																							
		<td width="60">����</td>          																																																																																																
		<td width="60">����</td>          																																																																																																
		<td width="80">����</td>          																																																																																																
		<td width="60">Ȯ���ݾ�</td>      																																																																																																																
		<td width="60">����ݾ�</td>																																																																																																																
		<td>��ü��</td>                   																																																																																																																
		<td width="30">����</td> 
		<td width="30">FileNo</td>																																																																																																																
     </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum
%>
	<input type=hidden name="checkone" value="<%= ojungsan.FItemList(i).FId %>">
	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojungsan.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ojungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ojungsan.FItemList(i).Fjungsan_date %></td>
		<td><a href="javascript:PopUpcheBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
		<td><%= ojungsan.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
		<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
		<td><%= ojungsan.FItemList(i).Fcompany_name %></td>
		<td>
		<a href="javascript:delbankingup('<%= ojungsan.FItemList(i).Fid %>')">
		x
		</a>
		</td>
		<td><%= ojungsan.FItemList(i).FipFileNo %></td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="1"></td>
		<td colspan="2"></td>
	</tr>
</table>

<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->