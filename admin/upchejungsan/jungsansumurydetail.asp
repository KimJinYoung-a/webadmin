<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
Dim USEERP : USEERP = FALSE
Dim USEUPFILE : USEUPFILE = TRUE
IF (USEUPFILE=TRUE) then USEERP= FALSE

dim research, ck_stype, designer, groupid
dim ck_SeTyp, ck_Mibus, ck_Up, topn, jgubun

ck_stype = request("ck_stype")
research = request("research")
designer = RequestCheckVar(request("designer"),32)
groupid  = RequestCheckVar(request("groupid"),32)
ck_SeTyp = RequestCheckVar(request("ck_SeTyp"),10)
ck_Mibus = RequestCheckVar(request("ck_Mibus"),10)
ck_Up    = RequestCheckVar(request("ck_Up"),10)
topn     = RequestCheckVar(request("topn"),10)
jgubun   = RequestCheckVar(request("jgubun"),10)

dim yyyymm : yyyymm = RequestCheckVar(request("yyyymm"),7)
if LEN(yyyymm)<>7 then yyyymm=""

if (research="") then ck_stype=""
if (research="") then ck_SeTyp="S"

if (topn="") then topn=500  ''2017/02/09 by eastone

dim i,premonth
dim sum1, sum2, sum3, sum4
dim allsum1, allsum2, allsum3, allsum4

dim ipsum,osum

if (ck_SeTyp="W") then USEERP=FALSE
%>
<script language='javascript'>


<!-- ������ ����� ���ݰ�꼭 ��� -->
function PopTaxPrint(itax_no,ibizno){
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + itax_no + "&cur_biz_no=2118700620&s_biz_no=" + ibizno + "&b_biz_no=2118700620","taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}


function PopTaxPrintReDirect(itax_no, makerid){
	var popwinsub = window.open("/admin/upchejungsan/red_taxprint.asp?tax_no=" + itax_no + "&makerid=" + makerid,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}


function CkeckAll(frm,bool){
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled!=true){
    			e.checked=bool;
    			AnCheckClick2(e);
			}
		}
	}
}

function CkeckNsubmit(frm){
	var pass = false;

	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")&&(e.checked)) {
			pass = true;
		}
	}

    if (!pass){
        alert('���� ������ �����ϴ�.');
        return;
    }

	if (pass){
	    <% IF (USEUPFILE) THEN %>
	    //2011-12 ����
    	var iURI = '/admin/upchejungsan/popIpkumUpFile.asp?targetGbn=<%= CHKIIF(ck_SeTyp="W","WN","ON") %>&frmName=' + frm.name;
    	var popwin=window.open(iURI,'popIpkumUpFile','width=800,height=600,scrollbars=yes,resizable=yes');
    	popwin.focus();
	    <% ELSEIF (NOT USEERP) THEN %>
	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.UseErp.value="";
    		frm.submit();
    		return;
    	}
    	<% ELSE %>
    	//2011-12 ����
    	var iURI = '/admin/approval/comm/popPayRequestSelect.asp?frmName=' + frm.name
    	var popwin=window.open(iURI,'popReqPayRequest','width=800,height=600,scrollbars=yes,resizable=yes');
    	popwin.focus();
    	<% end if %>
	}
}

function jsPopSubmit(frmName,ireqIcheDate,ipayRequestIdx){
    var frm = eval(frmName);
    if (confirm('���� ��û ��ü ������ ���� �Ͻðڽ��ϱ�?')){
	    frm.reqIcheDate.value = ireqIcheDate;
	    frm.payRequestIdx.value = ipayRequestIdx;
		frm.submit();
	}
}

function jsPopSubmitFile(frmName,ireqIcheDate,ipFileNo){
    var frm = eval(frmName);
    var ijgubun = '<%=jgubun%>';
    if ((ijgubun=='')&&(ipFileNo=='')){
		//alert('���� ���� (������,����) ����� �������� �ʾҽ��ϴ�.');
		//return;
        // �ڽ�â���� �Լ� ȣ��� confirm���� ũ�ҿ��� ������ �Ǿ��ٰ� �ȵǾ��ٰ� �ϴ� ���װ� ����.
		//if (!confirm('���� ���� (������,����) ����� �������� �ʾҽ��ϴ�. ��� �Ͻİڽ��ϱ�??')){
        //    return;
        //}
    }
	// �ڽ�â���� �Լ� ȣ��� confirm���� ũ�ҿ��� ������ �Ǿ��ٰ� �ȵǾ��ٰ� �ϴ� ���װ� ����.
    //if (confirm('���� ��û ��ü ������ ���� �Ͻðڽ��ϱ�??')){
	    frm.reqIcheDate.value = ireqIcheDate;
	    frm.ipFileNo.value = ipFileNo;
		frm.submit();
	//}
	//alert('ok');
}

function AnCheckClick2(e){
	if (e.checked)
		hL2(e);
	else
		dL2(e);
}

function hL2(E){
	while (E.tagName!="TR")
	{
		E=E.parentElement;
	}

    if (E.bgColor=="<%= LCASE(adminColor("dgray")) %>"){

    }else{
	    E.className = "H";
	}
}

function dL2(E){
	while (E.tagName!="TR"){
		E=E.parentElement;
	}

	E.className = "";
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
   	<form name="frm2" method=get>
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">

   	<tr  >
   	    <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>" align="center">�˻�<br>����</td>
        <td valign="top" bgcolor="F4F4F4">
            �귣��ID : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
			��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" > &nbsp;&nbsp;
			���ε� ���� :
			<input type="radio" name="ck_Up" value="" <% if ck_Up="" then response.write "checked" %> >��ü
			<input type="radio" name="ck_Up" value="N" <% if ck_Up="N" then response.write "checked" %> >���ε�������
			<input type="radio" name="ck_Up" value="Y" <% if ck_Up="Y" then response.write "checked" %> >���ε�ϷḸ

			�˻�����:<select name="topn">
			<option value="">��ü
			<option value="500" <%= CHKIIF(topn="500","selected","") %> >500��
			<option value="700" <%= CHKIIF(topn="700","selected","") %> >700��
			<option value="1000" <%= CHKIIF(topn="1000","selected","") %> >1000��
			<option value="1500" <%= CHKIIF(topn="1500","selected","") %> >1500��
			<option value="2000" <%= CHKIIF(topn="2000","selected","") %> >2000��
			</select>


        </td>
        <td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
           <img src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="document.frm2.submit();" style="cursor:pointer">
            <br><br>
        	<input type="button" value="���ó��� UPLOAD" onclick="CkeckNsubmit(frmlist)">
        </td>
	</tr>
	<tr>
	    <td bgcolor="F4F4F4">
	    �����ı��� :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
	    ��꼭 ���� :
			<!--
			<input type="radio" name="ck_SeTyp" value="" <% if ck_SeTyp="" then response.write "checked" %> >��ü
			-->
			<input type="radio" name="ck_SeTyp" value="S" <% if ck_SeTyp="S" then response.write "checked" %> >(����)��꼭
			<input type="radio" name="ck_SeTyp" value="W" <% if ck_SeTyp="W" then response.write "checked" %> >��õ¡��
			<input type="radio" name="ck_SeTyp" value="K" <% if ck_SeTyp="K" then response.write "checked" %> >���̰���(���������) <!-- 2013/08/02 �߰� gilzahousemd ���� ���̰������� �ŷ��ϸ� �ȵ� -->
	    </td>
	</tr>
	<tr>
	    <td bgcolor="F4F4F4">

		    ������ ���� :
			<input type="radio" name="ck_stype" value="" <% if ck_stype="" then response.write "checked" %> >��ü
			<input type="radio" name="ck_stype" value="SS" <% if ck_stype="SS" then response.write "checked" %> >�������(����) 	<!-- ���� ���곻�� �� �������� ���� & ������ ����/15��-->
			<input type="radio" name="ck_stype" value="AA" <% if ck_stype="AA" then response.write "checked" %> >�������(����/15��) 	<!-- ���� ���곻�� �� �������� ���� & ������ ����/15��-->
		    <input type="radio" name="ck_stype" value="BB" <% if ck_stype="BB" then response.write "checked" %> >�������(����)			<!-- ���� ���곻�� �� �������� ���� & ������ ����-->
			<input type="radio" name="ck_stype" value="CC" <% if ck_stype="CC" then response.write "checked" %> >�������(�̿�����)		<!-- ������ ���� ���곻�� �� �������� ����-->
			<input type="radio" name="ck_stype" value="DD" <% if ck_stype="DD" then response.write "checked" %> >�̿����� 				<!-- �������� ����� �̻�-->
			<input type="radio" name="ck_stype" value="ZZ" <% if ck_stype="ZZ" then response.write "checked" %> >��Ÿ					<!-- �������� ���̰ų�, �� �� ��¥ -->
			/
			<input type="radio" name="ck_stype" value="NN" <% if ck_stype="NN" then response.write "checked" %> >������� (<%=LEFT(now(),7)%>)

			&nbsp;&nbsp;
			����:<input type="text" name="yyyymm" value="<%=yyyymm%>" size="7">
	    </td>
	</tr>
	<tr>
	    <td bgcolor="F4F4F4">
	    ���̳ʽ� ����:
			<input type="radio" name="ck_Mibus" value="" <% if ck_Mibus="" then response.write "checked" %> >��ü
			<input type="radio" name="ck_Mibus" value="MJ" <% if ck_Mibus="MJ" then response.write "checked" %> >���̳ʽ� ����
			<input type="radio" name="ck_Mibus" value="MI" <% if ck_Mibus="MI" then response.write "checked" %> >���̳ʽ����Ծ�ü
			&nbsp;&nbsp;
			<input type="radio" name="ck_Mibus" value="CX" <% if ck_Mibus="CX" then response.write "checked" %> >���ó�� ���ɰ˻�
	    </td>
	</tr>
	</form>
</table>
<p>
<!-- ǥ ��ܹ� ��-->

<%
dim ojungsan
set ojungsan = new CUpcheJungsan

ojungsan.FRectGubun              = ck_stype
ojungsan.FRectMinusGubnu         = ck_Mibus
ojungsan.FRectBankingupflag      = ck_Up
ojungsan.FRectNotIncludeWonChon  = "on"
ojungsan.FRectOnlyIncludeWonChon = ""
ojungsan.FRectDesigner = designer
ojungsan.FRectGroupid  = groupid
ojungsan.FRectbankingupFile = "Y"
ojungsan.FRectTopN = topn
ojungsan.FRectJGubun= jgubun
ojungsan.FRectYYYYMM = yyyymm

IF (ck_SeTyp="S") or (ck_SeTyp="K") THEN
    if (ck_SeTyp="K") then
        ojungsan.FRectNotIncludeWonChon  = ""
        ojungsan.FRectOnlyIncludeKani = "on"
    end if
    ojungsan.JungsanFixedList
EnD If

dim ooffjungsan
set ooffjungsan = new COffJungsan

ooffjungsan.FRectGubunCd            = ck_stype
ooffjungsan.FRectMinusGubnu         = ck_Mibus
ooffjungsan.FRectBankingupflag      = ck_Up
ooffjungsan.FRectNotIncludeWonChon  = "on"
ooffjungsan.FRectOnlyIncludeWonChon = ""
ooffjungsan.FRectmakerid = designer
ooffjungsan.FRectGroupid  = groupid
ooffjungsan.FRectbankingupFile = "Y"
ooffjungsan.FRectJGubun= jgubun
IF (ck_SeTyp="S") and (ck_Mibus="CX") THEN
    ooffjungsan.FRectMinusGubnu = "CX1"
    ooffjungsan.JungsanFixedList
EnD If

ipsum =0

dim ipsumOff : ipsumOff=0
%>
<% IF (ck_SeTyp="") THEN %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td align="center">��꼭 ������ ���� ���� �ϼ���.</td>
    </tr>
</table>
<% ENd IF %>

<p />

* �������� : �귣�� �߰��������� �Ǵ� �׷��ڵ��� ���������� ������(��ü���� �ۼ��� �Է�)

<p />

<% IF (ck_SeTyp="S") or (ck_SeTyp="K") THEN %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr valign="top" bgcolor="F4F4F4" height="5">
		<td colspan="18" bgcolor="FFFFFF"><b>����(����)��꼭</b>
		<%= FormatNumber(ojungsan.FresultCount,0) %> ��
		</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="20"><input type="checkbox" name="ckall" onclick="CkeckAll(frmlist,this.checked)"></td>
      <td width="50">����</td>
      <td width="80">������</td>
      <td width="40">����</td>
      <td width="30">����</td>
      <td width="30">����</td>
      <td width="20"><img src="/images/icon_print02.gif" width="14" height="14"></td>
      <td width="80">�����<br>��������</td>
      <td width="120">�귣��ID</td>
      <td width="150">������</td>
      <td width="50">����</td>
      <td width="50">����</td>
      <td width="100">����</td>
      <td width="70">����ݾ�</td>
	  <td width="50">�׷��ڵ�</td>
      <td>��ü��</td>
      <td width="30">UP</td>
      <td width="30">FileNo</td>
     </tr>
     <form name="frmlist" method=post action="dobankingupflag.asp">
     <input type="hidden" name="mode" value="bankingupload">
     <input type="hidden" name="reqIcheDate" value=""> <!-- 2011-12 �߰� ��ü������ -->
     <input type="hidden" name="payRequestIdx" value=""> <!-- 2011-12 �߰� ���� ��û�� IDX : �űԽ�-1-->
     <input type="hidden" name="UseErp" value="<%= CHKIIF(UseErp=TRUE,"1","") %>">
     <input type="hidden" name="ipFileNo" value="">
     <input type="hidden" name="UseUpFile" value="<%= CHKIIF(UseUpFile=TRUE,"1","") %>">
     <input type="hidden" name="ck_Mibus" value="<%= ck_Mibus %>">
     <input type="hidden" name="targetGbn" value="<%= CHKIIF(ck_SeTyp="W","WN","ON") %>">
     <input type="hidden" name="jgubun" value="<%= jgubun %>">
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

    <% if NOT IsNULL(ojungsan.FItemList(i).FholdGroupid) then %>
    <tr align="center" bgcolor="#77AA77">
	<% elseif ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
	<td><abbr title="<%= ojungsan.FItemList(i).Fholdcause %>"><input type="checkbox" name="checkone" value="<%= ojungsan.FItemList(i).Fid %>" onClick="AnCheckClick2(this)" <%= CHKIIF(IsNULL(ojungsan.FItemList(i).FipFileNo) and IsNULL(ojungsan.FItemList(i).FholdGroupid),"","disabled") %>></td>
	<td><%= ojungsan.FItemList(i).Fyyyymm %></td>
	<td>
		<% if Left(ojungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
		<font color="red"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
		<% else %>
		<font color="blue"><%= ojungsan.FItemList(i).Ftaxregdate %></font>
		<% end if %>
	</td>
	<td><%= ojungsan.FItemList(i).getJGubunName %></td>
	<td><%= ojungsan.FItemList(i).Fjungsan_date %></td>
	<td><font color="<%= ojungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ojungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>

	<td><%= ojungsan.FItemList(i).Fbillsitecode %></td>
	<td>
	    <% if ojungsan.FItemList(i).Fjungsan_gubun="���̰���" then %>
	    <font color="blue"><%= ojungsan.FItemList(i).Fjungsan_gubun %></font>
	    <% else %>
	    <%= ojungsan.FItemList(i).Fjungsan_gubun %>
	    <% end if %>
	</td>
	<td>
		<a href="javascript:PopUpcheBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a>
	</td>
	<td><%= ojungsan.FItemList(i).Fjungsan_acctname %></td>
	<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
	<td><%= ojungsan.FItemList(i).Fjungsan_bank %></td>
	<td><%= ojungsan.FItemList(i).Fjungsan_acctno %></td>
	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
	<td><%= ojungsan.FItemList(i).Fgroupid %></td>
	<td><%= ojungsan.FItemList(i).Fcompany_name %></td>
	<td><% if ojungsan.FItemList(i).Fbankingupflag<>"N" then response.write ojungsan.FItemList(i).Fbankingupflag %></td>
	<td><%= ojungsan.FItemList(i).FipFileNo %>
	<% if (ojungsan.FItemList(i).FtargetGbn="OF") then %>
	<b>OF</b>
	<% end if %>
	</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="4"></td>
	</tr>
<% IF (ooffjungsan.FresultCount>0) then %>
    <tr bgcolor="#FFFFFF">
        <td colspan="18"><b>�������� ��� ó�� ���� ����</b>
        <%= FormatNumber(ooffjungsan.FresultCount,0) %> ��
        </td>
    </td>

<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsumOFF = ipsumOFF + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>

	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
	<td><abbr title="<%= ooffjungsan.FItemList(i).Fholdcause %>">
	    <input type="checkbox" name="checkoneEx" value="<%= ooffjungsan.FItemList(i).Fidx %>" onClick="AnCheckClick2(this)" <%= CHKIIF(ooffjungsan.FItemList(i).Fbankingupflag="Y" or Not IsNULL(ooffjungsan.FItemList(i).FholdGroupid),"disabled","") %> ></abbr></td>
	<td><%= ooffjungsan.FItemList(i).Fyyyymm %></td>
	<td>
		<% if Left(ooffjungsan.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
		<font color="red"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
		<% else %>
		<font color="blue"><%= ooffjungsan.FItemList(i).Ftaxregdate %></font>
		<% end if %>
	</td>
	<td><%= ooffjungsan.FItemList(i).getJGubunName %></td>
	<td><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
	<td><font color="<%= ooffjungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ooffjungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
	<td><%= ooffjungsan.FItemList(i).Fbillsitecode %></td>
	<td></td>
	<td>
		<a href="javascript:PopUpcheBrandInfoEdit('<%= ooffjungsan.FItemList(i).Fmakerid %>')"><%= ooffjungsan.FItemList(i).Fmakerid %></a>
	</td>
	<td><%= ooffjungsan.FItemList(i).Fjungsan_acctname %></td>
	<td><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
	<td><%= ooffjungsan.FItemList(i).Fjungsan_bank %></td>
	<td><%= ooffjungsan.FItemList(i).Fjungsan_acctno %></td>
	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
	<td><%= ooffjungsan.FItemList(i).FholdGroupid %></td>
	<td><%= ooffjungsan.FItemList(i).Fcompany_name %></td>
	<td><% if ooffjungsan.FItemList(i).Fbankingupflag<>"N" then response.write ooffjungsan.FItemList(i).Fbankingupflag %></td>
	<td><%= ooffjungsan.FItemList(i).FipFileNo %>
	<% if (ooffjungsan.FItemList(i).FtargetGbn="OF") then %>
	<b>OF</b>
	<% end if %></td>
	</tr>
<% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="13"></td>
		<td align="right"><%= FormatNumber(ipsumOFF,0) %></td>
		<td colspan="4"></td>
	</tr>
<% end if %>
	</form>
</table>
<% END IF %>
<%
set ooffjungsan = Nothing
set ojungsan = Nothing
%>


<%
dim ojungsanEtc
set ojungsanEtc = new CUpcheJungsan

ojungsanEtc.FRectGubun           = ck_stype
ojungsanEtc.FRectMinusGubnu         = ck_Mibus
ojungsanEtc.FRectBankingupflag      = ck_Up
ojungsanEtc.FRectNotIncludeWonChon = ""
ojungsanEtc.FRectOnlyIncludeWonChon = "on"
ojungsanEtc.FRectDesigner = designer
ojungsanEtc.FRectGroupid  = groupid
ojungsanEtc.FRectbankingupFile = "Y"

IF (ck_SeTyp="W") THEN
    ojungsanEtc.JungsanFixedList
EnD If

ipsum = 0
osum  = 0
%>
<br>

<% IF (ck_SeTyp="W") THEN %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr valign="top" bgcolor="F4F4F4" height="5">
		<td colspan="18" bgcolor="FFFFFF"><b>��õ¡�� �����</b>
		<%= FormatNumber(ojungsanEtc.FresultCount,0) %> ��
		</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="20"><input type="checkbox" name="ckall" onclick="CkeckAll(frmlist,this.checked)"></td>
      <td width="50">����</td>
      <td width="60">������</td>
      <td width="40">����</td>
      <td width="30">����</td>
      <td width="30">����</td>
      <td width="20"><img src="/images/icon_print02.gif" width="14" height="14"></td>
      <td width="120">�귣��ID</td>
      <td width="100">������</td>
      <td width="50">����</td>
      <td width="50">����</td>
      <td width="100">����</td>
      <td width="60">Ȯ���ݾ�</td>
      <td width="60">(��õ)����ݾ�</td>
	  <td width="50">�׷��ڵ�</td>
      <td>��ü��</td>
	  <td width="30">UP</td>
	  <td width="30">FileNo</td>
     </tr>
     <form name="frmlist" method=post action="dobankingupflag.asp">
     <input type="hidden" name="mode" value="bankingupload">
     <input type="hidden" name="reqIcheDate" value=""> <!-- 2011-12 �߰� ��ü������ -->
     <input type="hidden" name="payRequestIdx" value=""> <!-- 2011-12 �߰� ���� ��û�� IDX : �űԽ�-1-->
     <input type="hidden" name="UseErp" value="<%= CHKIIF(UseErp=TRUE,"1","") %>">
     <input type="hidden" name="ipFileNo" value="">
     <input type="hidden" name="UseUpFile" value="<%= CHKIIF(UseUpFile=TRUE,"1","") %>">
     <input type="hidden" name="ck_Mibus" value="<%= ck_Mibus %>">
     <input type="hidden" name="targetGbn" value="<%= CHKIIF(ck_SeTyp="W","WN","ON") %>">
     <input type="hidden" name="jgubun" value="<%= jgubun %>">
<% for i=0 to ojungsanEtc.FresultCount-1 %>
<%
osum = osum + fix(ojungsanEtc.FItemList(i).GetTotalSuplycash)
ipsum = ipsum + ojungsanEtc.FItemList(i).GetTotalWithHoldingJungSanSum
%>
	<% if ojungsanEtc.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td width="20"><abbr title="<%= ojungsanEtc.FItemList(i).Fholdcause %>"><input type="checkbox" name="checkone" value="<%= ojungsanEtc.FItemList(i).Fid %>" <%= CHKIIF(Not IsNULL(ojungsanEtc.FItemList(i).FholdGroupid),"disabled","") %> onClick="AnCheckClick2(this)"></abbr></td>
		<td ><%= ojungsanEtc.FItemList(i).Fyyyymm %></td>
		<td>
			<% if Left(ojungsanEtc.FItemList(i).Ftaxregdate,7) = Left(CStr(now()),7) then %>
			<font color="red"><%= ojungsanEtc.FItemList(i).Ftaxregdate %></font>
			<% else %>
			<font color="blue"><%= ojungsanEtc.FItemList(i).Ftaxregdate %></font>
			<% end if %>
		</td>
		<td><%= ojungsanEtc.FItemList(i).getJGubunName %></td>
		<td><%= ojungsanEtc.FItemList(i).Fjungsan_date %></td>
		<td><font color="<%= ojungsanEtc.FItemList(i).GetTaxtypeNameColor %>"><%= ojungsanEtc.FItemList(i).GetSimpleTaxtypeName %></font></td>

		<td><%= ojungsanEtc.FItemList(i).Fbillsitecode %></td>

		<td><%= ojungsanEtc.FItemList(i).Fdesignerid %></td>
		<td><%= ojungsanEtc.FItemList(i).Fjungsan_acctname %></td>
		<td><font color="<%= ojungsanEtc.FItemList(i).GetStateColor %>"><%= ojungsanEtc.FItemList(i).GetStateName %></font></td>
		<td><%= ojungsanEtc.FItemList(i).Fjungsan_bank %></td>
		<td><%= ojungsanEtc.FItemList(i).Fjungsan_acctno %></td>
		<td align="right"><%= FormatNumber(ojungsanEtc.FItemList(i).GetTotalSuplycash,0) %></td>
		<td align="right"><%= FormatNumber(ojungsanEtc.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
		<td><%= ojungsanEtc.FItemList(i).Fgroupid %></td>
		<td><%= ojungsanEtc.FItemList(i).Fcompany_name %></td>
		<td><% if ojungsanEtc.FItemList(i).Fbankingupflag<>"N" then response.write ojungsanEtc.FItemList(i).Fbankingupflag %></td>
		<td><%= ojungsanEtc.FItemList(i).FipFileNo %>
		<% if (ojungsanEtc.FItemList(i).FtargetGbn="OF") then %>
    	<b>OF</b>
    	<% end if %>
		</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12"></td>
		<td align="right"><%= FormatNumber(osum,0) %></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="4"></td>
	</tr>
	</form>
</table>
<% End IF %>
<%
set ojungsanEtc = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
