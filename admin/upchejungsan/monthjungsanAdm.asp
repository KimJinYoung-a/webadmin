<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

dim ojungsanTaxCC
set ojungsanTaxCC = new CUpcheJungsanTax
ojungsanTaxCC.FRectMakerid = makerid
ojungsanTaxCC.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxCC.FRectJGubun = "CC"
ojungsanTaxCC.getMonthUpcheJungsanList

dim ojungsanTaxCE
set ojungsanTaxCE = new CUpcheJungsanTax
ojungsanTaxCE.FRectMakerid = makerid
ojungsanTaxCE.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxCE.FRectJGubun = "CE"
ojungsanTaxCE.getMonthUpcheJungsanList

dim ojungsanTaxMM
set ojungsanTaxMM = new CUpcheJungsanTax
ojungsanTaxMM.FRectMakerid = makerid
ojungsanTaxMM.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTaxMM.FRectJGubun = "MM"
ojungsanTaxMM.getMonthUpcheJungsanList

dim i
%>
<script language='javascript'>
function PopDetail(iidx,tg,igroupid){
    var uri = 'jungsandetailsumONAdm.asp?id=' + iidx + '&groupid='+igroupid;
    if (tg=="OF") uri = 'jungsandetailsumOFAdm.asp?idx=' + iidx + '&groupid='+igroupid;
	var popwin = window.open(uri+'&makerid=<%=makerid%>','PopDetail','width=1280, height=800, scrollbars=1, resizable=yes');
	popwin.focus();
}

function PopTaxRegPrdCommission(makerid, yyyy1, mm1, onoffGubun, jidx) {
	<% 'var popwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=640 height=700 scrollbars=yes resizable=yes"); %>
    var popwin = window.open("/admin/upchejungsan/popTaxRegAdminapi.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=1024 height=768 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=800,height=700,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function PopConfirm(mnupos,iidx){
	//var popwin = window.open('jungsanmaster.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=900, height=540, scrollbars=1');
	//popwin.focus();
}

function PopTaxReg(v){
	//var popwin = window.open("poptaxreg.asp?id=" + v,"poptaxreg","width=640 height=700 scrollbars=yes resizable=yes");
	//popwin.focus();
}

function PopTaxRegOff(v){
	//var popwin = window.open("poptaxregoff.asp?idx=" + v,"poptaxregoff1","width=640 height=680 scrollbars=yes resizable=yes");
	//popwin.focus();
}
<% if (ojungsanTaxCC.FresultCount>0) then %>
//alert('2014�� 1�� ������� ������ ����п� ���ؼ���\n\n�ٹ����ٿ��� ��꼭�� ���� �Ͽ���\n\n�̼��� ���� ���� ���� �������� ���� �ֽñ� �ٶ��ϴ�.');
<% end if %>
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���� ��� ��� :&nbsp;<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		�귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>&nbsp;&nbsp;
		<!--
		<span ><strong>������ ���޴��� (��ۺ�), �߰������ (��Ÿ) �и� ����</strong></span>
        -->
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<p>

<% if (ojungsanTaxCC.FresultCount<1) and (ojungsanTaxCE.FresultCount<1) and (ojungsanTaxMM.FresultCount<1) then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td align="left"><strong>* ���� ����</strong></td>
</tr>
<tr height="30">
    <td align="center" bgcolor="#FFFFFF"> �˻� ����� ���� ���� �ʽ��ϴ�.</td>
</tr>
</table>
<% else %>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="15" align="left"><strong>* ������ ���� ����</strong> <font color=red>(������ ���� ��꼭�� <b>�ٹ�����</b>���� <b>�ϰ� ����</b>�մϴ�.)</font></td>
</tr>
<% if (ojungsanTaxCC.FresultCount>0) then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >�����</td>
    <td width="60" >����ó</td>
    <td width="50" >����<br>����</td>
    <td width="90" >�귣��ID</td>
    <td width="180" >���곻��</td>
    <td width="90" >���»� �����<br>��ǰ</td>
    <td width="80" >������</td>
    <td width="80" >���»� �����<br>��ۺ�</td>
    <td width="100">���޴���<br>(��ǰ)</td>
  	<td width="80">���޴���<br>(��ۺ�)</td>
  	<td width="80">�߰������<br>(��Ÿ)</td>
  	<td width="80">���޿�����</td>
    <!--td width="60" >���޿�����</td-->
    <td width="90" >��꼭����</td>
    <td width="80" >��꼭</td>
    <td >����ȸ</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(A)</td>
    <td>(B)</td>
    <td>(C)</td>
    <td>(D)</td>
    <td>(E)</td>
    <td>(F)</td>
    <td>(G)</td>
    <td>(H)</td>
    <td>(I)</td>
    <td>(J)</td>
</tr>
<% for i=0 to ojungsanTaxCC.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxCC.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxCC.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxCC.FItemList(i).getItemVatTypeName%></td>
    <td align="left"><%=ojungsanTaxCC.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxCC.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FdlvMeachulsum + ojungsanTaxCC.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FdlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).FetJungsanSum,0) %></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCC.FItemList(i).getToTalJungsanSum,0)%></td>
    <!--td><%= ojungsanTaxCC.FItemList(i).getMayIpkumdateStr %></td -->
    <td><%=ojungsanTaxCC.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxCC.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxCC.FItemList(i).Fneotaxno %>');">���
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
		<% else %>
      	<!--<a href="javascript:PopTaxRegPrdCommission('<%'=ojungsanTaxCC.FItemList(i).Fmakerid %>', '<%'= yyyy1 %>', '<%'= mm1 %>', '<%'= ojungsanTaxCC.FItemList(i).FtargetGbn %>','<%'= ojungsanTaxCC.FItemList(i).Fid %>');">����-->
      	<!--<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">-->
      	<!--</a>-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxCC.FItemList(i).FId %>','<%= ojungsanTaxCC.FItemList(i).FtargetGbn%>','<%= ojungsanTaxCC.FItemList(i).Fgroupid%>');">����<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="14">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
<p><br>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="14" align="left"><strong>* ��Ÿ ���� ����</strong> <font color=red>(��Ÿ ���� ���� ��꼭�� <b>�ٹ�����</b>���� <b>�ϰ� ����</b>�մϴ�.)</font></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >�����</td>
    <td width="60" >����ó</td>
    <td width="50" >����<br>����</td>
    <td width="90" >�귣��ID</td>
    <td width="180" >���곻��</td>
    <td width="90" ></td>
    <td width="80" >���θ��<br>(���»� �δ�)</td>
    <td width="80" ></td>
    <td width="100"></td>
  	<td width="80"></td>
  	<td width="80">���޿�����</td>
    <td width="90" >��꼭����</td>
    <td width="80" >��꼭</td>
    <td >����ȸ</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(B)</td>
    <td></td>
    <td></td>
    <td></td>
    <td>(F)</td>
    <td>(G)</td>
    <td>(H)</td>
    <td>(I)</td>
</tr>
<% if (ojungsanTaxCE.FresultCount>0) then %>
<% for i=0 to ojungsanTaxCE.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxCE.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxCE.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxCE.FItemList(i).getItemVatTypeName%></td>
    <td align="left"><%=ojungsanTaxCE.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxCE.FItemList(i).Ftitle%></td>
    <td align="right"></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCE.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"></td>
    <td align="right"></td>
    <td align="right"></td>
    <td align="right"><%=FormatNumber(ojungsanTaxCE.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTaxCE.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxCE.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxCE.FItemList(i).Fneotaxno %>');">���
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
		<% else %>
      	<!--<a href="javascript:PopTaxRegPrdCommission('<%'=ojungsanTaxCE.FItemList(i).Fmakerid %>', '<%'= yyyy1 %>', '<%'= mm1 %>', '<%'= ojungsanTaxCE.FItemList(i).FtargetGbn %>','<%'= ojungsanTaxCE.FItemList(i).Fid %>');">����-->
      	<!--<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">-->
      	<!--</a>-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxCE.FItemList(i).FId %>','<%= ojungsanTaxCE.FItemList(i).FtargetGbn%>','<%= ojungsanTaxCE.FItemList(i).Fgroupid%>');">����<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="14">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
<p><br>



<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="13" align="left"><strong>* ���� ���� ����</strong> (���»翡�� �ٹ��������� ������ �ּž� �մϴ�.) (�Ե����� �Ǹ� ���� �� ������ �Ǹ� ������ ������������ ó�� �˴ϴ�.)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" >�����</td>
    <td width="60" >����ó</td>
    <td width="50" >����<br>����</td>
    <td width="90" >�귣��ID</td>
    <td width="170" >���곻��</td>
    <td width="90" >���»�<br>��ǰ���޾�</td>
    <td width="80" >��ۺ�/��Ÿ</td>
    <td width="100">���޴���<br>(��ǰ)</td>
  	<td width="80">���޴���<br>(��ۺ�/��Ÿ)</td>
  	<td width="80">���»�����<br>(���޿�����)</td>
    <!--td width="60" >���޿�����</td-->
    <td width="90" >��꼭����</td>
    <td width="80" >��꼭</td>
    <td >����ȸ</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" >
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>(a)</td>
    <td>(b)</td>
    <td>(c)</td>
    <td>(d)</td>
    <td>(e)</td>
    <td>(f)</td>
    <td>(g)</td>
    <td>(h)</td>
</tr>
<% if (ojungsanTaxMM.FresultCount>0) then %>
<% for i=0 to ojungsanTaxMM.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=ojungsanTaxMM.FItemList(i).Fyyyymm%></td>
    <td><%=ojungsanTaxMM.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTaxMM.FItemList(i).getTaxtypeName%></td>
    <td align="left"><%=ojungsanTaxMM.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTaxMM.FItemList(i).Ftitle%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FdlvMeachulsum + ojungsanTaxMM.FItemList(i).FetMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).FdlvJungsanSum + ojungsanTaxMM.FItemList(i).FetJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTaxMM.FItemList(i).getToTalJungsanSum,0)%></td>
    <!--td><%= ojungsanTaxMM.FItemList(i).getMayIpkumdateStr %></td-->
    <td><%=ojungsanTaxMM.FItemList(i).GetTaxEvalStateName%></td>
    <td>
        <% if ojungsanTaxMM.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTaxMM.FItemList(i).Fneotaxno %>');">���
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsanTaxMM.FItemList(i).IsCommissionTax then %>
      	</a>
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg<%=CHKIIF(ojungsanTaxMM.FItemList(i).FtargetGbn="OF","Off","")%>('<%= ojungsanTaxMM.FItemList(i).FId %>');">����
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecFreeTaxCase then %>
      	<!--
      	<a href="javascript:PopTaxReg('<%= ojungsanTaxMM.FItemList(i).FId %>');">����
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% elseif ojungsanTaxMM.FItemList(i).IsElecSimpleBillCase then %>
      	<!--
      	<a href="javascript:PopConfirm('<%= menupos %>','<%= ojungsanTaxMM.FItemList(i).FId %>');">����Ȯ��
      	<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% end if %>
    </td>
    <td>
    <a href="javascript:PopDetail('<%= ojungsanTaxMM.FItemList(i).FId %>','<%= ojungsanTaxMM.FItemList(i).FtargetGbn%>','<%= ojungsanTaxMM.FItemList(i).Fgroupid %>');">����<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a>
    </td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="#FFFFFF" height="30">
<td align="center" colspan="13">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<% end if %>


<p><br><br>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">�� ���������곻��</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">����</td>
    <td>�� ��</td>
    <td width="200">��Ÿ</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td width="200">���Ǹűݾ�(���»�����)</td>
    <td width="40" align="center" >(A)</td>
    <td >���»簡 �ٹ����� ����Ʈ�� ���� �Ǹ��� �����(�ΰ����Ű�� ����Ű�ݾ�)</td>
    <td width="220">��꼭 �������� ����</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>������</td>
    <td align="center">(B)</td>
    <td>�ǸŴ�������� ���ٹ����� �����(�ٹ�����>>���»�� ���ݰ�꼭 ����)</td>
    <td>���ݰ�꼭 ����</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��ۺ�/��Ÿ�Ǹűݾ�</td>
    <td align="center">(C)</td>
    <td>�ٹ��������� �Աݵ� ��ۺ� + ��Ÿ�����	</td>
    <td>��꼭 �������� ����</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���޴���(��ǰ)</td>
    <td align="center">(D)</td>
    <td>(D)=(A)-(B) ��ǰ�Ǹſ� �ݾ׿� ���� �����-������</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���޴���(��ۺ�/��Ÿ)</td>
    <td align="center">(E)</td>
    <td>�ٹ����ٿ��� ���»�� �����ؾ��� ��ۺ� + ��Ÿ�����</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���޿�����</td>
    <td align="center">(F)</td>
    <td>(F)=(A)-(B)+(E) ���»�� ������ �Ѿ�(��ü�����)</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭����</td>
    <td align="center">(G)</td>
    <td>���»� ���� �� Ȯ�� >> ����Ȯ�� ���ݰ�꼭	�Ϳ� 5�� �ϰ������</td>
    <td>�Ϳ� 5�� �ϰ������</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭</td>
    <td align="center">(H)</td>
    <td>�ٹ�����>>���»� ����� ���ݰ�꼭 �ǹ� Ȯ�� �� ���</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>����ȸ</td>
    <td align="center">(I)</td>
    <td>���꿡 ���� �󼼳��� ��ȸ</td>
    <td></td>
</tr>
</table>
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">�� ��Ÿ���곻��</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">����</td>
    <td>�� ��</td>
    <td width="200">��Ÿ</td>
</tr>

<tr align="left" bgcolor="#FFFFFF">
    <td>���θ�� (���»� �δ�)</td>
    <td align="center">(B)</td>
    <td>��ü �δ� ���θ�� ���</td>
    <td>���ݰ�꼭 ����</td>
</tr>

<tr align="left" bgcolor="#FFFFFF">
    <td>���޿�����</td>
    <td align="center">(F)</td>
    <td>���꿡�� ������ �ݾ�</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭����</td>
    <td align="center">(G)</td>
    <td>���»� ���� �� Ȯ�� >> ����Ȯ�� ���ݰ�꼭	�Ϳ� 5�� �ϰ������</td>
    <td>�Ϳ� 5�� �ϰ������</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭</td>
    <td align="center">(H)</td>
    <td>�ٹ�����>>���»� ����� ���ݰ�꼭 �ǹ� Ȯ�� �� ���</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>����ȸ</td>
    <td align="center">(I)</td>
    <td>���꿡 ���� �󼼳��� ��ȸ</td>
    <td></td>
</tr>
</table>
<p>


<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="4" align="left">�� �������곻��</td>
</tr>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" width="240">����</td>
    <td>�� ��</td>
    <td width="200">��Ÿ</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td width="200">���»� ��ǰ���޾�</td>
    <td width="40" align="center" >(a)</td>
    <td >���»翡�� �ٹ��������� ������ ��ǰ����</td>
    <td width="220"></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��ۺ�/��Ÿ</td>
    <td align="center">(b)</td>
    <td>�ٹ��������� �Աݵ� ��ۺ� + ��Ÿ�����</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���޴���(��ǰ)</td>
    <td align="center">(c)</td>
    <td>��ǰ���޿� ���� �����</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���޴���(��ۺ�/��Ÿ)</td>
    <td align="center">(d)</td>
    <td>�ٹ����ٿ��� ���»�� �����ؾ��� ��ۺ� + ��Ÿ�����</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>���»�����(���޿�����)</td>
    <td align="center">(e)</td>
    <td>(e)=(c)+(d) ���»�� ������ �Ѿ�(��ü�����)</td>
    <td>���ݰ�꼭 ����(���»�>>�ٹ�����)</td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭����</td>
    <td align="center">(f)</td>
    <td>��üȮ�� �� ���ݰ�꼭 ����� : ����Ȯ�� / �̹���� : ��üȮ�δ��</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>��꼭</td>
    <td align="center">(g)</td>
    <td>���»�>>�ٹ����� ����� ���ݰ�꼭 �ǹ� Ȯ�� �� ���</td>
    <td></td>
</tr>
<tr align="left" bgcolor="#FFFFFF">
    <td>����ȸ</td>
    <td align="center">(h)</td>
    <td>���꿡 ���� �󼼳��� ��ȸ</td>
    <td></td>
</tr>
</table>
<p>

<%
set ojungsanTaxCC = Nothing
set ojungsanTaxCE = Nothing
set ojungsanTaxMM = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
