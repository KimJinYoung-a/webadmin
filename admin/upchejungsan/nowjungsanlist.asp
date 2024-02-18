<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰����
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
const CUNDERMargin = 10

dim yyyy1,mm1,designer,rectorder, groupid, finishflag, taxtype, chkmargin, vPurchaseType, jgubun
dim research, page, ix, targetGbn, jacctcd, differencekey, companynoYN
dim searchType, searchText, jungsanGubun

designer = requestCheckVar(request("designer"),32)
yyyy1    = requestCheckVar(request("yyyy1"),4)
mm1      = requestCheckVar(request("mm1"),2)
rectorder = requestCheckVar(request("rectorder"),32)
groupid  = requestCheckVar(request("groupid"),32)
research = requestCheckVar(request("research"),32)
finishflag = requestCheckVar(request("finishflag"),10)
page     = requestCheckVar(request("page"),10)
taxtype  = requestCheckVar(request("taxtype"),10)
chkmargin= requestCheckVar(request("chkmargin"),10)
vPurchaseType = requestCheckVar(request("purchasetype"),2)
jgubun   = requestCheckVar(request("jgubun"),10)
targetGbn = requestCheckVar(request("targetGbn"),10)
jacctcd = requestCheckVar(request("jacctcd"),10)
differencekey = requestCheckVar(request("differencekey"),10)
searchType = requestCheckVar(request("searchType"), 32)
searchText = requestCheckVar(request("searchText"), 32)
companynoYN = requestCheckVar(request("companynoYN"), 1)
jungsanGubun = requestCheckVar(request("jungsanGubun"), 12)

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

if page="" then page=1

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize  = 100
ojungsan.FCurrPage  = page
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsan.FRectDesigner = designer
ojungsan.FRectGroupID = groupid
ojungsan.FrectOrder = rectorder
ojungsan.Frectfinishflag = finishflag
ojungsan.FRectTaxType = taxtype
ojungsan.FRectPurchaseType = vPurchaseType
ojungsan.FRectJGubun = jgubun
ojungsan.FRecttargetGbn = targetGbn
ojungsan.FRectjacctcd = jacctcd
ojungsan.FRectdifferencekey = differencekey
ojungsan.FRectSearchType = searchType
ojungsan.FRectSearchText = searchText
ojungsan.FRectCompanynoYN = companynoYN
ojungsan.FRectJungsanGubun = jungsanGubun


IF (chkmargin="on") then
    ojungsan.FRectUnderMargin = CStr(CUNDERMargin)
end if

if (research<>"") then
    ojungsan.JungsanMasterList
end if

dim i
dim tot1,tot2,tot3,tot4,tot5, totcom, totdlv,totsum
tot1 = 0
tot2 = 0
tot3 = 0
tot4 = 0
tot5 = 0
totsum = 0
%>
<script language='javascript'>
function NextPage(ipage){
    document.frm.page.value=ipage;
    document.frm.submit();
}

function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function PopUpchebrandInfo(v){
	var popwin = window.open("/admin/lib/popupchebrandinfo.asp?designer=" + v,"popupchebrandinfo","width=640 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

function popDetail(v){
	window.open('popdetail.asp?id=' + v );
}

function dellThis(v){
	var upfrm = document.frmarr;
	var ret = confirm('��� ���� �����͸� ���� �Ͻðڽ��ϱ�?');
	if (ret){
		upfrm.idx.value = v;
		upfrm.mode.value = "dellall";
		upfrm.submit();
	}
}

function NextStep(idx){
	<%if groupid = "G02856" or groupid = "g02856" then %>
	alert('�ش���� ����Ұ��մϴ�.'); return;
	<%end if%>
 //   if ((idx=="294398")||(idx=="312521")||(idx=="314361")){ alert('�ش���� ����Ұ��մϴ�.'); return; } //2016/09/29��� ��û
    if ((idx=="294398")){ alert('�ش���� ����Ұ��մϴ�.'); return; } //2016/12/01��� ��û(312461 ����)
      if((idx=="354608") || (idx=="354186") || (idx=="380557")){ alert('�ش���� ����Ұ��մϴ�.'); return; } //2017/07/07
	var upfrm = document.frmarr;
	upfrm.mode.value= "statechange";
	upfrm.idx.value= idx;
	upfrm.rd_state.value="1";

	var ret = confirm('Ȯ�� ��� ���·� ���� �Ͻðڽ��ϱ�?<%=groupid%>');
	if (ret){
		upfrm.submit();
	}
}

function MakeBrandBatchJungsan(frm){
    if (frm.jgubun.value.length<1){
        alert('���� ��� ������ ���� �ϼ���.');
        frm.jgubun.focus();
        return;
    }

    if (frm.differencekey.value.length<1){
        alert('���� ������ ���� �ϼ���.');
        frm.differencekey.focus();
        return;
    }

    if (frm.itemvatYN.value.length<1){
        alert('��ǰ ���� ������ ���� �ϼ���.');
        frm.itemvatYN.focus();
        return;
    }

    if (confirm('���곻���� �ۼ� �Ͻðڽ��ϱ�?')){
        var queryurl = 'dodesignerjungsan.asp?mode=brandbatchprocess&jgubun='+frm.jgubun.value+'&designer=' + frm.makerid.value + '&yyyy1=' + frm.yyyy.value + '&mm1=' + frm.mm.value + '&differencekey=' + frm.differencekey.value + '&itemvatYN=' + frm.itemvatYN.value+'&ipchulArr='+frm.ipchulArr.value;

        var popwin = window.open(queryurl ,'on_jungsan_process','width=200, height=200, scrollbars=yes, resizable=yes');
    }
}

//��ü ����
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
	   	   	if(frm.chkitem.disabled==false){
		   	 	frm.chkitem.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					 	if(frm.chkitem[i].disabled==false){
					frm.chkitem[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

 //���� ���� ���º���
function jsMultiStateChange(){
	var frm = document.frmList;
	if(typeof(frm.chkitem) !="undefined"){
	 	if(!frm.chkitem.length){
	 		if(!frm.chkitem.checked){
	 			alert("������ ���� ����� �����ϴ�. ������ �ּ���");
	 			return;
	 		}
	 	}
        else{
            for(i=0;i<frm.chkitem.length;i++){
                if(frm.chkitem[i].checked) {
                    frm.idxarr.value = frm.idxarr.value + frm.chkitem[i].value + ",";
                }
            }
	 		if(frm.idxarr.value==""){
	 			alert("������ ���� ����� �����ϴ�. ������ �ּ���");
	 			return;
	 		}else{
                //alert(frm.idxarr.value);
                frm.submit();
            }
        }
	}
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="<%=rectorder%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">

   	<tr align="center" bgcolor="#FFFFFF" >
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
        <td align="left">
	        	�������� : <% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
				�귣��ID : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
				��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
				<input type="button" class="button" value="Code�˻�" onclick="popSearchGroupID(this.form.name,'groupid');" >&nbsp;&nbsp;
                ���������ڵ� : <input type="text" class="text" name="jacctcd" value="<%= jacctcd %>" size="7" >
	        </td>
        <td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
    		<a href="javascript:document.frm.rectorder.value='';document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
    	</td>
    </tr>
	<tr>
        <td bgcolor="#FFFFFF" >
        	�������� : 
            <% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
			&nbsp;&nbsp;
			����
			<select name="finishflag" >
			<option value="">��ü
			<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >������
			<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >��üȮ�δ��
			<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >��üȮ�οϷ�
			<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >����Ȯ��
			<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >�ԱݿϷ�
			</select>
			&nbsp;&nbsp;
			��꼭��������
			<select name="taxtype" >
			<option value="">��ü
			<option value="01" <%= CHKIIF(taxtype="01","selected","") %> >����
			<option value="02" <%= CHKIIF(taxtype="02","selected","") %> >�鼼
			<option value="03" <%= CHKIIF(taxtype="03","selected","") %> >��õ
			</select>
			&nbsp;&nbsp;
			<input type="checkbox" name="chkmargin" <%= CHKIIF(chkmargin="on","checked","") %>> ���� <%= CUNDERMargin %> % �̸�
			&nbsp;&nbsp;
			�˻�����:
			<select class="select" name="searchType">
				<option></option>
				<option value="socname" <% if (searchType = "socname") then %>selected<% end if %> >��ü��</option>
				<option value="socno" <% if (searchType = "socno") then %>selected<% end if %> >����ڹ�ȣ</option>
			</select>
			&nbsp;
			<input type="text" class="text" name=searchText value="<%= searchText %>" size="15" maxlength="20">
        </td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" >
        �����ı��� :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;&nbsp;
        ON/AC ���� :
        <select name="targetGbn" >
		<option value="">��ü
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>
		&nbsp;&nbsp;
		����
		<input type="text" class="text" name="differencekey" value="<%= differencekey %>" size="2" >
		&nbsp;&nbsp;
		* �ٹ����� ����� ���� : 
        <select name="companynoYN" class="select">
			<option value="">��ü
			<option value="Y" <%= CHKIIF(companynoYN="Y","selected","") %> >����ڸ�
			<option value="N" <%= CHKIIF(companynoYN="N","selected","") %> >���������
		</select>
		&nbsp;&nbsp;
		* ��ü�������� : 
		<select name="jungsanGubun" class="select">
			<option value="" <% if jungsanGubun="" then response.write "selected" %>>��ü</option>
			<option value="�Ϲݰ���" <% if jungsanGubun="�Ϲݰ���" then response.write "selected" %>>�Ϲݰ���</option>
			<option value="���̰���" <% if jungsanGubun="���̰���" then response.write "selected" %>>���̰���</option>
			<option value="��õ¡��" <% if jungsanGubun="��õ¡��" then response.write "selected" %>>��õ¡��</option>
			<option value="�鼼" <% if jungsanGubun="�鼼" then response.write "selected" %>>�鼼</option>
			<option value="����(�ؿ�)" <% if jungsanGubun="����(�ؿ�)" then response.write "selected" %>>����(�ؿ�)</option>
		</select>
        </td>
    </tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p>
<% if (designer<>"") and (yyyy1<>"") and (mm1<>"") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="brandbatch" >
<input type="hidden" name="makerid" value="<%= designer %>">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<tr bgcolor="#FFFFFF">
    <td>
        <select name="jgubun">
        <option value="">���� ��� ����
        <option value="MM">����
        <option value="CC">������
        <option value="CE">��Ÿ����
        </select>
        <select name="differencekey">
        <option value="">���� ����
        <option value="0">0��
        <option value="1">1��
        <option value="2">2��
        <option value="3">3��
        <option value="4">4��
        <option value="5">5��
        <option value="6">6��
        <option value="7">7��
        <option value="8">8��
        <option value="9">9��
        </select>
        <select name="itemvatYN">
        <option value="">��ǰ ���� ���� ����
        <option value="Y">����
        <option value="N">�鼼
        </select>
        <input type="hidden" name="ipchulArr" value="">
        <input type="button" value=" <%= designer %> &nbsp;&nbsp;<%= yyyy1 %>�� <%= mm1 %>�� ���� �ۼ� " onClick="MakeBrandBatchJungsan(document.brandbatch);">
    </td>
</form>
</tr>
</table>
<% end if %>
<% if taxtype="03" then %>
<input type="button" value="���� ����Ȯ��" onclick="jsMultiStateChange();">
<% end if %>
<form name="frmList" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="mode" value="multistatechange">
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" class="a" bgcolor=#BABABA>
    <tr bgcolor="#FFFFFF">
      <td colspan="30" >
      <%= page %>/<%= ojungsan.FTotalPage %> page �� <%=ojungsan.FTotalCount %>��
      </td>
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
      <% if taxtype="03" then %>
      <td width="70"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
      <% end if %>
      <td width="70">�����</td>
      <td width="40">����</td>
      <td width="50">����<br>���</td>
      <td width="50">����<br>����</td>
      <td width="30">����</td>
      <td width="30">����<br>(��꼭)</td>
      <td width="30">����<br>(��ǰ)</td>
      <td width="90"><a href="javascript:research(frm,'designer')">�귣��ID</a></td>
      <td>ȸ���</td>
      <td width="60">��ü���</td>
      <td width="30">����</td>
      <td width="60">�����Ѿ�</td>
      <td width="30">����</td>
      <td width="60">��Ź�Ѿ�</td>
      <td width="30">����</td>
      <td width="60">��Ÿ�Ǹ�</td>
      <td width="30">����</td>
      <td width="70">�Ѽ�����</td>
      <td width="70">��ۺ�/��Ÿ</td>
      <td width="80">�������</td>
      <td width="80"><a href="javascript:research(frm,'state')">����</a></td>
      <td width="70">���ݰ�꼭<br>�����</td>
      <td width="70"><a href="javascript:research(frm,'segum')">���ݹ�����</a></td>
      <td width="70">�Ա���</td>
      <td width="20">E</td>
      <td width="20">S</td>
      <td width="50"><a href="javascript:research(frm,'tax')">��������</a></td>
      <td width="30">����</td>
      <td width="30">���</td>
    </tr>
<% if ojungsan.FResultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="30" align="center" height="30">
        <% if research="" then %>
            [�˻� ��ư�� �����ּ���.]
        <% else %>
            [�˻� ����� �����ϴ�.]
        <% end if %>
        </td>
    </tr>
<% else %>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	tot1 = tot1 + ojungsan.FItemList(i).Fub_totalsuplycash
    	tot2 = tot2 + ojungsan.FItemList(i).Fme_totalsuplycash
    	tot3 = tot3 + ojungsan.FItemList(i).Fwi_totalsuplycash
    	tot4 = tot4 + ojungsan.FItemList(i).Fet_totalsuplycash
    	tot5 = tot5 + ojungsan.FItemList(i).Fsh_totalsuplycash

    	totcom = totcom + ojungsan.FItemList(i).Ftotalcommission
    	totdlv = totdlv + ojungsan.FItemList(i).Fdlv_totalsuplycash

    %>
   <tr align="center" bgcolor="#FFFFFF">
      <% if taxtype="03" then %>
      <td ><input type="checkbox" name="chkitem" value="<%= ojungsan.FItemList(i).FId %>"<% if ojungsan.FItemList(i).Ffinishflag="1" or ojungsan.FItemList(i).Ffinishflag="2" then %><% else %> disabled<% end if %>></td>
      <% end if %>
      <td ><a target=_blank href="nowjungsanmasteredit.asp?id=<%= ojungsan.FItemList(i).FId %>"><%= ojungsan.FItemList(i).Fyyyymm %>&nbsp;<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
      <td ><%= ojungsan.FItemList(i).FtargetGbn %></td>
      <td ><%= ojungsan.FItemList(i).getJGubunName %></td>
      <td ><%= ojungsan.FItemList(i).Fjacc_nm %></td>
      <td ><%= ojungsan.FItemList(i).Fdifferencekey %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ojungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      <td ><%= ojungsan.FItemList(i).GetItemVatTypeName %></td>
      <td ><a href="javascript:PopBrandInfoEdit('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
      <td align="left"><a href="javascript:PopUpcheInfoEdit('<%= ojungsan.FItemList(i).FGroupID %>')"><%= ojungsan.FItemList(i).Fcompany_name %></a></td>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=upche"><%= FormatNumber(ojungsan.FItemList(i).Fub_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fub_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fub_totalsuplycash/ojungsan.FItemList(i).Fub_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=maeip"><%= FormatNumber(ojungsan.FItemList(i).Fme_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fme_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fme_totalsuplycash/ojungsan.FItemList(i).Fme_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witaksell"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fwi_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fwi_totalsuplycash/ojungsan.FItemList(i).Fwi_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witakchulgo"><%= FormatNumber(ojungsan.FItemList(i).Fet_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fet_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fet_totalsuplycash/ojungsan.FItemList(i).Fet_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).Ftotalcommission,0) %></td>
      <td align="right"><a target=_blank href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=DL"><%= FormatNumber(ojungsan.FItemList(i).Fdlv_totalsuplycash,0) %></a></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font>
	  <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      <a href="javascript:NextStep('<%= ojungsan.FItemList(i).FId %>');">
     <img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      </a>
      <% end if %>
      </td>
	    <% if IsNULL(ojungsan.FItemList(i).Ftaxinputdate) then %>
	    <td ></td>
  	    <% else %>
 	    <td ><%= Left(Cstr(ojungsan.FItemList(i).Ftaxinputdate),10) %></td>
  	    <% end if %>
      <% if isNull(ojungsan.FItemList(i).Ftaxregdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Ftaxregdate),10) %></td>
      <% end if %>
      <% if isNull(ojungsan.FItemList(i).Fipkumdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Fipkumdate),10) %></td>
      <% end if %>
      <td ><a href="javascript:PopCSMailSend('<%= ojungsan.FItemList(i).FDesignerEmail %>','','');"><% if ojungsan.FItemList(i).FDesignerEmail<>"" then response.write "E" %></a></td>
      <td ><a href="javascript:PopCSSMSSend('<%= ojungsan.FItemList(i).Fjungsan_hp %>','','','');"><% if ojungsan.FItemList(i).Fjungsan_hp<>"" then response.write "S" %></a></td>
      <td ><%= ojungsan.FItemList(i).Fjungsan_gubun %></td>
      <td ><%= ojungsan.FItemList(i).Fjungsan_date %></td>
      <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      	<td ><a href="javascript:dellThis('<%= ojungsan.FItemList(i).FId %>')">x</a></td>
      <% else %>
        <td >
            <% if Not IsNULL(ojungsan.FItemList(i).FTaxLinkidx) then %>
      	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ojungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
      	   <% else %>
      	        <%= ojungsan.FItemList(i).Fbillsitecode %>
      	    <% end if %>

      	    <a href="/admin/upchejungsan/monthjungsanAdm.asp?makerid=<%= ojungsan.FItemList(i).Fdesignerid %>&yyyy1=<%= LEFT(ojungsan.FItemList(i).Fyyyymm,4) %>&mm1=<%= right(ojungsan.FItemList(i).Fyyyymm,2) %>" target="_blank">POP</a>
        </td>
      <% end if %>
    </tr>
    <% next %>
<% end if %>
    <% totsum = totsum + tot1 + tot2 + tot3 + tot4 + tot5 +totdlv %>
    <tr bgcolor="#FFFFFF" align="right">
      <% if taxtype="03" then %>
      <td></td>
      <% end if %>
      <td>�հ�</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td><%= FormatNumber(tot1,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot2,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot3,0) %></td>
      <td></td>
      <td><%= FormatNumber(tot4,0) %></td>
      <td></td>
      <td><%= FormatNumber(totcom,0) %></td>
      <td><%= FormatNumber(totdlv,0) %></td>
      <td><%= FormatNumber(totsum,0) %></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr bgcolor="#FFFFFF" >
        <td colspan="30" align="center">
            <% if ojungsan.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojungsan.StarScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojungsan.StarScrollPage to ojungsan.FScrollCount + ojungsan.StarScrollPage - 1 %>
				<% if ix>ojungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojungsan.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
    </tr>
</table>
</form>
<form name="frmarr" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="rd_state" value="">
</form>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
