<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���ݰ�꼭 ���� ��Ȳ
' Hieditor : 2019.04.04 ������ ����
'            2022.11.02 �ѿ�� ����(���ݰ�꼭 ���� ��36524 api -> ���ϰ� api �� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim page, makerid, yyyy1, mm1, finishflag, jgubun, groupid, targetGbn, jungsan_date, jungsan_gubun
page    = requestCheckvar(request("page"),10)
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
groupid = requestCheckvar(request("groupid"),10)
finishflag = requestCheckVar(request("finishflag"),10)
jgubun   = requestCheckVar(request("jgubun"),10)
targetGbn = requestCheckVar(request("targetGbn"),10)
jungsan_date = requestCheckvar(request("jungsan_date"),10)
jungsan_gubun = requestCheckvar(request("jungsan_gubun"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

if page="" then page=1

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 70
ojungsanTax.FCurrPage = page
ojungsanTax.FRectMakerid = makerid
ojungsanTax.FRectJGubun = jgubun
ojungsanTax.FRectgroupid = groupid
ojungsanTax.FRecttargetGbn = targetGbn
ojungsanTax.FRectFinishFlag = finishflag
ojungsanTax.FRectJungsanException = jungsan_gubun
ojungsanTax.FRectJungsanDate = jungsan_date

if (makerid="") then
    ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
end if
ojungsanTax.getJungsanTaxListAdm

dim i
dim commCnt : commCnt =0
dim isEvalEnabledTax

%>
<script type='text/javascript'>

function goMonthJungsan(yyyy,mm,jid,makerid){
    location.href='monthjungsanAdm.asp?menupos=1647&yyyy1='+yyyy+'&mm1='+mm +'&makerid='+makerid;
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("/designer/jungsan/red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=800,height=700,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}



function NextPage(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function PopTaxRegPrdCommission(makerid, yyyy1, mm1, onoffGubun, jidx) {
	var popwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx,"PopTaxRegPrdCommission","width=640 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkAll(comp){
    var frm = document.frmList;


    if (frm.chk){
        if (frm.chk.length){
            for(var i=0;i<frm.chk.length;i++){
                if (!frm.chk[i].disabled){
                    frm.chk[i].checked = comp.checked;
                }
            }
        }else{
            if (!frm.chk.disabled){
                frm.chk.checked = comp.checked;
            }
        }
    }

}

function evalOneTax(pp){
    var frm = document.frmList;
    var jidx=0;
    var makerid='';
    var yyyy1='';
    var mm1='';
    var onoffGubun='';
    var nextid = 0;

    if (frm.chk.length){
        for(var i=frm.chk.length-1;i>=0;i--){

            if ((frm.chk[i].checked)&&(frm.id[i].value==pp)){
                jidx = frm.id[i].value;
                makerid= frm.makerid[i].value;
                yyyy1= frm.yyyy1[i].value;
                mm1= frm.mm1[i].value;
                onoffGubun= frm.targetGbn[i].value;
                break;
            }else{
                if (frm.chk[i].checked) { nextid = frm.id[i].value; }
            }
        }
    }else{
        if ((frm.chk.checked)&&(frm.id.value==pp)){
            jidx = frm.id.value;
            makerid= frm.makerid.value;
            yyyy1= frm.yyyy1.value;
            mm1= frm.mm1.value;
            onoffGubun= frm.targetGbn.value;
        }
    }

    //alert(jidx)
    //alert(nextid)
    //alert(makerid);
    if ((jidx!=0)){
        var evalwin = window.open("/admin/upchejungsan/popTaxRegAdminapi.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx+"&isauto=on&nextjidx="+nextid,"PopTaxRegPrdCommissionAuto","width=1200 height=768 scrollbars=yes resizable=yes");
        <% 'var evalwin = window.open("popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx+"&isauto=on&nextjidx="+nextid,"PopTaxRegPrdCommissionAuto","width=640 height=700 scrollbars=yes resizable=yes"); %>
        evalwin.focus();
    }else{
        alert('Finish '+jidx+' , '+nextid);
        location.reload();
    }
}

function evalOneTax_V2(pp){
    var frm = document.frmList;
    var jidx=0;
    var makerid='';
    var yyyy1='';
    var mm1='';
    var onoffGubun='';
    var nextid = 0;

    if (frm.chk.length){
        for(var i=frm.chk.length-1;i>=0;i--){

            if ((frm.chk[i].checked)&&(frm.id[i].value==pp)){
                jidx = frm.id[i].value;
                makerid= frm.makerid[i].value;
                yyyy1= frm.yyyy1[i].value;
                mm1= frm.mm1[i].value;
                onoffGubun= frm.targetGbn[i].value;
                break;
            }else{
                if (frm.chk[i].checked) { nextid = frm.id[i].value; }
            }
        }
    }else{
        if ((frm.chk.checked)&&(frm.id.value==pp)){
            jidx = frm.id.value;
            makerid= frm.makerid.value;
            yyyy1= frm.yyyy1.value;
            mm1= frm.mm1.value;
            onoffGubun= frm.targetGbn.value;
        }
    }

    //alert(jidx)
    //alert(nextid)
    //alert(makerid);
    if ((jidx!=0)){
        var evalwin = window.open("/admin/upchejungsan/popUWehagotaxregapi.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx+"&isauto=on&nextjidx="+nextid,"PopTaxRegPrdCommissionAuto","width=1200 height=768 scrollbars=yes resizable=yes");
        evalwin.focus();
    }else{
        alert('Finish '+jidx+' , '+nextid);
        location.reload();
    }
}

// ���ó��� �ϰ� ����
function batchEval_V2(){
    var frm = document.frmList;
    var chkCNT = 0;
    var pp =-1;

    if (frm.chk){
        <% if ojungsanTax.FResultCount="1" then %>
            if (frm.chk.checked){
                pp=0;
                chkCNT++;
            }
        <% else %>
        if (frm.chk.length){
            for(var i=0;i<frm.chk.length;i++){
                if (frm.chk[i].checked){
                    if (pp==-1) {
                        pp=i;
                    }

                    chkCNT++;
                }
            }
        }else{
            if (frm.chk.checked){
                chkCNT++;
            }
        }
        <% end if %>
    }

    if (chkCNT<1){
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm(chkCNT+'�� �ϰ� ���� �Ͻðڽ��ϱ�?')){
        <% if ojungsanTax.FResultCount="1" then %>
            evalOneTax_V2(frm.id.value);
        <% else %>
            evalOneTax_V2(frm.id[pp].value);
        <% end if %>
    }
}

/*
// ���ó��� �ϰ� ����
function batchEval(){
    var frm = document.frmList;
    var chkCNT = 0;
    var pp =-1;

    if (frm.chk){
        <% if ojungsanTax.FResultCount="1" then %>
            if (frm.chk.checked){
                pp=0;
                chkCNT++;
            }
        <% else %>
        if (frm.chk.length){
            for(var i=0;i<frm.chk.length;i++){
                if (frm.chk[i].checked){
                    if (pp==-1) {
                        pp=i;
                    }

                    chkCNT++;
                }
            }
        }else{
            if (frm.chk.checked){
                chkCNT++;
            }
        }
        <% end if %>
    }

    if (chkCNT<1){
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm(chkCNT+'�� �ϰ� ���� �Ͻðڽ��ϱ�?')){
        <% if ojungsanTax.FResultCount="1" then %>
            evalOneTax(frm.id.value);
        <% else %>
        evalOneTax(frm.id[pp].value);
        <% end if %>
    }
}

// ���������� �ϰ� ����
function popCCbatchEval(){
    var yyyy1='';
    var mm1='';
    var jungsan_gubun='';
    var targetGbn='';
    yyyy1= frm.yyyy1.value;
    mm1= frm.mm1.value;
    targetGbn= frm.targetGbn.value;
    if (frm.jungsan_gubun.checked){
        jungsan_gubun='on';
    }

    var popwin = window.open("/admin/upchejungsan/evalCCBatch_utf8.asp?yyyy1="+yyyy1+"&mm1="+mm1+"&jungsan_gubun="+jungsan_gubun+"&targetGbn="+targetGbn,"evalCCBatch","width=1200 height=900 scrollbars=yes resizable=yes");
	popwin.focus()
 
}
*/

// ���������� �ϰ� ����
function popCCbatchEval_V2(){
    var yyyy1='';
    var mm1='';
    var jungsan_gubun='';
    var targetGbn='';
    yyyy1= frm.yyyy1.value;
    mm1= frm.mm1.value;
    targetGbn= frm.targetGbn.value;
    if (frm.jungsan_gubun.checked){
        jungsan_gubun='on';
    }

    var popwin = window.open("/admin/upchejungsan/evalCCBatch_V2_utf8.asp?yyyy1="+yyyy1+"&mm1="+mm1+"&jungsan_gubun="+jungsan_gubun+"&targetGbn="+targetGbn,"evalCCBatch","width=1200 height=900 scrollbars=yes resizable=yes");
	popwin.focus()
 
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		���� ��� ��� :&nbsp;<% DrawYMBox yyyy1,mm1 %>
		&nbsp;&nbsp;
		�귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>&nbsp;&nbsp;
		&nbsp;&nbsp;
        ��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >

	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">
    ����
		<select name="finishflag" >
		<option value="">��ü
		<option value="0" <%= CHKIIF(finishflag="0","selected","") %> >������
		<option value="1" <%= CHKIIF(finishflag="1","selected","") %> >��üȮ�δ��
		<option value="2" <%= CHKIIF(finishflag="2","selected","") %> >��üȮ�οϷ�
		<option value="3" <%= CHKIIF(finishflag="3","selected","") %> >����Ȯ��
		<option value="7" <%= CHKIIF(finishflag="7","selected","") %> >�ԱݿϷ�
		</select>
		&nbsp;
		�����ı��� :
        <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;
        ON/AC/OF ���� :
        <select name="targetGbn" >
		<option value="">��ü
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>

		&nbsp;
        ������ :
        <select name="jungsan_date">
        <option value="" <% if jungsan_date="" then response.write "selected" %> >����
        <option value="15��" <% if jungsan_date="15��" then response.write "selected" %> >15��
        <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
        <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
        </select>
        &nbsp;
        <input type="checkbox" name="jungsan_gubun"<% if jungsan_gubun="on" then response.write " checked"%>> �������� ����(�ؿ�) ����
    </td>
</tr>
</table>
</form>
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        <% '<input type="button" value="���ó����ϰ�����(�� ��36524)" onClick="batchEval();" class="button" > %>
        <% '<input type="button" value="���������� �ϰ� ����(�� ��36524)" onClick="popCCbatchEval();" class="button" > %>
        <input type="button" value="���ó����ϰ�����(���ϰ�)" onClick="batchEval_V2();" class="button" >
        <input type="button" value="���������� �ϰ� ����(���ϰ�)" onClick="popCCbatchEval_V2();" class="button" >
    </td>
    <td align="right">	
    </td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frmList" style="margin:0px;" >
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="16">
        �˻���� : <b><%= ojungsanTax.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %>/ <%= ojungsanTax.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20" ><input type="checkbox" name="chkALL" onClick="chkAll(this)"></td>
    <td width="60" >�����</td>
    <td width="60" >����ó</td>
    <td width="70" >������</td>
    <td width="100" >��꼭����</td>
    <td width="80" >�׷��ڵ�</td>
    <td width="80" >�귣��ID</td>
    <td width="120" >����</td>
    <td width="80" >������</td>
    <td width="80" >���ް���</td>
    <td width="80" >�ΰ���</td>
    <td width="80" >�հ�</td>
    <td width="90" >�������</td>
    <td width="70">���곻��</td>
    <td width="60">������<br>(������)</td>
    <td >���</td>

</tr>
<% for i=0 to ojungsanTax.FResultCount-1 %>
<%
if (ojungsanTax.FItemList(i).IsCommissionTax) then
    commCnt=commCnt+1
end if

isEvalEnabledTax = (ojungsanTax.FItemList(i).Ffinishflag=1 or ojungsanTax.FItemList(i).Ffinishflag=2) and ojungsanTax.FItemList(i).IsCommissionTax
isEvalEnabledTax = isEvalEnabledTax and ojungsanTax.FItemList(i).Fgroupid<>"G00617"
isEvalEnabledTax = isEvalEnabledTax and ojungsanTax.FItemList(i).Fgroupid<>"G00490"
isEvalEnabledTax = isEvalEnabledTax and ojungsanTax.FItemList(i).Fgroupid<>"G03633"
isEvalEnabledTax = isEvalEnabledTax and ojungsanTax.FItemList(i).Fgroupid<>"G04703"

isEvalEnabledTax = isEvalEnabledTax and ojungsanTax.FItemList(i).getJungsanTaxSum<>0

IF application("Svr_Info")="Dev" THEN
    isEvalEnabledTax=true
end if
%>
<tr bgcolor="#FFFFFF" align="center">
    <td>
        <input type="checkbox" name="chk" value="<%=i%>" <%=CHKIIF(isEvalEnabledTax,"","disabled")%> >
        <input type="hidden" name="id" value="<%= ojungsanTax.FItemList(i).Fid %>">
        <input type="hidden" name="makerid" value="<%= ojungsanTax.FItemList(i).Fmakerid %>">
        <input type="hidden" name="yyyy1" value="<%=LEFT(ojungsanTax.FItemList(i).FYYYYMM,4)%>">
        <input type="hidden" name="mm1" value="<%=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>">
        <input type="hidden" name="targetGbn" value="<%= ojungsanTax.FItemList(i).FtargetGbn %>">
    </td>
    <td><%=ojungsanTax.FItemList(i).FYYYYMM%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxJungsanGubun%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxTypeStrUpcheView%></td>
    <td><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td><%= ojungsanTax.FItemList(i).Ftaxregdate %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSuply,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxVat,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSum,0) %></td>
    <td><%= ojungsanTax.FItemList(i).GetTaxEvalStateName %></td>
    <td><img src="/images/icon_search.jpg" onClick="goMonthJungsan('<%=Left(ojungsanTax.FItemList(i).FYYYYMM,4)%>','<%=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>','<%=ojungsanTax.FItemList(i).Fid%>','<%=ojungsanTax.FItemList(i).Fmakerid%>');" style="cursor:pointer"></td>
    <td><%= ojungsanTax.FItemList(i).getTaxEvalStyleStr %></td>
    <td>
        <% if ojungsanTax.FItemList(i).IsElecTaxExists then %>
        <img style="cursor:pointer" src="/images/icon_print02.gif" onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ojungsanTax.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')">
        <!--
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsanTax.FItemList(i).Fneotaxno %>');">���
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	-->
      	<% else %>
      	<% 'if (ojungsanTax.FItemList(i).Ffinishflag=1 or ojungsanTax.FItemList(i).Ffinishflag=2) and (ojungsanTax.FItemList(i).IsCommissionTax) then %>
      	<!--<a href="javascript:PopTaxRegPrdCommission('<%'= ojungsanTax.FItemList(i).Fmakerid %>', '<%'=LEFT(ojungsanTax.FItemList(i).FYYYYMM,4)%>', '<%'=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>', '<%'= ojungsanTax.FItemList(i).FtargetGbn %>','<%'= ojungsanTax.FItemList(i).Fid %>');">����-->
      	<!--<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">-->
      	<% 'end if %>
      	</a>
      	<% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="16" align="center">
        <% if ojungsanTax.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ojungsanTax.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojungsanTax.StartScrollPage to ojungsanTax.FScrollCount + ojungsanTax.StartScrollPage - 1 %>
			<% if i>ojungsanTax.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojungsanTax.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
</table>
</form>
<%
set ojungsanTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->