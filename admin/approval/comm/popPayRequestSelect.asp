<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ڰ��� �ŷ�ó ����
' History : 2011.12.05 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payreqListCls.asp"--> 
<%
Dim frmName : frmName=requestCheckVar(request("frmName"),32)

Dim clsPay
Dim arrList, iTotCnt
set clsPay = new CPayReqList  
	clsPay.FpayRequestType	= 9
	clsPay.Fpayrequeststate = 1
	'clsPay.FisTakeDoc		=blnTakeDoc
	'clsPay.FRegID			=searchRegID  
	'clsPay.FSearchType		= searchType
	'clsPay.FSDate			=searchsdate
	'clsPay.FEDate			=searchedate
 	'clsPay.Farap_cd			=iarap_cd 
	clsPay.FCurrpage 		= 1
	clsPay.FPagesize		= 100
	arrList = clsPay.fnGetPayReqAllList
	iTotCnt = clsPay.FTotCnt 

set clsPay = nothing

dim intLoop
%>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript'>
function jsSubmit(){
    var frm = document.frm;
    
    if ((!frm.rdOpt[0].checked)&&(!frm.rdOpt[1].checked)){
        alert('�ۼ��� ������û����  �����ϼ���. (�ű� �Ǵ� ���� ������ �߰�)');
        return;
    }
    
    if (frm.rdOpt[0].checked){
        if (frm.yyyymmdd.value.length<1){
            alert('���� ��û���� �����ϼ���.');
            return;
        }
    }
    
    if (frm.rdOpt[1].checked){
        if (frm.rIdx.value.length<1){
            alert('�ۼ��� ������û����  �����ϼ���. (�ű� �Ǵ� ���� ������ �߰�)');
            frm.rIdx.focus();
            return;
        }
    }
    
    opener.jsPopSubmit('<%= frmName %>',frm.yyyymmdd.value,frm.rIdx.value);
    window.close();
}

function chkComp(comp){
    if (comp.value=='N'){
        frm.rIdx.disabled=true;
        frm.yyyymmdd.disabled=false;
    }
    
    if (comp.value=='P'){
        frm.rIdx.disabled=false;
        frm.yyyymmdd.disabled=true;
    }
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>������û�� ����</strong><br><hr width="100%"></td>
</tr>
<form name="frm" >
<tr>
    <td>
        <input type="radio" name="rdOpt" value="N" <%= chkIIF(isArray(arrList),"","checked") %> onClick="chkComp(this);"> �ű� ������û���� �ۼ�
        &nbsp;&nbsp; ������û��:
        <input type="text" name="yyyymmdd" size="10" maxlength=10 readonly value="">
        <a href="javascript:calendarOpen(frm.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>		
    </td>
</tr>
<tr>
	<td>
	    <input type="radio" name="rdOpt" value="P" onClick="chkComp(this);"> 
	    <select name="rIdx">
	    <option value="">���� ������û�� �� �߰�
	    <% IF isArray(arrList) THEN 
		    For intLoop = 0 To UBound(arrList,2) %>
	    <option value="<%=arrList(0,intLoop)%>">[<%=arrList(0,intLoop)%>]<%=arrList(3,intLoop)%> (<%=arrList(5,intLoop)%>)
	    <%  next 
	       End IF
	    %>
	    </select>
	</td>
</tr>
<tr>
	<td align="center">
	    <hr width="100%"><br>
	    <input type="button" class="button" value=" Ȯ �� " onClick="jsSubmit();">
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" --> 	
<!-- #include virtual="/admin/lib/poptail.asp"-->