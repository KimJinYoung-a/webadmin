<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �Ա����ϸ�ϻ���/����
' History : 2012.1.30 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
Dim targetGbn : targetGbn=requestCheckVar(request("targetGbn"),32)
Dim frmName : frmName=requestCheckVar(request("frmName"),32)


Dim intLoop
Dim arrList

Dim sqlStr, ipFileName

sqlStr = "select M.ipFileNo,M.ipFileName,M.ipFileRegdate,M.ipFileState, (select count(*) as CNT from db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail D with (nolock) where D.ipFileNo=M.ipFileNo) as CNT"
sqlStr = sqlStr & " From db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M with (nolock)"
sqlStr = sqlStr & " where M.ipFileState<1"
''if (targetGbn<>"") then
   sqlStr = sqlStr & " and M.ipfileGbn='"&targetGbn&"'" 
''end if
sqlStr = sqlStr & " order by M.ipFileNo desc"

rsget.Open sqlStr,dbget,1
IF Not rsget.Eof THEN
    arrList = rsget.getRows
ENd IF
rsget.Close
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
        if (frm.ipFileNo.value.length<1){
            alert('�ۼ��� ������û����  �����ϼ���. (�ű� �Ǵ� ���� ������ �߰�)');
            frm.ipFileNo.focus();
            return;
        }
    }
    
    opener.jsPopSubmitFile('<%= frmName %>',frm.yyyymmdd.value,frm.ipFileNo.value);
    window.close();
}

function chkComp(comp){
    if (comp.value=='N'){
        frm.ipFileNo.disabled=true;
        frm.yyyymmdd.disabled=false;
    }
    
    if (comp.value=='P'){
        frm.ipFileNo.disabled=false;
        frm.yyyymmdd.disabled=true;
    }
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�Ա�File ����</strong><br><hr width="100%"></td>
</tr>
<form name="frm" >
<input type="hidden" name="targetGbn" value="<%= targetGbn %>">
<input type="hidden" name="frmName" value="<%= frmName %>">
<tr>
    <td>
        <input type="radio" name="rdOpt" value="N" <%= chkIIF(isArray(arrList),"","checked") %> onClick="chkComp(this);"> �ű� �Ա�File�� �ۼ�
        &nbsp;&nbsp; ������û��:
        <input type="text" name="yyyymmdd" size="10" maxlength=10 readonly value="">
        <a href="javascript:calendarOpen(frm.yyyymmdd);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>		
    </td>
</tr>
<tr>
	<td>
	    <input type="radio" name="rdOpt" value="P" onClick="chkComp(this);"> 
	    <select name="ipFileNo">
	    <option value="">���� �Ա�File�� �߰�
	    <% IF isArray(arrList) THEN 
		    For intLoop = 0 To UBound(arrList,2) %>
	    <option value="<%=arrList(0,intLoop)%>">[<%=arrList(0,intLoop)%>]<%=arrList(1,intLoop)%> (<%=arrList(2,intLoop)%> �ۼ�) <%=arrList(4,intLoop)%>��
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