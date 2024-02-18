<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
function getURLhtml(url1, isendata,iencoding)
    dim objHttp, ret1
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
    objHttp.Open "GET", url1 & "?" & isendata , False
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.setTimeouts 3000,5000,5000,5000
    On Error Resume Next
    objHttp.Send 
    IF ERR then
        ret1 = "ERR:timeOUT? - "&url1
    end if
    On Error Goto 0


    if (ret1="") then
        if objHttp.status=200 then
            ret1 = Trim(BinaryToText(objHttp.ResponseBody,iencoding))
        else
            ret1 = "ERR:"&CStr(objHttp.status)
        end if
    END IF
    SET objHttp = Nothing

    getURLhtml = ret1
end function

Dim domaingbn : domaingbn = requestCheckvar(request("domaingbn"),10)
Dim orgsip    : orgsip = requestCheckvar(request("orgsip"),10)
Dim DomainArrWEB : DomainArrWEB= ARRAY("stgwww","www1","www2","www3","www4","www5","www6","www7","www8","www9","scm")
Dim DomainArrMOB : DomainArrMOB= ARRAY("stgm","m1","m2","m3","m4","m5","m6","m7","m8","m9","webadmin")

if (application("Svr_Info") = "Dev") then
    DomainArrWEB= ARRAY("2015www","testscm")
    DomainArrMOB= ARRAY("testm","testwebadmin")
end if

' response.write LBOUND(DomainArrWEB) & UBOUND(DomainArrWEB)
' response.write DomainArrWEB(LBOUND(DomainArrWEB)) & DomainArrWEB(UBOUND(DomainArrWEB))
' response.end
Dim retArrW : redim retArrW(UBOUND(DomainArrWEB))
Dim retArrM : redim retArrM(UBOUND(DomainArrMOB))


Dim chgip : chgip = requestCheckvar(request("chgip"),32)
Dim appname : appname = requestCheckvar(request("appname"),32)
Dim chghostname : chghostname = requestCheckvar(request("chghostname"),64)
''변경
IF (request("mode")="chgip") then
    if (chghostname="scm.10x10.co.kr") or (chghostname="webadmin.10x10.co.kr") or (chghostname="testscm.10x10.co.kr") or (chghostname="testwebadmin.10x10.co.kr") then
        Call getURLhtml("http://"&chghostname&"/admin/search/searchIPmanage_man.asp","mode=chgip&appname="&appname&"&chgip="&chgip,"euc-kr")
    else
        Call getURLhtml("http://"&chghostname&"/lib/searchIPmanage.asp","mode=chgip&appname="&appname&"&chgip="&chgip,"UTF-8")
    end if
end if 


Dim i
for i=LBound(DomainArrWEB) to UBound(DomainArrWEB)
    if (domaingbn="W" or domaingbn="") then
        if (DomainArrWEB(i)="scm") or (DomainArrWEB(i)="webadmin") or (DomainArrWEB(i)="testscm") or (DomainArrWEB(i)="testwebadmin") then
            retArrW(i) = getURLhtml("http://"&DomainArrWEB(i)&".10x10.co.kr/admin/search/searchIPmanage_man.asp","orgsip="&orgsip,"euc-kr")
        else
            retArrW(i) = getURLhtml("http://"&DomainArrWEB(i)&".10x10.co.kr/lib/searchIPmanage.asp","orgsip="&orgsip,"UTF-8")
        end if
    end if

    if (domaingbn="M" or domaingbn="") then
        if (DomainArrMOB(i)="scm") or (DomainArrMOB(i)="webadmin") or (DomainArrWEB(i)="testscm") or (DomainArrWEB(i)="testwebadmin") then
            retArrM(i) = getURLhtml("http://"&DomainArrMOB(i)&".10x10.co.kr/admin/search/searchIPmanage_man.asp","orgsip="&orgsip,"euc-kr")
        else
            retArrM(i) = getURLhtml("http://"&DomainArrMOB(i)&".10x10.co.kr/lib/searchIPmanage.asp","orgsip="&orgsip,"UTF-8")
        end if
    end if
next
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function chgSearchIP(ihostname,ichgipnm){
    var ichgip = document.frmchg.chgip.value;
    if (ichgip.length<1){
        alert('변경할 IP를 선택하세요.');
        document.frmchg.chgip.focus();
        return;
    }

    if (confirm("변경하시겠습니까?")){
        document.frm.mode.value="chgip";
        document.frm.appname.value=ichgipnm;
        document.frm.chgip.value=ichgip;
        document.frm.chghostname.value=ihostname;
        document.frm.submit();
    }
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
    <input type="hidden" name="mode" value="">
    <input type="hidden" name="appname" value="">
    <input type="hidden" name="chgip" value="">
    <input type="hidden" name="chghostname" value="">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
			사이트 : 
            <select name="domaingbn">
            <option value="">전체
            <option value="W" <%=CHKIIF(domaingbn="W","selected","")%> >www
            <option value="M" <%=CHKIIF(domaingbn="M","selected","")%>>M
            </select>

            &nbsp;
            검색서버IP
            <select name="orgsip">
            <option value="">전체
            <option value="206" <%=CHKIIF(orgsip="206","selected","")%>>206
            <option value="207" <%=CHKIIF(orgsip="207","selected","")%>>207
            <option value="208" <%=CHKIIF(orgsip="208","selected","")%>>208
            <option value="209" <%=CHKIIF(orgsip="209","selected","")%>>209
            <option value="210" <%=CHKIIF(orgsip="210","selected","")%>>210
            </select>
			
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value=" 검 색 " onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 액션 시작 -->
<form name="frmchg" method="post" >
<input type="hidden" name="mode" value="addbrandboostkey">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
        변경할 IP 
        <% if (application("Svr_Info") = "Dev") then %>
            <select name="chgip">
            <option value="">선택
            <option value="192.168.50.10" <%=CHKIIF(chgip="192.168.50.10","selected","")%>>192.168.50.10
            </select>
        <% else %>
            <select name="chgip">
            <option value="">선택
            <option value="192.168.0.206" <%=CHKIIF(chgip="192.168.0.206","selected","")%>>192.168.0.206
            <option value="192.168.0.207" <%=CHKIIF(chgip="192.168.0.207","selected","")%>>192.168.0.207
            <option value="192.168.0.208" <%=CHKIIF(chgip="192.168.0.208","selected","")%>>192.168.0.208
            <option value="192.168.0.209" <%=CHKIIF(chgip="192.168.0.209","selected","")%>>192.168.0.209
            <option value="192.168.0.210" <%=CHKIIF(chgip="192.168.0.210","selected","")%>>192.168.0.210
            </select>
        <% end if %>
		</td>
		<td align="right">
		</td>
	</tr>
</table>
</form>
<!-- 액션 끝 -->
<p>
<div></div>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% for i=Lbound(retArrW) to Ubound(retArrW) %>
<tr bgcolor="#FFFFFF">
    <td width="10%"><%=DomainArrWEB(i)%></td>
    <td width="40%"><%=retArrW(i)%></td>
    <td width="10%"><%=DomainArrMOB(i)%></td>
    <td width="40%"><%=retArrM(i)%></td>
</tr>
<% next %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
