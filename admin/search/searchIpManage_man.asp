<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("121.78.103.2","121.78.103.60","110.93.128.93","110.93.128.94","110.93.128.95")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    if NOT (Application("Svr_Info")="Dev") THEN
    response.write "ERR"
    response.end
    end if
end if

Dim orgsip : orgsip = requestCheckvar(request("orgsip"),10)
Dim ihostname : ihostname=Request.ServerVariables("HTTP_HOST")
Dim mode : mode = requestCheckvar(request("mode"),32)
Dim appname : appname = requestCheckvar(request("appname"),32)
Dim chgip : chgip = requestCheckvar(request("chgip"),15)
Dim ipchgOK : ipchgOK=FALSE
if (mode="chgip") then
    if (appname="G_ORGSCH_ADDR") or (appname="G_2NDSCH_ADDR") or (appname="G_4THSCH_ADDR") or (appname="G_3RDSCH_ADDR") or (appname="G_1STSCH_ADDR") or (appname="G_ZIPSCH_ADDR") then
        if (application("Svr_Info") = "Dev") then
            if (chgip="192.168.50.10") then
                application(appname) = chgip
                ipchgOK = True
            end if
        else
            if (chgip="192.168.0.206") or (chgip="192.168.0.207") or (chgip="192.168.0.208") or (chgip="192.168.0.209") or (chgip="192.168.0.210") then
                application(appname) = chgip
                ipchgOK = True
            end if
        end if
        
    end if
    IF NOT(ipchgOK) then
        response.write "ERR:chage IP"
    end if
end if
%>
<table border=1 cellspacing=0 cellpadding=3 class="a">
<tr>
    <td colspan="5"><%=Application("Svr_Info")%> / <%=ihostname%></td>
</tr>
<tr>
    <td width="100">원IP (hostName)</td>
    <td>현재IP</td>
    <td>비고</td>
    <td><font color="#CCCCCC">APPName</font></td>
    <td>Action</td>
</tr>
<% if (orgsip="") or (orgsip="206") then %>
<tr>
    <td>206 (Ten_Search01)</td>
    <td><%=Application("G_ORGSCH_ADDR")%></td>
    <td>인덱싱/자동완성</td>
    <td><font color="#CCCCCC">G_ORGSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_ORGSCH_ADDR')"></td>
</tr>
<% end if %>
<% if (orgsip="") or (orgsip="207") then %>
<tr>
    <td>207 (Ten_Search02)</td>
    <td><%=Application("G_2NDSCH_ADDR")%></td>
    <td>WEB(카테)</td>
    <td><font color="#CCCCCC">G_2NDSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_2NDSCH_ADDR')"></td>
</tr>
<% end if %>
<% if (orgsip="") or (orgsip="208") then %>
<tr>
    <td>208 (Ten_Search03)</td>
    <td><%=Application("G_4THSCH_ADDR")%></td>
    <td>MOB</td>
    <td><font color="#CCCCCC">G_4THSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_4THSCH_ADDR')"></td>
</tr>
<% end if %>
<% if (orgsip="") or (orgsip="209") then %>
<tr>
    <td>209 (Ten_Search04)</td>
    <td><%=Application("G_3RDSCH_ADDR")%></td>
    <td>APP</td>
    <td><font color="#CCCCCC">G_3RDSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_3RDSCH_ADDR')"></td>
</tr>
<% end if %>
<% if (orgsip="") or (orgsip="210") then %>
<tr>
    <td>210 (Ten_Search05)</td>
    <td><%=Application("G_1STSCH_ADDR")%></td>
    <td>WEB(검색)</td>
    <td><font color="#CCCCCC">G_1STSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_1STSCH_ADDR')"></td>
</tr>
<% end if %>
<% if (orgsip="") or (orgsip="206") then %>
<tr>
    <td>206 (Ten_Search01)</td>
    <td><%=Application("G_ZIPSCH_ADDR")%></td>
    <td>우편번호ZIP</td>
    <td><font color="#CCCCCC">G_ZIPSCH_ADDR</font></td>
    <td><input type="button" value="변경" onClick="chgSearchIP('<%=ihostname%>','G_ZIPSCH_ADDR')"></td>
</tr>
<% end if %>
</table>