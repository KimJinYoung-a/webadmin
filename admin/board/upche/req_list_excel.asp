<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��������
' History : ������ ����
'			2022.09.13 �ѿ�� ����(�����ٿ�ε�,�˻����� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim i, j,ix, page,gubun, onlymifinish, license_no, research, searchkey,catevalue, dispCate,maxDepth
dim ipjumYN , catemid ,catelarge, sellgubun, workid, iid, reqcomment, startdate, enddate
dim menupos, arrList
	page 			= requestCheckvar(getNumeric(request("pg")),10)
	gubun 			= requestCheckvar(request("gubun"),2)
	onlymifinish 	= requestCheckvar(request("onlymifinish"),3)
	research 		= requestCheckvar(request("research"),3)
	searchkey 		= requestCheckvar(request("searchkey"),32)
	catevalue		= requestCheckvar(request("catevalue"),3)
	ipjumYN			= requestCheckvar(request("ipjumYN"),1)
	catemid 		= requestCheckvar(request("catemidbox"),3)
	catelarge 		= requestCheckvar(request("catelargebox"),3)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2
	sellgubun			= requestCheckvar(request("sellgubun"),1)
	workid			= requestCheckvar(request("workid"),34)
	iid             = requestCheckVar(Request("iid"),9) 
	license_no		= requestCheckvar(request("license_no"),50)
	reqcomment		= requestCheckvar(request("reqcomment"),50)
	startdate = NullFillWith(requestCheckVar(request("startdate"),10),DateAdd("m",-1,date()))
	enddate = NullFillWith(requestCheckVar(request("enddate"),10),date())
    menupos 			= requestCheckvar(getNumeric(request("menupos")),10)

'// �⺻������ �����Ƿڼ�
if gubun="" then gubun="01"
if research="" and onlymifinish="" then onlymifinish="on"		
if (page = "") then page = "1"
If gubun = "01" Then 
	'sellgubun = ""
	workid = ""
End If

dim companyrequest
set companyrequest = New CCompanyRequest
	companyrequest.PageSize = 200000
	companyrequest.CurrPage = 1
	companyrequest.ScrollCount = 10
	companyrequest.FReqcd=gubun
	companyrequest.FOnlyNotFinish = onlymifinish
	companyrequest.FRectSearchKey = searchkey
	companyrequest.FRectCatevalue = catevalue
	companyrequest.FipjumYN = ipjumYN
	companyrequest.FRectDispCate = dispCate
	companyrequest.FRectSellgubun = sellgubun
	companyrequest.FRectWorkid = workid
	companyrequest.FRectID=iid
	companyrequest.FRectlicense_no=license_no
	companyrequest.FRectReqcomment=reqcomment
	companyrequest.FRectstartdate=startdate
	companyrequest.FRectenddate=DateAdd("d",+1,enddate)
	companyrequest.getReqListNotpaging

if companyrequest.resultCount>0 then
    arrList = companyrequest.fArrList
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENREQBOARD" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��

%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td colspan="10">
        �˻���� : <b><%= companyrequest.resultCount %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<td align="center">��ȣ</td>
<td align="center">��û��</td>
<td align="center">����</td>
<td align="center">ä��</td>
<td align="center">ó����</td>
<td align="center">��������</td>
<td align="center">ī�װ�����</td>
<td align="center">ȸ��URL</td>
<td align="center">�亯����</td>
    <td align="center">���</td>
</tr>
<% if isarray(arrList) then %>
    <% for i = 0 to ubound(arrList,2) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= arrList(0,i) %></td>
        <td align="center" ><%= FormatDate(arrList(15,i), "0000-00-00") %></td>
        <td>[<%= companyrequest.code2name(arrList(1,i)) %>] <%= arrList(2,i) %></td>
        <td align="center">
            <% if arrList(20,i)="Y" then %>�¶���/��������<%
            elseif arrList(20,i)="N" then %>�¶���<%
            elseif arrList(20,i)="F" then %>��������<%
            else %><%=arrList(20,i)%><%
            end if %>
        </td>
        <td align="center" >
            <% if (IsNull(arrList(16,i)) = true) then %>
                �̿Ϸ�
            <% else %>
            <%= FormatDate(arrList(16,i), "0000-00-00") %>
            <% end if %>
        </td>
        <td align="center">
            <%if arrList(21,i)="Y" then response.write "�����Ϸ�" %>
            <%if arrList(21,i)="N" then response.write "N" %>
            </td>
        <td align="center" >
            <div><%IF not isNull(arrList(35,i)) THEN%><%=arrList(36,i)%> > <%=arrList(37,i)%><%END IF%></div>
            <div><%=arrList(34,i)%></div>  
            </td>
        <td align="center">
            <%= arrList(10,i) %>
        </td> 
        <td align="center">
            <% if companyrequest.commentcheck(arrList(22,i))="Y" then %>
            Y
            <% else %>
            N
            <% end if %>
        </td>		
        <td align="center" >
        </td>
    </tr>   
    <%
    if i mod 1000 = 0 then
        Response.Flush		' ���۸��÷���
    end if
    next
    %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="10" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->