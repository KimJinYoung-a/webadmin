<%@ language="VBScript" %>
<% option explicit %> 
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<%
 
dim empno
empno = session("ssBctSn")

Dim iedmsIdx, tContents,iscmlinkno ,uid,divcd,uday,rday,tday,ddate,dday

iedmsIdx	=  requestCheckvar(Request("ieidx"),10)   

iscmlinkno		=  requestCheckvar(Request("iSL"),10)
uid		  =  requestCheckvar(Request("uid"),10)
divcd		=  requestCheckvar(Request("divcd"),60)
uday		=  requestCheckvar(Request("uday"),10)
rday		=  requestCheckvar(Request("rday"),10)
tday		=  requestCheckvar(Request("tday"),10)
ddate		=  ReplaceRequestSpecialChar(Request("ddate"))  
dday		=  requestCheckvar(Request("dday"),10)
 
	Function GetDivCDStr 
		if (divcd = "1") then
			GetDivCDStr = "����"
		elseif (divcd = "2") then
			GetDivCDStr = "����"
		elseif (divcd = "3") then
			GetDivCDStr = "����"
		elseif (divcd = "4") then
			GetDivCDStr = "����"
		elseif (divcd = "6") then
			GetDivCDStr = "������"
		elseif (divcd = "7") then
			GetDivCDStr = "���ϴ�ü"
		elseif (divcd = "5") then
			GetDivCDStr = "���"
		elseif (divcd = "8") then
			GetDivCDStr = "��Ÿ"
		elseif (divcd = "9") then
			GetDivCDStr = "����"		
		else
			GetDivCDStr = "===="
		end if
	end Function

 %> 
 <meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	 <!--���ڰ���--> 
	<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
	<input type="hidden" name="tC" value="">
	<input type="hidden" name="ieidx" value="<%=iedmsIdx%>">  
	<input type="hidden" name="iSL" value="<%=iscmlinkno%>">
	</form>
	<div id="divEapp" style="display:none;">
		<div style="text-align:center;padding:10px;color:blue;">
		<p>- ������ ���� �Ⱓ�� 2�� �̸��� ���, ��� �� ���������ϼ��� ������ �˴ϴ�.</p>
		</div>	
	<table width="500" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">idx</td>
		<td bgcolor="#FFFFFF" width="300">
			<div id="divSL"><%=iscmlinkno%></div>
		</td>
	</tr>
	<tr height="25">
		<td width=120 bgcolor="<%= adminColor("tabletop") %>">���� ���̵�</td>
		<td bgcolor="#FFFFFF">
			<%= uid%>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF">
			<%= GetDivCDStr%>
		</td>
	</tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����ϼ�/���δ��/���ϼ� </td>
    	<td bgcolor="#FFFFFF">
    		<%=uday%> / <%=rday%> / <%=tday%>
    	</td>
    </tr>
	<tr height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��û�Ⱓ</td>
    	<td bgcolor="#FFFFFF">
    		<div id="divDate"><%=ddate%> (<%=dday%>��)</div>
    	</td>
    </tr>
	</table>

	</div>
	 <%
	session.codePage = 949
%>
	<script type="text/javascript">  
		document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	 	document.frmEapp.submit();
		</script>
	<!--/���ڰ���-->

