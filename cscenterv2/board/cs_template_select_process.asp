<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%

dim mastergubun, gubun, contents
dim orderserial, userid, makerid, itemid, orderdetailidx
dim sqlStr
dim errMSG

mastergubun = RequestCheckvar(request("mastergubun"),2)
gubun = RequestCheckvar(request("gubun"),2)

orderserial = Trim(RequestCheckvar(request("orderserial"),16))
userid = Trim(RequestCheckvar(request("userid"),32))
makerid = Trim(RequestCheckvar(request("makerid"),32))
itemid = Trim(RequestCheckvar(request("itemid"),10))
orderdetailidx = Trim(RequestCheckvar(request("orderdetailidx"),10))


sqlStr = "select top 1 contents" + vbcrlf
sqlStr = sqlStr + " from [db_academy].[dbo].[tbl_ACA_cs_template] " + vbcrlf
sqlStr = sqlStr + " where isusing = 'Y'" + vbcrlf
sqlStr = sqlStr + " and gubun = '" + gubun + "'" + vbcrlf
sqlStr = sqlStr + " and mastergubun = '" + mastergubun + "'" + vbcrlf

rsACADEMYget.Open sqlStr,dbACADEMYget,1
if  not rsACADEMYget.EOF  then
	 contents = replace(replace(db2html(rsACADEMYget("contents")),vbcrlf,"<br>"), "'", "\'")
end if
rsACADEMYget.close


if (InStr(contents, "[��ü��ǰ�ּ�]") or InStr(contents, "[��ü��ǰ�����]") or InStr(contents, "[��ü��ǰ��ȭ]") or InStr(contents, "[��ü�ŷ��ù��]") or InStr(contents, "[��ü��Ʈ��Ʈ��]")) then
	if (makerid = "") then
		errMSG = "���� : �ڵ���ȯ ����[�귣������ ����]"
	else
		dim opartner
		set opartner = new CPartnerUser
		opartner.FRectDesignerID = makerid
		opartner.GetOnePartnerNUser

		dim OReturnAddr
		set OReturnAddr = new CCSReturnAddress

		OReturnAddr.FRectMakerid = makerid
		OReturnAddr.GetBrandReturnAddress

		if (opartner.FResultCount = 0) then
			errMSG = "���� : �ڵ���ȯ ����[�귣������ ����]"
		else
			contents = Replace(contents, "[��ü��ǰ�ּ�]", (OReturnAddr.FreturnZipaddr + " " + OReturnAddr.FreturnEtcaddr))
			contents = Replace(contents, "[��ü��ǰ�����]", OReturnAddr.FreturnName)
			contents = Replace(contents, "[��ü��ǰ��ȭ]", OReturnAddr.FreturnPhone)
			contents = Replace(contents, "[��ü�ŷ��ù��]", opartner.FOneItem.Ftakbae_name + "(" + opartner.FOneItem.Ftakbae_tel + ")")
			contents = Replace(contents, "[��ü��Ʈ��Ʈ��]", opartner.FOneItem.Fsocname_kor)
		end if
	end if
end if


if InStr(contents, "[���̵�]") then
	if (userid = "") then
		'''errMSG = "���� : �ڵ���ȯ ����[���̵����� ����]"
		contents = Replace(contents, "[���̵�]", "��")
	else
		contents = Replace(contents, "[���̵�]", userid)
	end if
end if


if InStr(contents, "[�̸�]") then
	if Not IsNull(session("ssBctCname")) then
		contents = Replace(contents, "[�̸�]", session("ssBctCname"))
	end if
end if

if InStr(contents, "[������ȭ]") then
	if Not IsNull(session("ssBctId")) then
		sqlStr = " select top 1 direct070 "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_partner.dbo.tbl_user_tenbyten "
		sqlStr = sqlStr + " where userid = '" + CStr(session("ssBctId")) + "' and IsNull(direct070, '') <> '' "

		rsget.Open sqlStr,dbget,1
		if  not rsget.EOF  then
			contents = Replace(contents, "[������ȭ]", rsget("direct070"))
		end if
		rsget.close
	end if
end if

%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html">
<script>

var source,convert,temp;

source = "<br>";
convert = "\n";
temp = '<% = contents %>';

while (temp.indexOf(source)>-1) {
	 pos= temp.indexOf(source);
	 temp = "" + (temp.substring(0, pos) + convert +
	 temp.substring((pos + source.length), temp.length));
}

	parent.TnCSTemplateGubunProcess(temp, "<%= errMSG %>");

</script>
</head>
<body>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
