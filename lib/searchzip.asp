<%@ Language=VBScript %>
<%option explicit%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->

<%

%>
<HTML>
<HEAD>
<TITLE>�����ȣ �˻� </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel=stylesheet type="text/css" href="http://www.10x10.co.kr/lib/css/2008ten.css">
<style>
.input {
	font-family: "����", "Verdana";
	font-size: 12px;
	line-height: 20px;
	color: #000000;
	background-color: #FFFFFF;
	border: 0px solid #FFFFFF;
}
.kindTab_on {width:50%;background-color:#DEF8D0;padding:5px;float:left;font-weight:bold;cursor:pointer;}
.kindTab_off {width:50%;background-color:#E0E0E0;padding:5px;float:left;font-weight:normal;cursor:pointer;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</HEAD>
<body bgcolor="white" text="black" link="black" vlink="black" alink="#464646" style="margin:0 0 0 0">
<script>
function SubmitForm(frm)
{
        if (frm.query.value.length < 2) { alert("�� �̸��� �α��� �̻� �Է��ϼ���."); return; }
        frm.submit();
}

function chgTab(dv) {
	if(dv=="a") {
		$("#kTab1").attr("class","kindTab_on");
		$("#kTab2").attr("class","kindTab_off");
		$("#dRow1").show();
		$("#dRow2").hide();
		$("#stype").val("addr");
	} else {
		$("#kTab1").attr("class","kindTab_off");
		$("#kTab2").attr("class","kindTab_on");
		$("#dRow1").hide();
		$("#dRow2").show();
		$("#stype").val("road");
	}
}
</script>
<%

' -------------------------------------
' ȸ���� �ּҸ� ã�� Popup Window ȭ��
' -------------------------------------
Dim strTarget, stype
Dim strQuery


	strTarget	= Request("target")
	strQuery	= (Request("query"))
	stype		= Request("stype")

If strQuery = "" then
%>
<table width="440" border="0" cellpadding="0" cellspacing="0">
<form action="/lib/searchzip.asp?" method="get" name="gil" onsubmit="SubmitForm(document.gil); return false;">
<input type="hidden" name="mode"	value="search">
<input type="hidden" name="target"	value="<%=strTarget%>">
<input type="hidden" name="form"	value="account">
<input type="hidden" name="post1"	value="post1">
<input type="hidden" name="post2"	value="post2">
<input type="hidden" name="add"		value="add">
<input type="hidden" name="stype"	id="stype" value="addr">
  <tr>
    <td background="http://fiximage.10x10.co.kr/web2007/zipcode/member_pop_bg.gif" height="55">
      <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/search_zip.gif" width="114" height="36"></div>
    </td>
  </tr>
	<tr>
		<td>
			<div align="center" id="kTab1" class="kindTab_on" onclick="chgTab('a')">�����˻�</div>
			<div align="center" id="kTab2" class="kindTab_off" onclick="chgTab('r')">���θ�˻�</div>
		</td>
	</tr>
	<tr id="dRow1">
		<td height="50">
			<div align="center">ã���� �ϴ� �ּ��� ��/��/�� �̸��� �Է��ϼ���.<br>
			(��: ��ġ��,���,�����)</div>
		</td>
	</tr>
	<tr id="dRow2" style="display:none;">
		<td height="50">
			<div align="center">ã���� �ϴ� �ּ��� ���θ� �̸��� �Է��ϼ���.<br>
			(��: ����1��, �������)</div>
		</td>
	</tr>
  <tr>
    <td height="37">
      <div align="center">
        <table border="0" cellpadding="0">
          <tr>
            <td>������ :</td>
            <td width="97">
              <input type="text" name="query" class="input_01" size="13" style="ime-mode:active">
            </td>
            <td width="61"><a href="javascript:SubmitForm(document.gil);"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/zip_search.gif" width="65" height="22"></a>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <tr>
    <td height="8">
      <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/pup_dotline.gif" width="144" height="8"></div>
    </td>
  </tr>
  <tr>
    <td height="37">
      <div align="center">
        <table width="380" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
          <tr>
            <td>
              <table width="380" border="0" cellpadding="2" cellspacing="1">
                <tr bgcolor="#f7f7f7">
                  <td class="a" width="109">
                    <div align="center"><b><font color="#666666">�����ȣ</font></b></div>
                  </td>
                  <td class="a" width="290">
                    <div align="center"><b><font color="#666666">�ּ�</font></b></div>
                  </td>
                  <td class="a" width="50" bgcolor="#f7f7f7">
                    <div align="center"><b><font color="#666666">����</font></b></div>
                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td colspan="3" class="a">
                    <div align="center"><font color="#999999">�������� �Է��� �� �˻��Ͽ��ּ���</font></div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <tr>
    <td style="border-bottom: 7px solid #e1e1e1;" height="38">
      <div align="center"><a href="javascript:self.close();"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/zip_close.gif" width="65" height="22" border="0"></a></div>
    </td>
  </tr>
 </form>
</table>

<% else %>
<SCRIPT language="JavaScript">
function CopyZip(t,post1,post2,add,dong)
{
	var frm = eval("opener." + t);
	// copy
	frm.zip1.value			= post1;
	frm.zip2.value			= post2;
	frm.addr1.value		= add;
	frm.addr2.value		= dong;
	
	

	// focus
	frm.addr2.focus();

	// close this window
	window.close();

}

function SubmitForm(frm)
{
        if (frm.query.value.length < 2) { alert("�� �̸��� �α��� �̻� �Է��ϼ���."); return; }
        frm.submit();
}

</SCRIPT>
<table width="440" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td background="http://fiximage.10x10.co.kr/web2007/zipcode/member_pop_bg.gif" height="55">
      <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/search_zip.gif" width="114" height="36"></div>
    </td>
  </tr>
	<tr>
		<td>
			<div align="center" id="kTab1" class="<%=chkIIF(stype="addr","kindTab_on","kindTab_off")%>" onclick="chgTab('a')">�����˻�</div>
			<div align="center" id="kTab2" class="<%=chkIIF(stype="road","kindTab_on","kindTab_off")%>" onclick="chgTab('r')">���θ�˻�</div>
		</td>
	</tr>
	<tr id="dRow1" style="<%=chkIIF(stype="addr","","display:none")%>">
		<td height="50">
			<div align="center">ã���� �ϴ� �ּ��� ��/��/�� �̸��� �Է��ϼ���.<br>
			(��: ��ġ��,���,�����)</div>
		</td>
	</tr>
	<tr id="dRow2" style="<%=chkIIF(stype="road","","display:none")%>">
		<td height="50">
			<div align="center">ã���� �ϴ� �ּ��� ���θ� �̸��� �Է��ϼ���.<br>
			(��: ����1��, �������)</div>
		</td>
	</tr>
<form action="/lib/searchzip.asp?" method="get" name="gil2" onsubmit="SubmitForm(document.gil2); return false;">
<input type="hidden" name="target" value="<%=strTarget%>">
<input type="hidden" name="stype"	id="stype" value="<%=stype%>">
  <tr>
    <td  height="37">
      <div align="center">
        <table border="0" cellpadding="0">
          <tr>
            <td>������ :</td>
            <td width="97">
              <input type="text" name="query" class="input_01" size="13" style="ime-mode:active">
            </td>
            <td width="61"><a href="javascript:SubmitForm(document.gil2);"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/zip_search.gif" width="65" height="22" border="0"></a>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</form>
  <tr>
    <td height="8">
      <div align="center"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/pup_dotline.gif" width="144" height="8"></div>
    </td>
  </tr>
  <tr>
    <td height="37">
      <div align="center">
        <table width="380" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
          <tr>
            <td>
              <table width="380" border="0" cellpadding="2" cellspacing="1">
                <tr>
                  <td class="a" width="109" bgcolor="#f7f7f7">
                    <div align="center"><b><font color="#666666">�����ȣ</font></b></div>
                  </td>
                  <td class="a" width="290" bgcolor="#f7f7f7">
                    <div align="center"><b><font color="#666666">�ּ�</font></b></div>
                  </td>
                  <td class="a" width="50" bgcolor="#f7f7f7">
                    <div align="center"><b><font color="#666666">����</font></b></div>
                  </td>
                </tr>
<%

Dim strSql
Dim nRowCount

Dim strAddress

dim useraddr01, useraddr02

dim lstr
        lstr = CStr(Len(strQuery))

	if stype="addr" then
		strSql = "SELECT   ADDR_ZIP1, ADDR_ZIP2, ADDR_SI,ADDR_GU,ADDR_DONG,ADDR_ETC,ADDR_Fulltext FROM [db_zipcode].[dbo].ADDR080TL WHERE ADDR_Fulltext like '%" & strQuery & "%' and ADDR_sortNo<>'999' "
	elseif stype="road" then
		strSql = "SELECT   ADDR_ZIP1, ADDR_ZIP2, ADDR_SI,ADDR_GU,ADDR_DONG,ADDR_ROAD,ADDR_BLDNO1,ADDR_BLDNO2,ADDR_ETC,ADDR_Fulltext " &_
				" FROM [db_zipcode].[dbo].ROAD010 " &_
				" WHERE ADDR_ROAD like '" & strQuery & "%' and ADDR_sortNo<>'999' " &_
				" order by addr_zip1, addr_Gu, addr_Road, Addr_BldNo1 "
	end if
	rsget.Open strSQL,dbget,1
	'oRs.Open strSQL,oCnn,1

	if not rsget.eof then
		do while not rsget.EOF and nRowCount < rsget.PageSize

			if stype="addr" then
				strAddress = trim(rsget("ADDR_Fulltext"))

				useraddr01 = trim(rsget("ADDR_SI")) & " " & trim( rsget("ADDR_GU"))
				useraddr02 = trim( rsget("ADDR_DONG")) & " " & trim( rsget("ADDR_ETC"))
				useraddr02 = Replace(useraddr02,"'","\'")
			elseif stype="road" then
				strAddress = trim(rsget("ADDR_Fulltext")) & " " & trim( rsget("ADDR_BLDNO1"))
				if Not(rsget("ADDR_BLDNO2")="" or isNull(rsget("ADDR_BLDNO2"))) then
					strAddress = strAddress & " ~ " & trim(rsget("ADDR_BLDNO2"))
				end if

				useraddr01 = trim(rsget("ADDR_SI")) & " " & trim( rsget("ADDR_GU"))
				useraddr02 = trim( rsget("ADDR_ROAD"))
				if Not(rsget("ADDR_ETC")="" or isNull(rsget("ADDR_ETC"))) then
					'�ٷ� ���ó�� �ִ� ���� ���� �ǹ�
					useraddr02 = useraddr02 & " " & trim(rsget("ADDR_BLDNO1")) & " " & trim(rsget("ADDR_ETC"))
				end if
				useraddr02 = Replace(useraddr02,"'","\'")
			end if

%>
				<tr bgcolor="#FFFFFF">
                  <td class="a" width="109" align="center" onclick="CopyZip('<%= strTarget %>','<%=rsget("ADDR_ZIP1")%>','<%=rsget("ADDR_ZIP2")%>','<% = useraddr01 %>', '<% = useraddr02 %>')" style="cursor:hand">
						<input type="text" name="post1" size="3" value='<%=rsget("ADDR_zip1")%>' class="input" style="cursor:hand"> -
						<input type="text" name="post2" size="3" value='<%=rsget("ADDR_zip2")%>' class="input" style="cursor:hand">
                  </td>
                  <td class="a" width="290"  align="center">
						<INPUT type="text" name="add" value='<%=strAddress%>' size="38" class="input">
                  </td>
                  <td class="a" width="50" >
                    <div align="center"><a href="javascript:CopyZip('<%= strTarget %>','<%=rsget("ADDR_ZIP1")%>','<%=rsget("ADDR_ZIP2")%>','<% = useraddr01 %>', '<% = useraddr02 %>')">����</a></div>
                  </td>
                </tr>
<%
		rsget.MoveNext
		loop
	end if
%>
              </table>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <tr>
    <td style="border-bottom: 7px solid #e1e1e1;" height="38">
      <div align="center"><a href="javascript:self.close();"><img src="http://fiximage.10x10.co.kr/web2007/zipcode/zip_close.gif" width="65" height="22" border="0"></a></div>
    </td>
  </tr>
</table>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
