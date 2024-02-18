<%@ language="VBScript" %>
<% option explicit %>
<%
'###########################################################
' Description :
' History : 2011.03.14 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
Dim clseapp,clsMem
Dim iarapcd, iedmsIdx
Dim sedmsname,sedmscode,sarap_cd,sarap_nm,sacc_cd,sacc_nm,sacc_use_cd,sACC_GRP_CD
Dim sEappName, mReportPrice
Dim spartname ,slastApprovalid,sjob_name,sscmLink,iscmlinkno
Dim tContents, blnPayEapp
Dim sCurrencyPrice,ipaytype,sCurrencyType
Dim idepartmentid,sdepartmentname, icid1, icid2, icid3, icid4
Dim isAgreeNeed, isAgreeNeedTarget
Dim insTable : insTable = False
Dim addFileName, addFileNamePh

iarapcd =  requestCheckvar(Request("iAidx"),13)
if iarapcd = "" then iarapcd = 0
iedmsIdx	=  requestCheckvar(Request("ieidx"),10)
'sacc_nm =  requestCheckvar(Request("sAN"),30)
tContents  = ReplaceRequestSpecialChar(Request("tC"))
iscmlinkno		=  requestCheckvar(Request("iSL"),10)
mReportPrice=  requestCheckvar(Request("mRP"),20)

'get default form
set clseapp = new CEApproval
	clseapp.Farap_cd  = iarapcd
	clseapp.Fedmsidx = iedmsIdx
	clseapp.fnGetEAppForm

	iedmsIdx        = clseapp.FedmsIdx
	sedmsname       = clseapp.Fedmsname
	sedmscode				= clseapp.Fedmscode
	sscmLink   			= clseapp.FscmLink
	slastApprovalid = clseapp.FlastApprovalid
	sjob_name				= clseapp.Fjob_name
	sarap_cd 			= clseapp.Farap_cd
	sarap_nm    	= clseapp.Farap_nm
	sacc_cd    		= clseapp.Facc_cd
	sacc_nm				= clseapp.Facc_nm
	sacc_use_cd		= clseapp.Facc_use_cd
	blnPayEapp		= clseapp.FisPayEapp
	sACC_GRP_CD		= clseapp.FACC_GRP_CD
	isAgreeNeed		= clseapp.FisAgreeNeed
	isAgreeNeedTarget = clseapp.FisAgreeNeedTarget
	addFileName		= getEdmsFileName(clseapp.FedmsCode, clseapp.FedmsName, clseapp.FedmsFile, addFileNamePh)

    if (sarap_cd = "0") or (sarap_cd = "") or (IsNull(sarap_cd)) then
        if (iedmsIdx = "102") or (iedmsIdx = "103") or (iedmsIdx = "104") then
            sacc_use_cd = "15100"
            sacc_nm = "��ǰ"
            insTable = True
        end if

        if (iedmsIdx = "102") then
            '// ��ǰ����
            iarapcd = "106"
            sarap_nm = "��ǰ���Ա�"
        elseif (iedmsIdx = "103") then
            '// ��ǰ����
            iarapcd = "108"
            sarap_nm = "���Ի�ǰ���"
        elseif (iedmsIdx = "104") then
            '// ��ǰ����
            iarapcd = "107"
            sarap_nm = "��ǰ���۴��"
        end if
    end if

	IF tContents ="" THEN
	    tContents		= clseapp.FedmsForm
	END IF
set clseapp = nothing

'tContents = replace(tContents,"</p><p>&nbsp;</p>","<BR><BR></p>")
'get partname
set clsMem = new CTenByTenMember
	clsMem.Fuserid = session("ssBctId")
	clsMem.fnGetDepartmentInfo
	idepartmentid		= clsMem.Fdepartment_id
 	sdepartmentname = clsMem.FdepartmentNameFull
 	icid1						= clsMem.Fcid1
 	icid2						= clsMem.Fcid2
 	icid3						= clsMem.Fcid3
 	icid4						= clsMem.Fcid4
 set clsMem = nothing

 IF iarapcd > 0 THEN
     IF sedmsname <> "" THEN
 	     sEappName = sedmsname&"_"&sarap_nm
     ELSE
         sEappName = sarap_nm
     END IF
 ELSE
 	sEappName = sedmsname
 END IF

Dim tmpAgree, tmpAgreelist, tmpAgreename, tmpAgreeTxt, tmpAgreejobnm
If isAgreeNeed = "Y" Then
	set tmpAgree = new CTenByTenMember
		tmpAgree.Fuserid = isAgreeNeedTarget
		tmpAgreelist = tmpAgree.fnGetInIDOutName
	IF isArray(tmpAgreelist) THEN
		tmpAgreename = tmpAgreelist(1,0)
		tmpAgreejobnm = tmpAgreelist(5,0)
		If isnull(tmpAgreejobnm) OR tmpAgreejobnm = "" Then
			tmpAgreejobnm = ""
		Else
			tmpAgreejobnm = " " & tmpAgreejobnm
		End If
		tmpAgreeTxt = tmpAgreename & tmpAgreejobnm & " ["&isAgreeNeedTarget&"]"
	End If
	set tmpAgree = nothing
End If
%>

<%
 IF sscmLink <> "" and iscmlinkno ="" THEN
 	Call Alert_return ("���԰�ο� ������ �߻��Ͽ����ϴ�.")
response.end
END IF
%>

<html>
<head>

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<!-- daumeditor head ------------------------->
 <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
 <meta http-equiv="X-UA-Compatible" content="IE=10" />
 <link rel="stylesheet" type="text/css" href="/webfonts/CoreSansC.css">
 <link rel="stylesheet" href="/css/scm.css" type="text/css">
 <link rel="stylesheet" href="/lib/util/daumeditor/css/editor.css" type="text/css" charset="euc-kr"/>
 <script type="text/javascript" src="/js/common.js"></script>
 <script src="/lib/util/daumeditor/js/editor_loader.js" type="text/javascript" charset="euc-kr"></script>
 <script src="/lib/util/daumeditor/js/editor_creator.js" type="text/javascript" charset="euc-kr"></script>
 <script type="text/javascript">
    var config = {
        initializedId: "",
        wrapper: "tx_trex_container",
        form: 'frm',
        txIconPath: "/lib/util/daumeditor/images/icon/editor/",
        txDecoPath: "/lib/util/daumeditor/images/deco/contents/",
        events: {
            preventUnload: false
        },
        sidebar: {
            attachbox: {
                show: true
            },
            attacher: {
                 image: {
                    popPageUrl: "/lib/util/daumeditor/pages/trex/image.asp"
                }
            }
        },
		toolbar: {
			fontfamily: {
				options: [
					{ label: ' ���� (<span class="tx-txt">������ABC123</span>)', title: '����', data: 'Gulim,����,AppleGothic,sans-serif', klass: 'tx-gulim' },
					{ label: ' ���� (<span class="tx-txt">������ABC123</span>)', title: '����', data: 'Batang,����', klass: 'tx-batang' },
					{ label: ' ���� (<span class="tx-txt">������ABC123</span>)', title: '����', data: 'Dotum,����', klass: 'tx-dotum' },
					{ label: ' CoreSansC (<span class="tx-txt">������ABC123</span>)', title: 'CoreSansC', data: 'CoreSansC-45Regular,malgun Gothic,�������', klass: 'tx-CoreSansC' },
					{ label: ' �ü� (<span class="tx-txt">������ABC123</span>)', title: '�ü�', data: 'Gungsuh,�ü�', klass: 'tx-gungseo' },
					{ label: ' Arial (<span class="tx-txt">ABC123</span>)', title: 'Arial', data: 'Arial', klass: 'tx-arial' },
					{ label: ' Verdana (<span class="tx-txt">ABC123</span>)', title: 'Verdana', data: 'Verdana', klass: 'tx-verdana' },
					{ label: ' Courier New (<span class="tx-txt">ABC123</span>)', title: 'Courier New', data: 'Courier New', klass: 'tx-courier-new' },
					{ label: ' Tahoma (<span class="tx-txt">ABC123</span>)', title: 'Tahoma', data: 'Tahoma', klass: 'tx-tahoma' }
				]
			}
		},
		canvas: {
			styles: {
				fontFamily: "CoreSansC", /* �⺻ ����ü */
                fontSize: "8pt",         /* ��Ʈ ������ */
			}
		}
	};
 </script>
<!-- //daumeditor head ------------------------->
<script language="javascript" src="eapp.js?t=<%=left(now(),10)%>" charset="euc-kr"></script>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#F4F4F4">
<table width="840" height="100%" cellpadding="0" cellspacing="0"  border="0" align="center">
<tr>
	<td valign="top">
		<table width="100%" cellpadding="1" cellspacing="0" class="a">
		<tr>
			<td>
				<form name="frm" method="post" action="proceapp.asp">
				<input type="hidden" name="hidM" value="I">
				<input type="hidden" name="hidRS" value="0">
				<input type="hidden" name="iaidx" value="<%=iarapcd%>"><!--iAIdx ��������� -->
				<input type="hidden" name="sACC" value="<%=sacc_cd%>"><!-- �߰� 2013/10/30-->
				<input type="hidden" name="ieIdx" value="<%=iedmsIdx%>">
				<input type="hidden" name="iAP" value="1">
				<input type="hidden" name="hidAid" value="<%=session("ssBctId")%>">
				<input type="hidden" name="hidPS" value="<%=idepartmentid%>">
				<input type="hidden" name="hidcid1" value="<%=icid1%>">
				<input type="hidden" name="hidcid2" value="<%=icid2%>">
				<input type="hidden" name="hidcid3" value="<%=icid3%>">
				<input type="hidden" name="hidcid4" value="<%=icid4%>">
				<input type="hidden" name="hidUN" value="<%=session("ssBctCname")%>">
				<input type="hidden" name="iLAID" value="<%=slastApprovalid%>">
				<input type="hidden" name="hidJN" value="<%=sjob_name%>">
				<input type="hidden" name="hidAI" id="hidAI" value=""><!--���缱���̵�(�������->)-->
				<input type="hidden" name="hidATxt" id="hidATxt" value="">
				<input type="hidden" name="hidAJ" id="hidAJ" value="">
				<input type="hidden" name="hidALI" id="hidALI" value=""><!--���������ھ��̵�-->
				<input type="hidden" name="hidALN" id="hidALTxt" value="">
				<input type="hidden" name="hidALJ" id="hidALJ" value="">
				<input type="hidden" name="hidAHI" id="hidAHI" value=""><!--���������ھ��̵�-->
				<input type="hidden" name="hidAHN" id="hidAHTxt" value="">
				<input type="hidden" name="hidAHJ" id="hidAHJ" value="">
				<input type="hidden" name="hidRfI" id="hidRfI" value=""><!--�������̵�-->
				<input type="hidden" name="blnL" value="0"> <!--���������� ��Ͽ���-->
				<% If isAgreeNeed = "Y" Then %>
				<input type="hidden" name="hidAI_H" id="hidAI_H" value="<%=isAgreeNeedTarget%>"><!--������ ���̵�-->
				<input type="hidden" name="hidATxt_H" id="hidATxt_H" value="<%=tmpAgreeTxt%>">
				<% Else %>
				<input type="hidden" name="hidAI_H" id="hidAI_H" value=""><!--������ ���̵�-->
				<input type="hidden" name="hidATxt_H" id="hidATxt_H" value="">
				<% End If %>
				<input type="hidden" name="tmpisAgreeNeed" value="<%=isAgreeNeed%>">
				<input type="hidden" name="tmpisAgreeNeedTarget" value="<%=isAgreeNeedTarget%>">
				<input type="hidden" name="hidPS_H" id="hidPS_H" value="">
				<input type="hidden" name="iRM" value="M010">
				<input type="hidden" name="hidPE" value="<%=blnPayEapp%>">


				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
				<tr>
					<td>
						<table width="100%" cellpadding="5" cellspacing="1" class="a" >
						<tr>
							<td class="verdana-large"><b><%=sEappName%> </b></td>
							<td align="right" width="100"><img src="/images/admin_logo_10x10.jpg"></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td align="right">
						<input type="button" class="button" style="color:blue" value="���缱���" onClick="jsRegID(1);">
						<input type='button' class='button' value='�����������' onClick='popDecision();'>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tR>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" width="60" nowrap>�����ڵ�</td>
							<td bgcolor="#FFFFFF" width="130"><%=sedmscode%></td>
							<td rowspan="5" bgcolor="#FFFFFF" valign="top" width="520">
								<div id="dAP">
								<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0">
								<tr align="center">
									<td valign="top" >
										<div id="dAP1">
										<table width="100%" cellpadding="5" cellspacing="0" class="a" border="0">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td></tr>
										<tr><td height="100" valign="bottom"></td></tr>
										</table>
										</div>
									</td>
								    <% If isAgreeNeed = "Y" Then %>
									<td valign='top'  width='180'  height='100%'>
										<div id='dAP_H'>
										<table width='100%' height='100%' cellpadding='5' cellspacing='0' class='a' border=0>
										<tr><td align='Center' bgcolor='#E6E6E6' height='20'>����</td></tr>
										<tr><td align='Center'>���δ��</td></tr>
										<tr><td align='Center'><%= tmpAgreeTxt %></td></tr>
										<tr><td align='Center'>&nbsp;</td></tr>
										<tr><td align='Center'><input type='checkbox' value='1' name='chkSms_H' checked> SMS����</td></tr>
										</table>
										</div>
									</td>
								    <% Else %>
									<td valign="top"  width="180">
									    <div id="dAP_H">
									    <table width="100%" cellpadding="5" cellspacing="0" class="a">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">����</td></tr>
										<tr><td align="Center">&nbsp;</td></tr>
										<tr><td align="Center"></td></tr>
										</table>
									    </div>
								    </td>
								    <% End If %>
									<td valign="top"  width="180">
										<div id="dAP0">
										<table width="100%" cellpadding="5" cellspacing="0" class="a">
										<tr><td align="Center" bgcolor="<%= adminColor("tabletop") %>">��������</td></tr>
										<tr><td align="Center">&nbsp;</td></tr>
										<tr><td align="Center"><%=sjob_name%></td></tr>
										</table>
										</div>
									</td>
								</tr>
								</table>
							</div>
							</td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >��/�μ�</td>
							<td bgcolor="#FFFFFF"><%=sdepartmentname%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >�ۼ���</td>
							<td bgcolor="#FFFFFF"><%= session("ssBctCname")%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >�ۼ���</td>
							<td bgcolor="#FFFFFF"><%=date()%></td>
						</tr>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" align="center" >����</td>
							<td bgcolor="#FFFFFF"><input class="input" type="text" name="sRfN" id="sRfN" value="" size="30" style="border:0;" readonly></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td width="60" rowspan="4"  align="center">ǰ�ǳ���</td>
							<td>IDX</td>
							<td>ǰ�Ǽ���</td>
							<td>ǰ�Ǳݾ�</td>
							<td>����Ÿ��</td>
							<td>SCM<br>������ȣ</td>
						</tr>
						<tr bgcolor="#FFFFFF" align="center">
							<td></td>
							<td><input type="text" class="text" name="sRN" size="40" maxlength="50" value="<%=sEappName%>"></td>
							<td><input type="text" class="text" name="mRP" size="15" maxlength="20" style="text-align:right;" value="<%=mReportPrice%>" <%=chkIIF(mReportPrice<>"","readonly","")%> onKeypress="num_check()" onkeyup="auto_amount(this.form,this)" onblur="jsIsHundred();"></td>
							<td>
								<select name="selPT" onChange="jsChFC();" class="select" <%IF not blnPayEapp THEN%>disabled<%END IF%>>
								<%sboptPayType ipaytype%>
							</select>
							<div  id="spCurr" style="display:<%IF ipaytype<>"1" or isNull(ipaytype) THEN%>none<%END IF%>;"> <%DrawexchangeRate "selCT",sCurrencyType,""%><input type="text" name="sCP" value="<%=sCurrencyPrice%>" size="10" style="text-align:right;"> </div>
							</td>
							<td><input type="hidden" name="iSL" value="<%=iscmlinkno%>"><%=iscmlinkno%> <%IF sscmLink <> "" THEN%>><A href="javascript:jsGoScm('<%=sscmLink%>','<%=iscmlinkno%>');">>�󼼺���</a><%END IF%></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr align="center">
								<td  bgcolor="<%= adminColor("tabletop") %>" width="60" rowspan="3">����</td>
							<td bgcolor="#FFFFFF" height="100">
							<textarea name="editor" id="editor" style="width: 100%; height: 490px;"><%=tContents%></textarea>
                				<!-- daumeditor  -->
                				<script type="text/javascript">
                				    EditorCreator.convert(document.getElementById("editor"), '/lib/util/daumeditor/teneditor/editorForm.html', function () {
                				        EditorJSLoader.ready(function (Editor) {
                				            new Editor(config);
                				            Editor.modify({
                				                content:  '<%=tContents%>'
                				            });
                				        });
                				    });

                				</script>
                				<!-- daumeditor   -->
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<% if addFileName<>"" then %>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr>
							<td bgcolor="<%= adminColor("tabletop") %>" width="60" align="center">���ü���</td>
							<td bgcolor="#FFFFFF"><span onclick="jsEdmsDownload('<%=uploadImgUrl%>','<%=addFileName%>','<%=addFileNamePh%>');" style="cursor:pointer;" title="���ü��� ��� �ٿ�ε�">�� <%=addFileName%></span></td>
						</tr>
						</table>
					</td>
				</tr>
				<% end if %>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspaciNg="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="60">÷�μ���</td>
							<td>÷������</td>
							<td>���ø�ũ</td>
						</tr>
						<tr  bgcolor="#FFFFFF">
							<td align="center" valign="top">
								<input type="button" value="����÷��" class="button" onClick="jsAttachFile('');">
								<div id="dFile"></div>
								<input type="hidden" name="sFile" value="">
							</td>
							<td><input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120"><br>
								<input type="text" name="sL" size="60" maxlength="120">
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
							<td rowspan="2" width="60">��������</td>
							<td>�����׸�</td>
							<td>�����������</td>
						</tr>
						<tr bgcolor="#FFFFFF"  align="center">
							<!--
							<td>[<%=iarapcd%>] <%=sarap_nm%></td>
							<td>[<%=sacc_use_cd%>] <%=sacc_nm%></td>
							<input type="hidden" name="iaidx" value="<%=iarapcd%>" class="text">
							<input type="hidden" name="sACC" value="<%=sacc_cd%>"  class="text">
							-->
							<td><input type="text" name="sANM" value="<%=CHKIIF(isNULL(sarap_nm),"","["&iarapcd&"]"&sarap_nm)%>" style="border:0;width:100%" readonly ></td>
					        <td><input type="text" name="sACCNM" value="<%=CHKIIF(isNULL(sacc_nm),"","["&sacc_use_cd&"]"&sacc_nm)%>" style="border:0;width:100%" readonly></td>

						</tr>
						<tr bgcolor="#FFFFFF">
        					<td colspan="3"><input type="button" class="button" value="�����׸� ����" onClick="jsGetARAP();"></td>
        				</tr>
						</table>
					</td>
				</tr>
				<%IF (blnPayEapp) THEN%>
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="0" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
						<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
								<td width="60" rowspan="2" style="padding:5px">�μ���<br>�ڱݱ���</td>
							<td width="300"  style="padding:5px"> �μ�</td>
							<td width="205" style="padding:5px"> �ݾ�</td>
							<td width="205" style="padding:5px"> %</td>
						</tr>
						<tr>
							<td colspan="3" bgcolor="#FFFFFF" valign="top">
                                <% if insTable = True then %>
                                <div id="divPM"><table border="0" cellpadding="3" cellspacing="0" class="a" width="760"><tbody><tr><td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;">�¶��λ����</td><td width="140" align="center" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;">PB</td><td align="center" width="200" style="border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;"><input type="text" name="mPM" id="mPM" value="" size="20" style="text-align:right;" onkeyup="jsSetMoney('m','0','1');auto_amount(this.form,this)" onkeypress="num_check()">��</td><td align="center" width="200" style="border-bottom:1px solid #BABABA;"><input type="text" name="iPM" id="iPM" value="" size="4" style="text-align:right;" onkeyup="jsSetMoney('i','0','1')">%</td></tr></tbody></table></div>
                                <input type="hidden" name="iP" id="iP" value="FDBCCEFBFF">
							    <input type="hidden" name="sP" id="sP" value="PB">
							    <input type="hidden" name="mP" id="mP" value="">
                                <% else %>
							    <div id="divPM"></div><br>
                                <input type="hidden" name="iP" id="iP" value="">
							    <input type="hidden" name="sP" id="sP" value="">
							    <input type="hidden" name="mP" id="mP" value="">
                                <% end if %>

							    &nbsp;<input type="button" value="�μ����" onClick="jsSetPartMoney(1,'<%=sacc_use_cd%>','<%=sACC_GRP_CD%>');" class="button" ><Br><Br>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<%END IF%>
				<tr>
					<td align="center" width="100%">
						<table border="0" cellpadding="5" cellspacing="0" width="100%">
							<tr>
								<% if iarapcd<>"351" then '��Ÿ�ν�û(�޿�)�� �ӽ����� ���� %>
								<td align="left"><input type="button" value="�ӽ�����" class="button" onclick="jsEappSubmit(0);"></td>
								<% end if %>
								<td align="right"><input id="btnSm" type="button" value="������" style="color:blue;" class="button" onclick="jsEappSubmit(1);"></td>
							</tr>
						</table>
					</td>
				</tr>
				</table>
				</form>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
		<!-- #include virtual="/lib/db/dbclose.asp" -->
</body>
</html>
