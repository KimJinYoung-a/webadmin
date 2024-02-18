<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->

<%
	''#####################################################################################################################################
	''	작업중인 파일입니다. 수정하려면 강준구에게 먼저 알려주세요.
	''#####################################################################################################################################
	
	Dim cDisp, vWidth, vDepth, vCateCode, vCateName, vUseYN, vSortNo, vCurrpage, vInputName
	vCurrpage 	= NullFillWith(Request("cpg"), "1")
	vDepth 		= NullFillWith(Request("depth"), "1")
	vCateCode 	= Request("catecode")
	vCateName	= Request("catename_s")
	vUseYN 		= Request("useyn_s")
	vSortNo 	= Request("sortno_s")
	vInputName	= Request("inputname")
	
	'vWidth = CInt((100/vDepth))

	SET cDisp = New cDispCate
	cDisp.FCurrPage = vCurrpage
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = vDepth
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateList()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
document.domain = "10x10.co.kr";

function jsThis(cname,c){
	opener.$("input[name=<%=vInputName%>]").val(cname);
	opener.$("input[name=<%=vInputName%>code]").val(c);
	window.close();
}
</script>
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:7px 10px;}
.box2 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:5px; margin-top:5px;}
.box3 {width:<%=(286*vDepth)%>px; margin-top:5px;}
.box3 .subFirstBox {width:260px; border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; float:left; margin-left:0px;}
.box3 .subBox {width:260px; border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; float:left; margin-left:5px;}
.box3 .subTTBox {border:0; border-radius: 6px; padding:3px 0; text-align:center; background-color:#888; color:#FFF; font-weight:bold;}
.box3 .subListBox {margin-top:5px;}
.box4 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FAFAFA; padding:7px 10px; ; margin-top:5px;}
.ttDep1 {background-color:#FAFAFA;}
.ttDep2 {background-color:#F5F5F5;}
.ttDep3 {background-color:#EFEFEF;}
.ttDep4 {background-color:#ECECEC;}
.ttDep5 {background-color:#E8E8E8;}
.ttDep6 {background-color:#E0E0E0;}
</style>

<div class="box3">
	<div class="subFirstBox ttDep1">
		<div class="subTTBox">1 Depth</div>
		<div class="subListBox">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
			<%
			Dim i, vTempDepth, vIsThisLine, vNowCateName, vBGcolor
			For i=0 To cDisp.FResultCount-1
				vIsThisLine = fnIsThisLine(cDisp.FItemList(i).FDepth,cDisp.FItemList(i).FCateCode,vCateCode)
				If vIsThisLine = "o" Then
					vNowCateName = vNowCateName & "[" & Right(cDisp.FItemList(i).FCateCode,3) & "]" & cDisp.FItemList(i).FCateName & " - "
				End IF
				
				If i=0 Then
					vTempDepth = cDisp.FItemList(i).FDepth
				End IF
	
				If vTempDepth <> cDisp.FItemList(i).FDepth Then
					Response.Write "	</table>" & vbCrLf &_
								"	</div>" & vbCrLf &_
								"</div>" & vbCrLf &_
								"<div class='subBox ttDep" & cDisp.FItemList(i).FDepth & "'>" & vbCrLf &_
								"	<div class='subTTBox'>" & cDisp.FItemList(i).FDepth & " Depth</div>" & vbCrLf &_
								"	<div class='subListBox'>" & vbCrLf &_
								"	<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' class='a' bgcolor='#CCCCCC'>"
				End If
				
				vBGcolor = "#FFFFFF"
				If vIsThisLine = "o" Then
					vBGcolor = "#FFF0F0"
				End If
				If cDisp.FItemList(i).FUseYN = "N" Then
					vBGcolor = "#CFCFCF"
				End If
			%>
			<tr>
				<td bgcolor="<%=vBGcolor%>" width="260">
					<table width="100%" border=0 class=a>
					<tr>
						<td>
							<a href="?menupos=<%=Request("menupos")%>&depth=<%=cDisp.FItemList(i).FDepth+1%>&catecode=<%=cDisp.FItemList(i).FCateCode%>&inputname=<%=vInputName%>"><%=cDisp.FItemList(i).FCateName%></a>
							[<a href="javascript:jsThis('<%=cDisp.FItemList(i).FCateName%>','<%=cDisp.FItemList(i).FCateCode%>');"><font color="blue"><b>선택하기</b></font></a>]
						</td>
						<td width="5" align="right"><%=CHKIIF(vIsThisLine="o","▶","")%></td>
					</tr>
					</table>
				</td>
			</tr>
			<%
				vTempDepth = cDisp.FItemList(i).FDepth
			Next
			%>
			</table>
		</div>
	</div>
</div>
<br>
<center><input type="button" value="닫   기" onClick="window.close();"></center>
<br style="clear:both;">
<% SET cDisp = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->