<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 자금관리 부서
' History : 2011.04.21 정윤정  생성
'	itype  = 1-품의서, 2-결제요청서, 9-업무포지셔닝(비율)
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim clsBS
Dim arrList, intLoop
Dim sUSEYN,sBS_NM,iType,sACC_GRP_CD,sACC_USE_CD
Dim blnView, blnSale
sBS_NM = requestCheckvar(Request("sBS_NM"),100)
sUSEYN = requestCheckvar(Request("sUSEYN"),3)
iType = requestCheckvar(Request("iType"),1)
sACC_GRP_CD = requestCheckvar(Request("sACCGRP"),3)
sACC_USE_CD = requestCheckvar(Request("sAUCD"),15)
blnView = "Y"
blnSale = "Y"
''blnSale = fnCheckBizSale(sACC_USE_CD,sACC_GRP_CD)
blnSale = "N" ''전체표시로 변경 2013/11/14

Set clsBS = new CBizSection
	clsBS.FBS_NM 	= sBS_NM
	clsBS.FUSE_YN = "Y"
	clsBS.FView		= blnView
	clsBS.FSale		= blnSale
	arrList = clsBS.fnGetBizSectionList
Set clsBS = nothing
%>
 <script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--

//부서 등록
function jsSetPart(){

	var strT  = "<table border=0 cellpadding=3 cellspacing=0 class=a width=760>"	;
 var iCount = 0;
   for(i=0;i<document.frmReg.chkV.length;i++){
   	if(document.frmReg.chkV[i].checked){
       if(iCount==0){
   		opener.document.frm.iP.value = document.frmReg.chkV[i].value;
   		opener.document.frm.sP.value = document.frmReg.sNM[i].value;
   		}else{
   		opener.document.frm.iP.value = opener.document.frm.iP.value +","+ document.frmReg.chkV[i].value;
   		opener.document.frm.sP.value = opener.document.frm.sP.value +","+ document.frmReg.sNM[i].value;
   		}
   		strT = strT+  "<tr><td  width='140' align='center' style='border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;'>"+document.frmReg.hidPM[i].value+"</td><td width='140'  align='center' style='border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;'>"+document.frmReg.sNM[i].value+"</td><td align='center' width='200'  style='border-bottom:1px solid #BABABA;border-right:1px solid #BABABA;'><input type='text' name='mPM' id='mPM'  value='' size='20' style='text-align:right;' onKeyUp=\"jsSetMoney('m','"+iCount+"','<%=iType%>');auto_amount(this.form,this)\" onKeypress=\"num_check()\" >원</td><td align='center' width=200  style='border-bottom:1px solid #BABABA;'><input type='text' name='iPM' id='iPM' value='' size='4'  style='text-align:right;' onKeyUp=jsSetMoney('i','"+iCount+"','<%=iType%>')>%</td></tr>";
   		iCount = iCount  + 1;
   	}
	}
	strT = strT+"</table>";

	opener.document.all.divPM.innerHTML = strT;
	self.close();
}


$(window).load(function(){ //페이지 로드시
	if($("#iP",window.opener.document).val() != ""){ //기존 선택값 있을 경우
		var arrI = $("#iP",window.opener.document).val().split(",");
		var arrN = $("#sP",window.opener.document).val().split(",");

		for(i=0;i<arrI.length;i++){
			 for(j=0;j<document.frmReg.chkV.length;j++){
			 	if(document.frmReg.chkV[j].value == arrI[i]){
			 		document.frmReg.chkV[j].checked = true;
			 	}
			}
		}
	}
});


   //검색
   function jsSearch(){
    document.frmS.submit();
   }
//-->
</script>
<form name="frmS" method="get" action="popGetBiz.asp" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
	<tr>
	<td><strong>부서  선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
		<tr align="center" bgcolor="#FFFFFF" >
			<td align="left">&nbsp;
			 부서명: <input type="text" name="sBS_NM" size="20" value="<%=sBS_NM%>">
			</td>
			<td rowspan="2" width="50" bgcolor="#EEEEEE">
				<input type="button" class="button_s" value="검색" onClick="jsSearch();">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="right"><input type="button" class="button" value="선택추가" onClick="jsSetPart();"></td>
</tr>
</table>
</form>
<form name="frmReg" method="post" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td>부서명</td>
			</tr>
			<%  Dim oldPCD
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
					IF oldPCD <> arrList(2,intLoop) THEN
				%>
				<tr bgcolor="#FFFFFF"  height=30 >
					<td><%=arrList(2,intLoop)%>&nbsp; <%=arrList(4,intLoop)%></td>
				</tr>
				<%END IF%>
			<tr height=30 align="center" bgcolor="<%IF arrList(3,intLoop) ="N" THEN%>#EFEFEF<%ELSE%>#FFFFFF<%END IF%>">
				<td align="left"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					 └ <input type="checkbox" value="<%=arrList(0,intLoop)%>" name="chkV" <%= CHKIIF(arrList(7,intLoop)="0000009010","disabled","") %> ><input type="hidden" name="sNM" value="<%=arrList(1,intLoop)%>">
					 <%=arrList(7,intLoop)%>&nbsp; <%=arrList(1,intLoop)%>
						<% if arrList(7,intLoop)<>arrList(0,intLoop) then %>
						&nbsp;<font color="#CCCCCC">(<%=arrList(0,intLoop)%>)</font>
						<% end if %>
					 <%IF arrList(4,intLoop) <> "" THEN%><input type="hidden" name="hidPM" value="<%=arrList(4,intLoop)%>"><%END IF%>
					 </td>
			</tr>
		<%		oldPCD  = arrList(2,intLoop)
				Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">
				<td colspan="2">등록된 내용이 없습니다.</td>
			</tr>
		<%END IF%>
		</table>
	</td>
</tr>
</table>
</form>
<!-- 페이지 끝 -->
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" -->
