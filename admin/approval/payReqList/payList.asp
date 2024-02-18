<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 급여 리스트
' History : 2012.02.15 정윤정  생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->  
<!-- #include virtual="/lib/classes/approval/payreqListCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->  
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
Dim clsPay 
Dim ipayrequeststate ,sadminId
Dim ipayrequestidx
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim arrList,intLoop 
Dim ipayRequestType,spayRequestTitle  
Dim searchType, searchsdate, searchedate, blnTakeDoc, sUserName
Dim arrAccount, intA ,arrAccConts, intAC
Dim iarap_cd,sarap_nm, research, iOutBank, selBiz, payReqType, PayReqPRice, notIncEtc,sCustNm

	iPageSize = 30
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1 
	  
	sadminId =  session("ssBctId")
	ipayrequestIdx	= requestCheckvar(Request("ipridx"),10)  
	
	searchType = requestCheckvar(Request("selST"),1)  
	searchsdate= requestCheckvar(Request("selSD"),10) 
	searchedate= requestCheckvar(Request("selED"),10) 
	iarap_cd		= requestCheckvar(Request("iaidx"),13) 
	sarap_nm		= requestCheckvar(Request("selarap"),50) 
    ipayrequeststate= requestCheckvar(Request("selPRS"),4)  
	blnTakeDoc= requestCheckvar(Request("selTD"),1)  
	sUserName= requestCheckvar(Request("sUnm"),30)  
	research = requestCheckvar(Request("research"),10)  
	iOutBank = requestCheckvar(Request("selOB"),30)
	selBiz   = requestCheckvar(Request("selBiz"),30)
	payReqType = requestCheckvar(Request("payReqType"),30)
	PayReqPRice = requestCheckvar(Request("PayReqPRice"),10)
	notIncEtc   = requestCheckvar(Request("notIncEtc"),10)
	sCustNm			= requestCheckvar(Request("sCnm"),50)
if (research="") and (ipayrequeststate="") then ipayrequeststate="255"
 notIncEtc="in"

'결재 기본 폼 정보 가져오기
set clsPay = new CPayReqList  
	clsPay.FpayRequestType	        = ipayRequestType
	clsPay.FSearchType		        = searchType
	clsPay.FSDate					= searchsdate
	clsPay.FEDate					= searchedate
 	clsPay.Farap_cd					= iarap_cd 
	clsPay.Fpayrequeststate         = ipayrequeststate
	clsPay.FisTakeDoc				= blnTakeDoc
	clsPay.FUsername				= sUserName  
	clsPay.FOutBank				    = iOutBank  
	clsPay.FBizSection_CD           = selBiz
	clsPay.FpayRequestType          = payReqType
	clsPay.Fpayrequestprice         = PayReqPRice
	clsPay.FnotIncEtc               = notIncEtc
	clsPay.FCustNm               = sCustNm
	clsPay.FCurrpage 				= iCurrpage
	clsPay.FPagesize				= ipagesize
	arrList = clsPay.fnGetPayReqAllList
	iTotCnt = clsPay.FTotCnt 

set clsPay = nothing

 
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수 

Dim TotSum : TotSum=0
%>  


					
  
 <script language="javascript" src="/admin/approval/eapp/eapp.js"></script> 
<script language="javascript">
<!--
	function jsNewReg(){
		var winR = window.open("regPayRequest.asp","popR","width=1000, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}
	
	function jsMod(ipridx){
		var winR = window.open("regPayRequest.asp?ipridx="+ipridx,"popR","width=1000, height=600, resizable=yes, scrollbars=yes");
		winR.focus();
	}
	
	function jsSearch(){
	 document.frm.submit();
	}
	
	// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	} 
 
	//수지항목 불러오기
 	function jsGetARAP(){ 
 			var winARAP = window.open("/admin/linkedERP/arap/popGetARAP.asp","popARAP","width=600,height=600,resizable=yes, scrollbars=yes");
 			winARAP.focus();
 	}
 	
 	function jsReSetARAP(){
 			document.frm.iaidx.value = 0;
 			document.frm.selarap.value = "";
 	}
 	
 	//선택 수지항목 가져오기
 	function jsSetARAP(dAC, sANM,sACC,sACCNM){ 
 		document.frm.iaidx.value = dAC; 
 		document.frm.selarap.value = sANM;  
 	}
 
 
function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

 
function popConfirmPayrequest(iridx,pidx){
    var iURI = '/admin/approval/eapp/confirmpayrequest.asp?iridx='+iridx+'&ipridx='+pidx+'&ias=1'; //ias 확인..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popModPayDoc(iridx,pidx){
	 var iURI = '/admin/approval/eapp/modeappPayDoc.asp?iridx='+iridx+'&ipridx='+pidx ; //ias 확인..
    var popwin = window.open(iURI,'popConfirmPayrequest','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus(); 
}
//-->
</script>  
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">   
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<form name="frm" method="get" action="payList.asp">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iPS" value="">
			<input type="hidden" name="research" value="on">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="3" width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left"> 
						<select name="selST">
							<option value="1" <%IF searchType="1" THEN%>selected<%END IF%>>결제요청일</option>
							<option value="2" <%IF searchType="2" THEN%>selected<%END IF%>>결제예정일</option>
							<!--<option value="3" <%IF searchType="3" THEN%>selected<%END IF%>>결제(입금)일</option>-->
						</select>
						<input type="text" name="selSD" size="10" value="<%=searchSDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selSD');"  style="cursor:hand;">
						~
						<input type="text" name="selED" size="10" value="<%=searchEDate%>"><img src="/images/calicon.gif" align="absmiddle" border="0" onClick="jsPopCal('selED');"  style="cursor:hand;">
						&nbsp;&nbsp;
					작성자:
					<input type="text" name="sUnm" size="8" value="<%=sUserName%>">
					&nbsp;&nbsp;
					거래처:
						<input type="text" name="sCnm" size="20" value="<%=sCustNm%>">
				</td>
				<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
				</td>
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>
					수지항목 :
					<input type="hidden" name="iaidx" value="<%=iarap_cd%>" >
					<input type="text" name="selarap" value="<%=sarap_nm%>" size="13" onClick="jsGetARAP();" readonly> 
					&nbsp;&nbsp; 
					결제상태:
					<select name="selPRS">
						<option value="">----</option>
						 <%sbOptPayRequestState ipayrequeststate%>
					</select>
					&nbsp;&nbsp;
					출금예정은행:
					<select name="selOB">
					<option value="">--선택--</option>
					<%
					Dim clsComm
					
					set clsComm = new CcommCode
					clsComm.Fparentkey = 2 '은행명 가져오기
					clsComm.Fcomm_cd = iOutBank
					clsComm.sbOptCommCD
					set clsComm = nothing
					%>
					</select>
					&nbsp;&nbsp;
					사업부문:
					<%
					Dim clsBS, arrBizList
					Set clsBS = new CBizSection 
                    	clsBS.FUSE_YN = "Y"  
                    	clsBS.FOnlySub = "Y"  
                    	arrBizList = clsBS.fnGetBizSectionList  
                    Set clsBS = nothing
                    %>
                    <select name="selBiz">
                    <option value="">--선택--</option>
                    <% For intLoop = 0 To UBound(arrBizList,2)	%>
                		<option value="<%=arrBizList(0,intLoop)%>" <%IF Cstr(selBiz) = Cstr(arrBizList(0,intLoop)) THEN%> selected <%END IF%>><%=arrBizList(1,intLoop)%></option>
                	<% Next %>
                    </select>
                    
					&nbsp;&nbsp;
					서류제출여부:
					<select name="selTD">
						<option value="">----</option>
						<option value="1" <%IF blnTakeDoc = "1" THEN%>selected<%END IF%>>Y</option>
						<option value="0" <%IF blnTakeDoc ="0" THEN%>selected<%END IF%>>N</option>
					</select>
				</td> 
			</tr>
			<tr  bgcolor="#FFFFFF" >
				<td>  
					결제요청금액:
					<input type="text" name="PayReqPRice" size="10" value="<%= PayReqPRice %>">
					 
				</td> 
			</tr>
			</form>
		</table>
	</td>
</tr>  
<tr>
	<td>
	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	        <td align="left"><input type="button" class="button" value="신규등록" onClick="jsNewReg();"></td> 
	    </tr>
	    </table>
	</td>
</tr> 
<tr>
	<td> 
		<!-- 상단 띠 시작 -->
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0">
		<Form name="frmAct" method="post" action="erpLink_Process.asp">
		<input type="hidden" name="LTp" value="A">
				<tr bgcolor="<%= adminColor("tabletop") %>" align="center"> 
				    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
					<td>결제<br>Idx</td>
					<td width="120">자금용도</td>
					<td>수지항목</td>
					<td>부서구분</td> 
					<td>거래처</td> 
					<td>출금은행</td> 
					<td>결제요청금액</td>
					<td>결제요청일</td> 
					<td>결제예정일</td> 
					<!--<td>결제(입금)일</td>  -->
					<td>작성자</td>
					<td>작성일</td>
					<td>결제상태</td>  
					<td>서류<br>종류</td> 
					<td>서류<br>제출</td>  
					<td>결제<br>확인</td>
				</tr>
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
					TotSum = TotSum + arrList(6,intLoop)
				%> 
				<tr bgcolor="#FFFFFF" align="center">  
				    <td><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%= CHKIIF((arrList(16,intLoop)="7") AND (TRUE),"","disabled") %> ></td>
					<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></a></td>
					<td width="120"><%=arrList(3,intLoop)%></td>
					<td><%=arrList(19,intLoop)%><br>[<%=arrList(27,intLoop)%>]&nbsp;<font color=gray><%=arrList(28,intLoop)%></font></td>
					<td><%=arrList(25,intLoop)%></td> 
					<td><%=arrList(29,intLoop)%></td> 
					<td><%=arrList(24,intLoop)%></td> 
					<td><%=formatnumber(arrList(6,intLoop),0)%></td>
					<td><%=arrList(5,intLoop)%></td>  
					<td><%=arrList(10,intLoop)%></td>  
					<!--<td><%=arrList(12,intLoop)%></td>-->
					<td><%=arrList(21,intLoop)%></td>  
					<td><%=Replace(arrList(18,intLoop)," 오","<br>오")%></td>
					<td><%=fnGetPayRequestState(arrList(16,intLoop))%></td>  
					<td><%=arrList(26,intLoop)%><br>
							<img src="/images/icon_arrow_link.gif" onClick="popModPayDoc(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>);" style="cursor:pointer">
						</td> 
					<td><%IF arrList(14,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>   
					<td>
					<% if arrList(16,intLoop)=1 Then %>
					<img src="/images/icon_arrow_link.gif" onClick="popConfirmPayrequest(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>);" style="cursor:pointer">
					<% end if %>
					</td>
				</tr>
				<%	
					Next
				%>
				<tr>
				    <td></td>
				    <td colspan="6"></td>
				    <td><%=formatnumber(TotSum,0) %></td>
				    <td colspan="10"></td>
				</tr>
				<%
					ELSE	
				%>
				<tr bgcolor="#FFFFFF">
					<td colspan="16" align="center">등록된 내역이 없습니다.</td>
				</tr>
				<%END IF%>
				</table>
			</td>
		</tr>
        </form>
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>        
					    </tr>        
					</table>
				</td>
			</tr>
			</table>
	</td> 
</tr> 	 
</table>  
</body>
</html> 

<!-- #include virtual="/lib/db/dbclose.asp" -->  