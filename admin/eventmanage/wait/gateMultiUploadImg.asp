<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결재라인 등록
' History : 2011.03.16 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />

<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<%
dim sFilePath1, sFileName1, sType, pvWidth
dim sFilePath2, sFileName2,sFilePath3, sFileName3

  sFilePath1 = requestCheckVar(Request("SFP1"),128) 
  sFileName1 = requestCheckVar(Request("SFN1"),60) 	 
  sFilePath2 = requestCheckVar(Request("SFP2"),128) 
  sFileName2 = requestCheckVar(Request("SFN2"),60) 	 
  sFilePath3 = requestCheckVar(Request("SFP3"),128) 
  sFileName3 = requestCheckVar(Request("SFN3"),60) 	 
  sType = requestCheckVar(Request("sType"),10) 
  pvWidth = requestCheckVar(Request("pvWidth"),10) 
  if pvWidth="" then pvWidth=105
   
  if sType ="tms" then
%>
<div id="pvImg">
		<div class="slide">
			<%if sFilePath1 <>"" then %><div id="tms1Img"><img src="<%=sFilePath1%>"  /></div><%end if%>
			<%if sFilePath2 <>"" then %><div id="tms2Img"><img src="<%=sFilePath2%>"  /></div><%end if%>
			<%if sFilePath3 <>"" then %><div id="tms3Img"><img src="<%=sFilePath3%>" /></div><%end if%>
		</div>
</div>
<div id="pvNm"> 											
	<%if sFilePath1 <>"" then %>
		<span id="tms1Nm"><p class="tMar05 fs11" ><%=sFileName1%><button type="button" onclick="jsDelimg('tms1');" >X</button></p></span>		
		<%end if%>
	<%if sFilePath2 <>"" then %>
	<span id="tms2Nm"><p class="tMar05 fs11" ><%=sFileName2%><button type="button" onclick="jsDelimg('tms2');" >X</button></p></span>		
		<%end if%>
	<%if sFilePath3 <>"" then %>
	<span id="tms3Nm"><p class="tMar05 fs11" ><%=sFileName3%><button type="button" onclick="jsDelimg('tms3');" >X</button></p></span>		
		<%end if%>
</div>
<script type="text/javascript" charset="euc-kr">
<!--
$(document).ready(function(){  	 
 	 var sValue = $("#pvImg").html();  
 	 var sName = $("#pvNm").html();  
 	 $(opener.document).find(".fullTemplatev17 .fullContV17 .slide").remove(); 
	// $(opener.document).find("#tmsNm").empty();    
 	  	 
	 $(opener.document).find(".fullTemplatev17 .fullContV17").append(sValue);   
	// $(opener.document).find("#tmsNm").append(sName);   	 
	 
	 $(opener.document).find("#hidtms1").val("<%=sFilePath1%>")
	 $(opener.document).find("#hidtms2").val("<%=sFilePath2%>")
	 $(opener.document).find("#hidtms3").val("<%=sFilePath3%>")
	  
	 opener.jsRollingbg('B');
 self.close();
});
//-->
</script>
<%else%>
<div id="pvImg">
	<div class="swiper-container">
		<div class="swiper-wrapper" id="tmsmImg">  	
			<%if sFilePath1 <>"" then %><div class="swiper-slide" id="tmsm1Img"><div class="thumbnail"><img src="<%=sFilePath1%>"  /></div></div><%end if%>
			<%if sFilePath2 <>"" then %><div class="swiper-slide" id="tmsm2Img"><div class="thumbnail"><img src="<%=sFilePath2%>"  /></div></div><%end if%>
			<%if sFilePath3 <>"" then %><div class="swiper-slide" id="tmsm3Img"><div class="thumbnail"><img src="<%=sFilePath3%>" /></div></div><%end if%>
		</div> 
		<div class="pagination-line"></div>
		<button type="button" class="btnNav btnPrev">이전</button>
		<button type="button" class="btnNav btnNext">다음</button>
	</div>
</div>
<div id="pvNm"> 											
	<%if sFilePath1 <>"" then %>
		<span id="tmsm1Nm"><p class="tMar05 fs11" ><%=sFileName1%><button type="button" onclick="jsDelimg('tmsm1');" >X</button></p></span>		
		<%end if%>
	<%if sFilePath2 <>"" then %>
	<span id="tmsm2Nm"><p class="tMar05 fs11" ><%=sFileName2%><button type="button" onclick="jsDelimg('tmsm2');" >X</button></p></span>		
		<%end if%>
	<%if sFilePath3 <>"" then %>
	<span id="tmsm3Nm"><p class="tMar05 fs11" ><%=sFileName3%><button type="button" onclick="jsDelimg('tmsm3');" >X</button></p></span>		
		<%end if%>
</div>
<script type="text/javascript" charset="euc-kr">
<!--
$(document).ready(function(){  	 
 	 var sValue = $("#pvImg").html();  
 	 var sName = $("#pvNm").html();  
 	 $(opener.document).find("#mdRolling .swiper-container").remove(); 
	// $(opener.document).find("#tmsmNm").empty();    
 	  	 
	 $(opener.document).find("#mdRolling").append(sValue);   
	// $(opener.document).find("#tmsmNm").append(sName);   	 
	 
	  $(opener.document).find("#hidtmsm1").val("<%=sFilePath1%>")
	 $(opener.document).find("#hidtmsm2").val("<%=sFilePath2%>")
	 $(opener.document).find("#hidtmsm3").val("<%=sFilePath3%>")
	 opener.jsRollingbgM('B');
 self.close();
});
//-->
</script>
<%end if %>
 
 	
															 