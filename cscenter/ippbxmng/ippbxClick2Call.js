
function escapeRegExp(str) {
    return str.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}

function replaceAll(str, find, replace) {
	return str.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function fnClick2Call(comp) {
    var iphoneNum = comp.value;

	iphoneNum = replaceAll(iphoneNum, "-", "");

    if (iphoneNum.length<1) {
        alert('전화번호를 입력하세요.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

	var contents = null;
	var currObj;

	try {
		if (top.opener) {
			try {
				// 헤더에서 주문내역조회창 연 경우
				if (top.opener.name == "i_ippbxmng") {
					top.opener.click2call(iphoneNum);
					return;
				}
			} catch (err) {
				alert("에러[0] : " + err.message);
				return;
			}
		}

		// =====================================================================
		// 시작점
		// =====================================================================
		if (top.opener) {
			currObj = top.opener;
		} else if (window.parent) {
			currObj = window.parent;
		} else {
			alert("잘못된 접근입니다.[0]");
			return;
		}

		// =====================================================================
		// 상위 윈도우 탐색
		// =====================================================================
		for (var i = 0; i < 20; i++) {
			if (currObj.name == "contents") {
				contents = currObj;
				break;
			}

			if (currObj.opener) {
				currObj = currObj.opener;
			} else if (currObj.parent) {
				currObj = currObj.parent;
			} else {
				alert("잘못된 접근입니다.[1]");
				return;
			}

			if (currObj.name == "") {
				alert("잘못된 접근입니다.[2]");
				return;
			}
		}

		if (contents == null) {
			alert("시스템팀 문의\n\ncontents 찾기에 실패!!");
			return;
		}

		if (contents.parent.header) {
			if (contents.parent.header.i_ippbxmng) {
				// 고객센터 접속인 경우
				contents.parent.header.i_ippbxmng.click2call(iphoneNum);
				return;
			} else {
				// 그 이외
				alert("고객센터에서만 사용가능한 기능입니다.");
				return;
			}
		}

	} catch (err) {
		// TODO : 도메인이 http - https 로 각각 다르게 되면 엑세스 제한 에러난다.
		alert("에러[1] : " + err.message);
	}
}

function fnClick2Call_TEST(comp) {
    var iphoneNum = comp.value;

	iphoneNum = replaceAll(iphoneNum, "-", "");

    if (iphoneNum.length<1) {
        alert('전화번호를 입력하세요.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

	var contents = null;
	var currObj;

	try {
		if (top.opener) {
			try {
				// 헤더에서 주문내역조회창 연 경우
				if (top.opener.name == "i_ippbxmng") {
					top.opener.click2call(iphoneNum);
					return;
				}
			} catch (err) {
				alert("에러[0] : " + err.message);
				return;
			}
		}

		// =====================================================================
		// 시작점
		// =====================================================================
		if (top.opener) {
			currObj = top.opener;
		} else if (window.parent) {
			currObj = window.parent;
		} else {
			alert("잘못된 접근입니다.[0]");
			return;
		}

		// =====================================================================
		// 상위 윈도우 탐색
		// =====================================================================
		for (var i = 0; i < 20; i++) {
			if (currObj.name == "contents") {
				contents = currObj;
				break;
			}

			if (currObj.opener) {
				currObj = currObj.opener;
			} else if (currObj.parent) {
				currObj = currObj.parent;
			} else {
				alert("잘못된 접근입니다.[1]");
				return;
			}

			if (currObj.name == "") {
				alert("잘못된 접근입니다.[2]");
				return;
			}
		}

		if (contents == null) {
			alert("시스템팀 문의\n\ncontents 찾기에 실패!!");
			return;
		}

		if (contents.parent.header) {
			if (contents.parent.header.i_ippbxmng) {
				// 고객센터 접속인 경우
				contents.parent.header.i_ippbxmng.click2call(iphoneNum);
				return;
			} else {
				// 그 이외
				alert("고객센터에서만 사용가능한 기능입니다.");
				return;
			}
		}

	} catch (err) {
		// TODO : 도메인이 http - https 로 각각 다르게 되면 엑세스 제한 에러난다.
		alert("에러[1] : " + err.message);
	}
}
