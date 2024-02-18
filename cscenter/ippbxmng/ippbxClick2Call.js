
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
        alert('��ȭ��ȣ�� �Է��ϼ���.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

	var contents = null;
	var currObj;

	try {
		if (top.opener) {
			try {
				// ������� �ֹ�������ȸâ �� ���
				if (top.opener.name == "i_ippbxmng") {
					top.opener.click2call(iphoneNum);
					return;
				}
			} catch (err) {
				alert("����[0] : " + err.message);
				return;
			}
		}

		// =====================================================================
		// ������
		// =====================================================================
		if (top.opener) {
			currObj = top.opener;
		} else if (window.parent) {
			currObj = window.parent;
		} else {
			alert("�߸��� �����Դϴ�.[0]");
			return;
		}

		// =====================================================================
		// ���� ������ Ž��
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
				alert("�߸��� �����Դϴ�.[1]");
				return;
			}

			if (currObj.name == "") {
				alert("�߸��� �����Դϴ�.[2]");
				return;
			}
		}

		if (contents == null) {
			alert("�ý����� ����\n\ncontents ã�⿡ ����!!");
			return;
		}

		if (contents.parent.header) {
			if (contents.parent.header.i_ippbxmng) {
				// ������ ������ ���
				contents.parent.header.i_ippbxmng.click2call(iphoneNum);
				return;
			} else {
				// �� �̿�
				alert("�����Ϳ����� ��밡���� ����Դϴ�.");
				return;
			}
		}

	} catch (err) {
		// TODO : �������� http - https �� ���� �ٸ��� �Ǹ� ������ ���� ��������.
		alert("����[1] : " + err.message);
	}
}

function fnClick2Call_TEST(comp) {
    var iphoneNum = comp.value;

	iphoneNum = replaceAll(iphoneNum, "-", "");

    if (iphoneNum.length<1) {
        alert('��ȭ��ȣ�� �Է��ϼ���.');
        if (!comp.disabled) { comp.focus(); };
        return;
    }

	var contents = null;
	var currObj;

	try {
		if (top.opener) {
			try {
				// ������� �ֹ�������ȸâ �� ���
				if (top.opener.name == "i_ippbxmng") {
					top.opener.click2call(iphoneNum);
					return;
				}
			} catch (err) {
				alert("����[0] : " + err.message);
				return;
			}
		}

		// =====================================================================
		// ������
		// =====================================================================
		if (top.opener) {
			currObj = top.opener;
		} else if (window.parent) {
			currObj = window.parent;
		} else {
			alert("�߸��� �����Դϴ�.[0]");
			return;
		}

		// =====================================================================
		// ���� ������ Ž��
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
				alert("�߸��� �����Դϴ�.[1]");
				return;
			}

			if (currObj.name == "") {
				alert("�߸��� �����Դϴ�.[2]");
				return;
			}
		}

		if (contents == null) {
			alert("�ý����� ����\n\ncontents ã�⿡ ����!!");
			return;
		}

		if (contents.parent.header) {
			if (contents.parent.header.i_ippbxmng) {
				// ������ ������ ���
				contents.parent.header.i_ippbxmng.click2call(iphoneNum);
				return;
			} else {
				// �� �̿�
				alert("�����Ϳ����� ��밡���� ����Դϴ�.");
				return;
			}
		}

	} catch (err) {
		// TODO : �������� http - https �� ���� �ٸ��� �Ǹ� ������ ���� ��������.
		alert("����[1] : " + err.message);
	}
}
