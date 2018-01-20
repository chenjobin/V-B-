
function centerModal(mdlId) { //居中显示遮罩层
	var modalToCenter = document.getElementById(mdlId), currH = 0, currW = 0;
	modalToCenter.style.display = 'block';
	currH = modalToCenter.offsetHeight;
	currW = modalToCenter.offsetWidth;
	if ((document.documentElement && Number(document.documentElement.clientHeight) !== 0 && document.documentElement.clientHeight < currH + 20) || (document.body.offsetHeight && document.body.offsetHeight < currH + 20) || (window.innerHeight && window.innerHeight < currH + 20)) {
		modalToCenter.style.position = 'absolute';
		modalToCenter.style.top = "20px";
		modalToCenter.style.marginTop = '0px';
		window.scroll(0, 0);
	} else {
		modalToCenter.style.position = 'fixed';
		modalToCenter.style.top = "50%";
		modalToCenter.style.marginTop = -(currH / 2) + 'px';
	}
	modalToCenter.style.left = "50%";
	modalToCenter.style.marginLeft = -(currW / 2) + 'px';
}