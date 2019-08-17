	//размеры окна
	var winWidth=550; // ширина окна
	var winHeight=400; // высота окна
	// изменяем размер
	window.resizeTo(winWidth, winHeight);
	// окно в центр экрана
	var winPosX=screen.width/2-winWidth/2;
	var winPosY=screen.height/2-winHeight/2;
	window.moveTo(winPosX, winPosY);
