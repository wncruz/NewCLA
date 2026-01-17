<!--
     NS4 = (document.layers);
     IE4 = (document.all);
    ver4 = (NS4 || IE4);
   isMac = (navigator.appVersion.indexOf("Mac") != -1);
  isMenu = (NS4 || (IE4 && !isMac));
  function popUp(){return};
  function popDown(){return};
  function startIt(){return};
  if (!ver4) event = null;

  if (isMenu) {
menuVersion = 3;
menuWidth = 200;
borWid = 1;
borSty = "double";
borCol = "#000000";
separator = 1;
separatorCol = "#000000";
fntFam = "arial";
fntBold = false;
fntItal = false;
fntSiz = 8;
fntCol = "#000000";
overFnt = "#FFFFFF";
itemPad = 2;
backCol = "#f1f1f1";
overCol = "#31659c";
imgSrc = "../imagens/tri.gif";
imgSiz = 10;
childOffset = 1;
childOverlap = 6;
perCentOver = null;

clickStart = false;
clickKill = false;
secondsVisible = 0.8;

keepHilite = false;
NSfontOver = true;

isFrames = false;
navFrLoc = "top";
mainFrName = "main";
}
//-->