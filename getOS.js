var os = navigator.appVersion;
var comp = "Unrecognized operating system: " + os;
var isComp = 1
if (os.indexOf("X11") != -1) {
  comp = "EBA Command Center is not supported on Unix!";
  isComp = 0;
}
if (os.indexOf("Windows NT 10") != -1) {
  comp = "EBA Command Center is su]O
pported in Windows 10";
  isComp = 2;
}
if (os.indexOf("Windows NT 8") != -1) {
  comp = "EBA Command Center is supported in Windows 8";
  isComp = 2;
}
if (os.indexOf("Windows NT 6.1") != -1) {
  comp = "You're using Windows 7. You must use the Windows 7 edition of EBA Command Center";
  isComp = 2;
}
if (os.indexOf("Windows NT Vista") != -1) {
  comp = "You're using Windows Vista. You must use the Windows 7 edition of EBA Command Center";
  isComp = 2;
}
if (os.indexOf("Windows NT XP") != -1) {
  comp = "You're using Windows XP. You must use the Windows XP edition of EBA Command Center";
  isComp = 2;
}
if (os.indexOf("Mac") != -1) {
  comp = "EBA Command Center is not supported on Mac!";
  isComp = 0;
}
if (os.indexOf("iPhone") != -1) {
  comp = "EBA Command Center is not supported on iPhone!";
  isComp = 0;
}
if (os.indexOf("Linux") != -1) {
  comp = "EBA Command Center is not supported on Unix!";
  isComp = 0;
}
if (os.indexOf("CrOs") != -1) {
  comp = "EBA Command Center is not supported on ChromeOS!";
  isComp = 0;
}
if (os.indexOf("Android") != -1) {
  comp = "EBA Command Center is not supported on Android!";
  isComp = 0;
}
if (os.indexOf("SMART-TV") != -1) {
  comp = "EBA Command Center is not supported on SmartTV!";
  isComp = 0;
}

//Display
if (isComp === 0) {
  document.getElementByID("OSRed").innerHTML = comp;
}
if (isComp === 1) {
  document.getElementByID("OSYellow").innerHTML = comp;
}
if (isComp === 2) {
  document.getElementByID("OSGreen").innerHTML = comp;
}
