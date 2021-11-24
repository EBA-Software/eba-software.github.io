var os = navigator.appVersion;
var comp = "Unrecognized operating system: " + os;
var isComp = 1
if (os.indexOf("X11") != -1) {
  comp = "This program is not supported on Unix!";
  isComp = 0;
}
if (os.indexOf("Windows NT 10") != -1) {
  comp = "This program is supported in Windows 10";
  isComp = 2;
}
if (os.indexOf("Windows NT 6.3") != -1) {
  comp = "This program is supported in Windows 8.1";
  isComp = 2;
}
if (os.indexOf("Windows NT 6.2") != -1) {
  comp = "This program is supported in Windows 8";
  isComp = 2;
}
if (os.indexOf("Windows NT 6.1") != -1) {
  comp = "You're using Windows 7. You must use the Windows 7 edition of this program";
  isComp = 2;
}
if (os.indexOf("Windows NT Vista") != -1) {
  comp = "You're using Windows Vista. You must use the Windows 7 edition of this program";
  isComp = 2;
}
if (os.indexOf("Windows NT XP") != -1) {
  comp = "You're using Windows XP. You must use the Windows XP edition of this program";
  isComp = 2;
}
if (os.indexOf("Mac") != -1) {
  comp = "This program is not supported on Mac!";
  isComp = 0;
}
if (os.indexOf("iPhone") != -1) {
  comp = "This program is not supported on iPhone!";
  isComp = 0;
}
if (os.indexOf("Linux") != -1) {
  comp = "This program is not supported on Unix!";
  isComp = 0;
}
if (os.indexOf("CrOS") != -1) {
  comp = "This program is not supported on ChromeOS!";
  isComp = 0;
}
if (os.indexOf("Android") != -1) {
  comp = "This program is not supported on Android!";
  isComp = 0;
}
if (os.indexOf("SMART-TV") != -1) {
  comp = "This program is not supported on SmartTV!";
  isComp = 0;
}

//Display
if (isComp === 0) {
  document.getElementById("OSRed").innerHTML = comp;
}
if (isComp === 1) {
  document.getElementById("OSYellow").innerHTML = comp;
}
if (isComp === 2) {
  document.getElementById("OSGreen").innerHTML = comp;
}
