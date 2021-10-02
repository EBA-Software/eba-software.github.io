var e = document.getElementById('counter');
var b = document.getElementById('sec');
var a = parseInt(e.innerHTML);
function countdown() {
  e = document.getElementById('counter');
  a = parseInt(e.innerHTML);
  if (a <= 0) {
    document.location = 'https://eba-software.github.io';
  } else {
    a = a - 1;
    e.innerHTML = a;
  }
  countdownSec();
}

function countdownSec() {
  b = document.getElementById('sec');
  if (a == 1) {
    b.innerHTML = 'second';
  } else {
    b.innerHTML = 'seconds';
  }
}
setInterval(countdown, 1000);
