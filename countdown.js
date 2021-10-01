var i = document.getElementByID('counter');
function countdown() {
  i = document.getElementByID('counter');
  if (parseInt(i.innerHTML) <= 0) {
    location.href = 'index.html';
  }
  if (parseInt(i.innerHTML) != 0) {
    i.innerHTML = parseInt(i.innerHTML) - 1;
  }
  countdownSec();
}

function countdownSec() {
  var x = document.getElementByID('sec');
  if (parseInt(i.innerHTML) = 1) {
    x.innerHTML = 'second';
  } else {
    x.innerHTML = 'seconds';
  }
}
setInternal(function(){ countdown(); },1000};
