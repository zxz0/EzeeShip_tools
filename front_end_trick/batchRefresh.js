function sleep(ms) { 	// refer: https://stackoverflow.com/a/39914235
  return new Promise(resolve => setTimeout(resolve, ms));
}

var buttons = $$('.getBtn') 	//no jQuery in page. or JS: document.getElementsByClassName("getBtn");
for(i = 0, len = buttons.length; i < len; i++) {
	buttons[i].click();
	await sleep(10);
}