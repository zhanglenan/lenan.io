String.prototype.trim = function(){ return this.replace(/(^\s*)|(\s*$)/g, "");}
function $(str){ return document.getElementById(str); }