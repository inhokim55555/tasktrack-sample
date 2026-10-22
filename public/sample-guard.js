// Original screens with local fictional data. No external navigation or uploads.
const external=a=>{try{return new URL(a,location.href).origin!==location.origin;}catch{return true;}};
document.addEventListener('click',e=>{const a=e.target.closest('a[href]');if(a&&external(a.href)){e.preventDefault();e.stopImmediatePropagation();alert('Sample only: external links are disabled.');}},true);
const open=window.open;window.open=function(url,...args){if(url&&url!=='about:blank'&&!String(url).startsWith('blob:')&&external(url))return null;return open.call(window,url,...args);};
function disableUploads(){document.querySelectorAll('input[type=file]').forEach(x=>{x.disabled=true;x.title='File imports are disabled in the public sample.';});}
document.addEventListener('DOMContentLoaded',()=>{disableUploads();new MutationObserver(disableUploads).observe(document.body,{childList:true,subtree:true});});
document.addEventListener('drop',e=>{e.preventDefault();e.stopImmediatePropagation();},true);
document.addEventListener('change',e=>{if(e.target.matches('input[type=file]')){e.target.value='';e.preventDefault();e.stopImmediatePropagation();}},true);
