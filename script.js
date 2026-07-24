
const toggle=document.querySelector('.mobile-toggle');
const nav=document.querySelector('.nav-links');
if(toggle&&nav){toggle.addEventListener('click',()=>nav.classList.toggle('open'));}
document.querySelectorAll('.nav-links a').forEach(a=>a.addEventListener('click',()=>nav?.classList.remove('open')));
