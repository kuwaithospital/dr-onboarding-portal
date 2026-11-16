/* Kuwait Hospital onboarding – shared header: shrinks into a sticky bar on scroll, live Kuwait date & time */
(function(){var h=document.getElementById('hdr'),on=false;function f(){var y=window.scrollY;if(!on&&y>60)on=true;else if(on&&y<8)on=false;h.classList.toggle('compact',on);}window.addEventListener('scroll',f,{passive:true});f();})();


(function(){var d=document.getElementById('ckD'),t=document.getElementById('ckT');
function tick(){var n=new Date();d.textContent=n.toLocaleDateString('en-GB',{weekday:'long',day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Kuwait'});
t.textContent=n.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit',timeZone:'Asia/Kuwait'});}
tick();setInterval(tick,15000);})();
