let guests=[];

fetch("template.txt",{cache:"no-store"})
.then(r=>r.text())
.then(t=>{templateBox.value=t});

templateBox.addEventListener("input",updatePreview);

excel.addEventListener("change",e=>{
const r=new FileReader();
r.onload=()=>{
const wb=XLSX.read(r.result,{type:"binary"});
const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1});
guests=rows.slice(1).map(x=>({
 name:(x[0]||"").trim(),
 number:(x[1]||"").trim(),
 type:(x[2]||"").trim(),
 sent:false
}));
generate();
updatePreview();
};
r.readAsBinaryString(e.target.files[0]);
});

function applyVars(t,g){
return t.replace(/{{NAME}}/g,g.name).replace(/{{INVITATION_TYPE}}/g,g.type);
}

function updatePreview(){
const g=guests.find(x=>!x.sent)||guests[0];
if(!g)return;
preview.textContent=applyVars(templateBox.value,g);
previewPhone.textContent="+91 "+g.number;
}

function showPreview(i){
const g=guests[i];
preview.textContent=applyVars(templateBox.value,g);
previewPhone.textContent="+91 "+g.number;
}

function generate(){
table.innerHTML=`<tr>
<th>#</th><th>Name</th><th>Mobile</th><th>Invite Type</th><th>Status</th><th>Send</th>
</tr>`;
guests.forEach((g,i)=>{
table.innerHTML+=`<tr id="row${i}">
<td>${i+1}</td>
<td class="clickable" onclick="showPreview(${i})">${g.name}</td>
<td>${g.number}</td>
<td>${g.type}</td>
<td>${g.sent?"Sent":"Pending"}</td>
<td><button onclick="send(${i})">Send</button></td>
</tr>`;
});
updateStats();
}

function send(i){
const g=guests[i];
window.open("https://wa.me/91"+g.number+"?text="+encodeURIComponent(applyVars(templateBox.value,g)));
g.sent=true;
generate();
updatePreview();
}

function updateStats(){
const sent=guests.filter(g=>g.sent).length;
stats.innerText="Total: "+guests.length+" | Sent: "+sent+" | Pending: "+(guests.length-sent);
}

/* Drag */
let isDown=false,offX=0,offY=0;
phoneHeader.onmousedown=e=>{
 isDown=true;
 offX=e.clientX-phone.offsetLeft;
 offY=e.clientY-phone.offsetTop;
};
document.onmouseup=()=>isDown=false;
document.onmousemove=e=>{
 if(!isDown)return;
 phone.style.left=(e.clientX-offX)+"px";
 phone.style.top=(e.clientY-offY)+"px";
};

/* Auto snap on scroll */
window.addEventListener("scroll",()=>{
 if(window.pageYOffset>200){
   phone.classList.add("floating");
 }else{
   phone.classList.remove("floating");
 }
});
