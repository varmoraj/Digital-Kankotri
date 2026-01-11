let guests=[];
let rawTemplate="";

fetch("template.txt",{cache:"no-store"})
.then(r=>r.text())
.then(t=>{
 rawTemplate=t;
 templateBox.value=t;
 updatePreview();
});

templateBox.addEventListener("input",updatePreview);

excel.addEventListener("change",function(e){
const reader=new FileReader();
reader.onload=function(evt){
const wb=XLSX.read(evt.target.result,{type:"binary"});
const sheet=wb.Sheets[wb.SheetNames[0]];
const rows=XLSX.utils.sheet_to_json(sheet,{header:1});
guests = rows.slice(1).map(r=>({
 name:(r[0]||"").toString().trim(),
 number:(r[1]||"").toString().trim(),
 type:(r[2]||"").toString().trim(),
 sent:false
}));
alert("Loaded "+guests.length+" guests");
updatePreview();
}
reader.readAsBinaryString(e.target.files[0]);
});

function applyVars(t,g){
return t.replace(/{{NAME}}/g,g?.name||"")
        .replace(/{{INVITATION_TYPE}}/g,g?.type||"");
}

function updatePreview(){
const g=guests.find(x=>!x.sent) || guests[0];
if(!g) return;
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
table.innerHTML+=`
<tr id="row${i}">
<td>${i+1}</td>
<td class="clickable" onclick="showPreview(${i})">${g.name}</td>
<td>${g.number}</td>
<td>${g.type}</td>
<td id="s${i}">Pending</td>
<td><button onclick="send(${i})">Send</button></td>
</tr>`;
});
updateStats();
}

function send(i){
const g=guests[i];
window.open("https://wa.me/91"+g.number+"?text="+encodeURIComponent(applyVars(templateBox.value,g)),"_blank");
g.sent=true;
document.getElementById("s"+i).innerText="Sent";
document.getElementById("row"+i).className="sent";
updateStats();
updatePreview();
}

function updateStats(){
const sent=guests.filter(g=>g.sent).length;
stats.innerText="Total: "+guests.length+" | Sent: "+sent+" | Pending: "+(guests.length-sent);
}

/* Drag */
let isDown=false,offX=0,offY=0;
phoneHeader.addEventListener("mousedown",e=>{
 isDown=true;
 offX=e.clientX-phone.offsetLeft;
 offY=e.clientY-phone.offsetTop;
});
document.addEventListener("mouseup",()=>isDown=false);
document.addEventListener("mousemove",e=>{
 if(!isDown) return;
 phone.style.left=(e.clientX-offX)+"px";
 phone.style.top=(e.clientY-offY)+"px";
 phone.classList.remove("side","floating");
});

/* Scroll snap */
window.addEventListener("scroll",()=>{
 if(window.scrollY>200){
   phone.classList.remove("side");
   phone.classList.add("floating");
 }else{
   phone.classList.remove("floating");
   phone.classList.add("side");
 }
});
