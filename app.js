let guests=[];
let rawTemplate="";

fetch("template.txt",{cache:"no-store"})
.then(r=>r.text())
.then(t=>{
 rawTemplate=t;
 document.getElementById("templateBox").value=t;
 updatePreview();
});

document.getElementById("templateBox").addEventListener("input",updatePreview);

document.getElementById("excel").addEventListener("change",function(e){
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
return t
.replace(/{{NAME}}/g,g?.name||"")
.replace(/{{INVITATION_TYPE}}/g,g?.type||"");
}

function updatePreview(){
const g=guests[0];
const txt=applyVars(document.getElementById("templateBox").value,g);
document.getElementById("preview").textContent=txt;
}

function generate(){
const table=document.getElementById("table");
table.innerHTML=`<tr>
<th>#</th><th>Name</th><th>Mobile</th><th>Invite Type</th><th>Status</th><th>Send</th>
</tr>`;

guests.forEach((g,i)=>{
table.innerHTML+=`
<tr id="row${i}">
<td>${i+1}</td>
<td>${g.name}</td>
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
const msg=applyVars(document.getElementById("templateBox").value,g);
window.open("https://wa.me/91"+g.number+"?text="+encodeURIComponent(msg),"_blank");
g.sent=true;
document.getElementById("s"+i).innerText="Sent";
document.getElementById("row"+i).className="sent";
updateStats();
}

function updateStats(){
const sent=guests.filter(g=>g.sent).length;
document.getElementById("stats").innerText=
"Total: "+guests.length+" | Sent: "+sent+" | Pending: "+(guests.length-sent);
}
