let guests=[];
let templates=[];

fetch("template.txt",{cache:"no-store"})
.then(r=>r.text())
.then(t=>{
 templates = t.split("===TEMPLATE===").map(x=>x.trim()).filter(x=>x);
 templateSelect.innerHTML="";
 templates.forEach((t,i)=>{
  const o=document.createElement("option");
  o.value=i;
  o.textContent="Template "+(i+1);
  templateSelect.appendChild(o);
 });
 templateBox.value=templates[0]||"";
 updatePreview();
});

templateSelect.onchange=()=>{
 templateBox.value = templates[templateSelect.value];
 updatePreview();
};

templateBox.addEventListener("input",updatePreview);

excel.addEventListener("change",e=>{
 const r=new FileReader();
 r.onload=()=>{
  const wb=XLSX.read(r.result,{type:"binary"});
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1});
  guests = rows.slice(1).filter(r=>r[1]).map(x=>({
    name:(x[0]||"").trim(),
    number:(x[1]||"").toString().replace(/\D/g,""),
    type:(x[2]||"Guest").trim(),
    sent:false
  }));
  generate();
  updatePreview();
 };
 r.readAsBinaryString(e.target.files[0]);
});

function applyVars(t,g){
 return t.replace(/{{NAME}}/g,g.name)
         .replace(/{{INVITATION_TYPE}}/g,g.type);
}

function updatePreview(){
 if(!guests.length) return;
 const g = guests.find(x=>!x.sent) || guests[0];
 preview.textContent = applyVars(templateBox.value,g);
 previewPhone.textContent="+91 "+g.number;
}

function showPreview(i){
 const g=guests[i];
 preview.textContent=applyVars(templateBox.value,g);
 previewPhone.textContent="+91 "+g.number;
}

function generate(){
 table.innerHTML=`<tr>
 <th>#</th><th>Name</th><th>Mobile</th><th>Type</th><th>Status</th><th>Send</th></tr>`;
 guests.forEach((g,i)=>{
 table.innerHTML+=`<tr class="${g.sent?'sent':''}">
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
 const url="https://wa.me/91"+g.number+"?text="+encodeURIComponent(applyVars(templateBox.value,g));
 window.open(url,"_blank");
 setTimeout(()=>{
  g.sent=true;
  generate();
  updatePreview();
 },2000);
}

function updateStats(){
 const sent=guests.filter(g=>g.sent).length;
 stats.innerText=`Total: ${guests.length} | Sent: ${sent} | Pending: ${guests.length-sent}`;
}

/* Drag Phone */
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
window.addEventListener("scroll",()=>{
 if(window.pageYOffset>200) phone.classList.add("floating");
 else phone.classList.remove("floating");
});
