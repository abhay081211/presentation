(function(){
var KEY="recordsheet-v1";
var S={cols:["Naam","Phone","Amount","Date","Note"],rows:[["","","","",""],["","","","",""],["","","","",""]]};
try{var raw=localStorage.getItem(KEY);if(raw){var p=JSON.parse(raw);if(p&&p.cols&&p.rows)S=p;}}catch(e){}
var $=function(i){return document.getElementById(i)};
var t=$("t"),msg=$("msg"),q=$("q");
function save(){try{localStorage.setItem(KEY,JSON.stringify(S));}catch(e){}}
function say(m){msg.textContent=m;}
function render(){
  var f=q.value.trim().toLowerCase(),h="<thead><tr><th class='n'>#</th>";
  S.cols.forEach(function(c,i){h+="<th><input data-h='"+i+"' value=\""+esc(c)+"\"></th>";});
  h+="<th class='n'></th></tr></thead><tbody>";
  S.rows.forEach(function(r,ri){
    if(f&&!r.join(" ").toLowerCase().includes(f))return;
    h+="<tr><td class='n'>"+(ri+1)+"</td>";
    S.cols.forEach(function(c,ci){h+="<td><input data-r='"+ri+"' data-c='"+ci+"' value=\""+esc(r[ci]||"")+"\"></td>";});
    h+="<td class='n'><button class='x' data-d='"+ri+"' title='Row hatao'>✕</button></td></tr>";
  });
  h+="</tbody><tfoot><tr><td class='n'>Σ</td>";
  S.cols.forEach(function(c,ci){
    var s=0,k=0;S.rows.forEach(function(r){var v=parseFloat(String(r[ci]||"").replace(/,/g,""));if(!isNaN(v)&&/^[\s\d.,-]+$/.test(String(r[ci]))){s+=v;k++;}});
    h+="<td>"+(k?"Total: "+(Math.round(s*100)/100):"")+"</td>";
  });
  h+="<td class='n'></td></tr></tfoot>";
  t.innerHTML=h;
}
function esc(s){return String(s).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;");}
t.addEventListener("input",function(e){
  var d=e.target.dataset;
  if(d.h!==undefined)S.cols[+d.h]=e.target.value;
  else if(d.r!==undefined)S.rows[+d.r][+d.c]=e.target.value;
  save();
});
t.addEventListener("change",function(){render();});
t.addEventListener("keydown",function(e){
  if(e.key==="Enter"&&e.target.dataset.r!==undefined){
    var n=t.querySelector("input[data-r='"+(+e.target.dataset.r+1)+"'][data-c='"+e.target.dataset.c+"']");
    if(n){n.focus();}else{$("addRow").click();}
  }
});
t.addEventListener("click",function(e){
  var d=e.target.dataset.d;
  if(d!==undefined&&confirm("Yeh row hata dein?")){S.rows.splice(+d,1);save();render();}
});
$("addRow").onclick=function(){S.rows.push(S.cols.map(function(){return"";}));save();render();
  var l=t.querySelectorAll("tbody tr:last-child input");if(l.length)l[0].focus();};
$("addCol").onclick=function(){S.cols.push("Column "+(S.cols.length+1));S.rows.forEach(function(r){r.push("");});save();render();};
$("delCol").onclick=function(){if(S.cols.length>1&&confirm("Aakhri column hata dein?")){S.cols.pop();S.rows.forEach(function(r){r.pop();});save();render();}};
$("clr").onclick=function(){if(confirm("Saara data hat jayega. Pakka?")){S.rows=[S.cols.map(function(){return"";})];save();render();}};
q.oninput=render;
function aoa(){return [S.cols].concat(S.rows);}
function out(name,data){
  try{
    var blob=data instanceof Blob?data:new Blob([data],{type:name.endsWith(".csv")?"text/csv;charset=utf-8":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    var u=URL.createObjectURL(blob),a=document.createElement("a");
    a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();
    setTimeout(function(){URL.revokeObjectURL(u);},1000);
    say("✅ File download ho gayi: "+name);
  }catch(e){say("Download nahi ho paya.");}
}
$("xlsx").onclick=function(){
  if(!window.XLSX){say("Excel library load nahi hui. CSV try karein.");return;}
  var ws=XLSX.utils.aoa_to_sheet(aoa()),wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Records");
  out("records.xlsx",XLSX.write(wb,{bookType:"xlsx",type:"array"}));
};
$("csv").onclick=function(){
  var c=aoa().map(function(r){return r.map(function(v){v=String(v==null?"":v);return/[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;}).join(",");}).join("\n");
  out("records.csv","\ufeff"+c);
};
$("imp").onchange=function(e){
  var f=e.target.files[0];if(!f)return;
  if(!window.XLSX){say("Import ke liye library load nahi hui.");return;}
  var r=new FileReader();
  r.onload=function(){
    try{
      var wb=XLSX.read(r.result,{type:"array"}),a=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
      if(!a.length){say("File khaali hai.");return;}
      var w=Math.max.apply(null,a.map(function(x){return x.length;}));
      S.cols=a[0].concat(Array(w-a[0].length).fill("")).map(String);
      S.rows=a.slice(1).map(function(x){return x.concat(Array(w-x.length).fill("")).map(String);});
      save();render();say("✅ File khul gayi: "+f.name);
    }catch(err){say("File padh nahi paye.");}
  };
  r.readAsArrayBuffer(f);e.target.value="";
};
render();
})();
