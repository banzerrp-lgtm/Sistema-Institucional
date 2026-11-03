const CLAVE = "1234"; // Cambia esta clave antes de publicar.
const HOJAS = {
  Estudiantes: ["id","nombre","documento","telefono","tutor","turno","semestre","estado","fecha_cambio"],
  Notas: ["id","estudiante","codigo","fecha","nota"],
  Pagos: ["id","estudiante","gestion","tipo","semestre","fecha","anulado","motivo"]
};
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function preparar_(){
  const ss=ss_();
  Object.keys(HOJAS).forEach(n=>{
    let sh=ss.getSheetByName(n); if(!sh) sh=ss.insertSheet(n);
    if(sh.getLastRow()===0) sh.getRange(1,1,1,HOJAS[n].length).setValues([HOJAS[n]]);
    sh.getDataRange().setNumberFormat("@");
  });
}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doGet(e){
  preparar_();
  if(!e || !e.parameter || e.parameter.clave!==CLAVE) return json_({error:"clave"});
  const data={};
  Object.keys(HOJAS).forEach(n=>{
    const sh=ss_().getSheetByName(n), vals=sh.getDataRange().getDisplayValues();
    const heads=vals.shift()||[];
    data[n]=vals.filter(r=>r.some(v=>v!=="")).map(r=>{
      const obj={}; heads.forEach((h,i)=>obj[h]=String(r[i]??"")); return obj;
    });
  });
  return json_(data);
}
function doPost(e){
  let lock;
  try{
    const body=JSON.parse(e.postData.contents||"{}");
    if(body.clave!==CLAVE) return json_({error:"clave"});
    preparar_(); lock=LockService.getScriptLock(); lock.waitLock(20000);
    const tab=body.tabla, action=body.accion;
    if(!HOJAS[tab]) return json_({error:"tabla"});
    const sh=ss_().getSheetByName(tab), headers=HOJAS[tab];
    const result={ok:true};
    if(action==="add"){
      const id=Utilities.getUuid().replace(/-/g,"").slice(0,12);
      const row=body.fila||{}; sh.appendRow(headers.map(h=>h==="id"?id:String(row[h]??"")));
      result.id=id;
    } else if(action==="addmany"){
      const filas=body.filas||[];
      const values=filas.map(f=>{const id=Utilities.getUuid().replace(/-/g,"").slice(0,12); return headers.map(h=>h==="id"?id:String(f[h]??""));});
      if(values.length) sh.getRange(sh.getLastRow()+1,1,values.length,headers.length).setValues(values);
      result.ids=values.map(r=>r[0]);
    } else if(action==="upd"){
      const id=String(body.id||""); const vals=sh.getDataRange().getDisplayValues();
      const ix=headers.indexOf("id"), rownum=vals.findIndex((r,i)=>i>0&&r[ix]===id);
      if(rownum<1) return json_({error:"no_encontrado"});
      const current=vals[rownum]; const patch=body.fila||{};
      headers.forEach((h,j)=>{if(h!=="id"&&Object.prototype.hasOwnProperty.call(patch,h)) current[j]=String(patch[h]??"");});
      sh.getRange(rownum+1,1,1,headers.length).setValues([current]); result.id=id;
    } else if(action==="updmany"){
      const vals=sh.getDataRange().getDisplayValues(), ix=headers.indexOf("id");
      const indexes=new Map(); vals.forEach((r,i)=>{if(i>0) indexes.set(r[ix],i+1);});
      (body.filas||[]).forEach(p=>{
        const rn=indexes.get(String(p.id)); if(!rn)return;
        const row=sh.getRange(rn,1,1,headers.length).getDisplayValues()[0];
        headers.forEach((h,j)=>{if(h!=="id"&&Object.prototype.hasOwnProperty.call(p.fila||{},h))row[j]=String(p.fila[h]??"");});
        sh.getRange(rn,1,1,headers.length).setValues([row]);
      });
      result.actualizados=(body.filas||[]).length;
    } else if(action==="cleandemo" && tab==="Estudiantes"){
      const all=ss_().getSheets();
      const studentRows=sh.getDataRange().getDisplayValues(), hmap=HOJAS.Estudiantes;
      const docix=hmap.indexOf("documento"), idix=hmap.indexOf("id");
      const ids=studentRows.slice(1).filter(r=>/^DEMO/.test(r[docix]||"")).map(r=>r[idix]);
      // Limpia exclusivamente los expedientes marcados DEMO y sus datos asociados.
      ["Notas","Pagos"].forEach(n=>{
        const sheet=ss_().getSheetByName(n), rows=sheet.getDataRange().getDisplayValues(), estix=HOJAS[n].indexOf("estudiante");
        for(let i=rows.length-1;i>=1;i--) if(ids.includes(rows[i][estix])) sheet.deleteRow(i+1);
      });
      for(let i=studentRows.length-1;i>=1;i--) if(/^DEMO/.test(studentRows[i][docix]||"")) sh.deleteRow(i+1);
      result.eliminados=ids.length;
    } else return json_({error:"accion"});
    return json_(result);
  }catch(err){return json_({error:String(err)});}
  finally{if(lock)lock.releaseLock();}
}