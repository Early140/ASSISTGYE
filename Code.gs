const SPREADSHEET_ID = "PEGAR_ID_DE_TU_GOOGLE_SHEET_AQUI";
const SHEET_PLAYERS = "Jugadores";
const SHEET_ATT = "Asistencias";

function doGet(e){
  try{
    const a = String(e.parameter.accion || "");
    if(a === "jugadores"){
      return json(listarJugadores(
        String(e.parameter.categoria || ""),
        String(e.parameter.estado || "ACTIVOS")
      ));
    }
    if(a === "asistenciaDia"){
      return json(asistenciaDia(
        String(e.parameter.categoria || ""),
        String(e.parameter.fecha || "")
      ));
    }
    if(a === "resumenMes"){
      return json(resumenMes(
        String(e.parameter.categoria || ""),
        String(e.parameter.mes || "")
      ));
    }
    return json({ok:false,error:"Acción GET inválida"});
  }catch(err){
    return json({ok:false,error:String(err.message || err)});
  }
}

function doPost(e){
  try{
    const d = JSON.parse(e.postData.contents || "{}");
    switch(String(d.accion || "")){
      case "crearJugador": return json(crearJugador(d));
      case "editarJugador": return json(editarJugador(d));
      case "bajaJugador": return json(cambiarEstadoJugador(d.id,false));
      case "reactivarJugador": return json(cambiarEstadoJugador(d.id,true));
      case "guardarAsistencia": return json(guardarAsistencia(d));
      default: return json({ok:false,error:"Acción POST inválida"});
    }
  }catch(err){
    return json({ok:false,error:String(err.message || err)});
  }
}

function setup(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let p = ss.getSheetByName(SHEET_PLAYERS);
  if(!p) p = ss.insertSheet(SHEET_PLAYERS);
  if(p.getLastRow() === 0){
    p.appendRow(["ID","Categoria","Nombre","Apellido","Nacimiento","Dorsal","Observaciones","Activo","FechaAlta","FechaModificacion"]);
    p.setFrozenRows(1);
  }

  let a = ss.getSheetByName(SHEET_ATT);
  if(!a) a = ss.insertSheet(SHEET_ATT);
  if(a.getLastRow() === 0){
    a.appendRow(["IDRegistro","Timestamp","Fecha","Categoria","Profesor","JugadorID","Nombre","Apellido","Estado"]);
    a.setFrozenRows(1);
  }

  return "OK";
}

function sheets(){
  setup();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    players:ss.getSheetByName(SHEET_PLAYERS),
    att:ss.getSheetByName(SHEET_ATT)
  };
}

function listarJugadores(categoria,estado){
  if(!categoria) return {ok:false,error:"Falta categoría"};
  const sh = sheets().players;
  const data = sh.getDataRange().getValues();
  if(data.length < 2) return {ok:true,jugadores:[]};

  const out = [];
  for(let i=1;i<data.length;i++){
    const r=data[i];
    if(String(r[1]).trim() !== categoria) continue;

    const activo = normalizarActivo(r[7]);
    if(estado === "ACTIVOS" && !activo) continue;
    if(estado === "BAJAS" && activo) continue;

    out.push({
      id:String(r[0]),
      categoria:String(r[1]||""),
      nombre:String(r[2]||""),
      apellido:String(r[3]||""),
      nacimiento:fechaISO(r[4]),
      dorsal:String(r[5]||""),
      observaciones:String(r[6]||""),
      activo:activo
    });
  }
  out.sort((a,b)=>(a.apellido+" "+a.nombre).localeCompare(b.apellido+" "+b.nombre,"es"));
  return {ok:true,jugadores:out};
}

function crearJugador(d){
  validarJugador(d);
  const sh=sheets().players;
  const id=Utilities.getUuid();
  const now=new Date();
  sh.appendRow([
    id,
    String(d.categoria).trim(),
    String(d.nombre).trim(),
    String(d.apellido).trim(),
    fechaParaSheet(d.nacimiento),
    String(d.dorsal||"").trim(),
    String(d.observaciones||"").trim(),
    "SI",
    now,
    now
  ]);
  return {ok:true,id:id};
}

function editarJugador(d){
  validarJugador(d);
  if(!d.id) return {ok:false,error:"Falta ID del jugador"};

  const sh=sheets().players;
  const row=encontrarFilaPorId(sh,d.id);
  if(!row) return {ok:false,error:"Jugador no encontrado"};

  sh.getRange(row,2,1,6).setValues([[
    String(d.categoria).trim(),
    String(d.nombre).trim(),
    String(d.apellido).trim(),
    fechaParaSheet(d.nacimiento),
    String(d.dorsal||"").trim(),
    String(d.observaciones||"").trim()
  ]]);
  sh.getRange(row,10).setValue(new Date());
  return {ok:true};
}

function cambiarEstadoJugador(id,activo){
  if(!id) return {ok:false,error:"Falta ID del jugador"};
  const sh=sheets().players;
  const row=encontrarFilaPorId(sh,id);
  if(!row) return {ok:false,error:"Jugador no encontrado"};

  sh.getRange(row,8).setValue(activo ? "SI" : "NO");
  sh.getRange(row,10).setValue(new Date());
  return {ok:true};
}

function asistenciaDia(categoria,fecha){
  if(!categoria || !fecha) return {ok:false,error:"Faltan datos"};

  const fechaObjetivo=fechaISO(fecha);
  const sh=sheets().att;
  const data=sh.getDataRange().getValues();
  const regs=[];

  for(let i=1;i<data.length;i++){
    const r=data[i];
    if(fechaISO(r[2])===fechaObjetivo && String(r[3]).trim()===categoria){
      regs.push({
        jugadorId:String(r[5]),
        estado:String(r[8]||"").trim().toUpperCase()
      });
    }
  }
  return {ok:true,registros:regs};
}

function guardarAsistencia(d){
  if(!d.fecha || !d.categoria || !d.profesor || !Array.isArray(d.registros)){
    return {ok:false,error:"Datos incompletos"};
  }

  const sh=sheets().att;
  const fecha=fechaISO(d.fecha);
  const categoria=String(d.categoria).trim();

  // Borra registros existentes de la misma fecha y categoría.
  // Esto evita duplicados y permite corregir una asistencia ya guardada.
  const last=sh.getLastRow();
  if(last>1){
    const data=sh.getRange(2,1,last-1,9).getValues();
    const filasABorrar=[];

    for(let i=0;i<data.length;i++){
      const f=fechaISO(data[i][2]);
      const c=String(data[i][3]).trim();
      if(f===fecha && c===categoria){
        filasABorrar.push(i+2);
      }
    }

    // borrar de abajo hacia arriba
    filasABorrar.sort((a,b)=>b-a).forEach(row=>sh.deleteRow(row));
  }

  const now=new Date();
  const rows=d.registros.map(r=>[
    Utilities.getUuid(),
    now,
    fechaParaSheet(fecha),
    categoria,
    String(d.profesor).trim(),
    String(r.jugadorId),
    String(r.nombre||""),
    String(r.apellido||""),
    String(r.estado||"").trim().toUpperCase()
  ]);

  if(rows.length){
    sh.getRange(sh.getLastRow()+1,1,rows.length,9).setValues(rows);
  }

  return {ok:true,guardados:rows.length,fecha:fecha,categoria:categoria};
}

function resumenMes(categoria,mes){
  if(!categoria || !mes) return {ok:false,error:"Faltan categoría o mes"};

  const jugadores=listarJugadores(categoria,"TODOS");
  if(!jugadores.ok) return jugadores;

  const sh=sheets().att;
  const data=sh.getDataRange().getValues();
  const map={};
  const practiceDates={};
  let totalRecords=0;

  jugadores.jugadores.forEach(p=>{
    map[p.id]={
      id:p.id,
      nombre:p.nombre,
      apellido:p.apellido,
      presente:0,
      ausente:0,
      justificado:0,
      tarde:0,
      total:0
    };
  });

  for(let i=1;i<data.length;i++){
    const r=data[i];
    const f=fechaISO(r[2]);
    const cat=String(r[3]||"").trim();

    if(cat!==categoria) continue;
    if(!f || f.substring(0,7)!==mes) continue;

    totalRecords++;
    practiceDates[f]=true;

    const id=String(r[5]||"");
    if(!map[id]){
      map[id]={
        id:id,
        nombre:String(r[6]||""),
        apellido:String(r[7]||""),
        presente:0,
        ausente:0,
        justificado:0,
        tarde:0,
        total:0
      };
    }

    const x=map[id];
    const estado=String(r[8]||"").trim().toUpperCase();
    x.total++;

    if(estado==="PRESENTE") x.presente++;
    else if(estado==="AUSENTE") x.ausente++;
    else if(estado==="JUSTIFICADO") x.justificado++;
    else if(estado==="TARDE") x.tarde++;
  }

  const detalle=Object.values(map)
    .filter(x=>x.total>0)
    .map(x=>{
      // PRESENTE y TARDE cuentan como asistencia efectiva.
      x.porcentaje=x.total
        ? Math.round(((x.presente+x.tarde)/x.total)*100)
        : 0;
      return x;
    })
    .sort((a,b)=>(a.apellido+" "+a.nombre).localeCompare(b.apellido+" "+b.nombre,"es"));

  const attendCount=detalle.reduce((s,x)=>s+x.presente+x.tarde,0);
  const considered=detalle.reduce((s,x)=>s+x.total,0);
  const avg=considered ? Math.round((attendCount/considered)*100) : 0;

  return {
    ok:true,
    practicas:Object.keys(practiceDates).length,
    promedioAsistencia:avg,
    jugadores:detalle.length,
    registros:totalRecords,
    detalle:detalle
  };
}

/* ---------- FECHAS ROBUSTAS ---------- */

function fechaISO(v){
  if(v===null || v===undefined || v==="") return "";

  // Si Sheets lo entrega como Date real
  if(Object.prototype.toString.call(v)==="[object Date]" && !isNaN(v.getTime())){
    return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd");
  }

  const s=String(v).trim();

  // yyyy-MM-dd o yyyy-MM-ddTHH...
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m){
    return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  }

  // dd/MM/yyyy
  m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m){
    return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  }

  // dd-MM-yyyy
  m=s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if(m){
    return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  }

  // Último intento con Date
  const d=new Date(s);
  if(!isNaN(d.getTime())){
    return Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd");
  }

  return "";
}

function fechaParaSheet(v){
  const iso=fechaISO(v);
  if(!iso) return "";
  const p=iso.split("-");
  return new Date(Number(p[0]),Number(p[1])-1,Number(p[2]),12,0,0);
}

function pad2(v){
  return String(v).padStart(2,"0");
}

/* ---------- HELPERS ---------- */

function validarJugador(d){
  if(!d.nombre || !d.apellido || !d.categoria){
    throw new Error("Nombre, apellido y categoría son obligatorios");
  }
}

function encontrarFilaPorId(sh,id){
  const last=sh.getLastRow();
  if(last<2) return 0;
  const vals=sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<vals.length;i++){
    if(String(vals[i][0])===String(id)) return i+2;
  }
  return 0;
}

function normalizarActivo(v){
  if(v===true) return true;
  const s=String(v||"").trim().toUpperCase();
  return s==="SI" || s==="SÍ" || s==="TRUE" || s==="1" || s==="ACTIVO";
}

function json(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
