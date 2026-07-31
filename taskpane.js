/* ============================================================
   ENGERP Levantamento — Task Pane
   Abre a ferramenta de levantamento (dialog.html) numa janela grande
   e, quando o usuário clica em "Enviar para o Excel" lá dentro, recebe
   os dados via messageParent e escreve nas abas da planilha ABERTA
   usando a API JavaScript do Excel (Office.js / Excel.run).
============================================================ */

let dialog = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
          document.getElementById('btnOpen').addEventListener('click', abrirLevantamento);
          document.getElementById('btnClear').addEventListener('click', ()=>log('Dica: para começar do zero, use o botão "Limpar levantamento salvo" dentro da própria ferramenta.'));
          log('Pronto. Clique em "Abrir Levantamento" para começar.');
    }
});

function log(msg){
    const el = document.getElementById('log');
    const ts = new Date().toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'});
    el.innerHTML = `<div class="log-line"><span class="ts">${ts}</span> ${msg}</div>` + el.innerHTML;
}

function abrirLevantamento(){
    const url = new URL('dialog.html', window.location.href).toString();
    Office.context.ui.displayDialogAsync(url, {height:90, width:90, promptBeforeOpen:false}, (asyncResult)=>{
          if(asyncResult.status === Office.AsyncResultStatus.Failed){
                  log('❌ Não foi possível abrir a ferramenta: ' + asyncResult.error.message);
                  return;
          }
          dialog = asyncResult.value;
          log('Ferramenta aberta. Faça o levantamento e clique em "Enviar para o Excel" quando terminar.');
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, onMessageFromDialog);
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg)=>{
                  log('A janela da ferramenta foi fechada.');
          });
    });
}

function onMessageFromDialog(arg){
    let payload;
    try{ payload = JSON.parse(arg.message); }catch(e){
          log('❌ Mensagem inválida recebida da ferramenta.');
          return;
    }
    if(payload.type === 'STATUS'){
          log(payload.text);
          return;
    }
    if(payload.type === 'LEVANTAMENTO'){
          log('Recebido: ' + payload.rooms.length + ' ambiente(s), ' + payload.walls.length + ' parede(s), ' +
                      payload.esquadrias.length + ' esquadria(s), ' + payload.itens.length + ' ite(m/ns). Escrevendo no Excel...');
          escreverNoExcel(payload)
            .then(()=>{
                      log('✅ Planilha atualizada com sucesso.');
                      if(dialog){
                                  try{ dialog.messageChild(JSON.stringify({type:'ACK', ok:true})); }catch(e){}
                      }
            })
            .catch(err=>{
                      log('❌ Erro ao escrever no Excel: ' + err.message);
                      console.error(err);
            });
    }
}

const COR_NAVY = '#13324F';
const COR_AZUL = '#2E4DFF';
const COR_CINZA_CLARO = '#F2F4F7';

async function escreverNoExcel(payload){
    await Excel.run(async (context) => {
          await writeAmbientesSheet(context, payload.rooms);
          await writeParedesSheet(context, payload.walls);
          await writeEsquadriasSheet(context, payload.esquadrias);
          await writeItensSheet(context, payload.itens);
          await context.sync();
    });
}

async function getCleanSheet(context, name){
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItemOrNullObject(name);
    await context.sync();
    if(sheet.isNullObject){
          sheet = sheets.add(name);
    } else {
          const used = sheet.getUsedRangeOrNullObject();
          used.load('address');
          await context.sync();
          if(!used.isNullObject) used.clear(Excel.ClearApplyTo.all);
    }
    return sheet;
}

function writeTable(sheet, startRow, headers, rows){
    const nCols = headers.length;
    const headerRange = sheet.getRangeByIndexes(startRow, 0, 1, nCols);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = 'white';
    headerRange.format.fill.color = COR_NAVY;
    headerRange.format.horizontalAlignment = 'Center';

  if(rows.length>0){
        const dataRange = sheet.getRangeByIndexes(startRow+1, 0, rows.length, nCols);
        dataRange.values = rows;
        dataRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        for(let i=0;i<rows.length;i++){
                if(i%2===1){
                          sheet.getRangeByIndexes(startRow+1+i, 0, 1, nCols).format.fill.color = COR_CINZA_CLARO;
                }
        }
  }
    sheet.getUsedRangeOrNullObject().format.autofitColumns();
    return startRow + 1 + rows.length;
}

async function writeAmbientesSheet(context, rooms){
    const sheet = await getCleanSheet(context, 'Memorial de Cálculo');
    const titleRange = sheet.getRangeByIndexes(0,0,1,13);
    titleRange.merge();
    titleRange.values = [['MEMORIAL DE CÁLCULO — PROJETO ARQUITETÔNICO (ENGERP)', '', '', '', '', '', '', '', '', '', '', '', '']];
    titleRange.format.font.bold = true;
    titleRange.format.font.size = 13;
    titleRange.format.font.color = 'white';
    titleRange.format.fill.color = COR_NAVY;

  const headers = ['Ambiente','Pavimento','Área (m²)','Perímetro (m)','Pé-direito (m)','Piso (especificação)',
                       'Rodapé (m)','Soleira (m)','Peitoril (m)','Revest. Parede (m²)','Pintura (m²)','Forro (m²)','Impermeabilização (m²)'];
    const rows = rooms.map(r=>[
          r.name, r.floor, r.area, r.perim, r.height, r.specPiso||'—',
          r.rodape, r.soleira, r.peitoril, r.revpar, r.pinturaParede, r.area, r.imperm
        ]);
    const lastRow = writeTable(sheet, 2, headers, rows);

  if(rows.length){
        const totalRange = sheet.getRangeByIndexes(lastRow, 0, 1, headers.length);
        totalRange.values = [['TOTAL','',
                                    rooms.reduce((s,r)=>s+r.area,0),
                                    rooms.reduce((s,r)=>s+r.perim,0),
                                    '', '',
                                    rooms.reduce((s,r)=>s+r.rodape,0),
                                    rooms.reduce((s,r)=>s+r.soleira,0),
                                    rooms.reduce((s,r)=>s+r.peitoril,0),
                                    rooms.reduce((s,r)=>s+r.revpar,0),
                                    rooms.reduce((s,r)=>s+r.pinturaParede,0),
                                    rooms.reduce((s,r)=>s+r.area,0),
                                    rooms.reduce((s,r)=>s+r.imperm,0)
                                  ]];
        totalRange.format.font.bold = true;
        totalRange.format.fill.color = '#DCE4F5';
  }
    sheet.activate();
}

async function writeParedesSheet(context, walls){
    const sheet = await getCleanSheet(context, 'Paredes');
    const headers = ['Tipo','Descrição','Comprimento (m)','Pé-direito (m)','Faces','Área (m²)'];
    const rows = walls.map(w=>[w.tipoLabel, w.desc||'—', w.largura, w.peDireito, w.qtyFaces, round2(w.largura*w.peDireito*w.qtyFaces)]);
    const lastRow = writeTable(sheet, 0, headers, rows);
    if(rows.length){
          const totalRange = sheet.getRangeByIndexes(lastRow, 0, 1, headers.length);
          totalRange.values = [['TOTAL','','','','', rows.reduce((s,r)=>s+r[5],0)]];
          totalRange.format.font.bold = true;
          totalRange.format.fill.color = '#DCE4F5';
    }
}

async function writeEsquadriasSheet(context, esquadrias){
    const sheet = await getCleanSheet(context, 'Esquadrias');
    const headers = ['Código','Tipo','Abertura','Largura (m)','Altura (m)','Qtd','Ambiente'];
    const rows = esquadrias.map(e=>[e.ident||'—', e.tipo, e.abertura||'—', e.largura, e.altura, e.qty, e.roomName||'—']);
    writeTable(sheet, 0, headers, rows);
}

async function writeItensSheet(context, itens){
    const sheet = await getCleanSheet(context, 'Louças-Metais-Luminárias');
    const headers = ['Categoria','Descrição','Qtd','Largura (m)','Profundidade (m)','Ambiente'];
    const rows = itens.map(i=>[i.categoria, i.desc||'—', i.qty, i.largura||'—', i.profundidade||'—', i.roomName||'—']);
    writeTable(sheet, 0, headers, rows);
}

function round2(n){ return Math.round(n*100)/100; }
